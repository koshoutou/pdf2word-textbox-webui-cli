"""
shape_builder.py - 矢量图形还原器
将 PDF 中的矢量图形 (drawings) 还原为 DOCX 内嵌的 VML/DrawingML 形状:
  - 矩形 (wsp with prstGeom rect)
  - 直线 (wsp with prstGeom line)
  - 椭圆/圆 (wsp with prstGeom ellipse)
  - 折线/多边形 (custom geom path)
  - 文本框/带填充的矩形 (含填充色)
对每个 drawing 分类: 矩形/线/曲线, 还原位置、尺寸、填充、边框颜色
"""
from __future__ import annotations
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

# OOXML 命名空间
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# EMU 转换: 1 pt = 12700 EMU
PT_TO_EMU = 12700


@dataclass
class ShapeInfo:
    """形状信息 (统一抽象)"""
    shape_type: str  # 'rect' / 'line' / 'ellipse' / 'path' / 'group'
    x: float          # pt (页面坐标)
    y: float
    w: float
    h: float
    fill_color: Optional[str] = None    # hex 'RRGGBB'
    line_color: Optional[str] = None
    line_width: float = 1.0   # pt
    line_dash: Optional[str] = None  # 'solid' / 'dash' / 'dot'
    points: Optional[List[Tuple[float, float]]] = None  # 折线/路径点
    rotation: float = 0.0     # 度


class ShapeClassifier:
    """矢量图形分类器 (从 DrawingData 提取形状)"""

    @staticmethod
    def classify_drawing(dr) -> List[ShapeInfo]:
        """从单个 DrawingData 提取 1~N 个 ShapeInfo"""
        from extractor import DrawingData
        shapes = []
        r = dr.rect
        x0, y0, x1, y1 = r
        w = x1 - x0
        h = y1 - y0
        # 颜色
        fill_color = None
        if dr.fill:
            fill_color = '%02X%02X%02X' % (
                int(max(0, min(1, dr.fill[0])) * 255),
                int(max(0, min(1, dr.fill[1])) * 255),
                int(max(0, min(1, dr.fill[2])) * 255))
        line_color = '000000'
        if dr.color:
            line_color = '%02X%02X%02X' % (
                int(max(0, min(1, dr.color[0])) * 255),
                int(max(0, min(1, dr.color[1])) * 255),
                int(max(0, min(1, dr.color[2])) * 255))
        line_width = max(0.25, dr.width or 1.0)
        line_dash = 'solid'
        if dr.dash:
            if len(dr.dash) >= 2:
                on, off = dr.dash[0], dr.dash[1]
                if on < 2 and off < 2:
                    line_dash = 'dot'
                else:
                    line_dash = 'dash'

        # 按 items 类型分类
        items = dr.items or []
        if not items:
            return []
        # 矩形 item: ('re', rect)
        rect_items = [it for it in items if it[0] == 're']
        line_items = [it for it in items if it[0] == 'l']
        curve_items = [it for it in items if it[0] in ('c', 'qu')]

        for it in rect_items:
            r_obj = it[1]
            shapes.append(ShapeInfo(
                shape_type='rect',
                x=r_obj.x0, y=r_obj.y0,
                w=r_obj.x1 - r_obj.x0,
                h=r_obj.y1 - r_obj.y0,
                fill_color=fill_color,
                line_color=line_color,
                line_width=line_width,
                line_dash=line_dash,
            ))
        # 线条: 单条作为独立形状
        for it in line_items:
            p1, p2 = it[1], it[2]
            lx = min(p1.x, p2.x)
            ly = min(p1.y, p2.y)
            lw = abs(p2.x - p1.x)
            lh = abs(p2.y - p1.y)
            if lw < 0.5 and lh < 0.5:
                continue
            # 水平线
            if lh < 1.5 and lw > 2:
                shapes.append(ShapeInfo(
                    shape_type='line',
                    x=lx, y=ly, w=lw, h=line_width,
                    line_color=line_color, line_width=line_width,
                    line_dash=line_dash,
                ))
            elif lw < 1.5 and lh > 2:
                shapes.append(ShapeInfo(
                    shape_type='line',
                    x=lx, y=ly, w=line_width, h=lh,
                    line_color=line_color, line_width=line_width,
                    line_dash=line_dash,
                ))
            else:
                # 斜线
                shapes.append(ShapeInfo(
                    shape_type='line',
                    x=p1.x, y=p1.y, w=p2.x - p1.x, h=p2.y - p1.y,
                    line_color=line_color, line_width=line_width,
                    line_dash=line_dash,
                    points=[(p1.x, p1.y), (p2.x, p2.y)],
                ))
        # 曲线/路径 → 多边形点集
        if curve_items and not rect_items:
            # 收集所有曲线控制点
            pts = []
            for it in curve_items:
                if it[0] == 'c':
                    # ('c', p1, p2, p3, p4) bezier
                    pts.append((it[1].x, it[1].y))
                    pts.append((it[4].x, it[4].y))
                elif it[0] == 'qu':
                    pts.append((it[1].x, it[1].y))
                    pts.append((it[2].x, it[2].y))
                    pts.append((it[3].x, it[3].y))
                    pts.append((it[4].x, it[4].y))
            if pts:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                shapes.append(ShapeInfo(
                    shape_type='path',
                    x=min(xs), y=min(ys),
                    w=max(xs) - min(xs), h=max(ys) - min(ys),
                    fill_color=fill_color, line_color=line_color,
                    line_width=line_width, line_dash=line_dash,
                    points=pts,
                ))
        return shapes


class DocxShapeBuilder:
    """将 ShapeInfo 转换为 DOCX 内嵌 DrawingML 形状 (浮动定位)"""

    def __init__(self):
        self._shape_id = 0

    def _next_id(self) -> int:
        self._shape_id += 1
        return self._shape_id + 10000

    def build_shape_anchor(self, shape: ShapeInfo) -> etree._Element:
        """构建 wp:anchor XML 元素 (绝对定位形状)"""
        self._shape_id += 1
        sid = self._shape_id + 10000
        cx = int(abs(shape.w) * PT_TO_EMU)
        cy = int(abs(shape.h) * PT_TO_EMU)
        pos_x = int(shape.x * PT_TO_EMU)
        pos_y = int(shape.y * PT_TO_EMU)
        # 颜色填充
        fill_xml = ''
        if shape.fill_color:
            fill_xml = f'<a:solidFill><a:srgbClr val="{shape.fill_color}"/></a:solidFill>'
        else:
            # 无填充
            fill_xml = '<a:noFill/>'
        # 边框
        line_xml = ''
        if shape.line_color:
            line_w = int(max(0.25, shape.line_width) * 12700)
            dash_xml = ''
            if shape.line_dash == 'dash':
                dash_xml = f'<a:prstDash val="dash"/>'
            elif shape.line_dash == 'dot':
                dash_xml = f'<a:prstDash val="dot"/>'
            line_xml = f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{shape.line_color}"/></a:solidFill>{dash_xml}</a:ln>'
        else:
            line_xml = '<a:ln><a:noFill/></a:ln>'
        # prstGeom 选择
        prst = 'rect'
        if shape.shape_type == 'ellipse':
            prst = 'ellipse'
        elif shape.shape_type == 'line' and shape.points and len(shape.points) == 2:
            prst = 'line'
        elif shape.shape_type == 'path':
            prst = 'rect'  # 简化: 路径用 rect 包围盒表示
        # 旋转
        rot_attr = f' rot="{shape.rotation * 60000}"' if shape.rotation else ''
        # line 端点坐标 (cxnSp 元素, 不是 wsp)
        if prst == 'line' and shape.points and len(shape.points) == 2:
            p1, p2 = shape.points
            # 相对坐标
            flipH_attr = ' flipH="1"' if p2[0] < p1[0] else ''
            flipV_attr = ' flipV="1"' if p2[1] < p1[1] else ''
            ext_cx = int(abs(p2[0] - p1[0]) * PT_TO_EMU) or 9525
            ext_cy = int(abs(p2[1] - p1[1]) * PT_TO_EMU) or 9525
            pos_h = int(min(p1[0], p2[0]) * PT_TO_EMU)
            pos_v = int(min(p1[1], p2[1]) * PT_TO_EMU)
            # 使用 cxnSp (connector)
            xml = f'''<wp:anchor xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page"><wp:posOffset>{pos_h}</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>{pos_v}</wp:posOffset></wp:positionV>
  <wp:extent cx="{ext_cx}" cy="{ext_cy}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="{sid}" name="线{sid}"/>
  <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
  <a:graphic>
    <a:graphicData uri="{A_NS}/cxnSp">
      <a:cxnSp>
        <a:nvCxnSpPr><a:cNvPr id="{sid}" name="line{sid}"/><a:cNvCxnSpPr/></a:nvCxnSpPr>
        <a:spPr{rot_attr}{flipH_attr}{flipV_attr}>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          {line_xml}
        </a:spPr>
      </a:cxnSp>
    </a:graphicData>
  </a:graphic>
</wp:anchor>'''
        else:
            xml = f'''<wp:anchor xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page"><wp:posOffset>{pos_x}</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>{pos_y}</wp:posOffset></wp:positionV>
  <wp:extent cx="{cx or 9525}" cy="{cy or 9525}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="{sid}" name="形状 {sid}"/>
  <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
  <a:graphic>
    <a:graphicData uri="{WP_NS}/sp">
      <a:sp>
        <a:nvSpPr><a:cNvPr id="{sid}" name="shape{sid}"/><a:cNvSpPr/></a:nvSpPr>
        <a:spPr{rot_xml}>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx or 9525}" cy="{cy or 9525}"/></a:xfrm>
          <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>
          {fill_xml}
          {line_xml}
        </a:spPr>
      </a:sp>
    </a:graphicData>
  </a:graphic>
</wp:anchor>'''
        return etree.fromstring(xml)

    def add_shape_to_paragraph(self, p, shape: ShapeInfo):
        """把形状加到段落的 run 中"""
        anchor = self.build_shape_anchor(shape)
        drawing = OxmlElement('w:drawing')
        drawing.append(anchor)
        # 加到段落的 run
        run = p.add_run()
        run._element.append(drawing)


def filter_meaningful_drawings(drawings, min_size=3.0) -> List:
    """过滤掉无意义的微小/全页背景 drawings"""
    result = []
    for dr in drawings:
        r = dr.rect
        w = r[2] - r[0]
        h = r[3] - r[1]
        # 全页背景 (尺寸接近整页)
        if w > 580 and h > 740:
            continue
        # 太小
        if w < min_size and h < min_size:
            continue
        # 无 items 且无 fill
        if not dr.items and not dr.fill:
            continue
        result.append(dr)
    return result
