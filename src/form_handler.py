"""
form_handler.py - PDF 表单域识别与还原
识别 PDF AcroForm 表单域, 还原为 DOCX 对应元素:
  - 签名域 (Signature): 还原为浮动图片(已有逻辑)或签名占位符
  - 文本框 (Text): 还原为带边框的段落/表格单元格
  - 复选框 (CheckBox): 还原为 ☐/☑ 符号
  - 单选按钮 (RadioButton): 还原为 ○/● 符号
  - 下拉列表 (ComboBox/ListBox): 还原为文本+下拉标记
  - 按钮 (Button): 还原为按钮样式的文本
"""
from __future__ import annotations
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field
import fitz


@dataclass
class FormField:
    """表单域信息"""
    field_type: str       # 'Signature' / 'Text' / 'CheckBox' / 'RadioButton' / 'ComboBox' / 'Button'
    field_name: str
    field_value: str = ''
    rect: Tuple[float, float, float, float] = (0, 0, 0, 0)
    page_no: int = 0
    flags: int = 0
    # 派生信息
    is_checked: bool = False       # 复选框是否勾选
    display_text: str = ''         # 显示文本
    placeholder: str = ''          # 占位符文本


class FormFieldDetector:
    """表单域检测器"""

    def __init__(self, pdf_path: str, verbose: bool = False):
        self.pdf_path = pdf_path
        self.verbose = verbose
        self.doc = None

    def detect_all_fields(self) -> Dict[int, List[FormField]]:
        """检测所有页的表单域, 返回 {page_no: [FormField]}"""
        result = {}
        try:
            self.doc = fitz.open(self.pdf_path)
            for pno in range(len(self.doc)):
                page = self.doc[pno]
                widgets = list(page.widgets())
                if not widgets:
                    continue
                fields = []
                for w in widgets:
                    f = self._parse_widget(w, pno)
                    if f:
                        fields.append(f)
                if fields:
                    result[pno] = fields
                    if self.verbose:
                        print(f'[Form] p{pno+1}: {len(fields)} fields')
            self.doc.close()
        except Exception as e:
            if self.verbose:
                print(f'[Form] error: {e}')
            if self.doc:
                self.doc.close()
        return result

    def _parse_widget(self, widget, page_no: int) -> Optional[FormField]:
        """解析单个 widget"""
        try:
            ftype = widget.field_type_string or 'Unknown'
            fname = widget.field_name or ''
            fvalue = widget.field_value or ''
            rect = tuple(widget.rect)
            flags = int(widget.field_flags) if hasattr(widget, 'field_flags') else 0
            f = FormField(
                field_type=ftype,
                field_name=fname,
                field_value=fvalue,
                rect=rect,
                page_no=page_no,
                flags=flags,
            )
            # 派生信息
            if ftype == 'CheckBox':
                # 复选框值: 'Yes'/'Off' 或 True/False
                f.is_checked = fvalue in ('Yes', 'yes', 'On', 'on', 'True', 'true', '1')
                f.display_text = '☑' if f.is_checked else '☐'
            elif ftype == 'RadioButton':
                f.is_checked = bool(fvalue) and fvalue != 'Off'
                f.display_text = '●' if f.is_checked else '○'
            elif ftype == 'Text':
                f.display_text = fvalue
                f.placeholder = widget.field_label or '' if hasattr(widget, 'field_label') else ''
            elif ftype == 'Signature':
                f.display_text = '[签名]'
                f.placeholder = '签名域'
            elif ftype == 'ComboBox':
                f.display_text = fvalue
            elif ftype == 'Button':
                f.display_text = fvalue or fname
            return f
        except Exception as e:
            if self.verbose:
                print(f'[Form] widget parse error: {e}')
            return None

    def is_signature_field(self, field: FormField) -> bool:
        """判断是否是签名域 (印章类)"""
        return field.field_type == 'Signature'

    def get_form_fields_for_page(self, page_no: int,
                                   fields_by_page: Dict[int, List[FormField]]) -> List[FormField]:
        """获取指定页的表单域"""
        return fields_by_page.get(page_no, [])

    @staticmethod
    def is_widget_in_signature_area(widget_rect, signature_fields: List[FormField]) -> bool:
        """判断某widget是否在签名域区域 (印章覆盖的)"""
        for sf in signature_fields:
            if (sf.rect[0] - 5 <= widget_rect[0] and widget_rect[2] <= sf.rect[2] + 5 and
                sf.rect[1] - 5 <= widget_rect[1] and widget_rect[3] <= sf.rect[3] + 5):
                return True
        return False
