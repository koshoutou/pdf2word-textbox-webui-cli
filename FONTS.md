# 字体安装指南

## 免责声明

> **本工具不附带任何字体文件。** 字体版权归各自所有者。本工具仅提供下载链接参考，不对字体授权负责。商业使用请确认字体的授权许可。

## 为什么需要安装字体？

PDF 中的文字使用特定字体（如仿宋、黑体、宋体等），转换后的 DOCX 保留了这些字体名。如果系统未安装对应字体，渲染时会被替换为其他字体，导致视觉差异。

## 快速安装

```bash
# 方式1: 使用安装脚本
chmod +x setup_fonts.sh
./setup_fonts.sh

# 方式2: 手动下载并安装
# 下载字体文件 → 放入 ~/.fonts/ → 执行 fc-cache -f
```

## 字体下载来源

### 推荐来源

| 来源 | 字体 | 说明 |
|------|------|------|
| [StellarCN/scp_zh](https://github.com/StellarCN/scp_zh/tree/master/fonts) | SimSun, SimHei | 宋体、黑体 |
| [zhyounger/FontsFromWindows](https://github.com/zhyounger/FontsFromWindows/tree/master/fonts) | 仿宋, 楷体等 | Windows 常用字体集合 |
| [DoveOutland/Common-Chinese-office-fonts](https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-) | 多种中文办公字体 | 常用中文办公字体库 |
| [晋城市财政局字体包](http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip) | 完整字体包 | 政府公文标准字体 |

### 字体回退映射

| PDF 字体 | 回退字体 | 说明 |
|----------|----------|------|
| 宋体 (SimSun) | SimSun | 已安装 |
| 黑体 (SimHei) | SimHei | 已安装 |
| 仿宋 (FangSong) | Noto Serif SC | 回退（serif 风格近似） |
| 楷体 (KaiTi) | LXGW WenKai | 回退（楷体风格近似） |
| 微软雅黑 | SimHei | 回退 |

## 字体安装详细步骤

### Linux

```bash
# 1. 下载字体文件
mkdir -p ~/.fonts
# 从上述来源下载 .ttf/.ttc 文件到 ~/.fonts/

# 2. 刷新字体缓存
fc-cache -f

# 3. 验证
fc-match "宋体"
fc-match "黑体"
```

### Windows

直接将字体文件双击安装，或复制到 `C:\Windows\Fonts\` 目录。

### macOS

将字体文件复制到 `~/Library/Fonts/` 或 `/Library/Fonts/`。

## 字体授权说明

- **SimSun/SimHei**: 微软公司开发，随 Windows 系统分发。非 Windows 系统使用需确认授权。
- **仿宋/楷体**: 中国国家标准字体，多家厂商有各自版本。
- **Noto Serif SC / Noto Sans SC**: Google 开源字体，SIL Open Font License。
- **LXGW WenKai**: 开源字体，SIL Open Font License。
- **FangSong_GB2312**: 国标字体，使用需确认授权。

> 商业用途请使用正版授权字体。本工具不对字体侵权承担责任。
