#!/bin/bash
# setup_fonts.sh - 中文字体安装脚本
# 下载并安装 PDF→DOCX 转换所需的常用中文字体
#
# 字体来源:
#   - SimSun/SimHei: https://github.com/StellarCN/scp_zh
#   - 仿宋/楷体等: https://github.com/zhyounger/FontsFromWindows
#   - 晋城市财政局字体包: http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip
#
# 免责声明: 本工具不附带任何字体文件, 用户需自行下载安装。
# 字体版权归各自所有者, 本工具仅提供下载链接参考, 不对字体授权负责。
# 商业使用请确认字体的授权许可。

set -e

FONTS_DIR="${HOME}/.fonts"
CONFIG_DIR="${HOME}/.config/fontconfig/conf.d"

echo "=========================================="
echo "  PDF→DOCX 转换器 - 中文字体安装脚本"
echo "=========================================="
echo ""

# 创建目录
mkdir -p "$FONTS_DIR"
mkdir -p "$CONFIG_DIR"

# 下载字体
download_font() {
    local name="$1"
    local url="$2"
    local outfile="$FONTS_DIR/$3"

    if [ -f "$outfile" ]; then
        echo "  ✓ $name 已存在, 跳过"
        return 0
    fi

    echo "  下载 $name ..."
    if curl -sL "$url" -o "$outfile" 2>/dev/null; then
        local size=$(stat -c%s "$outfile" 2>/dev/null || echo 0)
        if [ "$size" -gt 100000 ]; then
            echo "    ✓ $name 下载成功 ($(( size / 1024 ))KB)"
        else
            echo "    ✗ $name 下载失败或文件不完整"
            rm -f "$outfile"
            return 1
        fi
    else
        echo "    ✗ $name 下载失败"
        return 1
    fi
}

echo "=== 下载常用中文字体 ==="
echo ""
echo "提示: 如果自动下载失败, 请手动从以下地址下载:"
echo "  1. https://github.com/StellarCN/scp_zh/tree/master/fonts"
echo "  2. https://github.com/zhyounger/FontsFromWindows/tree/master/fonts"
echo "  3. https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-"
echo "  4. http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip (晋城市财政局)"
echo ""

# 从 GitHub 下载
download_font "宋体 (SimSun)" \
    "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimSun.ttf" \
    "SimSun.ttf" || true

download_font "黑体 (SimHei)" \
    "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimHei.ttf" \
    "SimHei.ttf" || true

# 楷体/仿宋 可从其他源下载
# download_font "楷体 (KaiTi)" "URL" "KaiTi.ttf" || true
# download_font "仿宋 (FangSong)" "URL" "FangSong.ttf" || true

echo ""
echo "=== 配置字体回退 ==="
cat > "$CONFIG_DIR/99-chinese-fonts.conf" << 'FONTCFG'
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <match target="pattern"><test name="family"><string>宋体</string></test><edit name="family" mode="assign" binding="strong"><string>SimSun</string></edit></match>
  <match target="pattern"><test name="family"><string>SimSun</string></test><edit name="family" mode="assign" binding="strong"><string>SimSun</string></edit></match>
  <match target="pattern"><test name="family"><string>黑体</string></test><edit name="family" mode="assign" binding="strong"><string>SimHei</string></edit></match>
  <match target="pattern"><test name="family"><string>SimHei</string></test><edit name="family" mode="assign" binding="strong"><string>SimHei</string></edit></match>
  <match target="pattern"><test name="family"><string>仿宋</string></test><edit name="family" mode="assign" binding="strong"><string>Noto Serif SC</string></edit></match>
  <match target="pattern"><test name="family"><string>FangSong</string></test><edit name="family" mode="assign" binding="strong"><string>Noto Serif SC</string></edit></match>
  <match target="pattern"><test name="family"><string>楷体</string></test><edit name="family" mode="assign" binding="strong"><string>LXGW WenKai</string></edit></match>
  <match target="pattern"><test name="family"><string>KaiTi</string></test><edit name="family" mode="assign" binding="strong"><string>LXGW WenKai</string></edit></match>
  <match target="pattern"><test name="family"><string>微软雅黑</string></test><edit name="family" mode="assign" binding="strong"><string>SimHei</string></edit></match>
</fontconfig>
FONTCFG

echo "=== 刷新字体缓存 ==="
fc-cache -f 2>/dev/null

echo ""
echo "=== 验证字体安装 ==="
for f in 宋体 黑体 仿宋 楷体 微软雅黑; do
    match=$(fc-match "$f" 2>/dev/null | cut -d: -f1 | xargs basename 2>/dev/null)
    echo "  $f → $match"
done

echo ""
echo "=========================================="
echo "  字体安装完成!"
echo "=========================================="
echo ""
echo "如需更多字体, 请从以下地址手动下载并放入 ~/.fonts/:"
echo "  - Windows 字体集合: https://github.com/zhyounger/FontsFromWindows"
echo "  - 常用中文办公字体: https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-"
echo "  - 晋城市财政局字体包: http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip"
echo ""
echo "免责声明: 字体版权归各自所有者。商业使用请确认授权许可。"
