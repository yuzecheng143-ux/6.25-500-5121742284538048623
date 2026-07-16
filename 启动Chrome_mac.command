#!/bin/bash
# 蒲公英达人采集 — Chrome调试模式启动器 (macOS)
# 双击此文件即可启动Chrome（首次可能需要右键→打开）

CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
PROFILE="$HOME/chrome-auto-profile"

# 检查Chrome是否已运行在调试端口
if curl -s http://127.0.0.1:9222/json > /dev/null 2>&1; then
    echo "Chrome调试端口已就绪，无需重复启动。"
    echo "现在可以运行 蒲公英采集 了。"
    read -p "按回车关闭此窗口..."
    exit 0
fi

if [ ! -f "$CHROME" ]; then
    echo "未找到Chrome，请确认已安装 Google Chrome。"
    read -p "按回车关闭此窗口..."
    exit 1
fi

echo "正在启动Chrome调试模式..."
echo "注意：这个Chrome窗口不要关闭，采集完成后再关。"
echo ""

nohup "$CHROME" \
  --remote-debugging-port=9222 \
  --remote-allow-origins=* \
  --user-data-dir="$PROFILE" \
  > /dev/null 2>&1 &

sleep 3

if curl -s http://127.0.0.1:9222/json > /dev/null 2>&1; then
    echo "Chrome已启动，现在可以运行 蒲公英采集 了。"
else
    echo "Chrome可能还在启动中，稍等片刻..."
fi

read -p "按回车关闭此窗口..."
