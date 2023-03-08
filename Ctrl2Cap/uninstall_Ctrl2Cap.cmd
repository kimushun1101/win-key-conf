@echo off
chcp 65001 >nul
net session >nul 2>&1
if not %errorLevel% == 0 (
    echo 右クリックから管理者として実行をクリックしてください。
    pause > nul
    exit
)

cd /d %~dp0Ctrl2Cap
ctrl2cap /uninstall
echo Ctrl2cap uninstalled. You must reboot for this to take effect. と出ていたら再起動してください。
pause > nul