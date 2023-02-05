@echo off
chcp 65001 >nul
net session >nul 2>&1
if not %errorLevel% == 0 (
    echo 右クリックから管理者として実行をクリックしてください。
    pause > nul
    exit
)

if exist "%~dp0Ctrl2Cap" (goto :execute_ctrl2cap) else (goto :confirm_download)

:confirm_download
@REM start https://learn.microsoft.com/en-us/sysinternals/downloads/ctrl2cap
set /p res="Ctrl2Cap をダウンロードしてよいですか？(y=yes / n=no)？"
if /i {%res%}=={y} (goto :download_ctrl2cap)
if /i {%res%}=={yes} (goto :download_ctrl2cap)
echo yes でなかったため終了します。
pause > nul
exit

:download_ctrl2cap
cd /d %~dp0
curl https://download.sysinternals.com/files/Ctrl2Cap.zip --output "%~dp0Ctrl2Cap.zip"
powershell -Command "Expand-Archive -Path "%~dp0Ctrl2Cap.zip" -Destination "%~dp0Ctrl2Cap""

:execute_ctrl2cap
cd /d "%~dp0Ctrl2Cap"
ctrl2cap /install
echo Ctrl2cap successfully installed. You must reboot for it to take effect. と出ていたら再起動してください。
pause > nul