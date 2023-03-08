@echo off
chcp 65001 >nul

if exist "%~dp0muhenkan-switch" (goto :execute_muhenkan-switch) else (goto :confirm_download)

:confirm_download
set /p res="muhenkan-switch をダウンロードしてよいですか？(y=yes / n=no)？"
if /i {%res%}=={y} (goto :download_muhenkan-switch)
if /i {%res%}=={yes} (goto :download_muhenkan-switch)
echo yes でなかったため終了します。
pause > nul
exit

:download_muhenkan-switch
cd /d %~dp0
curl -OL https://github.com/kimushun1101/muhenkan-switch/releases/latest/download/muhenkan-switch.zip
powershell -Command "Expand-Archive -Path '%~dp0muhenkan-switch.zip' -Destination '%~dp0muhenkan-switch'"
del /q "%~dp0muhenkan-switch.zip"

:execute_muhenkan-switch
echo なにかのキーを入力後、フォルダが開きます。muhenkan.exe を起動してください。
pause > nul
start explorer "%~dp0muhenkan-switch"
