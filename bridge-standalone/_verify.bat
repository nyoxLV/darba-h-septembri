@echo off
rem Usage: _verify.bat [browser-args]   (default = same as real app)
cd /d "%~dp0"
if not exist build\webview2_data mkdir build\webview2_data

set "BRIDGE_BROWSER_ARGS=%1"
if "%BRIDGE_BROWSER_ARGS%"=="" set "BRIDGE_BROWSER_ARGS=--disable-gpu --remote-debugging-port=9333"
echo [verify] BRIDGE_BROWSER_ARGS = '%BRIDGE_BROWSER_ARGS%'

start "" /B powershell.exe -NoProfile -Command ^
  "Start-Process -FilePath 'build\Release\BridgeTest.exe' -WorkingDirectory '%~dp0.' -RedirectStandardOutput stdout.log -RedirectStandardError stderr.log"

set /a tries=0
:waitloop
timeout 2 >nul
node cdp_read.js > _cdp_out.txt 2>&1
findstr "rows" _cdp_out.txt >nul && goto :gotit
set /a tries+=1
if %tries% lss 8 goto :waitloop

:gotit
echo [verify] page state:
type _cdp_out.txt
echo.
echo [verify] C++ console log (stdout.log):
type stdout.log
