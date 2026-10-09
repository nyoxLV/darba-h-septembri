@echo off
rem ════════════════════════════════════════════════════════════════
rem  APEX MACRO v9 — standalone bridge harness: configure + build + run
rem ════════════════════════════════════════════════════════════════
setlocal

cd /d "%~dp0"

set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
if not exist "%VCVARS%" goto :err_vcvars

call "%VCVARS%" >nul || goto :err_vcvars

:have_toolchain
if exist build\CMakeCache.txt goto :do_build
echo === Configuring (downloads WebView2 SDK on first run) ===
cmake -S . -B build -G "Visual Studio 18 2026" -A x64 || goto :err_configure

:do_build
echo === Building Release ===
cmake --build build --config Release --target BridgeTest --parallel || goto :err_build

echo.
echo ════════════════════════════════════════════════
echo   Launching bridge test harness...
echo   Watch the window + console trace side by side.
echo ════════════════════════════════════════════════
build\Release\BridgeTest.exe

pause
exit /b 0

:err_vcvars
echo [ERROR] vcvars64.bat not found or failed at %VCVARS%
pause & exit /b 1

:err_configure
echo [ERROR] cmake configure failed
pause & exit /b 1

:err_build
echo [ERROR] build failed
pause & exit /b 1
