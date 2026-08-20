@echo off
chcp 936 >nul 2>&1
setlocal
cd /d "%~dp0"
set PORT=8731
set TUN=didi

REM ── 0. 检查 cloudflared ──
where cloudflared >nul 2>nul || (
  echo [错误] 未找到 cloudflared.exe。
  echo 请先看《外网稳定链接设置.md》第 1 步安装并登录（免费），然后重开本窗口。
  pause
  exit /b 1
)

REM ── 1. 选择 Python 并后台启动本地服务 ──
set "PY="
if exist "C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe" (
  set "PY=C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
) else (
  set "PY=python"
)
%PY% -c "import openpyxl" >nul 2>&1 || %PY% -m pip install openpyxl >nul 2>&1

netstat -ano | findstr :%PORT% | findstr LISTENING >nul 2>&1
if errorlevel 1 (
  echo 本地服务未运行，后台启动中...
  REM 释放端口
  for /f "tokens=5" %%a in ('netstat -ano ^| findstr :%PORT% ^| findstr LISTENING') do taskkill /PID %%a /F >nul 2>&1
  start "" %PY% server.py
  timeout /t 3 /nobreak >nul
) else (
  echo 本地服务已在 %PORT% 运行。
)

REM ── 2. 启动命名隧道（地址固定，跨重启不变）──
echo.
echo 正在建立稳定外网隧道（命名隧道：%TUN%）...
echo 首次使用需先执行一次（仅一次，按提示浏览器登录 Cloudflare 免费账号）：
echo   cloudflared tunnel login
echo   cloudflared tunnel create %TUN%
echo 创建后重跑本文件，即可获得固定地址： https://^<tunnel-id^>.cfargotunnel.com
echo （关闭本窗口即停止外网访问；本地服务仍在。）
echo.
cloudflared tunnel run --protocol http2 %TUN%
endlocal
