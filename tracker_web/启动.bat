@echo off
chcp 936 >nul 2>&1
setlocal
cd /d "%~dp0"
set "PORT=8731"

REM ===== 1. 释放端口残留 =====
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :%PORT% ^| findstr LISTENING') do (
  echo [释放] 结束占用 %PORT% 的旧进程 PID=%%a
  taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak >nul 2>&1

REM ===== 2. 定位 Python：逐个候选"实测能跑代码"才算数 =====
REM     避免被 Windows 商店的假 python 占位符骗到
set "PY="
for /f "delims=" %%i in ('where python 2^>nul') do (
  if not defined PY (
    "%%i" -c "print(1)" >nul 2>&1
    if not errorlevel 1 set PY="%%i"
  )
)
if defined PY goto :havepy

for /f "delims=" %%i in ('where py 2^>nul') do (
  if not defined PY (
    "%%i" -3 -c "print(1)" >nul 2>&1
    if not errorlevel 1 set "PY=%%i -3"
  )
)
if defined PY goto :havepy

if exist "C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe" set "PY=C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if defined PY goto :havepy

echo.
echo [错误] 没有找到可用的 Python！
echo 下载安装：https://www.python.org/downloads/
echo 安装时务必勾选 "Add python.exe to PATH"，装完后重新双击本文件。
echo.
pause
exit /b 1

:havepy
echo [Python] %PY%

REM ===== 3. 确保 openpyxl（Excel 导入需要）=====
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [安装] 缺少 openpyxl，正在自动安装（首次需要网络）...
  %PY% -m pip install openpyxl
  if errorlevel 1 (
    echo.
    echo [错误] openpyxl 安装失败。请检查网络后重试，或手动执行：
    echo        %PY% -m pip install openpyxl
    echo.
    pause
    exit /b 1
  )
)

echo ============================================================
echo   跟单进度跟踪 - 本地服务
echo   本机打开： http://127.0.0.1:%PORT%
echo   关闭本窗口 = 停止服务
echo   若下方立即报错，窗口会保持打开，请截图报错内容
echo ============================================================
timeout /t 2 /nobreak >nul 2>&1
start "" http://127.0.0.1:%PORT%
%PY% server.py
echo.
echo [提示] 服务已停止（exit code=%errorlevel%）。若页面打不开，请截图上方报错。
pause
endlocal
