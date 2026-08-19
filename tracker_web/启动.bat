@echo off
cd /d "%~dp0"
echo ============================================================
echo   跟单网页原型 - 启动器
echo   浏览器打开： http://127.0.0.1:8731
echo   关闭本窗口 = 停止服务 = 网页打不开
echo   启动日志写入同目录 server.log
echo ============================================================

REM 先释放 8731 端口上可能残留的旧进程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8731 ^| findstr LISTENING') do (
  echo 发现占用 8731 的旧进程 PID=%%a，正在结束...
  taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 >nul

echo 正在启动服务（新后台逻辑，支持跨页全选删除）...
"C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe" server.py > server.log 2>&1
echo.
echo 服务已停止。详见同目录 server.log
pause
