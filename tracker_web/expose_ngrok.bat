@echo off
REM 外网穿透（ngrok，需先免费注册并配置 authtoken）
REM 前置：把 ngrok.exe 放到 PATH，并先执行一次  ngrok config add-authtoken <你的token>
REM 确保 server.py 已在本机 8731 运行，再双击本文件
where ngrok >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到 ngrok.exe。请先到 https://ngrok.com 下载并配置 token，放到 PATH 后重开本窗口。
  pause
  exit /b 1
)
echo 正在建立外网隧道，稍候出现 https://xxxx.ngrok-free.app 即为外网链接...
echo 关闭本窗口即停止外网访问。
ngrok http 8731
pause
