@echo off
REM 外网穿透（Cloudflare 快速隧道，免注册）
REM 前置：把 cloudflared.exe 放到 PATH（例如 C:\Windows 或本目录）
REM 确保 server.py 已在本机 8731 运行，再双击本文件
where cloudflared >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到 cloudflared.exe。请先到 https://github.com/cloudflare/cloudflared/releases 下载，
  echo        放到 C:\Windows 或把它的目录加入系统 PATH，再重开本窗口。
  pause
  exit /b 1
)
echo 正在建立外网隧道，稍候出现 https://xxxx.trycloudflare.com 即为外网链接...
echo 关闭本窗口即停止外网访问。
cloudflared tunnel --url http://localhost:8731
pause
