@echo off
cd /d "%~dp0"
echo 正在启动跟单网页原型（首次会自动导入数据）...
"C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe" import_data.py
"C:\Users\beiyou201\.workbuddy\binaries\python\envs\default\Scripts\python.exe" server.py
pause
