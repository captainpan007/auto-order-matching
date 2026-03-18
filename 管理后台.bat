@echo off
rem 激活 Conda 环境
call conda activate admin

rem 切换到 app.py 文件所在的目录
cd /d D:\www\images_ai

rem 启动 Flask 应用
start python -m flask run -h 0.0.0.0 -p 8000 --debug

rem 这一行可以让终端保持打开，方便查看输出信息
pause
