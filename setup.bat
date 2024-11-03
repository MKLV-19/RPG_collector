@echo off
set VENV=venv

:: 检查虚拟环境是否存在
if not exist %VENV% (
    echo 创建虚拟环境...
    python -m venv %VENV%
)

:: 激活虚拟环境
call %VENV%\Scripts\activate

:: 安装依赖
echo 安装依赖...
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

echo 安装完成！
pause