@echo off
set VENV=venv

:: ������⻷���Ƿ����
if not exist %VENV% (
    echo �������⻷��...
    python -m venv %VENV%
)

:: �������⻷��
call %VENV%\Scripts\activate

:: ��װ����
echo ��װ����...
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

echo ��װ��ɣ�
pause