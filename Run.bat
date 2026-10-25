@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
title RPG 图片生成器

pushd "%~dp0"
if errorlevel 1 (
    echo 无法打开程序所在文件夹。
    pause
    exit /b 1
)

if not exist "RPGcollector.py" (
    echo 找不到 RPGcollector.py，请把 Run.bat 放回项目文件夹。
    goto failed
)
if not exist "requirements.txt" (
    echo 找不到 requirements.txt，请检查项目文件是否完整。
    goto failed
)

set "VENV_PY=%CD%\venv\Scripts\python.exe"
if not exist "%VENV_PY%" goto find_python
"%VENV_PY%" -c "import sys; assert sys.version_info >= (3, 8)" >nul 2>nul
if not errorlevel 1 goto check_dependencies
echo 现有 venv 环境无法使用。
echo 请将项目内的 venv 文件夹改名为 venv_old，再次双击 Run.bat。
goto failed

:find_python
echo 首次运行，正在准备运行环境……
py -3 -c "import sys; assert sys.version_info >= (3, 8)" >nul 2>nul
if not errorlevel 1 goto create_with_py
python -c "import sys; assert sys.version_info >= (3, 8)" >nul 2>nul
if not errorlevel 1 goto create_with_python
echo 未找到可用的 Python 3.8 或更高版本。
echo 请先安装 Python，安装时勾选 Add Python to PATH，然后重新双击 Run.bat。
goto failed

:create_with_py
py -3 -m venv "%CD%\venv"
if errorlevel 1 goto failed
goto check_dependencies

:create_with_python
python -m venv "%CD%\venv"
if errorlevel 1 goto failed

:check_dependencies
echo [1/3] 正在检查运行依赖……
"%VENV_PY%" -c "import pandas; from PIL import Image, ImageDraw, ImageFont; from openpyxl import load_workbook" >nul 2>nul
if not errorlevel 1 goto generate
echo 正在自动安装依赖，首次安装需要联网，请稍候……
"%VENV_PY%" -m pip install -r "%CD%\requirements.txt" --disable-pip-version-check --retries 2 --timeout 20
if not errorlevel 1 goto verify_dependencies
echo 安装未成功，正在尝试备用下载源……
"%VENV_PY%" -m pip install -r "%CD%\requirements.txt" -i https://mirrors.aliyun.com/pypi/simple/ --disable-pip-version-check --retries 2 --timeout 20
if errorlevel 1 goto failed

:verify_dependencies
"%VENV_PY%" -c "import pandas; from PIL import Image, ImageDraw, ImageFont; from openpyxl import load_workbook"
if errorlevel 1 goto failed

:generate
echo [2/3] 正在生成图片……
"%VENV_PY%" "%CD%\RPGcollector.py"
if errorlevel 1 goto failed
echo [3/3] 生成完成，正在打开 output 文件夹。
start "" "%CD%\output"
popd
endlocal
exit /b 0

:failed
echo.
echo 本次运行未完成，请查看上方的错误信息。
echo 如果安装依赖失败，请检查网络后再次双击 Run.bat。
pause
popd
endlocal
exit /b 1
