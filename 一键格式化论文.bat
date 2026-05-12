@echo off
chcp 65001 >nul
setlocal EnableExtensions

title 一键格式化论文

echo ========================================
echo 一键格式化论文
echo ========================================
echo.

REM 定位本工具目录：bat 放在哪里，就从哪里找 scripts
set "TOOL_DIR=%~dp0"
set "SCRIPT=%TOOL_DIR%scripts\format_paper_with_targeted_repair_loop.py"

if not exist "%SCRIPT%" (
  echo [错误] 找不到格式化脚本：
  echo %SCRIPT%
  echo.
  echo 请确认这个 bat 文件在工具根目录里，不要单独拿出去。
  pause
  exit /b 1
)

REM 自动寻找 Python。优先用 Windows Python Launcher: py -3
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set "PY=python"
  ) else (
    echo [错误] 没找到 Python。
    echo.
    echo 请先安装 Python 3，然后重新运行本工具。
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时建议勾选：Add python.exe to PATH
    pause
    exit /b 1
  )
)

REM 获取输入文件：支持把 docx 拖到 bat 上，也支持手动输入路径
if "%~1"=="" (
  echo 请把论文 docx 文件拖到这个窗口里，然后按回车：
  set /p "INPUT="
) else (
  set "INPUT=%~1"
)

REM 去掉用户手动输入时可能带上的引号
set "INPUT=%INPUT:"=%"

if not exist "%INPUT%" (
  echo.
  echo [错误] 找不到这个文件：
  echo %INPUT%
  echo.
  pause
  exit /b 1
)

for %%F in ("%INPUT%") do (
  set "OUT=%%~dpnF_格式化后%%~xF"
)

echo.
echo 使用 Python：%PY%
echo 输入文件：%INPUT%
echo 输出文件：%OUT%
echo.

REM 自动安装依赖 lxml
echo 正在检查依赖 lxml...
%PY% -c "import lxml" >nul 2>nul
if not %errorlevel%==0 (
  echo 未检测到 lxml，正在自动安装...
  %PY% -m pip install lxml
  if not %errorlevel%==0 (
    echo.
    echo [错误] lxml 安装失败。请检查网络，或手动运行：
    echo %PY% -m pip install lxml
    echo.
    pause
    exit /b 1
  )
)

echo.
echo 开始格式化，请稍等...
echo 将执行：格式化 + 定向修复循环 + 最终目录/三线表强制收尾
%PY% "%SCRIPT%" "%INPUT%" "%OUT%" --max-loops 8

if not %errorlevel%==0 (
  echo.
  echo [失败] 格式化过程中出现错误。
  echo 请把本窗口里的报错截图发给维护者。
  echo.
  pause
  exit /b 1
)

echo.
echo ========================================
echo 格式化完成！
echo 输出文件：
echo %OUT%
echo ========================================
echo.
echo 建议打开新文件，重点检查目录、摘要、标题、正文、表格、图片、页码。
echo.
pause
