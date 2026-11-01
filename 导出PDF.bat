@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set "FOLDER=%~1"
if "%FOLDER%"=="" set /p "FOLDER=请把保存网页的文件夹（或 html 文件）拖进来，然后回车: "
if "%FOLDER%"=="" goto :end

rem 优先用同目录下的 python.exe（便携版），其次系统 PATH 里的 python / py
set "PY="
if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  where py >nul 2>nul && set "PY=py"
)
if not defined PY (
  echo.
  echo 没找到 Python。请安装 Python 3.9+ 并在安装时勾选 "Add python.exe to PATH"，
  echo 或者把一个 python.exe 放到本 bat 旁边。
  goto :end
)

echo.
"%PY%" -X utf8 "%~dp0fenbi_export.py" "%FOLDER%"

:end
echo.
pause
