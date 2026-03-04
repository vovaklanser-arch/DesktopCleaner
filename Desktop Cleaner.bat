@echo off
chcp 65001 >nul
set "DIR=%~dp0"
if exist "%DIR%Lib" rd /s /q "%DIR%Lib" 2>nul
if exist "%DIR%pyvenv.cfg" del "%DIR%pyvenv.cfg" 2>nul
if exist "%USERPROFILE%\Lib" rd /s /q "%USERPROFILE%\Lib" 2>nul
if exist "%USERPROFILE%\pyvenv.cfg" del "%USERPROFILE%\pyvenv.cfg" 2>nul
set VIRTUAL_ENV=
set PYTHONHOME=
set PYTHONPATH=
cd /d "%TEMP%"
py -3 -m pip install -r "%DIR%requirements.txt" -q 2>nul
py -3 "%DIR%desktop_app.py"
if errorlevel 1 pause
