@echo off
chcp 936 >nul 2>nul
setlocal
set "SCRIPT=%~dp0install_menu.py"
if not exist "%SCRIPT%" goto NOSCRIPT
set "PYEXE=%~dp0python\python.exe"
if exist "%PYEXE%" goto RUN
where python >nul 2>nul
if errorlevel 1 goto TRYPY
set "PYEXE=python"
goto RUN
:TRYPY
where py >nul 2>nul
if errorlevel 1 goto NOPYTHON
py -3 "%SCRIPT%" install
goto DONE
:RUN
"%PYEXE%" "%SCRIPT%" install
goto DONE
:NOSCRIPT
echo ERROR - install_menu.py not found next to this batch file.
pause
exit /b 1
:NOPYTHON
echo ERROR - Python 3 not found. Keep the "python" folder, or install Python 3.
pause
exit /b 1
:DONE
echo.
pause
