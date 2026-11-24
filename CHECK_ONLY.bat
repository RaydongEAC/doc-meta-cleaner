@echo off
chcp 936 >nul 2>nul
setlocal
set "SCRIPT=%~dp0clean_doc_meta.py"
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
py -3 "%SCRIPT%" --dry-run --deep --neutralize --scan %*
goto DONE
:RUN
"%PYEXE%" "%SCRIPT%" --dry-run --deep --neutralize --scan %*
goto DONE
:NOSCRIPT
echo ERROR - clean_doc_meta.py not found next to this batch file.
pause
exit /b 1
:NOPYTHON
echo ERROR - Python 3 not found. Keep the "python" folder, or install Python 3.
pause
exit /b 1
:DONE
echo.
echo ---- CHECK ONLY, no file was changed. Press any key to close ----
pause >nul
