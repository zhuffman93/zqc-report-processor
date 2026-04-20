@echo off
REM Lastrada Report Processor - Compilation Script
REM Run this from inside the project folder to produce a standalone .exe

echo ========================================
echo  Lastrada Report Processor - Build
echo ========================================
echo.

REM Use the venv Python/pip so all packages are available
set VENV=C:\Users\tcbit\PyCharmMiscProject\.venv\Scripts

echo Installing / updating dependencies...
"%VENV%\pip.exe" install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: pip install failed
    pause & exit /b 1
)

echo Installing PyInstaller...
"%VENV%\pip.exe" install pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: PyInstaller install failed
    pause & exit /b 1
)

echo.
echo Compiling - this may take a minute...
echo.

"%VENV%\pyinstaller.exe" ^
    --onefile ^
    --windowed ^
    --name "Lastrada_Report_Processor" ^
    --hidden-import "pystray._win32" ^
    --hidden-import "PIL._tkinter_finder" ^
    fpc_processor.py

if errorlevel 1 (
    echo.
    echo ERROR: Compilation failed - check messages above
    pause & exit /b 1
)

echo.
echo ========================================
echo  BUILD SUCCESSFUL
echo ========================================
echo.
echo Executable: dist\Lastrada_Report_Processor.exe
echo.
pause
