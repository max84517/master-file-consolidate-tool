@echo off
setlocal

:: Build to C:\Temp to avoid OneDrive WinError 5
set BUILD_DIR=C:\Temp\MCTool
set DIST_DIR=%BUILD_DIR%\dist
set WORK_DIR=%BUILD_DIR%\build

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"

echo === Building MasterConsolidateTool ===
poetry run pyinstaller --noconfirm ^
    --distpath "%DIST_DIR%" ^
    --workpath "%WORK_DIR%" ^
    MasterConsolidateTool.spec

if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)

:: Copy dist back to project folder
if exist "dist" rmdir /s /q "dist"
xcopy /E /I "%DIST_DIR%\MasterConsolidateTool" "dist\MasterConsolidateTool" >nul

echo.
echo === Build complete ===
echo Executable: dist\MasterConsolidateTool\MasterConsolidateTool.exe
