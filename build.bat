@echo off
echo [Build] MSFS SimRate Monitor Build Script
echo ==========================================
echo.

echo [Build] Cleaning up old build files...
if exist dist del /q dist\* 2>nul
if exist build rmdir /s /q build 2>nul
if exist MSFS-SimRate-Monitor.spec del /q MSFS-SimRate-Monitor.spec 2>nul

echo [Build] Compiling MSFS-SimRate-Monitor using uv and PyInstaller...
uv run pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "MSFS-SimRate-Monitor" ^
    --icon "mini_gui_icon.ico" ^
    --add-data "mini_gui_icon.ico;." ^
    --add-data "SimConnect;SimConnect" ^
    --add-data "fonts;fonts" ^
    --hidden-import "winshell" ^
    --hidden-import "win32com.client" ^
    mini_gui.py

echo.
if exist dist\MSFS-SimRate-Monitor.exe (
    echo [Build] ✓ Success! Executable is in the dist folder.
) else (
    echo [Build] ✗ Build failed. Check the output above for errors.
)
echo.
pause