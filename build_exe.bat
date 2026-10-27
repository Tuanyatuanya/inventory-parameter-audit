@echo off
REM ==========================================================================
REM  Build the ASL Inventory Parameter Optimization Tool into a Windows .exe
REM ==========================================================================

echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building the executable with PyInstaller...
pyinstaller --onefile --windowed --name "ASL_Inventory_Parameter_Optimization_Tool" main.py

echo.
echo Done. The executable is in the "dist" folder:
echo     dist\ASL_Inventory_Parameter_Optimization_Tool.exe
echo.
pause
