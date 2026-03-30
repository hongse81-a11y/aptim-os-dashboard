@echo off
chcp 65001 >nul
echo.
echo ========================================================
echo Starting PyInstaller build...
echo This process will take 2-5 minutes.
echo ========================================================
echo.

.\venv\Scripts\pyinstaller --onefile --windowed --icon=NONE --name="APTIM_OS_Dashboard" --add-data="app.py;." --add-data="credentials.json;." --copy-metadata=streamlit --collect-data=streamlit --collect-data=plotly --collect-data=altair --collect-data=gspread run_dash.py

echo.
echo ========================================================
echo Build complete! Check the 'dist' folder.
echo You will find [APTIM_OS_Dashboard.exe] inside!
echo Please rename it to Korean if needed.
echo ========================================================
pause
