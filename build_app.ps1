Write-Host ""
Write-Host "========================================================"
Write-Host "APTIM OS 수강현황 데스크톱 앱(exe) 빌드를 시작합니다..."
Write-Host "이 작업은 약 2~5분 정도 소요될 수 있습니다."
Write-Host "========================================================"
Write-Host ""

& .\venv\Scripts\pyinstaller --onefile --windowed --icon=NONE --name="APTIM_OS_수강현황_대시보드" --add-data="app.py;." --add-data="credentials.json;." --copy-metadata=streamlit --collect-data=streamlit --collect-data=plotly --collect-data=altair --collect-data=gspread --hidden-import=streamlit.runtime.scriptrunner.magic_funcs run_dash.py

Write-Host ""
Write-Host "========================================================"
Write-Host "완벽합니다! 앱 생성이 성공적으로 끝났습니다."
Write-Host "현재 폴더 안의 'dist' 폴더에 들어가시면"
Write-Host "[APTIM_OS_수강현황_대시보드.exe] 파일이 있습니다!"
Write-Host "이 파일 1개를 팀원들의 PC로 전달하시면 됩니다."
Write-Host "========================================================"
