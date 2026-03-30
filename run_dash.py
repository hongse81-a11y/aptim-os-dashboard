import os
import sys
import streamlit.web.cli as stcli

# PyInstaller가 패키징할 때 아래 라이브러리들을 누락하지 않고 알아채도록(Tracing) 명시적으로 적어줍니다.
import streamlit
import pandas
import gspread
import google.oauth2.service_account
import plotly.express
import altair

# Streamlit 내부적으로 실행시키는 매직 함수들이 PyInstaller에서 종종 누락되므로 수동으로 추가합니다.
import streamlit.runtime.scriptrunner.magic_funcs

def resolve_path(path):
    # PyInstaller로 패키징된 실행 파일(.exe)인 경우, 임시 폴더(_MEIPASS)에서 파일을 찾습니다.
    if getattr(sys, 'frozen', False):
        resolved_path = os.path.join(sys._MEIPASS, path)
    else:
        resolved_path = os.path.abspath(path)
    return resolved_path

if __name__ == "__main__":
    # app.py 의 절대 경로를 가져와서 Streamlit 엔진으로 실행시킵니다.
    script_path = resolve_path("app.py")
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())
