import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
import os
import sys

# 1. 페이지 기본 설정
st.set_page_config(page_title="APTIM OS 수강현황", layout="wide")

# 구글 시트 이름 및 ID
SPREADSHEET_NAME = "APTIM OS 수강 현황"
AFFILIATION_SHEET_ID = "1dLbbqbgyPZGAeEZDLkO86Pr9lTvVkVFVDOMeMIl1CFw"

@st.cache_data(ttl=600)  # 데이터를 10분마다 새로 구글 시트에서 가져옴
def load_data():
    try:
        # 서비스 계정 인증
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # 실행 파일(.exe)로 빌드된 경우와 일반 파이썬 실행 경로 구분
        if getattr(sys, 'frozen', False):
            cred_path = os.path.join(sys._MEIPASS, 'credentials.json')
        else:
            cred_path = 'credentials.json'

        # credentials.json 파일이 있으면 파일로 인증
        if os.path.exists(cred_path):
            credentials = Credentials.from_service_account_file(cred_path, scopes=scopes)
        else:
            # 배포된 Cloud 환경(Streamlit Cloud)에서는 Secrets 저장소에서 인증키 정보를 불러옴
            credentials = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scopes
            )
            
        gc = gspread.authorize(credentials)
        spreadsheet = gc.open(SPREADSHEET_NAME)
        
        all_data = []
        
        # 모든 시트(탭)를 순회하며 데이터 수집
        for worksheet in spreadsheet.worksheets():
            ws_title = worksheet.title  # 탭 이름 (예: '3월 1주')
            try:
                values = worksheet.get_all_values()
                if len(values) > 1:
                    # '이름' 이라는 글자가 포함된 행을 찾아 헤더로 지정
                    header_index = 0
                    for i, row in enumerate(values):
                        if '이름' in row:
                            header_index = i
                            break
                            
                    headers = values[header_index]
                    
                    # 중복된 컬럼명 이름 변경 (Reindexing 오류 해결)
                    seen = {}
                    new_headers = []
                    for h in headers:
                        h_str = str(h).strip()
                        if not h_str:
                            h_str = "Unnamed"
                        if h_str in seen:
                            seen[h_str] += 1
                            new_headers.append(f"{h_str}_{seen[h_str]}")
                        else:
                            seen[h_str] = 0
                            new_headers.append(h_str)
                            
                    # 그 다음 줄부터 실제 데이터로 변환
                    df_temp = pd.DataFrame(values[header_index+1:], columns=new_headers)
                    
                    # 빈 행 삭제 및 '순위'나 '이름'이 없는 빈 줄 데이터 완벽 제외
                    if '이름' in df_temp.columns:
                        df_temp = df_temp[df_temp['이름'].astype(str).str.strip() != '']
                        
                        # 특정 인원 제외
                        excluded_names = ['박태상', '박진석', '이영호', '이정혁', '성준규', '최준영', '이용철', '석재호', '권세희', '이한구', '이정규', '이서후', '김현수']
                        df_temp = df_temp[~df_temp['이름'].astype(str).str.strip().isin(excluded_names)]
                        
                    df_temp.dropna(how='all', inplace=True)
                    
                    if not df_temp.empty:
                        # '주차' 컬럼을 추가하여 탭 이름 기록
                        df_temp['주차'] = ws_title
                        all_data.append(df_temp)
            except Exception as inner_e:
                print(f"[{ws_title}] 탭 읽기 오류 무시: {inner_e}")
                
        if not all_data:
            return pd.DataFrame()
            
        # 모든 주차 데이터를 하나의 표(DataFrame)로 합치기
        df = pd.concat(all_data, ignore_index=True)
        
        # '강의' 글자 제거 (예: '44개 강의' -> '44개')
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.replace(' 강의', '', regex=False).str.replace('강의', '', regex=False)
        
        # 텍스트와 숫자가 섞여있는 데이터에서 숫자(소수점 포함)만 깔끔하게 추출
        cols_to_clean = ['평균 등급', '참여율', '완료율']
        for col in cols_to_clean:
            if col in df.columns:
                df[col] = df[col].astype(str)
                # 정규표현식: 숫자와 소수점만 찾아서 첫번째 매칭값 추출 (예: '5.27\\nA-' -> '5.27', '90%' -> '90')
                df[col] = df[col].str.extract(r'([0-9.]+)', expand=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
        # (옵션) '완료율'이 100 넘어가는 경우 등이 있다면 100 이하로 고정
        if '완료율' in df.columns:
            df['완료율'] = df['완료율'].apply(lambda x: min(x, 100) if pd.notnull(x) else x)
            
        # '제출 과제', '수강 코스' 등에 있는 '개' 글자도 지우고 숫자형으로 변환해두면 좋습니다.
        for col in ['제출 과제', '수강 코스']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.extract(r'([0-9.]+)', expand=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
        return df
    
    except Exception as e:
        st.error(f"구글 시트 데이터 로드 중 오류 발생: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def load_affiliation_data():
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        if getattr(sys, 'frozen', False):
            cred_path = os.path.join(sys._MEIPASS, 'credentials.json')
        else:
            cred_path = 'credentials.json'

        if os.path.exists(cred_path):
            credentials = Credentials.from_service_account_file(cred_path, scopes=scopes)
        else:
            credentials = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scopes
            )
            
        gc = gspread.authorize(credentials)
        # ID로 시트 열기
        spreadsheet = gc.open_by_key(AFFILIATION_SHEET_ID)
        worksheet = spreadsheet.get_worksheet(0) # 첫 번째 시트
        
        values = worksheet.get_all_values()
        if len(values) > 1:
            # 이름(B열/index 1), 소속(G열/index 6), 기수(I열/index 8)
            data = []
            for row in values[1:]: # 헤더 제외
                name = row[1].strip() if len(row) > 1 else ""
                aff = row[6].strip() if len(row) > 6 else ""
                cohort = row[8].strip() if len(row) > 8 else ""
                if name:
                    data.append({'이름': name, '소속': aff, '기수': cohort})
            
            return pd.DataFrame(data).drop_duplicates(subset=['이름'])
        return pd.DataFrame()
    except Exception as e:
        st.sidebar.error(f"소속 및 기수 매핑 데이터를 불러오지 못했습니다: {e}")
        return pd.DataFrame()

# ==================== UI 구성 ==================== #
st.title("📊 APTIM OS 주간 수강현황 대시보드")
st.markdown("---")

df = load_data()
aff_df = load_affiliation_data()

# 소속 정보 병합
if not df.empty and not aff_df.empty:
    df = df.merge(aff_df, on='이름', how='left')
    # 소속 정보가 없는 경우 '미분류' 처리
    if '소속' in df.columns:
        df['소속'] = df['소속'].fillna('미분류')

# 데이터 로딩 실패 처리
if df.empty:
    st.warning(f"'{SPREADSHEET_NAME}' 시트에서 데이터를 찾을 수 없습니다. (1) 시트 이름을 정확히 확인하시고, (2) '공유'에 서비스 계정(credentials.json 이메일) 편집자 추가를 꼭 진행해 주세요!")
    st.stop()

# --- 사이드바 필터 --- #
st.sidebar.header("🔍 대시보드 필터")

# 이름 검색 드롭다운 (다중 선택)
all_names = sorted(df['이름'].dropna().unique().tolist())

# 소속별 그룹 선택 추가
if '소속' in df.columns:
    all_groups = sorted(df['소속'].dropna().unique().tolist())
    selected_groups = st.sidebar.multiselect("소속 그룹 선택 (그룹 단위 선택)", options=all_groups, default=[], placeholder="그룹 선택")
    
    # 선택된 그룹에 속한 이름들을 기본값으로 설정하기 위함
    group_member_names = df[df['소속'].isin(selected_groups)]['이름'].unique().tolist() if selected_groups else []
else:
    selected_groups = []
    group_member_names = []

# 개별 인원 선택 (그룹 선택에 따른 초기값 연동)
# multiselect의 default 값을 group_member_names로 설정하면 그룹 선택 시 자동으로 채워짐
selected_names = st.sidebar.multiselect(
    "개별/복수 인원 선택 (추세 분석)", 
    options=all_names, 
    default=group_member_names, 
    placeholder="선택 안 함 (전체 보기)"
)

# 필터링 적용 (아무도 선택하지 않은 경우 원본 전체 df 사용)
if not selected_names:
    filtered_df = df
else:
    filtered_df = df[df['이름'].isin(selected_names)]

# 주차 선택 (사이드바)
st.sidebar.markdown("---")
all_weeks = df['주차'].unique().tolist()
selected_week = st.sidebar.selectbox("📅 기준 주차 선택 (지표 및 표)", options=all_weeks, index=len(all_weeks)-1 if all_weeks else 0)

latest_week = selected_week
prev_week = ""
if latest_week in all_weeks:
    current_idx = all_weeks.index(latest_week)
    if current_idx > 0:
        prev_week = all_weeks[current_idx - 1]

# --- 1. 상단 주요 요약 지표 (선택된 인원의 지난 주차 대비) --- #
st.subheader("💡 주차별 요약")
if prev_week:
    st.write(f"**기준 주차:** {latest_week} (전주 '{prev_week}' 대비 변화량)")
else:
    st.write(f"**기준 주차:** {latest_week} (비교할 이전 주차 기록이 없습니다.)")

col1, col2, col3, col4, col5 = st.columns(5)

latest_data = filtered_df[filtered_df['주차'] == latest_week]
prev_data = filtered_df[filtered_df['주차'] == prev_week] if prev_week else pd.DataFrame()

if not latest_data.empty:
    avg_completion = latest_data['완료율'].mean() if '완료율' in latest_data.columns else None
    avg_score = latest_data['평균 등급'].mean() if '평균 등급' in latest_data.columns else None
    avg_participation = latest_data['참여율'].mean() if '참여율' in latest_data.columns else None
    avg_courses = latest_data['수강 코스'].mean() if '수강 코스' in latest_data.columns else None
    avg_assignments = latest_data['제출 과제'].mean() if '제출 과제' in latest_data.columns else None
    
    delta_completion = None
    delta_score = None
    delta_participation = None
    delta_courses = None
    delta_assignments = None
    
    # 전주 데이터가 존재하면 변화량(Delta) 계산
    if not prev_data.empty:
        prev_completion = prev_data['완료율'].mean() if '완료율' in prev_data.columns else None
        prev_score = prev_data['평균 등급'].mean() if '평균 등급' in prev_data.columns else None
        prev_participation = prev_data['참여율'].mean() if '참여율' in prev_data.columns else None
        prev_courses = prev_data['수강 코스'].mean() if '수강 코스' in prev_data.columns else None
        prev_assignments = prev_data['제출 과제'].mean() if '제출 과제' in prev_data.columns else None
        
        delta_completion = f"{avg_completion - prev_completion:.1f}%" if pd.notnull(avg_completion) and pd.notnull(prev_completion) else None
        delta_score = f"{avg_score - prev_score:.1f}점" if pd.notnull(avg_score) and pd.notnull(prev_score) else None
        delta_participation = f"{avg_participation - prev_participation:.1f}%" if pd.notnull(avg_participation) and pd.notnull(prev_participation) else None
        # 수강 코스(강의 개수)는 소수점 없이 정수로 표현
        delta_courses = f"{int(avg_courses - prev_courses)}개" if pd.notnull(avg_courses) and pd.notnull(prev_courses) else None
        delta_assignments = f"{int(avg_assignments - prev_assignments)}개" if pd.notnull(avg_assignments) and pd.notnull(prev_assignments) else None

    # Delta 값을 포함하여 화면에 표시 (표시할 때 값이 숫자형이 아닐 경우(N/A)도 안전하게 처리)
    val_comp = f"{avg_completion:.1f}%" if pd.notnull(avg_completion) else "N/A"
    val_score = f"{avg_score:.1f}점" if pd.notnull(avg_score) else "N/A"
    val_part = f"{avg_participation:.1f}%" if pd.notnull(avg_participation) else "N/A"
    val_courses = f"{int(avg_courses)}개" if pd.notnull(avg_courses) else "N/A"
    val_assignments = f"{int(avg_assignments)}개" if pd.notnull(avg_assignments) else "N/A"

    col1.metric("📊 평균 완료율", val_comp, delta=delta_completion)
    col2.metric("📋 평균 수강 코스", val_courses, delta=delta_courses)
    col3.metric("🎯 평균 평가 점수", val_score, delta=delta_score)
    col4.metric("🤝 주간 참여율", val_part, delta=delta_participation)
    col5.metric("📝 평균 제출 과제", val_assignments, delta=delta_assignments)

st.markdown("---")

# --- 2. 시각화: 개인별 주간 변화 추이 (꺾은선 그래프) --- #
st.subheader("📈 주간 학습 성과 변화 가시화")
st.markdown("특정 인원에 대한 매주의 학습 진행 경과를 추적합니다.")

# 탭을 나눠서 덜 복잡하게 표현
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["수강 코스(강의 개수) 추이", "완료율 진척도 추이", "평균 점수 추이", "등급참여율 추이", "제출 과제 추이", "⚠️ 기수별 최하위 관리 리포트"])

with tab1:
    if '수강 코스' in filtered_df.columns:
        fig_course = px.bar(filtered_df, x='주차', y='수강 코스', color='이름', text_auto=True,
                            title='수강을 완료한 코스(강의) 개수 변화', barmode='group')
        st.plotly_chart(fig_course, use_container_width=True)

with tab2:
    if '완료율' in filtered_df.columns:
        fig_completion = px.line(filtered_df, x='주차', y='완료율', color='이름', markers=True,
                                 title='과정 완료율 변화 (%)')
        fig_completion.update_layout(yaxis=dict(range=[0, 105]))
        st.plotly_chart(fig_completion, use_container_width=True)

with tab3:
    if '평균 등급' in filtered_df.columns:
        fig_score = px.line(filtered_df, x='주차', y='평균 등급', color='이름', markers=True,
                            title='평가 평균 점수 변화 추세')
        fig_score.update_layout(yaxis=dict(range=[0, 105]))
        st.plotly_chart(fig_score, use_container_width=True)

with tab4:
    if '참여율' in filtered_df.columns:
        fig_part = px.line(filtered_df, x='주차', y='참여율', color='이름', markers=True,
                           title='참여율 추이')
        st.plotly_chart(fig_part, use_container_width=True)

with tab5:
    if '제출 과제' in filtered_df.columns:
        fig_assign = px.bar(filtered_df, x='주차', y='제출 과제', color='이름', text_auto=True,
                            title='제출 과제 통합 변화', barmode='group')
        st.plotly_chart(fig_assign, use_container_width=True)

with tab6:
    st.markdown("### ⚠️ 기수별 최하위 2명 모니터링 리포트")
    st.markdown("선택된 **기준 주차**의 총 완료율 30%와 과제 제출 총량 70%를 가중치 순위(Weighted Rank Sum)로 반영하여 기수별 최하위 2명을 선정합니다.")
    st.markdown("* 1기 인원은 최하위 선정 대상에서 제외됩니다.")

    if '기수' in df.columns:
        # 기수 데이터 전처리
        df['기수'] = df['기수'].fillna('미분류').astype(str).str.strip().replace('', '미분류')
        
        # 1기를 제외한 유효 기수 추출
        cohorts = sorted([c for c in df['기수'].unique() if c and c != '1기' and c != 'Drop-out' and c != '미분류'])
        
        latest_all = df[df['주차'] == latest_week].copy()
        
        if not latest_all.empty:
            # 이번주의 완료율 및 제출과제 데이터 준비
            metrics = latest_all[['이름', '완료율', '제출 과제', '기수', '소속']].copy()
            
            # 수치 전처리
            metrics['완료율'] = pd.to_numeric(metrics['완료율'], errors='coerce').fillna(0)
            metrics['제출 과제'] = pd.to_numeric(metrics['제출 과제'], errors='coerce').fillna(0)
            
            bottom_members = []
            
            for cohort in cohorts:
                cohort_df = metrics[metrics['기수'] == cohort].copy()
                if cohort_df.empty:
                    continue
                
                # 순위 계산 (오름차순: 실적이 낮을수록 순위 점수가 낮아짐, 공동 최하위 발생 시 method='min' 적용)
                cohort_df['완료율_순위'] = cohort_df['완료율'].rank(method='min', ascending=True)
                cohort_df['과제제출_순위'] = cohort_df['제출 과제'].rank(method='min', ascending=True)
                
                # 두 순위의 가중 합산 점수 (완료율 30%, 과제제출 70%)
                cohort_df['순위합산'] = (cohort_df['완료율_순위'] * 0.3) + (cohort_df['과제제출_순위'] * 0.7)
                
                # 가중 순위합산이 낮은 순(실적이 가장 나쁜 순)으로 정렬
                cohort_df = cohort_df.sort_values(by=['순위합산', '완료율', '제출 과제'], ascending=[True, True, True])
                
                # 하위 2명 선정
                bottom_2 = cohort_df.head(2)
                bottom_members.append(bottom_2)
            
            if bottom_members:
                final_bottom_df = pd.concat(bottom_members, ignore_index=True)
                
                # 기수 정렬순으로 표시
                final_bottom_df = final_bottom_df.sort_values(by=['기수', '순위합산'], ascending=[True, True])
                
                st.markdown("#### 🚨 기수별 관리 대상 (하위 2명) 현황")
                
                # 표용 데이터 프레임 정비
                display_cols = ['기수', '이름', '소속', '완료율', '제출 과제']
                df_to_show = final_bottom_df[display_cols].copy()
                df_to_show.columns = ['기수', '이름', '소속', '총 완료율', '과제 제출 총량']
                
                # 단위 포맷팅
                df_to_show['총 완료율'] = df_to_show['총 완료율'].map(lambda x: f"{x:.1f}%")
                df_to_show['과제 제출 총량'] = df_to_show['과제 제출 총량'].map(lambda x: f"{int(x)}개")
                
                st.dataframe(df_to_show, use_container_width=True, hide_index=True)
                
                # 시각적 가시성을 위한 기수별 요약 카드 렌더링
                st.markdown("#### 📌 기수별 모니터링 대상자 코멘트")
                cols = st.columns(3)
                for idx, cohort in enumerate(cohorts):
                    cohort_bottom = final_bottom_df[final_bottom_df['기수'] == cohort]
                    if cohort_bottom.empty:
                        continue
                    
                    with cols[idx % 3]:
                        details_list = []
                        for _, row in cohort_bottom.iterrows():
                            details_list.append(
                                f"📍 **{row['이름']}** ({row['소속']})<br>"
                                f"&nbsp;&nbsp; 총 완료율: {row['완료율']:.1f}% | 과제 제출 총량: {int(row['제출 과제'])}개"
                            )
                        details_str = "<br><br>".join(details_list)
                        
                        st.markdown(
                            f"""
                            <div style="background-color: #fff5f5; border-left: 5px solid #ff4d4d; padding: 15px; border-radius: 6px; margin-bottom: 15px; min-height: 140px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                                <h4 style="margin: 0 0 10px 0; color: #d93838; font-weight: bold;">🏷️ {cohort}</h4>
                                <div style="font-size: 13.5px; color: #4a4a4a; line-height: 1.5;">
                                    {details_str}
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
            else:
                st.info("조건에 만족하는 하위 대상 데이터가 없습니다.")
        else:
            st.info("선택한 기준 주차에 해당하는 수강 데이터가 존재하지 않습니다.")
    else:
        st.warning("데이터에 기수 정보가 없습니다. 관리자에게 문의하세요.")


# --- 3. 상세 데이터 표 --- #
st.markdown("---")
st.subheader(f"📋 선택 인원 상세 수강 이력 ('{latest_week}' 기준)")

# 선택한 주차 데이터만 볼지 / 전체 주차를 다 볼지 토글
show_all_weeks = st.checkbox("과거 모든 주차의 이력도 함께 보기 (체크 시 지난 기록 펼침)", value=False)

try:
    if show_all_weeks:
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(latest_data, use_container_width=True)
except Exception as e:
    # 파이썬 데이터 타입 변환 시 발생할 수 있는 내부 오류(Arrow-compat 에러 등) 방지
    st.write("표 출력 오류 발생 시 단순 형식으로 표출합니다.")
    if show_all_weeks:
        st.table(filtered_df.astype(str))
    else:
        st.table(latest_data.astype(str))

# --- 4. PDF 다운로드 및 리포트 안내 --- #
st.sidebar.markdown("---")
st.sidebar.markdown("### 🖨️ 보고용 PDF/인쇄 출력")
st.sidebar.info("""
이 페이지를 그대로 PDF로 만들 수 있습니다!
1. 키보드의 `Ctrl + P` (맥은 `Cmd + P`) 입력
2. 인쇄 프린터를 **'PDF로 저장'** 으로 변경
3. 우측 하단의 **'배경 그래픽 표출'** 옵션을 체크☑️
4. 저장!
""")
