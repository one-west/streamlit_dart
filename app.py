import streamlit as st
import pandas as pd
import numpy as np
import re
from dotenv import load_dotenv
import OpenDartReader
import os
from datetime import datetime

# =========================
#  문자열 -> 숫자 변환 (DF 단계)
# =========================
def to_number_strict(x):
    if pd.isna(x):
        return np.nan
    s = str(x)
    # 공백/통화/쉼표 제거
    s = (s.replace("\u00a0", "")
           .replace("\ufeff", "")
           .replace(",", "")
           .replace("₩", "")
           .replace("원", "")
           .strip())
    # 괄호 음수
    if re.fullmatch(r"\(.*\)", s):
        s = "-" + s[1:-1].strip()
    # 유니코드 마이너스/대시 통일
    s = s.replace("\u2212", "-").replace("–", "-").replace("—", "-")
    # 앞의 + 제거
    if s.startswith("+"):
        s = s[1:]
    if s in ("", "-", "--", "+"):
        return np.nan
    return pd.to_numeric(s, errors="coerce")

# =========================
#  엑셀 저장 (단순 버전: xlsxwriter)
# =========================
def save_excel_with_comma_format(df: pd.DataFrame, file_name: str):
    # 1) 모든 amount 열 숫자화
    amount_cols = [c for c in df.columns if "amount" in c.lower()]
    for col in amount_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("float64")

    # 2) xlsxwriter로 저장 + 숫자 포맷 적용
    with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

        workbook  = writer.book
        worksheet = writer.sheets["Sheet1"]
        num_fmt = workbook.add_format({"num_format": "#,##0"})

        for idx, col in enumerate(df.columns):
            if "amount" in col.lower():
                worksheet.set_column(idx, idx, 18, num_fmt)

# =========================
#  앱 본문
# =========================
load_dotenv()
api_key = os.getenv("DART_API_KEY") or st.secrets.get("DART_API_KEY", None)
if not api_key:
    api_key = st.sidebar.text_input("API 키를 입력하세요", type="password")
if not api_key:
    st.warning("API 키가 필요합니다. 입력 후 다시 시도하세요.")
    st.stop()

dart = OpenDartReader(api_key)

st.title("📊 DART 재무제표 수집기")
st.markdown("종목코드를 입력하면 재무제표를 가져옵니다.")

code_name_map = {
    "006400": "삼성SDI",
    "373220": "LG에너지솔루션",
    "01592447": "에스케이온",
    "259630": "엠플러스",
    "137400": "피엔티",
    "222080": "씨아이에스",
    "267320": "나인테크",
    "196490": "디에이테크놀로지",
    "109740": "디에스케이",
    "299030": "하나기술",
    "240600": "유진테크놀로지",
    "148930": "에이치와이티씨",
}
company_names = list(code_name_map.values())

select_all = st.checkbox("✅ 전체 선택", value=True)
selected_names = st.multiselect(
    "조회할 기업 선택",
    options=company_names,
    default=company_names if select_all else [],
    key="corp_selector",
)
codes = [code for code, name in code_name_map.items() if name in selected_names]

current_year = datetime.now().year
year_range = list(range(current_year, current_year - 10, -1))
years = st.multiselect("조회 연도 (복수 선택 가능)", year_range, default=[current_year - 1])

report_map = {
    "사업보고서": "11011",
    "반기보고서": "11012",
    "3분기보고서": "11014",
    "1분기보고서": "11013",
}
report_label = st.selectbox("보고서 유형", list(report_map.keys()))
report_code = report_map[report_label]

if st.button("📥 재무제표 수집"):
    if not codes:
        st.info("선택된 기업이 없습니다.")
        st.stop()
    if not years:
        st.info("선택된 연도가 없습니다.")
        st.stop()

    result_list = []
    for year in years:
        for code in codes:
            try:
                df = dart.finstate_all(code, bsns_year=year, reprt_code=report_code)
                if isinstance(df, pd.DataFrame) and not df.empty:
                    df["조회기업"] = code_name_map.get(code, code)
                    df["조회연도"] = year
                    result_list.append(df)
                    st.success(f"{year} - {code} 수집 완료")
                else:
                    st.warning(f"{year} - {code} 데이터 없음")
            except Exception as e:
                st.error(f"{year} - {code} 오류: {e}")

    if not result_list:
        st.info("수집된 데이터가 없습니다.")
        st.stop()

    result_df = pd.concat(result_list, ignore_index=True)

    # 모든 amount 열 숫자 변환
    amount_like_cols = [c for c in result_df.columns if "amount" in c.lower()]
    for col in amount_like_cols:
        result_df[col] = result_df[col].apply(to_number_strict)
        result_df[col] = pd.to_numeric(result_df[col], errors="coerce").astype("float64")

    file_name = f"dart_finstate_{'_'.join(map(str, years))}.xlsx"
    save_excel_with_comma_format(result_df, file_name)

    with open(file_name, "rb") as f:
        st.download_button(
            label="📁 엑셀 다운로드",
            data=f,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.success("완료되었습니다!")
