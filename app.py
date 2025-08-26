import streamlit as st
import pandas as pd
from dotenv import load_dotenv
import OpenDartReader
import os
from datetime import datetime
from openpyxl import load_workbook


def save_excel_with_comma_format(df: pd.DataFrame, file_name: str):
    """
    DataFrame을 엑셀로 저장한 뒤, 컬럼명에 'amount'가 포함된 열의 표시형식을
    '#,##0'으로 지정
    """
    # 1) 우선 저장
    df.to_excel(file_name, index=False)

    # 2) openpyxl로 서식 적용
    wb = load_workbook(file_name)
    ws = wb.active

    # 'amount' 포함 열(1-based index)
    amount_cols = [i + 1 for i, col in enumerate(df.columns) if "amount" in col.lower()]

    if amount_cols:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for idx in amount_cols:
                cell = row[idx - 1]
                cell.number_format = "#,##0"  # 쉼표만, 소수점 없음

    wb.save(file_name)


# ✅ API 키 불러오기: 우선순위 = 사용자 입력 > .env > secrets.toml
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

# 1. 종목코드 → 기업명 매핑
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

# 2. 전체 선택 여부 체크박스
select_all = st.checkbox("✅ 전체 선택", value=True)

# 3. 멀티 선택: 기본값은 전체 or 비워두기
selected_names = st.multiselect(
    "조회할 기업 선택",
    options=company_names,
    default=company_names if select_all else [],
    key="corp_selector",
)

# 4. 선택된 기업명 → 종목코드 변환
codes = [code for code, name in code_name_map.items() if name in selected_names]

current_year = datetime.now().year
year_range = list(range(current_year, current_year - 10, -1))
years = st.multiselect(
    "조회 연도 (복수 선택 가능)", year_range, default=[current_year - 1]
)

report = st.selectbox(
    "보고서 유형",
    options={
        "사업보고서": "11011",
        "반기보고서": "11012",
        "3분기보고서": "11014",
        "1분기보고서": "11013",
    }.items(),
)

if st.button("📥 재무제표 수집"):
    result_list = []
    for year in years:
        for code in codes:
            try:
                df = dart.finstate_all(code, bsns_year=year, reprt_code=report[1])
                if isinstance(df, pd.DataFrame) and not df.empty:
                    df["조회기업"] = code_name_map.get(code, code)
                    df["조회연도"] = year
                    result_list.append(df)
                    st.success(f"{year} - {code} 수집 완료")
                else:
                    st.warning(f"{year} - {code} 데이터 없음")
            except Exception as e:
                st.error(f"{year} - {code} 오류: {e}")

    if result_list:
        result_df = pd.concat(result_list, ignore_index=True)

        # ✅ 'amount'가 들어간 모든 컬럼을 숫자로 변환 (thstrm_add_amount 포함)
        amount_like_cols = [c for c in result_df.columns if "amount" in c.lower()]

        def _to_number(x):
            s = str(x)
            # 다양한 공백/기호 제거
            s = (
                s.replace("\u00a0", "")  # NBSP
                .replace("\ufeff", "")  # BOM
                .replace(",", "")
                .replace("₩", "")
                .replace("원", "")
                .strip()
            )
            # 괄호표기 (1,234) -> -1234
            if s.startswith("(") and s.endswith(")"):
                s = "-" + s[1:-1]
            # +기호 제거
            if s.startswith("+"):
                s = s[1:]
            if s in ("", "-"):
                return None
            return pd.to_numeric(s, errors="coerce")

        for col in amount_like_cols:
            result_df[col] = result_df[col].apply(_to_number)

        # ✅ 엑셀 저장
        file_name = f"dart_finstate_{'_'.join(map(str, years))}.xlsx"
        save_excel_with_comma_format(result_df, file_name)  # ← 쉼표 서식 적용 저장
        with open(file_name, "rb") as f:
            st.download_button(
                label="📁 엑셀 다운로드",
                data=f,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("수집된 데이터가 없습니다.")
