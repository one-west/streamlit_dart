import streamlit as st
import pandas as pd
from dotenv import load_dotenv
import OpenDartReader
import os

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
st.markdown("종목코드를 입력하면 2023년 사업보고서 기준 재무제표를 가져옵니다.")

codes_input = st.text_area("종목코드를 쉼표(,)로 구분하여 입력", "006400,373220,259630")
codes = [c.strip() for c in codes_input.split(",") if c.strip()]

year = st.selectbox("조회 연도", options=[2024, 2023, 2022, 2021], index=1)
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
    for code in codes:
        try:
            df = dart.finstate(code, bsns_year=year, reprt_code=report[1])
            if isinstance(df, pd.DataFrame) and not df.empty:
                df["조회코드"] = code
                result_list.append(df)
                st.success(f"{code} - 수집 완료")
            else:
                st.warning(f"{code} - 데이터 없음")
        except Exception as e:
            st.error(f"{code} - 오류: {e}")

    if result_list:
        result_df = pd.concat(result_list, ignore_index=True)
        st.dataframe(result_df)

        # 엑셀로 다운로드
        file_name = f"dart_finstate_{year}.xlsx"
        result_df.to_excel(file_name, index=False)
        with open(file_name, "rb") as f:
            st.download_button(
                label="📁 엑셀 다운로드",
                data=f,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("수집된 데이터가 없습니다.")
