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
st.markdown("종목코드를 입력하면 재무제표를 가져옵니다.")

# 종목코드 → 기업명 매핑 (자주 사용하는 상장사 기준)
code_name_map = {
    "006400": "LG이노텍",
    "373220": "LG에너지솔루션",
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

# 멀티 선택 UI (기업명 기준으로 보여줌)
selected_names = st.multiselect(
    "✅ 종목 선택 (다중 선택 가능)", 
    options=list(code_name_map.values()), 
    default=["LG이노텍", "LG에너지솔루션", "엠플러스"]
)

# 기업명 → 종목코드 변환
codes = [code for code, name in code_name_map.items() if name in selected_names]

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
