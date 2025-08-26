import streamlit as st
import pandas as pd
import numpy as np
import re
from dotenv import load_dotenv
import OpenDartReader
import os
from datetime import datetime
from openpyxl import load_workbook

# =========================
#  문자열 정규화 유틸
# =========================
ZERO_WIDTHS = ["\u200b", "\u200c", "\u200d", "\u2060"]
THIN_SPACES = ["\u202f", "\u2009", "\u200a", "\u2007"]
OTHER_SPACES = ["\u00a0", "\ufeff"]


def _strip_all_spaces_and_currency(s: str) -> str:
    t = str(s)
    for z in ZERO_WIDTHS + THIN_SPACES + OTHER_SPACES:
        t = t.replace(z, "")
    t = t.replace(",", "").replace("₩", "").replace("원", "").strip()
    return t


def _normalize_minus_paren(s: str) -> str:
    # 유니코드 마이너스/대시 통일
    t = s.replace("\u2212", "-").replace("–", "-").replace("—", "-")
    # 괄호 음수 -> -기호
    if re.fullmatch(r"\(.*\)", t):
        t = "-" + t[1:-1].strip()
    # 앞의 + 제거
    if t.startswith("+"):
        t = t[1:]
    return t


# =========================
#  문자열 -> 숫자 변환 (DF 단계)
# =========================
def to_number_strict(x):
    if pd.isna(x):
        return np.nan
    s = _strip_all_spaces_and_currency(x)
    s = _normalize_minus_paren(s)
    # 빈/대시만 있는 값은 결측
    if s in ("", "-", "--", "+"):
        return np.nan
    # 숫자/부호/소수점 외 제거 (방어코드)
    s = re.sub(r"[^0-9\-\.\+]", "", s)
    return pd.to_numeric(s, errors="coerce")


# =========================
#  엑셀 저장 + 천단위 서식 + 마지막 강제 숫자화
# =========================
def save_excel_with_comma_format(df: pd.DataFrame, file_name: str):
    """
    1) 우선 df를 엑셀로 저장
    2) openpyxl로 모든 'amount' 열을 순회하며:
       - 셀 값이 문자열일 때 숫자처럼 보이면 실제 숫자로 강제 변환
       - number_format='#,##0' 적용
    """
    # 1) 우선 저장
    df.to_excel(file_name, index=False)

    # 2) openpyxl로 서식 및 강제 숫자화
    wb = load_workbook(file_name)
    ws = wb.active

    # 'amount' 포함 열(1-based index)
    amount_cols = [i + 1 for i, col in enumerate(df.columns) if "amount" in col.lower()]

    def _clean_to_number_like(s: str):
        if s is None:
            return None
        t = _strip_all_spaces_and_currency(s)
        t = _normalize_minus_paren(t)
        # 숫자/부호/소수점만 남기기
        t = re.sub(r"[^0-9\-\.\+]", "", t)
        if t in ("", "-", "--", "+"):
            return None
        return t

    if amount_cols:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for idx in amount_cols:
                cell = row[idx - 1]
                if isinstance(cell.value, str):
                    num_str = _clean_to_number_like(cell.value)
                    if num_str is not None:
                        try:
                            cell.value = (
                                int(num_str) if "." not in num_str else float(num_str)
                            )
                        except Exception:
                            # 변환 실패 시 그대로 둠
                            pass
                # 숫자 서식 강제
                cell.number_format = "#,##0"

    wb.save(file_name)


# =========================
#  앱 본문
# =========================
# ✅ API 키 로딩: 사용자 입력 > .env > secrets.toml
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
    "01592447": "에스케이온",  # DART 고유번호일 수도 있으니 그대로 둠
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

# 3. 멀티 선택
selected_names = st.multiselect(
    "조회할 기업 선택",
    options=company_names,
    default=company_names if select_all else [],
    key="corp_selector",
)

# 4. 선택된 기업명 → 종목코드 변환
codes = [code for code, name in code_name_map.items() if name in selected_names]

# 5. 연도 선택
current_year = datetime.now().year
year_range = list(range(current_year, current_year - 10, -1))
years = st.multiselect(
    "조회 연도 (복수 선택 가능)", year_range, default=[current_year - 1]
)

# 6. 보고서 유형
report_map = {
    "사업보고서": "11011",
    "반기보고서": "11012",
    "3분기보고서": "11014",
    "1분기보고서": "11013",
}
report_label = st.selectbox("보고서 유형", list(report_map.keys()))
report_code = report_map[report_label]

# 7. 수집 실행
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

    # ===== 합치기 =====
    result_df = pd.concat(result_list, ignore_index=True)

    # ===== 'amount'가 들어간 모든 컬럼 숫자화(DF 단계) =====
    amount_like_cols = [c for c in result_df.columns if "amount" in c.lower()]

    # (1) 문자열 정리 + 숫자 변환
    for col in amount_like_cols:
        result_df[col] = result_df[col].apply(to_number_strict)

    # (2) dtype을 확실히 float로 고정
    for col in amount_like_cols:
        result_df[col] = pd.to_numeric(result_df[col], errors="coerce").astype(
            "float64"
        )

    # ===== 엑셀 저장(마지막 방어: openpyxl에서 텍스트 숫자 강제 숫자화 + #,##0 서식) =====
    years_str = "_".join(map(str, years))
    file_name = f"dart_finstate_{years_str}.xlsx"
    save_excel_with_comma_format(result_df, file_name)

    # ===== 다운로드 버튼 =====
    with open(file_name, "rb") as f:
        st.download_button(
            label="📁 엑셀 다운로드",
            data=f,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.success("완료되었습니다!")
