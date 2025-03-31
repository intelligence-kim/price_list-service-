import streamlit as st
import pandas as pd
import msoffcrypto
from io import BytesIO
import re
import warnings

warnings.simplefilter("ignore")  # openpyxl 경고 무시

st.set_page_config(page_title="총원장 분석기", layout="wide")
st.title("📊 총원장 & 단가표 분석기")

# ======================
# 엑셀 처리 함수
# ======================
def read_excel_with_password(file, password, label, sheet=None):
    file_bytes = file.getvalue()
    file_copy = BytesIO(file_bytes)

    # 비암호화 파일 먼저 시도
    try:
        if sheet:
            return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
        else:
            return pd.read_excel(BytesIO(file_bytes))
    except Exception:
        pass

    # 비밀번호가 없는데 암호화된 경우
    if not password.strip():
        raise RuntimeError(f"❗ [{label}] 파일은 암호화되어 있습니다. 비밀번호를 입력해주세요.")

    # 암호화된 파일 복호화 시도
    try:
        office_file = msoffcrypto.OfficeFile(file_copy)
        office_file.load_key(password=password.strip())
        decrypted = BytesIO()
        office_file.decrypt(decrypted)
        if sheet:
            return pd.read_excel(decrypted, sheet_name=sheet)
        else:
            return pd.read_excel(decrypted)
    except Exception:
        raise RuntimeError(f"❗ [{label}] 파일의 비밀번호가 올바르지 않거나 복호화에 실패했습니다.")

# ======================
# 코드 추출 함수
# ======================
def extract_code(product_name):
    text = str(product_name)
    matches = re.findall(r"\((.*?)\)", text)
    if not matches:
        return None
    candidate = matches[0]
    if "/" in candidate:
        candidate = candidate.split("/")[0]
    if re.match(r"^[가-힣]", candidate):
        return None
    if re.match(r"^\d+단$", candidate):
        return None
    return candidate.strip()

# ======================
# UI 입력
# ======================
col1, col2 = st.columns(2)

with col1:
    ledger_file = st.file_uploader("🔐 총원장 파일 (.xlsx)", type=["xlsx"])
    ledger_pw = st.text_input("총원장 비밀번호 (없으면 비워두세요)", type="password")

with col2:
    price_file = st.file_uploader("💰 도매 단가표 파일 (.xlsx)", type=["xlsx"])
    price_pw = st.text_input("단가표 비밀번호 (없으면 비워두세요)", type="password")

# ======================
# 분석 시작
# ======================
with st.form("upload_form"):
    submitted = st.form_submit_button("📊 분석 시작")

if submitted:
    if ledger_file and price_file:
        try:
            with st.spinner("📂 총원장 파일 불러오는 중..."):
                df_ledger = read_excel_with_password(ledger_file, ledger_pw, "총원장")  # 기본 시트

            with st.spinner("📂 단가표 파일 불러오는 중..."):
                df_price = read_excel_with_password(price_file, price_pw, "도매 단가표", sheet="2024.데코단가모음")

            # 코드 추출 + 정제
            df_ledger["코드"] = df_ledger["주문상품명"].apply(extract_code)
            df_ledger["코드"] = df_ledger["코드"].astype(str).str.strip().str.upper()
            df_price["코드"] = df_price["코드"].astype(str).str.strip().str.upper()

            # 병합
            merged = pd.merge(df_ledger, df_price[["코드", "부가포함가"]], on="코드", how="left")

            # 최종 DataFrame 정제
            display_df = merged[["구분", "주문일", "분류", "거래처", "주문상품명", "코드", "부가포함가"]].copy()
            display_df["주문일"] = pd.to_datetime(display_df["주문일"], errors="coerce").dt.strftime("%Y-%m-%d")
            display_df["분류"] = display_df["분류"].astype(str)
            display_df["거래처"] = display_df["거래처"].astype(str)
            display_df["주문상품명"] = display_df["주문상품명"].astype(str)

            st.session_state["display_df"] = display_df
            st.success("✅ 데이터 병합 완료!")

        except Exception as e:
            st.error(str(e))
    else:
        st.warning("📎 두 개의 엑셀 파일을 모두 업로드해주세요.")

# ======================
# 병합 결과 & 검색
# ======================
if "display_df" in st.session_state:
    st.subheader("📋 병합된 전체 데이터")
    st.dataframe(st.session_state["display_df"], use_container_width=True, height=500)

    st.subheader("🔍 상품명 또는 코드 검색")
    with st.form("search_form"):
        keyword = st.text_input("검색어 입력 (상품명 일부 또는 코드)", key="search_input")
        do_search = st.form_submit_button("🔍 검색")

    if do_search and keyword:
        df = st.session_state["display_df"]
        result = df[
            df["주문상품명"].str.contains(keyword, na=False, case=False, regex=False) |
            df["코드"].str.contains(keyword, na=False, case=False, regex=False)
        ]
        st.success(f"🔎 {len(result)}개의 결과가 검색되었습니다.")
        st.dataframe(result.head(500), use_container_width=True, height=400)
