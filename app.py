# app.py
import streamlit as st
import pandas as pd
import msoffcrypto
from io import BytesIO
import re

st.set_page_config(page_title="총원장 분석기", layout="wide")
st.title("📊 총원장 & 단가표 분석기")

# 엑셀 파일 업로드
col1, col2 = st.columns(2)
with col1:
    encrypted_file = st.file_uploader("🔐 총원장 파일 업로드 (.xlsx, 암호: 4698)", type=["xlsx"])
with col2:
    price_file = st.file_uploader("💰 도매 단가표 업로드 (.xlsx)", type=["xlsx"])

@st.cache_data
def decrypt_excel(uploaded_file, password):
    office_file = msoffcrypto.OfficeFile(uploaded_file)
    office_file.load_key(password=password)
    decrypted = BytesIO()
    office_file.decrypt(decrypted)
    df = pd.read_excel(decrypted)
    return df

@st.cache_data
def read_excel(uploaded_file):
    return pd.read_excel(uploaded_file)

def extract_code(product_name):
    match = re.search(r"\((.*?)\)", str(product_name))
    return match.group(1) if match else None

if encrypted_file and price_file:
    try:
        df_ledger = decrypt_excel(encrypted_file, "4698")
        df_price = read_excel(price_file)

        # 상품코드 추출
        df_ledger["코드"] = df_ledger["주문상품명"].apply(extract_code)

        # 병합
        merged = pd.merge(
            df_ledger,
            df_price[["코드", "부가포함가"]],
            on="코드",
            how="left"
        )

        display_df = merged[["구분", "주문자", "분류", "거래처", "주문상품명", "코드", "부가포함가"]]

        st.subheader("📋 병합된 데이터 보기")
        st.dataframe(display_df, use_container_width=True, height=500)

        # 검색 기능
        st.subheader("🔍 상품명 또는 코드 검색")
        keyword = st.text_input("검색어를 입력하세요 (상품명 일부 또는 코드):")

        if keyword:
            result = display_df[
                display_df["주문상품명"].str.contains(keyword, na=False, case=False) |
                display_df["코드"].str.contains(keyword, na=False, case=False)
            ]
            st.write(f"🔎 {len(result)}개의 결과가 검색되었습니다.")
            st.dataframe(result, use_container_width=True, height=400)

    except Exception as e:
        st.error(f"오류 발생: {e}")
else:
    st.info("왼쪽 상단에서 두 개의 파일을 모두 업로드해주세요.")
