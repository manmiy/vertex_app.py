import streamlit as st
import os
import pandas as pd
from io import BytesIO
from PIL import Image
from datetime import date
from google import genai
from google.oauth2 import service_account
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- 自作モジュールの読み込み ---
from excel_builder import create_filled_excel
from ai_parser import process_image

# ==========================================
# 1. 初期設定
# ==========================================
st.set_page_config(page_title="日報 (Vertex AI)", page_icon="📝", layout="wide")

if 'extracted_df' not in st.session_state:
    st.session_state.extracted_df = None

if 'is_reading' not in st.session_state:
    st.session_state.is_reading = False

def clear_extracted_data():
    st.session_state.extracted_df = None
    st.session_state.is_reading = False

# ==========================================
# 2. パスワード認証
# ==========================================
CORRECT_PASSWORD = "1234"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.subheader("ログインが必要です")
    password_input = st.text_input("パスワード", type="password")

    if st.button("ログイン", type="primary"):
        if password_input == CORRECT_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    st.stop()

# ==========================================
# 3. メインUI画面 (ログイン後)
# ==========================================
st.title("📋 日報 ")

st.markdown("### ⚙️ 出力・対象年設定")
col1, col2, col3, col4 = st.columns(4)
with col1: target_year_val = st.number_input("📅 対象年", value=date.today().year, step=1)
with col2: filename_input = st.text_input("💾 保存ファイル名", value="日報データ.xlsx")
with col3: sheet1_input = st.text_input("📄 1枚目シート", value="人工集計")
with col4: sheet2_input = st.text_input("📄 2枚目シート", value="日報明細")

uploaded_files = st.file_uploader("ノートの画像をすべて選択してください", type=["jpg", "jpeg", "png","tif","tiff"], accept_multiple_files=True, on_change=clear_extracted_data)

if uploaded_files:
    if st.session_state.extracted_df is None:
        cols = st.columns(min(3, len(uploaded_files)) if len(uploaded_files) > 0 else 1)
        for i, file in enumerate(uploaded_files):
            with cols[i % 3]: st.image(Image.open(file), caption=file.name, use_container_width=True)
        
        if st.button("✨ 画像から読み取る！", type="primary", use_container_width=True, disabled=st.session_state.is_reading):
            st.session_state.is_reading = True
            st.rerun()

    if st.session_state.is_reading:
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_dfs = []
        
        try:
            credentials_info = st.secrets["gcp_service_account"]
            creds = service_account.Credentials.from_service_account_info(credentials_info, scopes=["https://www.googleapis.com/auth/cloud-platform"])
            client = genai.Client(vertexai=True, project=credentials_info["project_id"], location="us-central1", credentials=creds)

            status_text.text(f"並列処理中... 全 {len(uploaded_files)} 枚")

            with ThreadPoolExecutor(max_workers=5) as executor:
                # 分割した ai_parser.py から呼び出し
                futures = [executor.submit(process_image, f, client, target_year_val) for f in uploaded_files]

                for idx, future in enumerate(as_completed(futures)):
                    result = future.result()
                    if isinstance(result, pd.DataFrame):
                        all_dfs.append(result)
                    elif isinstance(result, dict) and "error" in result:
                        st.error(f"⚠️ {result['file_name']} の読み取りに失敗しました。（エラー: {result['error']}）")
                    progress_bar.progress((idx + 1) / len(uploaded_files))

            if all_dfs:
                st.session_state.extracted_df = pd.concat(all_dfs, ignore_index=True)
                st.session_state.is_reading = False
                st.rerun()
            else:
                st.session_state.is_reading = False
                st.error("抽出失敗")

        except Exception as e:
            st.session_state.is_reading = False
            st.error(f"システムエラー: {e}")

    else:
        if st.session_state.extracted_df is not None:
            col1, col2 = st.columns(2)
            with col1:
                for f in uploaded_files: st.image(Image.open(f), use_container_width=True)
            with col2:
                st.subheader("📝 読み取り結果")
                edited_df = st.data_editor(st.session_state.extracted_df, num_rows="dynamic", use_container_width=True, height=850)
                if st.button("リセット", use_container_width=True): 
                    st.session_state.extracted_df = None
                    st.rerun()
            st.write("---")
            
            # 分割した excel_builder.py から呼び出し
            wb = create_filled_excel(edited_df, sheet1_input, sheet2_input, target_year_val)
            output = BytesIO()
            wb.save(output)
            st.download_button("📥 ダウンロード", output.getvalue(), file_name=filename_input, use_container_width=True, type="primary")
