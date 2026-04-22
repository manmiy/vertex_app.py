import streamlit as st
import os
from io import BytesIO
from google import genai
from google.oauth2 import service_account

# 自作モジュールの読み込み
from excel_builder import create_filled_excel
from ai_parser import process_image

# パスワード・UI設定・メインロジックをここに記述
# ユーザーが操作する画面はここだけです
