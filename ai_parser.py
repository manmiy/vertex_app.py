import pandas as pd
from io import StringIO
from PIL import Image
import re
from google import genai
from google.oauth2 import service_account

PROMPT_TEXT = """
あなたは優秀なデータ入力アシスタントです。提供された画像は手書きの日報ノートです...
(中略：以前の長いプロンプト内容)
"""

def process_image(file, client, target_year):
    # 画像解析とDataFrame変換のロジック
    response = client.models.generate_content(model="gemini-2.5-flash", contents=[Image.open(file), PROMPT_TEXT])
    # ... (パース処理) ...
    return df
