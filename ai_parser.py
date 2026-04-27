import pandas as pd
from io import StringIO
from PIL import Image
import re

# ==========================================
# AIへの指示書（プロンプト）
# ==========================================
PROMPT_TEXT = """
あなたは建設業の手書き日報を正確に読み取るデータ入力のプロフェッショナルです。
提供された画像（手書きノート）を読み取り、以下の【ルール】に厳格に従って、指定された形式のテキストデータのみを出力してください。

【抽出する項目（左から順に、必ず8項目）】
月 | 日 | 工事コード | 現場名 | 業務内容 | 時間 | 休憩 | 区分

【最重要ルール：行の分割】
・同じ日、同じ工事コードの枠内に、複数の「業務内容」や「時間」が改行して書かれている場合があります。
・その場合、絶対に「1つの時間につき1行のデータ」として分割して出力してください。
（分割した際、月、日、工事コード、現場名が空欄になる場合は、そのまま空欄で出力してください。システム側で補完します）

【データ整形ルール】
・出力は必ず「|（パイプ）」区切りのテキストのみとしてください。
・月、日は数字のみ（「月」「日」という漢字は不要）にしてください。
・時間は「HH:MM」の形式に可能な限り整形してください（例: 8:00）。
・空欄の箇所は何も書かずに「|」を続けてください（例：休憩がない場合は「...|時間||区分」）。
・「〃」や「同上」といった記号も、そのまま空欄として出力してください。
・「区分」は、特に記載がなければ「日勤」とし、残業や休日出勤の記載があればそれに従ってください。

【出力フォーマット制限（厳守）】
・1行目のヘッダー（「月|日|工事コード...」）は絶対に出力しないでください。データ行から始めてください。
・Markdown記号（``` や ```csv など）は絶対に使用しないでください。
・データ行以外の説明文、挨拶、補足は一切不要です。
"""

# ==========================================
# 画像解析エンジン
# ==========================================
def process_image(file, client, target_year):
    try:
        # 1. 画像を開く
        image = Image.open(file)
        
        # 2. Gemini APIへ送信
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[image, PROMPT_TEXT]
        )

        raw_text = response.text.strip()

        # 3. Markdownの強制除去（AIがルールを破った場合のフェイルセーフ）
        match = re.search(r'`{3}(?:.*?)\n(.*?)`{3}', raw_text, re.DOTALL)
        base_text = match.group(1).strip() if match else raw_text

        # 4. 空行の除外
        clean_lines = [line for line in base_text.split('\n') if line.strip() != '']

        if not clean_lines:
            return {"file_name": file.name, "error": "文字が検出されませんでした"}

        # 5. PandasのDataFrame（表）に変換
        col_names = ['月', '日', '工事コード', '現場名', '業務内容', '時間', '休憩', '区分']
        df = pd.read_csv(
            StringIO('\n'.join(clean_lines)),
            sep='|',
            names=col_names,
            header=None,
            on_bad_lines='skip',
            engine='python' # Pandasのエラー警告を抑える設定
        )

        # 6. AIが指示を無視してヘッダーを出力した場合の除去（フェイルセーフ）
        if len(df) > 0 and str(df.iloc[0]['月']).strip() == "月":
            df = df.iloc[1:]

        return df

    except Exception as e:
        # エラーが起きた場合は、システム全体を止めずにエラーメッセージだけを返す
        return {"file_name": file.name, "error": str(e)}
