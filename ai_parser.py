import pandas as pd
from io import StringIO
from PIL import Image
import re
from datetime import date
import jpholiday

# ==========================================
# AIへの指示書（あなたの考案した最適化プロンプト）
# ==========================================
PROMPT_TEXT = """
あなたは優秀なデータ入力アシスタントです。提供された画像は手書きの日報ノートです。
画像内の表は左から順に以下の列構成になっています：
1. 日付 (月/日)
2. 工事コード (1397-00 など)
3. 時間 (10:00 など)
4. 人工 (0.5, 2 など - これは抽出しますがExcel出力用には使いません)
5. 業務内容 (ML-750打設準備 など)

【抽出ルール】
- 「月」「日」「工事コード」「現場名」「時間」「業務内容」の6項目を抽出してください。
- 画像の最上部にある「〇年〇月分」などのタイトル表記や、表の枠外にあるメモ書きは絶対にデータとして抽出しないでください。罫線で囲まれたメインの表の中身のみを対象とします。
- ⚠️ 業務内容に該当する欄に、作業内容と一緒に「現場名（場所の名前など）」が書き込まれている場合があります。現場名は「現場名」として独立させ、残った純粋な作業内容を「業務内容」として別々に抽出してください。現場名がない場合は空欄としてください。
- ⚠️ 工事コードの「〃」といった同上を示す記号が書かれている場合は、記号をそのまま出力せず、直上の行から同じ内容を補完して出力してください。
- ⚠️ 工事コードが完全に空欄の場合も、直前の行と同じ工事コードを補完して出力してください。
- 時間に「-」や「〜」などの記号が含まれている場合、それらを除外して数字とコロンのみ（例: 10:30）にしてください。
- 業務内容や現場名に含まれる改行はスペースに置き換えてください。
- 出力は必ず「|」（パイプ記号）で区切ったテキスト形式（挨拶不要）でお願いします。

⚠️【最重要ルール：行の分割】
ノート上で「同じ日付・同じ工事コード」の枠内に、時間が複数行に分かれて書かれている場合（例：15:00、16:00...）、絶対にカンマ等で1つのセルにまとめないでください。
必ず「月」「日」「工事コード」「現場名」を上の行から補って、1つの時間につき1行（独立したデータ）として出力してください。

■ 出力フォーマット
1行目: 月|日|工事コード|現場名|時間|業務内容
2行目以降: | 区切りのデータ（計6列）
"""

# 祝日判定ツール（パーサー用）
def is_holiday_jp_parser(d):
    return d.weekday() >= 5 or jpholiday.is_holiday(d)

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

        # 3. Markdownの強制除去
        match = re.search(r'`{3}(?:.*?)\n(.*?)`{3}', raw_text, re.DOTALL)
        base_text = match.group(1).strip() if match else raw_text

        # 4. データ行のクリーニング（あなたの考案したMarkdownテーブル線除去ロジック）
        clean_lines = []
        for line in base_text.split('\n'):
            line = line.strip()
            # 空行や、Markdownのテーブル区切り線（|---|---など）を除外
            if not line or re.match(r'^\|?[\s\-]+\|?[\s\-\|]*$', line): 
                continue
            # 先頭と末尾のパイプ記号を除去（Pandasで余分な空列ができるのを防ぐため）
            if line.startswith('|'): line = line[1:]
            if line.endswith('|'): line = line[:-1]
            clean_lines.append(line.strip())

        if not clean_lines:
            return {"file_name": file.name, "error": "文字が検出されませんでした", "raw_text": raw_text}

        # 5. PandasのDataFrame（表）に変換 (6列)
        col_names = ["月", "日", "工事コード", "現場名", "時間", "業務内容"]
        df = pd.read_csv(
            StringIO('\n'.join(clean_lines)),
            sep='|',
            names=col_names,
            header=None,
            on_bad_lines='skip',
            engine='python'
        )

        # 6. ヘッダー行の除外
        if not df.empty and str(df.iloc[0]['月']).strip() == "月":
            df = df.iloc[1:].reset_index(drop=True)

        # 7. 必須項目の欠損行を除去
        df = df.dropna(subset=['月', '日', '業務内容'])

        # 8. データの最終整形とカレンダー判定
        if not df.empty:
            if '工事コード' in df.columns:
                df['工事コード'] = df['工事コード'].astype(str).str.replace(r'[^a-zA-Z0-9-]', '', regex=True)
            if '時間' in df.columns:
                df['時間'] = df['時間'].astype(str).str.replace(r'[^0-9:]', '', regex=True)
            
            # 👇 あなたの素晴らしいロジック：Pythonによる確実な休出判定
            def get_cat(r):
                try:
                    m, d = int(r['月']), int(r['日'])
                    return "休日/残業" if is_holiday_jp_parser(date(target_year, m, d)) else "日勤"
                except: 
                    return "日勤"
            
            # UIとExcelビルダーに渡すための互換性確保（休憩と区分を追加して8列にする）
            df['休憩'] = None
            df['区分'] = df.apply(get_cat, axis=1)
            
        return df

    except Exception as e:
        # エラー発生時はディクショナリで返す
        return {"file_name": file.name, "error": str(e), "raw_text": raw_text if 'raw_text' in locals() else ""}
