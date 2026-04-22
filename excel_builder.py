import pandas as pd
import re
from datetime import timedelta, date
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
import jpholiday

def is_holiday_jp(d):
    return d.weekday() >= 5 or jpholiday.is_holiday(d)

def create_filled_excel(df_extracted, sheet1_name="人工集計", sheet2_name="日報明細", target_year=2024):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet1_name
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # (中略：以前提供された create_filled_excel の中身をここに配置)
    # ※ 依存関係を整理した完全なロジックをここに集約します
    return wb
