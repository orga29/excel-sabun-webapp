import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="温市Excel差分ツール", page_icon="🌿", layout="centered")
st.markdown("""
    <style>
        body {
            background-color: #e6f4e6;
        }
        .main {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

st.image("スクリーンショット 2025-04-24 054018.png", width=200)
st.title("温市 Excel差分比較ツール")

st.markdown("""
#### 📝 使い方：
1. 「暫定データファイル」と「確定データファイル」を選んでください。
2. 差分を自動で抽出して表示します。
3. Excelファイルとしてダウンロードも可能です。
""")

def build_map(df: pd.DataFrame) -> dict:
    m = {}
    for _, row in df.iterrows():
        code = row[0]
        if pd.isna(code):
            continue
        code = str(code).strip()
        name = row[1] if not pd.isna(row[1]) else ''
        qty = int(row[4]) if not pd.isna(row[4]) else 0
        after_sort = int(row[7]) if not pd.isna(row[7]) else 0
        m[code] = {'name': name, 'qty': qty, 'after_sort': after_sort}
    return m

def compute_diff(map1: dict, map2: dict) -> pd.DataFrame:
    rows = []
    for code, rec1 in map1.items():
        qty1 = rec1['qty']
        after1 = rec1['after_sort']
        if code in map2:
            rec2 = map2[code]
            qty2 = rec2.get('qty', 0)
            after = rec2.get('after_sort', after1)
        else:
            qty2 = 0
            after = after1
        diff = qty2 - qty1
        if diff != 0:
            rows.append({
                '商品コード': code,
                '商品名': rec1['name'],
                '暫定': qty1,
                '確定': qty2,
                '仕分後残': after,
                '増減数': f'+{diff}' if diff > 0 else str(diff)
            })
    for code, rec2 in map2.items():
        if code not in map1:
            qty2 = rec2['qty']
            after = rec2['after_sort']
            rows.append({
                '商品コード': code,
                '商品名': rec2['name'],
                '暫定': 0,
                '確定': qty2,
                '仕分後残': after,
                '増減数': f'+{qty2}'
            })
    df = pd.DataFrame(rows)
    df = df[~df['商品名'].str.startswith('■') & ~df['商品名'].str.endswith('◇')]
    df['__sort'] = df['商品名'].str[0].apply(lambda ch: 0 if ch == '■' else 1 if ch in ('□', '▢') else 2)
    df = df.sort_values(['__sort', '商品コード']).drop(columns='__sort')
    return df

def to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = '差分結果'

    # 書式設定用のスタイル
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4CAF50")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    headers = list(df.columns)
    ws.append(headers)

    for col_num, col in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = 14

    for row in df.itertuples(index=False):
        ws.append(row)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=ws.max_row, column=col_num)
            cell.border = border
            if isinstance(cell.value, int):
                cell.alignment = Alignment(horizontal='right')
            else:
                cell.alignment = Alignment(horizontal='left')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

file1 = st.file_uploader("暫定データファイル (.xlsx)", type="xlsx")
file2 = st.file_uploader("確定データファイル (.xlsx)", type="xlsx")

if file1 and file2:
    df1 = pd.read_excel(file1, header=None).iloc[4:].reset_index(drop=True)
    df2 = pd.read_excel(file2, header=None).iloc[4:].reset_index(drop=True)
    diff_df = compute_diff(build_map(df1), build_map(df2))

    st.success("差分を抽出しました！")
    st.dataframe(diff_df, use_container_width=True)

    excel_data = to_excel(diff_df)
    st.download_button("📥 差分をExcelでダウンロード", excel_data, file_name=f"差分_{datetime.date.today()}.xlsx")
