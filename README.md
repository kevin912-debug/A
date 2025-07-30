import pandas as pd
import numpy as np
from datetime import datetime
import os
import glob
import traceback
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ========== 第一步：補貨分析 ==========
file_path = '新增 Microsoft Excel 工作表.xlsx'
if not os.path.exists(file_path):
    print("❌ 找不到你的檔案.xlsx，請確認資料夾內有正確檔案")
    exit()

df = pd.read_excel(file_path)
df['amount_in_transit'] = df['amount_in_transit'].fillna(0)
df['目標數量'] = df['7天銷售量'] * 6

df['缺色是否缺貨'] = np.where(df['運送中'] == 1, 0, np.where(df['0的比例'] >= 0.1, 1, 0))
df['足夠銷售是否缺貨'] = np.where((df['庫存量'] + df['運送中']) - (df['7天銷售量'] * 6) <= 0, 1, 0)
df['基本量是否缺貨'] = np.where((df['庫存量'] + df['運送中']) - (df['Size'] * 2) <= 0, 1, 0)
df['是否需要補貨'] = np.where((df[['缺色是否缺貨', '足夠銷售是否缺貨', '基本量是否缺貨']].sum(axis=1)) >= 1, '需要補貨', '不需要補貨')

def assign_rank_groups(nonzero_df, main_col):
    nonzero_df = nonzero_df.sort_values(by=[main_col, '7天銷售量', '30天月銷量'], ascending=[False, False, False]).reset_index()
    nonzero_df['rank'] = nonzero_df.index + 1
    total = len(nonzero_df)
    bins = []
    for i in nonzero_df['rank']:
        p = i / total
        if p <= 0.1:
            bins.append('D10')
        elif p <= 0.2:
            bins.append('D9')
        elif p <= 0.3:
            bins.append('D8')
        elif p <= 0.4:
            bins.append('D7')
        elif p <= 0.5:
            bins.append('D6')
        elif p <= 0.6:
            bins.append('D5')
        elif p <= 0.7:
            bins.append('D4')
        elif p <= 0.8:
            bins.append('D3')
        elif p <= 0.9:
            bins.append('D2')
        else:
            bins.append('D1')
    nonzero_df['group'] = bins
    return nonzero_df.set_index('index')['group']

df['訪客數分位'] = '怪異'
df['轉換率分位'] = '怪異'
visitor_groups = assign_rank_groups(df[df['商品訪客數'] > 0].copy(), '商品訪客數')
conversion_groups = assign_rank_groups(df[df['轉換率(可出貨訂單)'] > 0].copy(), '轉換率(可出貨訂單)')
df.loc[visitor_groups.index, '訪客數分位'] = visitor_groups
df.loc[conversion_groups.index, '轉換率分位'] = conversion_groups

def cross_analysis(row):
    v, c = row['訪客數分位'], row['轉換率分位']
    if v in ['D8','D9','D10'] and c in ['D8','D9','D10']:
        return '🔥 熱賣潛力'
    elif v in ['D8','D9','D10'] and c in ['D1','D2','D3']:
        return '⚠️ 流量浪費'
    elif v in ['D1','D2','D3'] and c in ['D8','D9','D10']:
        return '🎯 精準小品'
    elif v in ['D1','D2','D3'] and c in ['D1','D2','D3']:
        return '❌ 待調整'
    else:
        return '可觀察'

df['流量轉換交叉分析'] = df.apply(cross_analysis, axis=1)

# ========== 第二步：清理補貨名單 ==========
drop_cols = [
    "Unnamed: 0", "暴衝值", "num_of_zero", "搜尋點擊", "買家(可出貨訂單)",
    "點擊數", "自然流量", "防禦款", "中小男童", "中小女童", "中大男童",
    "大類名稱", "中類名稱", "小類名稱"
]
df = df.drop(columns=[c for c in drop_cols if c in df.columns])
if "可捕獲" in df.columns:
    df = df[df["可捕獲"] != 0]
if "運送中" in df.columns:
    df = df[df["運送中"] != 1]
if "是否需要補貨" in df.columns:
    df = df[df["是否需要補貨"] != "不需要補貨"]

today = datetime.today().strftime("%Y%m%d")
clean_file = f"{today}_補貨名單.xlsx"
df.to_excel(clean_file, index=False)
print(f"✅ 已輸出補貨名單：{clean_file}")

# ========== 第三步：補貨明細 ==========
def normalize_color(s):
    if pd.isna(s): return ""
    return str(s).replace("（", "(").replace("）", ")").replace("，", ",").replace(" ", "")

try:
    sales_files = glob.glob("export_report.parentskudetail.*.xlsx")
    if len(sales_files) != 1:
        raise FileNotFoundError("找不到唯一的銷售資料（export_report.parentskudetail...xlsx）")
    sales_file = sales_files[0]

    inventory_files = [f"{i}.xlsx" for i in range(1, 7) if os.path.exists(f"{i}.xlsx")]
    if not inventory_files:
        raise FileNotFoundError("找不到任何庫存資料（1.xlsx～6.xlsx）")

    sales_df = pd.read_excel(sales_file)
    replenishment_df = pd.read_excel(clean_file)
    replenishment_df = replenishment_df.dropna(subset=["目標數量"]).rename(columns={
        "et_title_product_id": "商品ID",
        "目標數量": "補貨目標數量"
    })

    inventory_list = []
    for file in inventory_files:
        df_inv = pd.read_excel(file)
        df_inv = df_inv[df_inv["et_title_product_id"].isin(replenishment_df["商品ID"])]
        df_inv[["花色_raw", "尺碼_raw"]] = df_inv["et_title_variation_name"].str.split(",", n=1, expand=True)
        df_inv["花色"] = df_inv["花色_raw"].fillna("").apply(normalize_color)
        df_inv["尺碼"] = df_inv["尺碼_raw"].fillna("").str.replace("碼", "").str.strip()
        df_inv["庫存數"] = pd.to_numeric(df_inv["et_title_variation_stock"], errors="coerce")
        df_inv = df_inv.rename(columns={"et_title_product_id": "商品ID"})
        inventory_list.append(df_inv[["商品ID", "花色", "尺碼", "庫存數"]])
    inventory_all = pd.concat(inventory_list, ignore_index=True)

    color1 = PatternFill(start_color="E5F0FF", end_color="E5F0FF", fill_type="solid")
    color2 = PatternFill(start_color="FFFCE5", end_color="FFFCE5", fill_type="solid")
    wb = Workbook()
    wb.remove(wb.active)

    for _, row in replenishment_df.iterrows():
        product_id = row["商品ID"]
        target_amount = row["補貨目標數量"]
        subset = sales_df[sales_df["商品ID"] == product_id][["商品規格", "商品件數(可出貨訂單)"]].copy()
        if subset.empty:
            continue
        subset["商品件數(可出貨訂單)"] = pd.to_numeric(subset["商品件數(可出貨訂單)"], errors="coerce")
        subset["花色"] = subset["商品規格"].str.extract(r"^(.+?)\s*,\s*\d")[0].apply(normalize_color)
        subset["尺碼"] = subset["商品規格"].str.extract(r",\s*(\d+)\s*碼")[0].astype(str).str.strip()
        subset = subset.dropna(subset=["花色", "尺碼"])
        subset = subset[subset["尺碼"].str.match(r"^\d+$")]

        color_ratio = subset.groupby("花色")["商品件數(可出貨訂單)"].sum().reset_index(name="花色銷售數")
        color_ratio["花色比例"] = color_ratio["花色銷售數"] / color_ratio["花色銷售數"].sum()

        size_ratio = subset.groupby("尺碼")["商品件數(可出貨訂單)"].sum().reset_index(name="尺碼銷售數")
        size_ratio["尺碼比例"] = size_ratio["尺碼銷售數"] / size_ratio["尺碼銷售數"].sum()

        cross = color_ratio[["花色", "花色比例"]].merge(size_ratio[["尺碼", "尺碼比例"]], how="cross")
        cross["補貨目標總量"] = target_amount
        inventory_sub = inventory_all[inventory_all["商品ID"] == product_id]
        merged = pd.merge(cross, inventory_sub, on=["花色", "尺碼"], how="left")
        merged["庫存數"] = merged["庫存數"].fillna(0)

        ws = wb.create_sheet(title=str(product_id)[:31])
        headers = ["花色", "尺碼", "花色比例", "尺碼比例", "補貨目標總量", "預估補貨", "庫存數", "實際補貨", "優化後實際補貨"]
        ws.append(headers)

        for i, r in merged.iterrows():
            excel_row = i + 2
            ws.append([r["花色"], r["尺碼"], r["花色比例"], r["尺碼比例"], target_amount, None, r["庫存數"], None])
            ws[f"F{excel_row}"] = f"=ROUND(E{excel_row}*C{excel_row}*D{excel_row},0)"
            ws[f"H{excel_row}"] = f"=MAX(F{excel_row}-G{excel_row},0)"
            ws[f"I{excel_row}"] = f"=IF(H{excel_row}<2,IF(G{excel_row}=0,2,H{excel_row}),H{excel_row})"

        for i in range(2, ws.max_row + 1):
            ws[f"C{i}"].number_format = '0.00%'
            ws[f"D{i}"].number_format = '0.00%'

        max_col_width = 0
        last_color, toggle, fill = None, True, color1
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            current_color = row[0].value
            if current_color != last_color:
                toggle = not toggle
                fill = color1 if toggle else color2
                last_color = current_color
            for cell in row:
                if cell.value is not None:
                    max_col_width = max(max_col_width, len(str(cell.value)))
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_col_width + 6

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:I{ws.max_row}"

    detail_path = f"{today}_補貨明細.xlsx"
    wb.save(detail_path)
    print(f"✅ 已完成補貨明細：{detail_path}")

except Exception as e:
    print("❌ 錯誤：")
    traceback.print_exc()
