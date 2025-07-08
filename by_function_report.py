# By function report

import xlwings as xw 
import os 
import shutil
import tkinter as tk
import pandas as pd
from tkinter import filedialog, Tk, simpledialog, messagebox

# 建立 GUI 窗口
root = Tk()
root.withdraw()

＃要輸入年&月
year_month = simpledialog.askstring（"輸人年月"，"請輸人報表年月（例如2025005）"）
#選擇要進行分組/分割的總檔 
input_fle = filedialog.askopenfilename（title="選擇要進行分割的總檔 檔案"）
＃選擇分組/分割檔要移入的已選好&儲存標籤空檔
tag_blank_file = filedialog.askopenfilename（title="選擇空白模板（含標籤）檔案"）
＃選擇分組/分割檔要儲存的資料夾路徑
output_folder = fledialog.askdirectory（title="選擇 輸出資料夾"）
#選擇要mapping的表
mapping_file = = filedialog.askopenfilename(title="請選擇 Mapping 檔案（含 function 和 Sheet name）")

#確保輸出的資料夾存在
os.makedirs(output_folder, exist_ok=True)

#讀取mapping表
mapping_df = pd.read_excel(mapping_file)

#開啟總檔，啟動excel app，不開GUI，Excel只會在後台進行，程式在執行時不會打擾到用戶
app=xw.App(visible=False)
wb=app.books.open(input_file)

＃遍歷每一個唯一的function群組，每次只處理一個群組
for target_file in mapping_df[‘Function’].unique():
  output _filename = f"{target_file}_{year_month}.xIsx"
  output_path = os.path.join(output_folder, output_filename)

if not os.path.exists(output_path):
  shutil.copy(tag_blank_file, output_path)

# 開啟目標檔案
target_wb = app.books.open(output_path)

＃取出對應要加入的工作表名稱
selected_sheets = mapping_df[mapping_df[‘Function’] == target_file][‘Sheet name']
all_sheet_names = [s.name for s in wb.sheets]

for sheet_name in selected _sheets: 
	if sheet_name in all _sheet_names:
source_sheet = wb.sheets[sheet_name]
	#用Com保留格式
source_sheet.api.Copy(After = target_wb.sheets[-1].api)
print(f"{sheet_name} in {target_file)_{year_month}.xlsx")
else:
	print(f" no {sheet_name}")

if len(target_wb.sheets)>1: 
	target_wb.sheets[0].delete()

target_wb.save()
target_ wb.close()
print(f" Done: {target _file}")

wb.close()
app.quit()

print("All Done")
root = tk.Tk()
root.withdraw()
messagebox.showinfo（"完成通知”，"所有檔案已成功建立"）
