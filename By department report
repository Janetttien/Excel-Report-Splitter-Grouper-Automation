# By department report

import xlwings as xw 
import os 
import shutil
import tkinter as tk
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

#確保輸出的資料夾存在
os.makedirs(output_folder, exist_ok=True)
#啟動excel app，不開GUI，Excel只會在後台進行，程式在執行時不會打擾到用戶
app=xw.App(visible=False)
wb=app.books.open(input_file)

for sheet in wb.sheets:
  output_filename = f"{sheet.name}_{year_month}.xlsx"
  output_path = os.path.join(output_folder, output_filename)
  shutil.copy(tag_blank_file, output_path)

new_wb=app.books.open(output_path)
sheet.copy(before=new_wb.sheets[0])

if len (new_wb.sheets)>1: 
	new_wb.sheets|-1].delete()
try:
  new_wb.save()
  print(f"Done: {output_filename}")

except Exception as e:
	print(f"Failed:{output_path}, error:{e}")
new_wb.close()

wb.close()
app.quit()

print("All Done")

root = tk.Tk()
root.withdraw()
messagebox.showinfo（"完成通知"，"所有檔案已成功建立"）
