import pandas as pd

file_path = "data/SI变更验证参考表.xlsx"
df_main = pd.read_excel(file_path, sheet_name=0, header=None)

# 顯示變更明細 (C 欄, 第 3 列開始)
print("變更明細欄位內容:")
print(df_main.iloc[2:, 2].dropna())

# 確認是否為字串類型
print("資料類型:", df_main.iloc[2:, 2].dtype)
