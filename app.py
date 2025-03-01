import pandas as pd
import streamlit as st

# 上傳與解析 Excel 檔案
def load_excel(file):
    xls = pd.ExcelFile(file)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    return df

# 分析測試安排
def process_test_plan(df, selected_items, input_value):
    header = df.columns.tolist()
    matched_tests = []
    
    for _, row in df.iterrows():
        change_item = row.iloc[2]  # 假設 C 欄是變更項目
        if pd.isna(change_item):
            continue
        
        if change_item in selected_items or (input_value and input_value in str(change_item)):
            tests = []
            for i in range(3, len(row)):
                if row.iloc[i] > 0:
                    tests.append({
                        "department": header[i - 1],
                        "test": header[i],
                        "quantity": row.iloc[i],
                    })
            matched_tests.extend(tests)
    
    # 按部門合併樣機數量（取最大值）
    aggregated_results = {}
    for test in matched_tests:
        dept, test_name, qty = test["department"], test["test"], test["quantity"]
        if dept not in aggregated_results:
            aggregated_results[dept] = {}
        aggregated_results[dept][test_name] = max(aggregated_results[dept].get(test_name, 0), qty)
    
    return aggregated_results

# Streamlit 應用程式
st.title("電視測試安排系統")
uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    df = load_excel(uploaded_file)
    selected_items = st.multiselect("選擇變更項目", df.iloc[:, 2].dropna().unique())
    input_value = st.text_input("或手動輸入變更內容")
    
    if st.button("生成測試安排"):
        test_results = process_test_plan(df, selected_items, input_value)
        
        for department, tests in test_results.items():
            st.subheader(department)
            for test, qty in tests.items():
                st.write(f"{test}: {qty} 台")
