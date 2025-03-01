import pandas as pd
import streamlit as st
import io

# 上傳與解析 Excel 檔案
def load_excel(file):
    xls = pd.ExcelFile(file)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, dtype=str)  # 確保所有資料讀取為字串
    return df

# 分析測試安排
def process_test_plan(df, selected_items, input_value):
    header_department = df.iloc[0, 3:].tolist()  # D1:CC1 測試部門
    header_test = df.iloc[1, 3:].tolist()  # D2:CC2 測試項目
    data_rows = df.iloc[2:].reset_index(drop=True)  # 從第3行開始為測試數據
    matched_tests = []
    matched_items = []  # 存儲匹配的完整變更項目
    
    for _, row in data_rows.iterrows():
        change_item = row.iloc[2]  # C 欄為變更項目
        if pd.isna(change_item):
            continue
        
        if change_item in selected_items or (input_value and input_value.lower() in str(change_item).lower()):
            matched_items.append(change_item)  # 記錄完整變更項目名稱
            tests = {}
            for i in range(3, len(row)):
                department = header_department[i - 3]  # 測試部門
                test_name = header_test[i - 3]  # 測試項目
                
                try:
                    quantity = int(float(row.iloc[i]))  # 嘗試轉換數值
                except ValueError:
                    quantity = 0  # 如果轉換失敗，預設為 0
                
                if quantity > 0:
                    if department not in tests:
                        tests[department] = {}
                    if test_name not in tests[department]:
                        tests[department][test_name] = 0
                    
                    tests[department][test_name] = max(tests[department][test_name], quantity)
            
            matched_tests.append(tests)
    
    # 按測試部門合併數量（取最大值）
    aggregated_results = {}
    for tests in matched_tests:
        for department, test_dict in tests.items():
            if department not in aggregated_results:
                aggregated_results[department] = {}
            for test, qty in test_dict.items():
                aggregated_results[department][test] = max(aggregated_results[department].get(test, 0), qty)
    
    # 轉換為 DataFrame
    output_data = []
    for department, tests in aggregated_results.items():
        for test, qty in tests.items():
            output_data.append([department, test, qty])
    
    result_df = pd.DataFrame(output_data, columns=["測試部門", "測試項目", "測試數量"])
    return result_df, matched_items

# Streamlit 應用程式
st.title("電視測試安排系統")
uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    df = load_excel(uploaded_file)
    selected_items = st.multiselect("選擇變更項目", df.iloc[2:, 2].dropna().unique())  # 跳過標題行
    input_value = st.text_input("或手動輸入變更內容")
    
    if st.button("生成測試安排"):
        result_df, matched_items = process_test_plan(df, selected_items, input_value)
        
        if matched_items:
            st.subheader("匹配的變更項目")
            st.write(matched_items)  # 顯示匹配的完整變更項目
        
        st.subheader("測試安排結果")
        st.dataframe(result_df)
        
        # 生成 Excel 檔案供下載
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name="測試安排")
        output.seek(0)
        
        st.download_button(
            label="下載測試安排 Excel",
            data=output,
            file_name="測試安排.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
