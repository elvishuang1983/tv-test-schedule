import pandas as pd
import streamlit as st
import openai
import io

# Streamlit UI 設定
st.title("📺 電視測試安排系統")

# 📌 使用者手動輸入 OpenAI API Key
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

api_key = st.text_input("🔑 請輸入 OpenAI API Key", type="password", value=st.session_state.api_key)
if api_key:
    st.session_state.api_key = api_key  # 存入 session_state

# ✅ GPT 呼叫函式（使用 GPT-4）
def call_gpt(prompt):
    if not st.session_state.api_key:
        return "⚠️ 請輸入有效的 OpenAI API Key"
    
    try:
        client = openai.OpenAI(api_key=st.session_state.api_key)
        response = client.chat.completions.create(
            model="gpt-4",  # ✅ 確保使用 GPT-4
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content

    except Exception as e:
        return f"⚠️ GPT API 錯誤: {str(e)}"

# 📤 上傳 Excel 檔案
uploaded_file = st.file_uploader("📤 上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    # ✅ 讀取 Excel
    def load_excel(file):
        xls = pd.ExcelFile(file)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, dtype=str)
        return df

    df = load_excel(uploaded_file)

    # 🔧 讓使用者選擇變更項目
    selected_items = st.multiselect("🔧 選擇變更項目", df.iloc[2:, 2].dropna().unique())
    input_value = st.text_input("🔎 或手動輸入變更內容")

    if input_value:
        matched_items = [item for item in df.iloc[2:, 2].dropna().unique() if input_value.lower() in str(item).lower()]
        selected_matched_items = st.multiselect("🎯 篩選匹配的變更項目", matched_items, default=matched_items)
        selected_items.extend(selected_matched_items)
        selected_items = list(set(selected_items))  # 移除重複

    # 📊 生成測試安排
    if st.button("📊 生成測試安排"):
        def process_test_plan(df, selected_items):
            header_department = df.iloc[0, 3:].tolist()
            header_test = df.iloc[1, 3:].tolist()
            data_rows = df.iloc[2:].reset_index(drop=True)
            matched_tests = []

            for _, row in data_rows.iterrows():
                change_item = row.iloc[2]
                if pd.isna(change_item) or change_item not in selected_items:
                    continue

                tests = {}
                for i in range(3, len(row)):
                    department = header_department[i - 3]
                    test_name = header_test[i - 3]
                    try:
                        quantity = int(float(row.iloc[i]))
                    except ValueError:
                        quantity = 0
                    if quantity > 0:
                        if department not in tests:
                            tests[department] = {}
                        tests[department][test_name] = max(tests[department].get(test_name, 0), quantity)

                matched_tests.append(tests)

            aggregated_results = {}
            for tests in matched_tests:
                for department, test_dict in tests.items():
                    if department not in aggregated_results:
                        aggregated_results[department] = {}
                    for test, qty in test_dict.items():
                        aggregated_results[department][test] = max(aggregated_results[department].get(test, 0), qty)

            output_data = []
            for department, tests in aggregated_results.items():
                for test, qty in tests.items():
                    output_data.append([department, test, qty])

            result_df = pd.DataFrame(output_data, columns=["測試部門", "測試項目", "測試數量"])
            return result_df

        result_df = process_test_plan(df, selected_items)
        st.session_state.result_df = result_df  # 存入 session_state
        st.subheader("✅ 測試安排結果 (來自 Excel)")
        st.dataframe(result_df)

        # 📥 下載 Excel 測試計劃
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name="測試安排")
        output.seek(0)

        st.download_button(
            label="📥 下載測試安排 Excel",
            data=output,
            file_name="測試安排.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ⚡ GPT 額外測試建議（針對指定變更）
    use_gpt = st.checkbox("🤖 是否讓 GPT 額外提供測試建議？")

    if use_gpt and st.session_state.api_key and st.button("⚡ GPT 測試建議"):
        with st.spinner("📊 GPT 分析中，請稍候..."):
            prompt = f"""
            產品：電視
            變更項目：{selected_items}
            
            根據上述變更，請提供額外的測試項目建議。
            對於每個建議的測試項目，請提供：
            1. 測試目標
            2. 測試方法與步驟
            3. 測試驗收標準
            """
            gpt_suggestions = call_gpt(prompt)  # ✅ GPT 產生測試建議

        # 顯示 GPT 生成的測試方法
        st.subheader("📌 GPT 額外測試建議")
        st.text_area("🔍 GPT 建議內容", gpt_suggestions, height=300)

        # 轉換為 DataFrame 並提供 Excel 下載
        gpt_results = []
        for line in gpt_suggestions.split("\n"):
            if line.strip():
                gpt_results.append({"GPT 測試建議": line.strip()})

        gpt_df = pd.DataFrame(gpt_results)

        output_gpt = io.BytesIO()
        with pd.ExcelWriter(output_gpt, engine='xlsxwriter') as writer:
            gpt_df.to_excel(writer, index=False, sheet_name="GPT測試建議")
        output_gpt.seek(0)

        st.download_button(
            label="📥 下載 GPT 測試建議 Excel",
            data=output_gpt,
            file_name="GPT測試建議.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
