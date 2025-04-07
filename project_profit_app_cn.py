
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="项目利润统计系统", layout="wide")

st.title("📊 项目利润统计系统（2025年1~3月）")
st.write("上传包含“月份、项目名称、渠道、利润”字段的 Excel 表格，系统将自动统计 2025年1月至3月的项目利润。")

uploaded_files = st.file_uploader("📁 上传一个或多个 Excel 文件", accept_multiple_files=True, type=["xlsx"])

all_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)
            if all(col in df.columns for col in ['月份', '游戏', '渠道', '利润']):
                df_temp = df[['月份', '游戏', '渠道', '利润']].copy()
                df_temp = df_temp.rename(columns={
                    '月份': 'month',
                    '游戏': 'project_name',
                    '渠道': 'channel',
                    '利润': 'profit'
                })
                df_temp['source_file'] = uploaded_file.name
                all_data.append(df_temp)
            else:
                st.warning(f"⚠️ 文件 {uploaded_file.name} 缺少必要字段，已跳过。")
        except Exception as e:
            st.error(f"❌ 无法读取文件 {uploaded_file.name}，错误：{e}")

    if all_data:
        df_all = pd.concat(all_data, ignore_index=True)
        df_all['month'] = df_all['month'].fillna(method='ffill')

        # 仅保留 1~3 月数据
        df_q1 = df_all[df_all['month'].astype(str).str.contains("1月|2月|3月")]

        st.subheader("📌 原始数据预览")
        st.dataframe(df_q1)

        st.subheader("📈 项目利润汇总（按项目 + 月份）")
        df_summary = df_q1.groupby(['project_name', 'month'])['profit'].sum().reset_index()
        st.dataframe(df_summary)

        # 导出功能
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, index=False, sheet_name='汇总')
            df_q1.to_excel(writer, index=False, sheet_name='原始数据')
        st.download_button("📥 下载统计结果 Excel", data=output.getvalue(), file_name="利润汇总_2025Q1.xlsx")

    else:
        st.info("请上传包含有效字段的 Excel 文件。")
else:
    st.info("请上传 Excel 文件开始统计。")
