
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="项目利润统计系统", layout="wide")

st.title("📊 项目利润统计系统")
st.caption("上传包含“月份、游戏、渠道、利润”的 Excel 文件，系统将自动统计 2025年1月至3月的项目利润。")

uploaded_files = st.file_uploader("📁 上传Excel文件（可多选）", type=["xlsx"], accept_multiple_files=True)

all_data = []
missing_fields = []

if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file)
            required_cols = ['月份', '游戏', '渠道', '利润']
            if all(col in df.columns for col in required_cols):
                temp = df[['月份', '游戏', '渠道', '利润']].copy()
                temp.columns = ['month', 'project_name', 'channel', 'profit']
                temp['source_file'] = file.name
                all_data.append(temp)
            else:
                missing_fields.append(file.name)
        except Exception as e:
            st.error(f"❌ 无法读取文件 {file.name}，错误：{e}")

    if missing_fields:
        st.warning("以下文件缺少必要字段，已跳过：")
        for name in missing_fields:
            st.text(f"⚠️ {name}")

    if all_data:
        df_all = pd.concat(all_data, ignore_index=True)
        df_all['month'] = df_all['month'].fillna(method='ffill')

        # 过滤仅保留 2025年1~3月数据
        df_q1 = df_all[df_all['month'].astype(str).str.contains("1月|2月|3月")]

        st.subheader("📋 原始数据预览（仅1~3月）")
        st.dataframe(df_q1, use_container_width=True)

        # 汇总统计
        st.subheader("📈 项目利润汇总（按项目 + 月份）")
        summary = df_q1.groupby(['project_name', 'month'])['profit'].sum().reset_index()
        summary = summary.sort_values(by=['project_name', 'month'])
        st.dataframe(summary, use_container_width=True)

        # 图表展示
        st.subheader("📊 利润柱状图")
        chart_data = summary.pivot(index='project_name', columns='month', values='profit').fillna(0)
        st.bar_chart(chart_data)

        # 导出功能
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary.to_excel(writer, index=False, sheet_name='利润汇总')
            df_q1.to_excel(writer, index=False, sheet_name='原始数据')
        st.download_button("📥 下载汇总Excel文件", data=output.getvalue(), file_name="项目利润汇总_2025Q1.xlsx")

    else:
        st.info("请上传包含有效字段的 Excel 文件。")
else:
    st.info("请上传文件开始统计。")
