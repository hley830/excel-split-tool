import streamlit as st

st.write("程序已启动")
import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(
    page_title="Excel拆分工具",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Excel拆分工具")

uploaded_file = st.file_uploader(
    "上传Excel文件",
    type=["xlsx"]
)

chunk_size = st.number_input(
    "每个工作表放多少行",
    min_value=1,
    value=10000,
    step=1000
)

if uploaded_file is not None:

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        df = df.dropna(how="all")

        st.success(f"读取成功，共 {len(df):,} 行数据")

        if st.button("开始拆分"):

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                sheet_count = 0

                for i in range(0, len(df), chunk_size):

                    chunk = df.iloc[i:i + chunk_size]

                    sheet_count += 1

                    chunk.to_excel(
                        writer,
                        sheet_name=f"数据{sheet_count}",
                        index=False
                    )

            output.seek(0)

            st.success(
                f"拆分完成，共生成 {sheet_count} 个工作表"
            )

            file_name = (
                os.path.splitext(uploaded_file.name)[0]
                + "_拆分结果.xlsx"
            )

            st.download_button(
                label="📥 下载拆分结果",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(str(e))
