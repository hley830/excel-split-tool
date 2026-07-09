import streamlit as st
import pandas as pd
import io
import math
import tempfile
import os

st.set_page_config(
    page_title="Excel平均拆分工具",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Excel平均拆分工具")

uploaded_file = st.file_uploader(
    "请选择Excel文件",
    type=["xlsx"]
)

split_count = st.number_input(
    "平均拆分成多少份",
    min_value=2,
    value=10,
    step=1
)

if uploaded_file is not None:

    try:

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getbuffer())
            temp_file = tmp.name

        df = pd.read_excel(temp_file, engine="openpyxl")

        df = df.dropna(how="all")

        total_rows = len(df)

        st.success(f"读取成功，共 {total_rows:,} 行")

        rows_per_sheet = math.ceil(total_rows / split_count)

        st.info(
            f"将平均拆分成 {split_count} 个工作表，每个约 {rows_per_sheet:,} 行"
        )

        if st.button("开始拆分"):

            output = io.BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:

                for i in range(split_count):

                    start = i * rows_per_sheet
                    end = min((i + 1) * rows_per_sheet, total_rows)

                    chunk = df.iloc[start:end]

                    chunk.to_excel(
                        writer,
                        sheet_name=f"第{i+1}份",
                        index=False
                    )

            output.seek(0)

            filename = (
                os.path.splitext(uploaded_file.name)[0]
                + "_平均拆分.xlsx"
            )

            st.success("拆分完成！")

            st.download_button(
                "📥 下载拆分后的Excel",
                data=output.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"发生错误：{e}")
