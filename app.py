import streamlit as st
import os
import re
import shutil
import pandas as pd

st.set_page_config(page_title="跨境路由测算系统", layout="wide")

st.title("📦 跨境路由测算系统")
st.markdown("🛫 上传你要测算的国家路径 Excel 文件，系统将自动处理并输出分析结果。")

# 📤 文件上传
uploaded_file = st.file_uploader("📤 上传 Excel 输入文件（如：中乌.xlsx、中俄.xlsx）", type="xlsx")

if uploaded_file is not None:
    # 统一保存为英文文件名，规避中文兼容问题
    input_filename = "current_input.xlsx"
    with open(input_filename, "wb") as f:
        f.write(uploaded_file.read())
    st.success("✅ 文件上传成功，请点击下方按钮运行测算")

    # 提取上传文件的原始名称（用于结果命名展示）
    original_name = uploaded_file.name
    basename = re.sub(r"\.xlsx$", "", original_name)
    output_name = f"{basename}_路由结果.xlsx"

    if st.button("🚀 开始测算"):
        with st.spinner("🔄 正在运行测算，请稍候..."):

            # 子脚本路径
            scripts = [
                "scripts/generate_route_summary.py",
                "scripts/generate_pickup.py",
                "scripts/generate_transfer_and_line.py",
                "scripts/match_last_mile.py",
                "scripts/merge_outputs.py",
                "run_all.py"
            ]

            for script in scripts:
                result = os.system(f'python "{script}"')
                if result != 0:
                    st.error(f"❌ 脚本执行失败：{script}")
                    st.stop()

            # 支持识别默认输出文件并重命名为中文名
            final_output = "routing_analysis.xlsx"
            legacy_output = "中乌路由pro.xlsx"
            output_source = final_output if os.path.exists(final_output) else legacy_output

            if os.path.exists(output_source):
                try:
                    shutil.move(output_source, output_name)
                    st.success(f"✅ 测算完成，已生成：{output_name}")
                except PermissionError:
                    st.error("⚠️ 无法重命名输出文件，请关闭 Excel 后重试")
                    st.stop()
            else:
                st.error("⚠️ 未找到输出文件，请确认脚本是否生成结果")
                st.stop()

            # 📥 下载按钮与数据预览
            if os.path.exists(output_name):
                with open(output_name, "rb") as f:
                    st.download_button("📥 下载分析结果", f.read(), file_name=output_name)

                df = pd.read_excel(output_name)
                st.subheader("🔍 路由测算结果预览（前100行）")
                st.dataframe(df.head(100))
else:
    st.warning("请先上传一个 Excel 输入文件（如中乌.xlsx、中俄.xlsx）")


