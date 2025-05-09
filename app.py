import streamlit as st
import os

st.set_page_config(page_title="中乌路由测算系统", layout="wide")
st.title("📦 中乌路由测算系统")
st.markdown("点击下方按钮可一键生成全流程结果 Excel 文件。")

if st.button("🚀 开始执行全流程"):
    scripts = [
        "子文件/生成之前表的结构和属性.py",
        "子文件/生成揽收仓.py",
        "子文件/shiy11.py",
        "子文件/匹配末端配送.py",
        "子文件/合并表.py",
        "子文件/运行这个其他不管.py"
    ]
    for script in scripts:
        result = os.system(f'python "{script}"')
        if result != 0:
            st.error(f"❌ 脚本执行失败: {script}")
            st.stop()

    st.success("✅ 所有脚本执行完毕！")
    try:
        with open("中乌路由pro.xlsx", "rb") as f:
            st.download_button("📥 下载结果文件", f.read(), file_name="中乌路由pro.xlsx")
    except FileNotFoundError:
        st.error("未找到输出文件，请检查脚本是否正确执行。")

#在cmd/Powershell/Vscode终端执行 D:
#cd VsCodeProjects\中乌路由pro
#激活虚拟环境
#.\venv\Scripts\activate
#启动 Streamlit 应用
#streamlit run app.py

