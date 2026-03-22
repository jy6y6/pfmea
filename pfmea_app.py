import streamlit as st
import pandas as pd
import requests
import json
import io
import os
import re
import time
from PIL import Image as PILImage
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import base64
import hashlib
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

# =================配置与初始化=================
# 设置页面配置
st.set_page_config(
    page_title="智能工具箱 - PFMEA生成与图片处理",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS：淡绿色系主题
st.markdown("""
<style>
    /* 主题颜色：淡雅青瓷绿 */
    :root {
        --primary-color: #81c784; /* 主色调 */
        --bg-color: #f1f8e9;     /* 背景 */
        --card-bg: #ffffff;      /* 卡片背景 */
        --text-color: #2e7d32;   /* 深绿文字 */
        --border-color: #c8e6c9; /* 边框 */
    }

    /* 全局样式 */
    .stApp {
        background-color: var(--bg-color);
        color: var(--text-color);
    }

    /* 侧边栏 */
    .stSidebar {
        background-color: #e8f5e9 !important;
    }

    /* 按钮样式 */
    .stButton>button {
        background-color: var(--primary-color) !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #66bb6a !important;
        box-shadow: 0 4px 12px rgba(129, 199, 132, 0.4);
    }

    /* 卡片容器 */
    .card {
        background-color: var(--card-bg);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.05);
        border: 1px solid var(--border-color);
        margin-bottom: 20px;
    }

    /* 标题 */
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: var(--text-color);
        font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
    }

    /* 拖拽区域样式 */
    .drag-item {
        background-color: #e8f5e9;
        border: 2px dashed #81c784;
        border-radius: 8px;
        padding: 10px;
        margin: 5px 0;
        cursor: grab;
        transition: all 0.2s;
    }
    .drag-item:hover {
        background-color: #dcedc8;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# 会话状态初始化
if 'image_order' not in st.session_state:
    st.session_state.image_order = []
if 'pfmea_result' not in st.session_state:
    st.session_state.pfmea_result = ""
if 'api_key' not in st.session_state:
    st.session_state.api_key = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44" # 默认豆包API

# =================模块一：图片拖拽排序与Excel插入=================
def module_image_sorter():
    st.subheader("🖼️ 模块一：图片拖拽排序与Excel处理")

    # 1. 图片上传
    uploaded_files = st.file_uploader("上传多张图片 (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="img_upload")

    if uploaded_files:
        # 初始化顺序列表
        if len(st.session_state.image_order) != len(uploaded_files):
            st.session_state.image_order = list(range(len(uploaded_files)))

        # 显示实时预览与拖拽排序
        st.markdown("### 🔢 拖拽调整图片顺序 (长按拖动)")
        # 使用 Streamlit columns 模拟拖拽列表
        cols = st.columns(len(uploaded_files))
        new_order = st.session_state.image_order.copy()

        for idx in st.session_state.image_order:
            with cols[idx]:
                try:
                    img = PILImage.open(uploaded_files[idx])
                    st.image(img, caption=f"图片 {idx+1}", use_column_width=True)
                    # 模拟拖拽交互（Streamlit 本身不支持原生拖拽排序，这是最接近的UI体验）
                    target_pos = st.selectbox(f"位置调整 {idx+1}", 
                        options=list(range(1, len(uploaded_files)+1)), 
                        index=idx,
                        key=f"pos_{idx}")
                    new_order[idx] = target_pos - 1 # 转换为索引
                except Exception as e:
                    st.error(f"图片读取错误: {e}")

        if st.button("🔄 更新排序"):
            st.session_state.image_order = new_order
            st.success("顺序已更新！")

        # 2. Excel 文件上传与处理
        st.markdown("---")
        excel_file = st.file_uploader("上传Excel模板", type=["xlsx"], key="excel_upload")

        if excel_file and st.button("🚀 开始处理图片并插入Excel"):
            with st.spinner("正在处理图片并生成Excel..."):
                try:
                    # 读取Excel
                    wb = load_workbook(excel_file)
                    ws = wb.active

                    # 创建临时文件夹保存图片
                    if not os.path.exists("temp_imgs"):
                        os.makedirs("temp_imgs")

                    # 按新顺序保存图片
                    ordered_images = []
                    for pos in st.session_state.image_order:
                        img_data = uploaded_files[pos]
                        img_path = f"temp_imgs/{img_data.name}"
                        with open(img_path, "wb") as f:
                            f.write(img_data.getbuffer())
                        ordered_images.append(img_path)

                    # 插入图片到Excel (此处仅为示例逻辑，具体位置需根据你的模板调整)
                    for i, img_path in enumerate(ordered_images):
                        img = XLImage(img_path)
                        # 假设插入到 A 列，每张图片占 20 行
                        img_cell = f"A{5 + i*20}"
                        img.width = 200
                        img.height = 150
                        ws.add_image(img, img_cell)

                    # 保存
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.download_button(
                        label="✅ 下载处理后的Excel",
                        data=output,
                        file_name="处理后的图片Excel.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Excel处理失败: {e}")

# =================模块三：PFMEA生成器 (修复版)=================
def module_pfmea_generator():
    st.subheader("🤖 模块三：PFMEA 智能生成器")

    # API Key 输入
    api_key = st.text_input("API Key", value=st.session_state.api_key, type="password")
    st.session_state.api_key = api_key

    # 问题输入
    prompt = st.text_area(
        "输入产品或工艺描述 (例如：汽车刹车盘生产流程)",
        height=150,
        placeholder="请详细描述你需要生成PFMEA的对象..."
    )

    # 优化的 AI 调用函数
    def generate_pfmea_with_retries(product_desc, api_key):
        url = "https://api.doubao.com/v1/chat/completions" # 豆包官方API域名
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        # 构建更严谨的请求体
        data = {
            "model": "doubao-pro-32k", # 你可以根据需要更换模型
            "messages": [
                {"role": "system", "content": "你是一位资深的质量工程师，精通AIAG-VDA PFMEA标准。请用Markdown表格格式输出结果。"},
                {"role": "user", "content": f"请根据以下描述生成一份详细的PFMEA分析表：{product_desc}\n\n输出要求：包含过程步骤、功能要求、潜在失效模式、失效后果、严重度(S)、失效原因、频度(O)、现行控制、探测度(D)、AP等列。"}
            ],
            "temperature": 0.3,
            "max_tokens": 2000
        }

        # 使用 Session 和 Retry 机制防止超时
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
        session.mount('https://', HTTPAdapter(max_retries=retries))

        try:
            response = session.post(url, headers=headers, json=data, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                # 提取AI回复内容
                ai_reply = result.get("choices", [{}])[0].get("message", {}).get("content", "")
                # 尝试清洗 JSON 或 Markdown 格式
                cleaned_reply = re.sub(r"```markdown|```", "", ai_reply).strip()
                return cleaned_reply
            else:
                return f"API Error: {response.status_code}\n{response.text}"

        except requests.exceptions.RequestException as e:
            return f"网络连接错误: {str(e)}。请检查网络或稍后再试。"

    if st.button("🚀 生成 PFMEA"):
        if not prompt.strip():
            st.warning("请输入产品描述！")
        else:
            with st.spinner("AI 正在深度思考并生成报告，请稍候..."):
                # 清除旧结果
                st.session_state.pfmea_result = ""
                
                # 调用修复后的函数
                result = generate_pfmea_with_retries(prompt, api_key)
                
                if "错误" not in result and "Error" not in result:
                    st.session_state.pfmea_result = result
                    st.success("生成成功！")
                    # 自动滚动到下方显示结果
                    st.markdown("---")
                    st.markdown("### 📄 生成的PFMEA报告")
                    st.markdown(st.session_state.pfmea_result)
                else:
                    st.error(result)

    # 显示结果区域
    if st.session_state.pfmea_result:
        st.markdown("---")
        st.markdown("### 📄 生成的PFMEA报告")
        st.markdown(st.session_state.pfmea_result)

# =================主程序入口=================
def main():
    # 侧边栏
    with st.sidebar:
        st.image("https://streamlit.io/images/brand/streamlit-mark-color.png", width=50)
        st.title("🔧 功能导航")
        page = st.radio("选择模块", ["图片处理", "PFMEA生成"])
        st.markdown("---")
        st.caption("v1.0 · 智能工具箱")

    # 主页面逻辑
    if page == "图片处理":
        module_image_sorter()
    else: # PFMEA生成
        module_pfmea_generator()

if __name__ == "__main__":
    main()
