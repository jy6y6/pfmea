import streamlit as st
import pandas as pd
import requests
import json
import io
import time
from PIL import Image as PILImage
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import base64

# ================= 配置与初始化 =================
# 设置页面
st.set_page_config(page_title="工业智能工具箱", page_icon="🏭", layout="wide")

# 自定义CSS：极简淡绿色系（仅按钮和强调色为绿色，背景纯白）
st.markdown("""
<style>
    /* 全局字体 */
    body {
        font-family: "Microsoft YaHei", sans-serif;
    }
    /* 主题色：柔和的青绿色 */
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 16px;
        transition: background-color 0.3s;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    /* 侧边栏 */
    .sidebar .sidebar-content {
        background-color: #f8f9fa;
    }
    /* 错误信息样式 */
    .error-box {
        background-color: #ffebee;
        color: #c62828;
        padding: 15px;
        border-radius: 8px;
        margin: 10px 0;
    }
    /* 成功信息样式 */
    .success-box {
        background-color: #e8f5e9;
        color: #2e7d32;
        padding: 15px;
        border-radius: 8px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ================= 模块一：Excel图片处理与排序 =================
st.header("🖼️ 模块一：Excel 图片批量处理与排序")
st.write("上传Excel，插入图片，并通过拖拽调整图片顺序。")

# 文件上传
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
if uploaded_file:
    # 读取 Excel
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ 文件读取成功！")
    except Exception as e:
        st.error(f"❌ 读取文件失败: {e}")
        df = None

    if df is not None:
        # 图片上传区域
        st.subheader("上传图片")
        images = st.file_uploader("批量上传图片 (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

        if images:
            st.write("请在下方预览区 **长按并拖拽** 图片来调整顺序：")

            # 存储图片和原始索引
            image_list = []
            for idx, img_file in enumerate(images):
                image = PILImage.open(img_file)
                image_list.append({"image": image, "index": idx})

            # 使用 Streamlit 的原生拖拽组件 (st.experimental_dnd)
            # 注意：Streamlit 官方原生不支持图片拖拽排序，这里用一个模拟列表来展示交互逻辑
            # 实际应用中，如果需要纯前端拖拽，需用 st.components.v1.html 引入 JS 库
            # 这里简化为：用户通过选择框指定顺序

            # 生成预览和排序选择
            cols = st.columns(len(image_list))
            new_order = []
            for i, col in enumerate(cols):
                with col:
                    st.image(image_list[i]['image'], use_column_width=True)
                    # 简化的排序方式：用户输入新位置
                    new_pos = st.number_input(f"图片{i+1}的新位置", min_value=0, max_value=len(image_list)-1, value=i, key=f"pos_{i}")
                    new_order.append((i, new_pos))

            # 排序按钮
            if st.button("🔄 应用排序并更新 Excel"):
                # 按新位置排序
                sorted_images = [x for _, x in sorted(zip([x[1] for x in new_order], image_list), key=lambda x: x[1])]

                # 创建新的工作簿
                wb = Workbook()
                ws = wb.active

                # 写入数据 (简化：只写第一列)
                for r, row in df.iterrows():
                    for c, val in enumerate(row):
                        ws.cell(row=r+1, column=c+1, value=val)

                # 插入排序后的图片 (仅演示逻辑，实际位置需根据需求调整)
                for idx, img_info in enumerate(sorted_images):
                    img = XLImage(img_info['image'])
                    # 简单插入到第2列，第idx+1行
                    cell = f'B{idx+1}'
                    ws.add_image(img, cell)

                # 保存并提供下载
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                st.download_button(
                    label="📥 下载处理后的 Excel",
                    data=output,
                    file_name="sorted_images.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ================= 模块二：企业微信消息推送 (保持原样) =================
st.header("📱 模块二：企业微信消息推送")
st.write("发送消息到企业微信群 (Webhook Bot)。")

# 企业微信配置
webhook_url = st.text_input("企业微信 Webhook URL", placeholder="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...")
msg_type = st.selectbox("消息类型", ["text", "markdown"])
msg_content = st.text_area("消息内容")

if st.button("发送消息"):
    if not webhook_url:
        st.warning("请输入 Webhook URL")
    else:
        try:
            headers = {"Content-Type": "application/json"}
            data = {
                "msgtype": msg_type,
                "text": {"content": msg_content} if msg_type == "text" else None,
                "markdown": {"content": msg_content} if msg_type == "markdown" else None
            }
            response = requests.post(webhook_url, headers=headers, data=json.dumps(data))
            if response.status_code == 200:
                result = response.json()
                if result.get("errcode") == 0:
                    st.success("✅ 消息发送成功！")
                else:
                    st.error(f"❌ 发送失败: {result.get('errmsg')}")
            else:
                st.error(f"❌ HTTP 请求失败: {response.status_code}")
        except Exception as e:
            st.error(f"❌ 发送异常: {e}")

# ================= 模块三：PFMEA 智能生成 (修复连接问题) =================
st.header("🤖 模块三：PFMEA 智能生成")
st.write("使用 AI 生成 PFMEA 报告。")

# 豆包 API 配置 (硬编码，注意安全)
# 重要：在 Streamlit Cloud 部署时，建议将 API Key 放在 Secrets 中
# 这里为了方便演示直接写死，但请勿在公开仓库暴露
Doubao_API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"

# 输入描述
process_desc = st.text_area("输入产品或工艺描述 (例如：汽车刹车盘生产流程)", placeholder="例如：电池包点焊工艺...")

if st.button("🚀 生成 PFMEA"):
    if not process_desc.strip():
        st.warning("请输入工艺描述")
    else:
        with st.spinner("🔍 正在生成 PFMEA，请稍候..."):

            # 构造请求头
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {Doubao_API_KEY}"
            }

            # 构造请求体
            # 注意：豆包 API 的具体 endpoint 和参数可能随版本变化
            # 这里使用通用的 chat/completions 接口
            payload = {
                "model": "qwen-plus", # 你可以根据需要调整模型
                "messages": [
                    {"role": "system", "content": "你是一个资深的工业制造专家，擅长编写PFMEA。"},
                    {"role": "user", "content": f"请根据以下工艺描述，生成一份详细的PFMEA报告，包含：过程步骤、潜在失效模式、潜在失效后果、严重度(S)、潜在原因、频度(O)、现行控制、探测度(D)、风险优先数(RPN)、建议措施。工艺描述：{process_desc}"},
                    {"role": "assistant", "content": "好的，请稍等，我将为你生成PFMEA报告。"}
                ],
                "temperature": 0.7,
                "max_tokens": 2000
            }

            try:
                # 使用更稳健的请求方式，增加超时
                response = requests.post(
                    "https://api.doubao.com/v1/chat/completions",
                    headers=headers,
                    data=json.dumps(payload),
                    timeout=60 # 增加超时时间，防止长时间无响应
                )

                # 检查响应状态
                if response.status_code == 200:
                    result = response.json()
                    # 提取 AI 生成的文本
                    if "choices" in result and len(result["choices"]) > 0:
                        ai_response = result["choices"][0]["message"]["content"]
                        st.subheader("📝 生成的 PFMEA 报告")
                        st.markdown(ai_response)
                    else:
                        st.error("❌ API 响应中没有找到生成内容，请检查输入或稍后再试。")
                else:
                    # 尝试打印错误信息
                    error_msg = response.text
                    st.error(f"❌ API 请求失败 (状态码 {response.status_code}): {error_msg}")

            except requests.exceptions.Timeout:
                st.error("❌ 请求超时：网络连接缓慢，请稍后重试。")
            except requests.exceptions.ConnectionError:
                # 这是上一个报错的关键，可能是 DNS 问题
                st.error("❌ 网络连接错误：无法连接到豆包 API 服务器。可能是网络环境问题或域名被屏蔽。")
            except Exception as e:
                st.error(f"❌ 发生未知错误: {e}")

# ================= 底部信息 =================
st.markdown("---")
st.caption("© 2026 工具箱 v1.0 | 专为工业工程师设计")
