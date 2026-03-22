import streamlit as st
import pandas as pd
import requests
import json
import io
import os
import re
import time
import hashlib
import base64
from datetime import datetime, timedelta
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
from PIL import Image as PILImage
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# 导入本地知识库（100+ 工序，每个至少4条失效模式）
try:
    from knowledge_base import LOCAL_KNOWLEDGE_BASE
except ImportError:
    # 如果文件不存在，使用默认空字典
    LOCAL_KNOWLEDGE_BASE = {}
    st.warning("未找到 knowledge_base.py，请确保该文件存在。")

# ===================== 全局配置 =====================
st.set_page_config(
    page_title="工程工具箱",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS（精致淡绿色主题，与之前相同）
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f0f7e8 0%, #e9f3e0 100%);
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
    }
    .card {
        background: rgba(255, 255, 255, 0.96);
        border-radius: 28px;
        padding: 28px;
        margin-bottom: 28px;
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.05), 0 2px 4px rgba(0, 0, 0, 0.02);
        border: 1px solid rgba(212, 226, 193, 0.5);
        transition: all 0.3s cubic-bezier(0.2, 0, 0, 1);
    }
    .card:hover {
        box-shadow: 0 20px 32px rgba(0, 0, 0, 0.08);
        transform: translateY(-2px);
        border-color: #c2ddac;
    }
    .stButton button {
        background: linear-gradient(95deg, #6f9e6f 0%, #5a8a5a 100%);
        color: white;
        border-radius: 40px;
        border: none;
        padding: 0.6rem 1.8rem;
        font-weight: 500;
        font-size: 0.95rem;
        letter-spacing: 0.3px;
        transition: all 0.2s;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stButton button:hover {
        background: linear-gradient(95deg, #5a8a5a 0%, #4f784f 100%);
        transform: translateY(-1px);
        box-shadow: 0 6px 12px rgba(111, 158, 111, 0.2);
    }
    .stButton button:active {
        transform: translateY(0);
    }
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 20px;
        border: 1px solid #d4e2c1;
        background-color: #ffffff;
        padding: 0.6rem 1rem;
        transition: 0.2s;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #6f9e6f;
        box-shadow: 0 0 0 2px rgba(111,158,111,0.2);
        outline: none;
    }
    h1, h2, h3 {
        color: #2c5a2c;
        font-weight: 600;
        letter-spacing: -0.02em;
    }
    h1 {
        font-size: 2.2rem;
        margin-bottom: 0.2rem;
    }
    h2 {
        font-size: 1.6rem;
        margin-top: 0.5rem;
        border-left: 4px solid #6f9e6f;
        padding-left: 16px;
    }
    h3 {
        font-size: 1.2rem;
        margin-top: 0.5rem;
    }
    .dataframe {
        border-radius: 20px;
        overflow: hidden;
        font-size: 0.85rem;
        border-collapse: collapse;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    .dataframe th {
        background-color: #eef5e6;
        color: #2c5a2c;
        font-weight: 600;
        padding: 10px 12px;
    }
    .dataframe td {
        padding: 8px 12px;
    }
    hr {
        margin: 32px 0;
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #d4e2c1, transparent);
    }
    .stAlert {
        border-radius: 20px;
        border-left-width: 6px;
    }
    @media (max-width: 768px) {
        .card {
            padding: 20px;
        }
        h1 {
            font-size: 1.8rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# 初始化 session_state
def init_session():
    if "current_page" not in st.session_state:
        st.session_state.current_page = "home"
    if "push_history" not in st.session_state:
        st.session_state.push_history = []
    if "uploaded_images" not in st.session_state:
        st.session_state.uploaded_images = []
    if "user_knowledge_base" not in st.session_state:
        st.session_state.user_knowledge_base = {}      # 用户导入的知识库
    if "generated_pfmea_data" not in st.session_state:
        st.session_state.generated_pfmea_data = {}
    if "selected_ai_scheme" not in st.session_state:
        st.session_state.selected_ai_scheme = {}

init_session()

# ===================== 通用辅助函数 =====================
def compress_image_to_limit(image_bytes, max_size_mb=2, max_side=1024):
    img = PILImage.open(io.BytesIO(image_bytes))
    if img.mode in ('RGBA', 'LA', 'P'):
        rgb = PILImage.new('RGB', img.size, (255, 255, 255))
        rgb.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
        img = rgb
    img.thumbnail((max_side, max_side), PILImage.Resampling.LANCZOS)
    output = io.BytesIO()
    quality = 85
    while True:
        output.seek(0)
        output.truncate()
        img.save(output, format='JPEG', quality=quality, optimize=True)
        if output.tell() <= max_size_mb * 1024 * 1024 or quality <= 30:
            break
        quality -= 5
    return output.getvalue()

def send_to_wechat_robot(image_bytes_list, webhook_url, text_content=None):
    success_count = 0
    for idx, img_bytes in enumerate(image_bytes_list):
        try:
            compressed = compress_image_to_limit(img_bytes)
            b64 = base64.b64encode(compressed).decode('utf-8')
            md5 = hashlib.md5(compressed).hexdigest()
            payload = {"msgtype": "image", "image": {"base64": b64, "md5": md5}}
            if text_content and idx == 0:
                text_payload = {"msgtype": "text", "text": {"content": text_content}}
                requests.post(webhook_url, json=text_payload, timeout=10)
            response = requests.post(webhook_url, json=payload, timeout=10)
            if response.json().get("errcode") == 0:
                success_count += 1
        except Exception:
            continue
    return success_count

def clean_history_limit(history, max_total=200, keep=100):
    return history[-keep:] if len(history) > max_total else history

def reset_df_index(df):
    if not df.empty and not isinstance(df.index, pd.RangeIndex):
        df = df.reset_index(drop=True)
    return df

def export_history_to_excel(history_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        history_df.to_excel(writer, index=False, sheet_name="历史记录")
    output.seek(0)
    return output

# ===================== 模块一：Excel 图片工具（简化版）=====================
def excel_image_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📸 Excel 图片工具")
    st.markdown("将多张图片按上传顺序插入 Excel 表格指定区域，支持新建或加载现有文件。")

    # 单元格区域设置
    col1, col2 = st.columns(2)
    with col1:
        start_cell = st.text_input("起始单元格 (如 A1)", "A1")
    with col2:
        end_cell = st.text_input("结束单元格 (如 C5)", "C5")

    # 计算单元格数量
    try:
        start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
        end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
        if start_match and end_match:
            start_col = openpyxl.utils.column_index_from_string(start_match.group(1))
            start_row = int(start_match.group(2))
            end_col = openpyxl.utils.column_index_from_string(end_match.group(1))
            end_row = int(end_match.group(2))
            rows = end_row - start_row + 1
            cols = end_col - start_col + 1
            total_cells = rows * cols
            st.info(f"📐 区域共 {total_cells} 个单元格（{rows} 行 × {cols} 列）")
        else:
            st.warning("单元格格式错误，示例：A1")
            total_cells = 0
    except:
        total_cells = 0

    # 图片上传
    st.markdown("#### 🖼️ 选择图片")
    uploaded_files = st.file_uploader(
        "支持 JPG、PNG、BMP 格式，可多选",
        type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True,
        key="img_upload"
    )
    if uploaded_files:
        st.session_state.uploaded_images = [(f.name, f.getvalue()) for f in uploaded_files]
        st.success(f"✅ 已上传 {len(st.session_state.uploaded_images)} 张图片，将按上传顺序插入")

    # Excel 来源选择
    st.markdown("#### 📁 Excel 文件来源")
    excel_source = st.radio("选择新建或现有文件", ["新建空白工作簿", "上传现有 Excel 文件"], horizontal=True)
    existing_wb = None
    if excel_source == "上传现有 Excel 文件":
        existing_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xlsm"])
        if existing_file:
            try:
                existing_wb = load_workbook(io.BytesIO(existing_file.read()))
                st.success("已加载现有 Excel")
            except Exception as e:
                st.error(f"加载失败: {e}")

    # 生成下载
    if st.button("🚀 生成并下载 Excel 文件", type="primary", width='stretch'):
        if not st.session_state.uploaded_images:
            st.error("请先上传图片")
        elif not start_cell or not end_cell or total_cells == 0:
            st.error("请填写正确的起始和结束单元格")
        else:
            try:
                start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
                end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
                start_col = openpyxl.utils.column_index_from_string(start_match.group(1))
                start_row = int(start_match.group(2))
                end_col = openpyxl.utils.column_index_from_string(end_match.group(1))
                end_row = int(end_match.group(2))

                if existing_wb:
                    wb = existing_wb
                    ws = wb.active
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "图片表格"

                for row in range(start_row, end_row+1):
                    ws.row_dimensions[row].height = 150
                for col in range(start_col, end_col+1):
                    ws.column_dimensions[get_column_letter(col)].width = 15

                idx = 0
                for r in range(start_row, end_row+1):
                    for c in range(start_col, end_col+1):
                        if idx >= len(st.session_state.uploaded_images):
                            break
                        img_name, img_bytes = st.session_state.uploaded_images[idx]
                        try:
                            pil_img = PILImage.open(io.BytesIO(img_bytes))
                            max_w = 140
                            max_h = 140
                            ratio = min(max_w/pil_img.width, max_h/pil_img.height)
                            new_w = int(pil_img.width * ratio)
                            new_h = int(pil_img.height * ratio)
                            resized = pil_img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                            temp_buf = io.BytesIO()
                            resized.save(temp_buf, format='PNG')
                            temp_buf.seek(0)
                            xl_img = XLImage(temp_buf)
                            xl_img.width = new_w
                            xl_img.height = new_h
                            cell_coord = f"{get_column_letter(c)}{r}"
                            ws.add_image(xl_img, cell_coord)
                        except Exception as e:
                            st.warning(f"插入图片 {img_name} 失败: {e}")
                        idx += 1
                    if idx >= len(st.session_state.uploaded_images):
                        break

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                st.download_button(
                    label="📥 下载 Excel 文件",
                    data=output,
                    file_name=f"图片表格_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )
                st.success("✅ Excel 生成完成！")
            except Exception as e:
                st.error(f"生成失败: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 模块二：信息推送工具（保持不变）=====================
def image_push_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📱 信息推送工具")
    st.markdown("填写检测信息，拍照/选图推送至企业微信群")

    WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=bdf3c7d5-a7fd-4d5a-92bb-4ab15a32e042"

    with st.form("push_form"):
        col1, col2 = st.columns(2)
        with col1:
            model = st.text_input("型号 *", placeholder="请输入产品型号")
            line = st.text_input("线体 *", placeholder="请输入生产线体")
        with col2:
            detection_date = st.date_input("检测日期", value=datetime.now().date())
            inspector = st.text_input("检测人", placeholder="姓名")

        detection_desc = st.text_area("检测情况 *", placeholder="请详细描述检测情况...", height=100)
        remark = st.text_area("备注（选填）", placeholder="补充说明", height=60)

        st.markdown("**检测图片（可多选，最多2张）**")
        images = st.file_uploader("点击拍照 / 从相册选择", type=["jpg", "jpeg", "png", "bmp"], accept_multiple_files=True, key="push_images")
        if images and len(images) > 2:
            st.warning("最多选择2张图片，已自动限制")
            images = images[:2]

        submitted = st.form_submit_button("📤 提交并推送至企业微信", type="primary", width='stretch')

        if submitted:
            if not model or not line or not detection_desc:
                st.error("请填写带 * 的必填项")
            else:
                text_content = f"【检测报告】\n型号: {model}\n线体: {line}\n检测日期: {detection_date}\n检测人: {inspector}\n检测情况: {detection_desc}\n备注: {remark}"
                image_bytes_list = [img.getvalue() for img in images] if images else []
                success_cnt = send_to_wechat_robot(image_bytes_list, WEBHOOK_URL, text_content)
                if success_cnt == len(image_bytes_list):
                    st.success("✅ 推送成功！")
                    record = {
                        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "型号": model,
                        "线体": line,
                        "检测日期": detection_date.strftime("%Y-%m-%d"),
                        "检测人": inspector,
                        "检测情况": detection_desc,
                        "备注": remark,
                        "图片数量": len(image_bytes_list),
                        "推送结果": "成功"
                    }
                    st.session_state.push_history.append(record)
                    st.session_state.push_history = clean_history_limit(st.session_state.push_history)
                else:
                    st.error(f"❌ 推送失败（成功{success_cnt}/{len(image_bytes_list)}张）")
                    record = {
                        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "型号": model,
                        "线体": line,
                        "检测日期": detection_date.strftime("%Y-%m-%d"),
                        "检测人": inspector,
                        "检测情况": detection_desc,
                        "备注": remark,
                        "图片数量": len(image_bytes_list),
                        "推送结果": "失败"
                    }
                    st.session_state.push_history.append(record)
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # 历史记录管理
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📜 历史填报记录")
    if st.session_state.push_history:
        df_history = pd.DataFrame(st.session_state.push_history)

        # 时间筛选
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("起始日期", value=datetime.now().date() - timedelta(days=30))
        with col2:
            end_date = st.date_input("结束日期", value=datetime.now().date())
        if start_date <= end_date:
            df_history['时间日期'] = pd.to_datetime(df_history['时间']).dt.date
            mask = (df_history['时间日期'] >= start_date) & (df_history['时间日期'] <= end_date)
            filtered_df = df_history.loc[mask].drop(columns=['时间日期'])
        else:
            filtered_df = df_history
            st.error("起始日期不能大于结束日期")

        st.dataframe(filtered_df, use_container_width=True)

        if st.button("📥 导出当前筛选记录为 Excel", width='stretch'):
            if not filtered_df.empty:
                excel_file = export_history_to_excel(filtered_df)
                st.download_button(
                    label="点击下载 Excel",
                    data=excel_file,
                    file_name=f"推送记录_{start_date}_{end_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )
            else:
                st.warning("没有符合条件的记录")
    else:
        st.info("暂无历史记录")
    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 模块三：PFMEA 智能生成工具 =====================
# 注意：知识库来自 LOCAL_KNOWLEDGE_BASE（100+ 工序）和用户导入的 user_knowledge_base

# AI 生成函数（复用之前的代码，但移除了测试按钮）
def create_retry_session():
    session = requests.Session()
    retry = Retry(total=3, backoff_factor=1, status_forcelist=[429,500,502,503,504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    return session

def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_ENDPOINTS = [
        "https://api.doubao.com/v1/chat/completions",
        "https://api.doubaoai.com/v1/chat/completions",
        "https://open.doubao.com/v1/chat/completions"
    ]
    MODELS = ["doubao-pro-32k", "ep-20240805194357-jzrql", "doubao-lite-32k"]
    session = create_retry_session()
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    prompt = f"""
    你是专业的汽车电子行业PFMEA工程师，精通AIAG-VDA FMEA标准。
    针对【{process_name}】工序（产品类型：{product_type}），生成{scheme_count}组完全不同的PFMEA方案。
    每组方案包含3-5条失效模式，必须涵盖不同维度（人、机、料、法、环），确保内容差异化。
    返回严格的JSON格式：[{{"方案名称":"方案1：...","pfmea_list":[{{"失效模式":"...","失效后果":"...","失效原因":"...","预防措施":"...","探测措施":"...","严重度S":x,"频度O":x,"探测度D":x,"AP等级":"x"}}]}}]
    只返回JSON，不要其他文字。
    """
    last_error = ""
    for endpoint in API_ENDPOINTS:
        for model in MODELS:
            data = {
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.8,
                "max_tokens": 4000
            }
            try:
                response = session.post(endpoint, headers=headers, json=data, timeout=90)
                response.raise_for_status()
                result = response.json()
                if "choices" not in result or not result["choices"]:
                    last_error = f"端点 {endpoint} 模型 {model} 返回无choices"
                    continue
                content = result["choices"][0]["message"]["content"]
                content = re.sub(r'^```json\s*|\s*```$', '', content.strip())
                parsed = json.loads(content)
                if isinstance(parsed, list) and all("pfmea_list" in s for s in parsed):
                    return parsed, None
                else:
                    last_error = f"返回格式不正确: {content[:200]}"
            except json.JSONDecodeError as e:
                last_error = f"JSON解析失败: {e}\n原始内容: {content[:200] if 'content' in locals() else '无'}"
                continue
            except Exception as e:
                last_error = str(e)
                continue
    return None, f"所有API端点均失败，最后错误: {last_error}"

def parse_pfmea_excel(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        possible_cols = {
            "工序": ["工序", "过程步骤", "工序名称", "过程"],
            "失效模式": ["失效模式", "潜在失效模式"],
            "失效后果": ["失效后果", "潜在影响", "失效影响"],
            "失效原因": ["失效原因", "潜在原因", "失效起因"],
            "预防措施": ["预防措施", "预防控制"],
            "探测措施": ["探测措施", "探测控制"],
            "严重度S": ["严重度S", "严重度", "S"],
            "频度O": ["频度O", "频度", "O"],
            "探测度D": ["探测度D", "探测度", "D"],
            "AP等级": ["AP等级", "AP", "优先级"]
        }
        col_mapping = {}
        for target, candidates in possible_cols.items():
            for col in df.columns:
                if any(c in col for c in candidates):
                    col_mapping[target] = col
                    break
        if "工序" not in col_mapping:
            st.error("Excel中未找到工序列")
            return {}
        knowledge = {}
        for _, row in df.iterrows():
            process = str(row[col_mapping["工序"]]).strip()
            if not process or process == "nan":
                continue
            item = {}
            for target, col in col_mapping.items():
                if target != "工序":
                    val = row[col] if pd.notna(row[col]) else ""
                    if target in ["严重度S", "频度O", "探测度D"]:
                        try:
                            val = int(float(val))
                        except:
                            val = 5
                    item[target] = val
            if "AP等级" not in item:
                s = item.get("严重度S", 5)
                o = item.get("频度O", 3)
                d = item.get("探测度D", 4)
                if s >= 9 or (s >= 7 and (o >= 4 or d >= 8)):
                    ap = "高"
                elif s <= 4 and o <= 6:
                    ap = "低"
                else:
                    ap = "中"
                item["AP等级"] = ap
            if process not in knowledge:
                knowledge[process] = []
            knowledge[process].append(item)
        return knowledge
    except Exception as e:
        st.error(f"解析Excel失败: {e}")
        return {}

def merge_knowledge(knowledge_dict, existing_kb):
    for proc, items in knowledge_dict.items():
        if proc not in existing_kb:
            existing_kb[proc] = []
        existing_items = existing_kb[proc]
        existing_keys = {f"{i.get('失效模式','')}_{i.get('失效原因','')}" for i in existing_items}
        for item in items:
            key = f"{item.get('失效模式','')}_{item.get('失效原因','')}"
            if key not in existing_keys:
                existing_items.append(item)
                existing_keys.add(key)
    return existing_kb

def export_knowledge_to_json():
    return json.dumps(st.session_state.user_knowledge_base, ensure_ascii=False, indent=2)

def import_knowledge_from_json(json_str):
    try:
        data = json.loads(json_str)
        for proc, items in data.items():
            if proc not in st.session_state.user_knowledge_base:
                st.session_state.user_knowledge_base[proc] = []
            existing_keys = {f"{i.get('失效模式','')}_{i.get('失效原因','')}" for i in st.session_state.user_knowledge_base[proc]}
            for item in items:
                key = f"{item.get('失效模式','')}_{item.get('失效原因','')}"
                if key not in existing_keys:
                    st.session_state.user_knowledge_base[proc].append(item)
                    existing_keys.add(key)
        return True
    except:
        return False

def export_pfmea_excel(pfmea_data, product_type):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "PFMEA"
    headers = ["工序", "失效模式", "失效后果", "失效原因", "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, size=11, name="微软雅黑")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    row = 2
    for process, items in pfmea_data.items():
        for item in items:
            ws.cell(row=row, column=1, value=process)
            ws.cell(row=row, column=2, value=item.get("失效模式", ""))
            ws.cell(row=row, column=3, value=item.get("失效后果", ""))
            ws.cell(row=row, column=4, value=item.get("失效原因", ""))
            ws.cell(row=row, column=5, value=item.get("预防措施", ""))
            ws.cell(row=row, column=6, value=item.get("探测措施", ""))
            ws.cell(row=row, column=7, value=item.get("严重度S", ""))
            ws.cell(row=row, column=8, value=item.get("频度O", ""))
            ws.cell(row=row, column=9, value=item.get("探测度D", ""))
            ws.cell(row=row, column=10, value=item.get("AP等级", ""))
            for col in range(1, 11):
                cell = ws.cell(row=row, column=col)
                cell.font = Font(name="微软雅黑", size=9)
                cell.alignment = Alignment(horizontal="center" if col>=7 else "left", vertical="center", wrap_text=True)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            ap = item.get("AP等级", "")
            if ap == "高":
                fill = PatternFill(start_color="FF4D4F", end_color="FF4D4F", fill_type="solid")
            elif ap == "中":
                fill = PatternFill(start_color="FAAD14", end_color="FAAD14", fill_type="solid")
            elif ap == "低":
                fill = PatternFill(start_color="52C41A", end_color="52C41A", fill_type="solid")
            else:
                fill = None
            if fill:
                ws.cell(row=row, column=10).fill = fill
            row += 1
    col_widths = [18, 25, 30, 30, 35, 35, 8, 8, 8, 8]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[chr(64+i)].width = width
    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output

def pfmea_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title("⚡ PFMEA 智能生成系统")
    st.caption("符合 AIAG-VDA FMEA 第一版 | IATF16949:2016")
    st.divider()

    # 产品类型
    product_type = st.radio("产品类型", ["电池包", "充电器"], horizontal=True)

    # 合并本地知识库和用户知识库，得到所有工序列表
    # 注意：LOCAL_KNOWLEDGE_BASE 已经包含了电池包和充电器的所有工序
    all_processes = set(LOCAL_KNOWLEDGE_BASE.keys())
    if st.session_state.user_knowledge_base:
        all_processes.update(st.session_state.user_knowledge_base.keys())
    all_processes = sorted(all_processes)

    # 搜索过滤工序
    st.markdown("#### 🔍 搜索工序")
    search_term = st.text_input("输入关键词过滤工序", placeholder="例如：焊接、测试、来料")
    if search_term:
        filtered_processes = [p for p in all_processes if search_term.lower() in p.lower()]
    else:
        filtered_processes = all_processes

    st.markdown(f"共找到 **{len(filtered_processes)}** 个工序（总计 {len(all_processes)} 个）")

    # 可搜索的多选框（由于 streamlit 原生不支持搜索，我们使用文本框+多选框的组合）
    selected_processes = st.multiselect("选择工序（可多选）", filtered_processes, default=filtered_processes[:2] if filtered_processes else [])

    # 自定义工序添加
    col1, col2 = st.columns([3, 1])
    with col1:
        custom_process = st.text_input("自定义工序名称")
    with col2:
        if st.button("➕ 添加自定义工序", width='stretch'):
            if custom_process and custom_process not in selected_processes:
                selected_processes.append(custom_process)
                st.success(f"已添加工序：{custom_process}")
                st.rerun()

    # 知识库管理
    with st.expander("📚 知识库管理（可导入/编辑/导出）"):
        uploaded_kb_file = st.file_uploader("导入已有 PFMEA Excel 文件（.xlsx）", type=["xlsx"], key="kb_import")
        if uploaded_kb_file:
            with st.spinner("正在解析文件..."):
                kb_data = parse_pfmea_excel(uploaded_kb_file.getvalue())
                if kb_data:
                    st.session_state.user_knowledge_base = merge_knowledge(kb_data, st.session_state.user_knowledge_base)
                    st.success(f"成功导入 {len(kb_data)} 个工序的条目！")
                    st.rerun()
        if st.session_state.user_knowledge_base:
            for proc, items in st.session_state.user_knowledge_base.items():
                with st.expander(f"📁 {proc}（共{len(items)}条）"):
                    df_kb = pd.DataFrame(items)
                    df_kb = reset_df_index(df_kb)
                    edited_df = st.data_editor(df_kb, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"kb_edit_{proc}")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("✅ 更新此工序", key=f"kb_update_{proc}"):
                            st.session_state.user_knowledge_base[proc] = edited_df.to_dict("records")
                            st.success("已更新")
                    with col2:
                        if st.button("🗑️ 删除此工序", key=f"kb_del_{proc}"):
                            del st.session_state.user_knowledge_base[proc]
                            st.rerun()
            col1, col2 = st.columns(2)
            with col1:
                kb_json = export_knowledge_to_json()
                st.download_button("📤 导出知识库备份（JSON）", data=kb_json, file_name=f"PFMEA知识库_{datetime.now().strftime('%Y%m%d')}.json", mime="application/json", width='stretch')
            with col2:
                backup_file = st.file_uploader("📥 导入知识库备份（JSON）", type=["json"], key="kb_import_json")
                if backup_file:
                    try:
                        if import_knowledge_from_json(backup_file.read().decode('utf-8')):
                            st.success("导入成功")
                            st.rerun()
                        else:
                            st.error("导入失败，文件格式错误")
                    except:
                        st.error("导入失败")
        else:
            st.info("暂无用户知识库内容，可通过上传旧PFMEA文件或AI生成后自动入库来积累。")

    # 生成设置
    st.subheader("生成设置")
    gen_mode = st.radio("生成模式", ["本地标准库（含知识库）", "AI智能生成（多方案）"], horizontal=True)
    scheme_count = 3
    mix_knowledge = False
    if gen_mode == "AI智能生成（多方案）":
        scheme_count = st.slider("AI生成方案数量", 2, 5, 3)
        mix_knowledge = st.checkbox("混合知识库内容作为独立方案", value=True, help="勾选后，知识库中该工序的内容将作为一个额外方案供选择")

    # 生成按钮
    if st.button("🚀 生成PFMEA方案", type="primary", width='stretch') and selected_processes:
        st.session_state.generated_pfmea_data = {}
        st.session_state.selected_ai_scheme = {}
        progress_bar = st.progress(0)
        for idx, proc in enumerate(selected_processes):
            progress_bar.progress((idx) / len(selected_processes))
            if gen_mode == "本地标准库（含知识库）":
                # 从本地库（LOCAL_KNOWLEDGE_BASE）和用户库中获取
                lib_items = LOCAL_KNOWLEDGE_BASE.get(proc, [])
                kb_items = st.session_state.user_knowledge_base.get(proc, [])
                combined = lib_items + kb_items
                if not combined:
                    st.warning(f"工序【{proc}】无本地数据，请使用AI生成或先导入知识库")
                    continue
                st.session_state.generated_pfmea_data[proc] = [{"方案名称": "本地库+知识库方案", "pfmea_list": combined}]
            else:
                with st.spinner(f"AI正在生成【{proc}】的方案..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if err:
                        st.error(f"{proc} AI生成失败: {err}")
                        # 降级到本地库
                        lib_items = LOCAL_KNOWLEDGE_BASE.get(proc, [])
                        if lib_items:
                            st.warning(f"已使用本地库作为备用方案")
                            st.session_state.generated_pfmea_data[proc] = [{"方案名称": "本地库备用方案", "pfmea_list": lib_items}]
                        else:
                            continue
                    else:
                        if mix_knowledge and proc in st.session_state.user_knowledge_base:
                            kb_scheme = {"方案名称": "📁 我的知识库方案", "pfmea_list": st.session_state.user_knowledge_base[proc]}
                            schemes.append(kb_scheme)
                        st.session_state.generated_pfmea_data[proc] = schemes
                        st.session_state.selected_ai_scheme[proc] = 0
        progress_bar.progress(1.0)
        if st.session_state.generated_pfmea_data:
            st.success("生成完成！请选择方案")
            st.rerun()

    # 方案选择与预览
    if st.session_state.generated_pfmea_data:
        st.subheader("选择最终方案")
        final_data = {}
        for proc in selected_processes:
            if proc not in st.session_state.generated_pfmea_data:
                continue
            data = st.session_state.generated_pfmea_data[proc]
            if len(data) == 1:
                selected_scheme = data[0]
            else:
                scheme_names = [s["方案名称"] for s in data]
                selected_idx = st.radio(f"为【{proc}】选择方案", options=range(len(scheme_names)), format_func=lambda i: scheme_names[i], key=f"select_{proc}", index=st.session_state.selected_ai_scheme.get(proc, 0))
                st.session_state.selected_ai_scheme[proc] = selected_idx
                selected_scheme = data[selected_idx]
            st.markdown(f"**{selected_scheme['方案名称']}**")
            df_scheme = pd.DataFrame(selected_scheme["pfmea_list"])
            df_scheme = reset_df_index(df_scheme)
            edited_df = st.data_editor(df_scheme, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"edit_{proc}")
            final_data[proc] = edited_df.to_dict("records")
            if st.button(f"💾 将当前方案存入知识库", key=f"save_kb_{proc}"):
                if proc not in st.session_state.user_knowledge_base:
                    st.session_state.user_knowledge_base[proc] = []
                existing_keys = {f"{i.get('失效模式','')}_{i.get('失效原因','')}" for i in st.session_state.user_knowledge_base[proc]}
                for item in edited_df.to_dict("records"):
                    key = f"{item.get('失效模式','')}_{item.get('失效原因','')}"
                    if key not in existing_keys:
                        st.session_state.user_knowledge_base[proc].append(item)
                        existing_keys.add(key)
                st.success(f"已存入知识库（去重）")
                st.rerun()
            st.divider()

        if st.button("✅ 确认并导出 Excel 文件", type="primary", width='stretch'):
            excel_file = export_pfmea_excel(final_data, product_type)
            st.download_button(
                label="📥 下载 PFMEA Excel 文件",
                data=excel_file,
                file_name=f"{product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )

    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 主界面 =====================
def main():
    if st.session_state.current_page == "home":
        st.markdown("<div style='text-align: center; padding: 2rem 0 1rem;'><h1>🛠️ 工程工具箱</h1><p style='color:#6c7a6c;'>请选择要使用的工具模块</p></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        with col1:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("📊", unsafe_allow_html=True)
                st.subheader("Excel 图片工具")
                st.markdown("将多张图片按顺序插入 Excel 表格")
                if st.button("进入工具", key="btn_excel", width='stretch'):
                    st.session_state.current_page = "excel_image"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("💬", unsafe_allow_html=True)
                st.subheader("信息推送工具")
                st.markdown("拍照/选图推送至企业微信群")
                if st.button("进入工具", key="btn_push", width='stretch'):
                    st.session_state.current_page = "image_push"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col3:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("⚡", unsafe_allow_html=True)
                st.subheader("PFMEA 智能生成")
                st.markdown("符合 AIAG-VDA 标准的 FMEA 生成")
                if st.button("进入工具", key="btn_pfmea", width='stretch'):
                    st.session_state.current_page = "pfmea"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

    elif st.session_state.current_page == "excel_image":
        excel_image_tool()
        if st.button("🏠 返回首页", width='stretch'):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "image_push":
        image_push_tool()
        if st.button("🏠 返回首页", width='stretch'):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "pfmea":
        pfmea_tool()
        if st.button("🏠 返回首页", width='stretch'):
            st.session_state.current_page = "home"
            st.rerun()

if __name__ == "__main__":
    main()
