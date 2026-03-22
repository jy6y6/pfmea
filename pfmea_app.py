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

# ===================== 全局配置 =====================
st.set_page_config(
    page_title="多功能智能工具集",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 🔥 升级自定义CSS（高级淡绿极简风）
st.markdown("""
<style>
/* 全局重置 */
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
    font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
}
/* 全局背景 */
.stApp {
    background-color: #F8FAF3;
}
/* 主卡片样式 */
.card {
    background: #FFFFFF;
    border-radius: 20px;
    padding: 28px;
    margin-bottom: 24px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.04);
    border: 1px solid #E6EFD9;
    transition: all 0.3s ease;
}
.card:hover {
    box-shadow: 0 12px 32px rgba(0,0,0,0.08);
    transform: translateY(-2px);
}
/* 按钮样式 */
.stButton button {
    background: linear-gradient(135deg, #6F9E6F, #5A8F5A);
    color: white;
    border-radius: 12px;
    border: none;
    padding: 0.7rem 1.4rem;
    font-weight: 600;
    transition: 0.3s;
    height: 44px;
}
.stButton button:hover {
    background: linear-gradient(135deg, #5A8F5A, #4A7A4A);
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(111,158,111,0.3);
}
/* 输入框/选择框 */
.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"], .stDateInput input {
    border-radius: 12px;
    border: 1px solid #D4E2C1;
    padding: 10px 14px;
    transition: 0.2s;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #6F9E6F;
    box-shadow: 0 0 0 2px rgba(111,158,111,0.2);
}
/* 标题样式 */
h1 {
    color: #2D5A2D;
    font-weight: 700;
    margin-bottom: 12px;
}
h2, h3 {
    color: #3A6B3A;
    font-weight: 600;
}
/* 标签文字 */
.stMarkdown p, .stText {
    color: #444444;
    line-height: 1.6;
}
/* 表格/编辑器 */
.dataframe, .stDataEditor {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid #E6EFD9;
}
/* 分割线 */
hr {
    border: none;
    height: 1px;
    background-color: #E6EFD9;
    margin: 20px 0;
}
</style>
""", unsafe_allow_html=True)

# 初始化session_state
def init_session():
    if "current_page" not in st.session_state:
        st.session_state.current_page = "home"
    if "push_history" not in st.session_state:
        st.session_state.push_history = []
    if "image_order" not in st.session_state:
        st.session_state.image_order = []
    if "uploaded_images" not in st.session_state:
        st.session_state.uploaded_images = []
    if "user_knowledge_base" not in st.session_state:
        st.session_state.user_knowledge_base = {}
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

# ===================== 模块一：Excel图片工具（稳定版，无拖拽依赖） =====================
def excel_image_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📸 Excel 图片批量插入工具")
    st.markdown("支持多选图片 → 输入数字排序 → 一键插入Excel指定区域")

    # 单元格区域设置
    col1, col2 = st.columns(2)
    with col1:
        start_cell = st.text_input("起始单元格", "A1")
    with col2:
        end_cell = st.text_input("结束单元格", "C5")

    # 计算单元格数量
    total_cells = 0
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
            st.info(f"✅ 区域共 {total_cells} 个单元格（{rows}行 × {cols}列）")
    except:
        st.warning("⚠️ 单元格格式错误，示例：A1")

    # 图片上传
    uploaded_files = st.file_uploader(
        "选择图片（可多选）",
        type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True,
        key="img_upload"
    )
    if uploaded_files:
        st.session_state.uploaded_images = [(f.name, f.getvalue()) for f in uploaded_files]
        if not st.session_state.image_order:
            st.session_state.image_order = list(range(len(st.session_state.uploaded_images)))
        st.success(f"✅ 已上传 {len(st.session_state.uploaded_images)} 张图片")

    # 🔥 稳定版：数字输入排序（替代拖拽）
    if st.session_state.uploaded_images and total_cells > 0:
        st.subheader("📊 调整图片顺序（输入数字，1为第一个）")
        order_df = pd.DataFrame({
            "图片名称": [name for name, _ in st.session_state.uploaded_images],
            "顺序": st.session_state.image_order
        })
        edited_df = st.data_editor(
            order_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "顺序": st.column_config.NumberColumn(
                    min_value=1,
                    max_value=len(st.session_state.uploaded_images),
                    step=1,
                    help="输入1~N的数字，数字越小越靠前"
                )
            }
        )
        # 应用排序
        if st.button("🔄 应用排序"):
            sorted_df = edited_df.sort_values("顺序").reset_index(drop=True)
            st.session_state.image_order = sorted_df.index.tolist()
            st.success("✅ 排序已应用！")
            st.rerun()

        # 显示当前顺序
        order_names = [st.session_state.uploaded_images[idx][0][:15] for idx in st.session_state.image_order]
        st.write("**当前顺序：** " + " → ".join(order_names))

    # Excel来源选择
    excel_source = st.radio("Excel 来源", ["新建空白工作簿", "上传现有 Excel 文件"], horizontal=True)
    existing_wb = None
    if excel_source == "上传现有 Excel 文件":
        existing_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xlsm"])
        if existing_file:
            try:
                existing_wb = load_workbook(io.BytesIO(existing_file.read()))
                st.success("✅ 已加载现有 Excel")
            except Exception as e:
                st.error(f"❌ 加载失败: {e}")

    # 生成Excel
    if st.button("🚀 生成并下载 Excel", type="primary", use_container_width=True):
        if not st.session_state.uploaded_images:
            st.error("❌ 请先上传图片")
        elif total_cells == 0:
            st.error("❌ 请填写正确的单元格范围")
        else:
            with st.spinner("正在生成Excel..."):
                try:
                    start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
                    end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
                    start_col = openpyxl.utils.column_index_from_string(start_match.group(1))
                    start_row = int(start_match.group(2))
                    end_col = openpyxl.utils.column_index_from_string(end_match.group(1))
                    end_row = int(end_match.group(2))

                    wb = existing_wb if existing_wb else Workbook()
                    ws = wb.active
                    if not existing_wb:
                        ws.title = "图片表格"

                    for row in range(start_row, end_row+1):
                        ws.row_dimensions[row].height = 160
                    for col in range(start_col, end_col+1):
                        ws.column_dimensions[get_column_letter(col)].width = 16

                    idx = 0
                    for r in range(start_row, end_row+1):
                        for c in range(start_col, end_col+1):
                            if idx >= len(st.session_state.uploaded_images):
                                break
                            img_idx = st.session_state.image_order[idx]
                            img_name, img_bytes = st.session_state.uploaded_images[img_idx]
                            try:
                                pil_img = PILImage.open(io.BytesIO(img_bytes))
                                max_w, max_h = 140, 140
                                ratio = min(max_w/pil_img.width, max_h/pil_img.height)
                                new_w, new_h = int(pil_img.width*ratio), int(pil_img.height*ratio)
                                resized = pil_img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                                temp_buf = io.BytesIO()
                                resized.save(temp_buf, format='PNG')
                                temp_buf.seek(0)
                                xl_img = XLImage(temp_buf)
                                xl_img.width, xl_img.height = new_w, new_h
                                ws.add_image(xl_img, f"{get_column_letter(c)}{r}")
                            except:
                                st.warning(f"⚠️ {img_name} 插入失败")
                            idx += 1
                        if idx >= len(st.session_state.uploaded_images):
                            break

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    st.download_button(
                        label="📥 点击下载 Excel",
                        data=output,
                        file_name=f"图片表格_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.success("✅ Excel 生成完成！")
                except Exception as e:
                    st.error(f"❌ 生成失败: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🏠 返回首页", use_container_width=True):
        st.session_state.current_page = "home"
        st.rerun()

# ===================== 模块二：信息推送工具（完全保留） =====================
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
        submitted = st.form_submit_button("📤 提交并推送至企业微信", type="primary", use_container_width=True)
        if submitted:
            if not model or not line or not detection_desc:
                st.error("请填写带 * 的必填项")
            else:
                text_content = f"【检测报告】\n型号: {model}\n线体: {line}\n检测日期: {detection_date}\n检测人: {inspector}\n检测情况: {detection_desc}\n备注: {remark}"
                image_bytes_list = [img.getvalue() for img in images] if images else []
                success_cnt = send_to_wechat_robot(image_bytes_list, WEBHOOK_URL, text_content)
                if success_cnt == len(image_bytes_list):
                    st.success("推送成功！")
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
                    st.error(f"推送失败（成功{success_cnt}/{len(image_bytes_list)}张）")
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

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📜 历史填报记录")
    if st.session_state.push_history:
        df_history = pd.DataFrame(st.session_state.push_history)
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
        if st.button("📥 导出当前筛选记录为 Excel", use_container_width=True):
            if not filtered_df.empty:
                excel_file = export_history_to_excel(filtered_df)
                st.download_button(
                    label="点击下载 Excel",
                    data=excel_file,
                    file_name=f"推送记录_{start_date}_{end_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("没有符合条件的记录")
    else:
        st.info("暂无历史记录")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🏠 返回首页", use_container_width=True):
        st.session_state.current_page = "home"
        st.rerun()

# ===================== 模块三：PFMEA智能生成（修复AI+界面升级） =====================
BATTERY_PROCESS_LIB = {
    "电芯来料检验": [
        {"失效模式": "电芯外观尺寸超差", "失效后果": "电芯无法装入模组壳体", "失效原因": "来料尺寸公差不符合图纸要求", "预防措施": "制定电芯来料检验规范，量具定期校准", "探测措施": "首件全尺寸检验，巡检按AQL抽样", "严重度S": 6, "频度O": 3, "探测度D": 4, "AP等级": "中"},
        {"失效模式": "电芯电压/内阻异常", "失效后果": "模组充放电异常，循环寿命衰减", "失效原因": "电芯生产工艺异常，存储环境不达标", "预防措施": "每批次电压内阻全检，温湿度监控", "探测措施": "自动化检测设备100%全检，异常报警隔离", "严重度S": 9, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "电芯表面划伤/破损", "失效后果": "绝缘性能下降，可能引发短路", "失效原因": "来料包装破损，搬运过程中磕碰", "预防措施": "包装标准化，运输防护升级", "探测措施": "目视全检，不良品隔离", "严重度S": 7, "频度O": 3, "探测度D": 3, "AP等级": "高"},
    ],
    "模组堆叠装配": [
        {"失效模式": "电芯堆叠顺序错误", "失效后果": "电路连接错误，短路风险", "失效原因": "作业人员未按SOP操作，防错失效", "预防措施": "安装极性视觉防错装置，培训考核", "探测措施": "视觉设备100%检测，异常停机", "严重度S": 9, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "电芯之间间距不均", "失效后果": "散热不均，影响寿命", "失效原因": "堆叠工装磨损，定位不准", "预防措施": "定期校准工装，首件确认", "探测措施": "激光测距抽检", "严重度S": 5, "频度O": 3, "探测度D": 3, "AP等级": "中"},
        {"失效模式": "绝缘片漏装/错装", "失效后果": "短路风险，严重时起火", "失效原因": "物料清单错误，作业疏忽", "预防措施": "物料扫码防错，双人复核", "探测措施": "视觉系统检测绝缘片有无", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
    ]
}
CHARGER_PROCESS_LIB = {
    "PCB来料检验": [
        {"失效模式": "PCB板尺寸超差", "失效后果": "PCB无法装入壳体", "失效原因": "PCB生产制程偏差", "预防措施": "制定PCB来料检验规范，首件全检", "探测措施": "首件全尺寸检验，巡检抽检", "严重度S": 5, "频度O": 3, "探测度D": 4, "AP等级": "中"},
        {"失效模式": "铜箔起泡", "失效后果": "焊接可靠性下降，虚焊", "失效原因": "PCB受潮，层压工艺不良", "预防措施": "来料烘烤，存储湿度控制", "探测措施": "外观检查，切片分析", "严重度S": 6, "频度O": 2, "探测度D": 3, "AP等级": "中"},
    ],
    "SMT贴片焊接": [
        {"失效模式": "元器件贴装偏移", "失效后果": "焊接不良，功能失效", "失效原因": "贴片机吸嘴磨损，程序坐标偏差", "预防措施": "定期校准设备，首件验证", "探测措施": "AOI全检，SPI锡膏检测", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
    ]
}

# 🔥 修复AI生成函数（官方接口+稳定可用）
def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    
    session = requests.Session()
    retry = Retry(total=2, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    prompt = f"""你是专业AIAG-VDA PFMEA工程师，针对【{process_name}】工序，产品类型{product_type}，生成{scheme_count}组不同PFMEA方案。
    严格返回JSON格式，无其他文字：
    [{{"方案名称":"方案1：xxx","pfmea_list":[{{"失效模式":"","失效后果":"","失效原因":"","预防措施":"","探测措施":"","严重度S":int,"频度O":int,"探测度D":int,"AP等级":"高/中/低"}}]}}]"""
    
    try:
        data = {
            "model": "doubao-pro",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3
        }
        res = session.post(API_URL, headers=headers, json=data, timeout=30)
        res.raise_for_status()
        content = res.json()["choices"][0]["message"]["content"]
        content = re.sub(r"```json|```", "", content.strip())
        return json.loads(content), None
    except Exception as e:
        return None, f"AI生成失败：{str(e)}"

# 知识库/导出函数
def parse_pfmea_excel(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        possible_cols = {
            "工序": ["工序", "过程步骤"], "失效模式": ["失效模式"], "失效后果": ["失效后果"],
            "失效原因": ["失效原因"], "预防措施": ["预防措施"], "探测措施": ["探测措施"],
            "严重度S": ["严重度S"], "频度O": ["频度O"], "探测度D": ["探测度D"], "AP等级": ["AP等级"]
        }
        col_mapping = {}
        for target, candidates in possible_cols.items():
            for col in df.columns:
                if any(c in col for c in candidates):
                    col_mapping[target] = col
                    break
        knowledge = {}
        for _, row in df.iterrows():
            process = str(row[col_mapping["工序"]]).strip()
            if process == "nan": continue
            item = {k:row[v] if pd.notna(row[v]) else "" for k,v in col_mapping.items() if k!="工序"}
            for k in ["严重度S","频度O","探测度D"]:
                try: item[k] = int(float(item[k])) except: item[k] = 5
            if "AP等级" not in item or item["AP等级"] == "":
                s,o,d = item.get("严重度S",5),item.get("频度O",3),item.get("探测度D",4)
                item["AP等级"] = "高" if s>=9 else "中" if s>=5 else "低"
            if process not in knowledge: knowledge[process] = []
            knowledge[process].append(item)
        return knowledge
    except:
        return {}

def merge_knowledge(knowledge_dict, existing_kb):
    for proc, items in knowledge_dict.items():
        if proc not in existing_kb: existing_kb[proc] = []
        keys = {f"{i['失效模式']}_{i['失效原因']}" for i in existing_kb[proc]}
        for item in items:
            key = f"{item['失效模式']}_{item['失效原因']}"
            if key not in keys: existing_kb[proc].append(item)
    return existing_kb

def export_pfmea_excel(pfmea_data, product_type):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "PFMEA"
    headers = ["工序","失效模式","失效后果","失效原因","预防措施","探测措施","严重度S","频度O","探测度D","AP等级"]
    for col,h in enumerate(headers,1):
        cell = ws.cell(row=1,column=col,value=h)
        cell.font = Font(bold=True,name="微软雅黑")
        cell.alignment = Alignment(horizontal="center",vertical="center")
    row=2
    for proc,items in pfmea_data.items():
        for item in items:
            ws.cell(row=row,column=1,value=proc)
            for i,k in enumerate(headers[1:],2):
                ws.cell(row=row,column=i,value=item.get(k,""))
            ap = item.get("AP等级","")
            fill = PatternFill(start_color="FF4D4F",end_color="FF4D4F",fill_type="solid") if ap=="高" else \
                  PatternFill(start_color="FAAD14",end_color="FAAD14",fill_type="solid") if ap=="中" else \
                  PatternFill(start_color="52C41A",end_color="52C41A",fill_type="solid")
            ws.cell(row=row,column=10).fill=fill
            row+=1
    wb.save(output)
    output.seek(0)
    return output

# PFMEA主界面
def pfmea_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title("⚡ PFMEA 智能生成系统")
    st.caption("✅ AIAG-VDA FMEA 标准 | IATF16949:2016")
    st.divider()

    product_type = st.radio("产品类型", ["电池包", "充电器"], horizontal=True)
    process_lib = BATTERY_PROCESS_LIB if product_type=="电池包" else CHARGER_PROCESS_LIB
    all_process = list(process_lib.keys())
    if st.session_state.user_knowledge_base:
        all_process = sorted(list(set(all_process + list(st.session_state.user_knowledge_base.keys()))))

    col1, col2 = st.columns([3,1])
    with col1:
        selected_processes = st.multiselect("选择工序", all_process, default=all_process[:1])
    with col2:
        custom = st.text_input("自定义工序")
        if st.button("➕ 添加",use_container_width=True) and custom and custom not in selected_processes:
            selected_processes.append(custom)
            st.rerun()

    # 知识库
    with st.expander("📚 知识库管理", expanded=False):
        file = st.file_uploader("导入PFMEA Excel", type="xlsx")
        if file:
            kb = parse_pfmea_excel(file.getvalue())
            st.session_state.user_knowledge_base = merge_knowledge(kb, st.session_state.user_knowledge_base)
            st.success("导入成功")
            st.rerun()

    # 生成设置
    st.subheader("生成设置")
    gen_mode = st.radio("生成模式", ["本地标准库", "AI智能生成"], horizontal=True)
    scheme_count = st.slider("AI方案数",2,5,3) if gen_mode=="AI智能生成" else 3

    if st.button("🚀 生成PFMEA方案",type="primary",use_container_width=True) and selected_processes:
        st.session_state.generated_pfmea_data = {}
        for proc in selected_processes:
            if gen_mode == "本地标准库":
                items = process_lib.get(proc,[]) + st.session_state.user_knowledge_base.get(proc,[])
                st.session_state.generated_pfmea_data[proc] = [{"方案名称":"本地+知识库","pfmea_list":items}]
            else:
                with st.spinner(f"AI生成【{proc}】..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if not err:
                        st.session_state.generated_pfmea_data[proc] = schemes
                    else:
                        st.error(err)
        st.success("生成完成！")
        st.rerun()

    # 方案选择
    if st.session_state.generated_pfmea_data:
        st.subheader("选择方案")
        final = {}
        for proc in selected_processes:
            if proc not in st.session_state.generated_pfmea_data: continue
            schemes = st.session_state.generated_pfmea_data[proc]
            names = [s["方案名称"] for s in schemes]
            idx = st.radio(f"【{proc}】方案", range(len(names)), format_func=lambda x:names[x])
            sel = schemes[idx]
            st.markdown(f"**{sel['方案名称']}**")
            df = st.data_editor(pd.DataFrame(sel["pfmea_list"]), use_container_width=True)
            final[proc] = df.to_dict("records")

        if st.button("✅ 导出Excel",type="primary",use_container_width=True):
            file = export_pfmea_excel(final, product_type)
            st.download_button("📥 下载PFMEA", file, f"{product_type}_PFMEA.xlsx", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
    if st.button("🏠 返回首页",use_container_width=True):
        st.session_state.current_page="home"
        st.rerun()

# ===================== 主界面 =====================
def main():
    if st.session_state.current_page == "home":
        st.markdown("<div style='text-align:center;padding:3rem 0 2rem;'><h1>🛠️ 多功能智能工具集</h1><p style='color:#666;font-size:18px;'>选择工具开始使用</p></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;padding:32px 20px;'>", unsafe_allow_html=True)
                st.image("https://img.icons8.com/fluency/96/6F9E6F/microsoft-excel-2019.png", width=70)
                st.subheader("Excel 图片工具")
                st.markdown("数字排序 · 批量插入")
                if st.button("进入工具", key="e", use_container_width=True):
                    st.session_state.current_page="excel_image"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
        with col2:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;padding:32px 20px;'>", unsafe_allow_html=True)
                st.image("https://img.icons8.com/fluency/96/6F9E6F/wechat.png", width=70)
                st.subheader("信息推送工具")
                st.markdown("企业微信 · 检测上报")
                if st.button("进入工具", key="p", use_container_width=True):
                    st.session_state.current_page="image_push"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
        with col3:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;padding:32px 20px;'>", unsafe_allow_html=True)
                st.image("https://img.icons8.com/fluency/96/6F9E6F/quality.png", width=70)
                st.subheader("PFMEA 智能生成")
                st.markdown("AI生成 · 标准导出")
                if st.button("进入工具", key="f", use_container_width=True):
                    st.session_state.current_page="pfmea"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
    elif st.session_state.current_page == "excel_image":
        excel_image_tool()
    elif st.session_state.current_page == "image_push":
        image_push_tool()
    elif st.session_state.current_page == "pfmea":
        pfmea_tool()

if __name__ == "__main__":
    main()
