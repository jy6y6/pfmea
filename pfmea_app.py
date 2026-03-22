import streamlit as st
import pandas as pd
import requests
import json
import io
import os
import re
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

# 精美的淡绿色渐变主题CSS
st.markdown("""
<style>
    /* 全局背景渐变 */
    .stApp {
        background: linear-gradient(135deg, #f5f9f0 0%, #e8f0e2 100%);
    }
    /* 卡片样式 */
    .card {
        background-color: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(2px);
        border-radius: 24px;
        padding: 28px;
        margin-bottom: 24px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.08);
        border: 1px solid rgba(127, 183, 126, 0.3);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .card:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 28px rgba(0,0,0,0.12);
    }
    /* 按钮样式 */
    .stButton button {
        background: linear-gradient(90deg, #7fb77e, #6ca06b);
        color: white;
        border-radius: 40px;
        border: none;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
        transition: all 0.2s;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    .stButton button:hover {
        background: linear-gradient(90deg, #6ca06b, #5a8e59);
        transform: scale(1.02);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    /* 输入框样式 */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 16px;
        border: 1px solid #cbd5e0;
        transition: border 0.2s;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #7fb77e;
        box-shadow: 0 0 0 2px rgba(127,183,126,0.2);
    }
    /* 标题样式 */
    h1, h2, h3 {
        color: #3c6e3c;
        font-weight: 600;
    }
    /* 表格样式 */
    .dataframe {
        border-radius: 12px;
        overflow: hidden;
        border-collapse: separate;
        border-spacing: 0;
    }
    /* 滑块样式 */
    .stSlider > div {
        padding-top: 8px;
    }
    /* 图标样式 */
    .icon-emoj {
        font-size: 2rem;
        margin-right: 8px;
        vertical-align: middle;
    }
</style>
""", unsafe_allow_html=True)

# 初始化 session_state
if "current_page" not in st.session_state:
    st.session_state.current_page = "home"
if "push_history" not in st.session_state:
    st.session_state.push_history = []
if "uploaded_images" not in st.session_state:
    st.session_state.uploaded_images = []
if "image_order" not in st.session_state:
    st.session_state.image_order = []
if "user_knowledge_base" not in st.session_state:
    st.session_state.user_knowledge_base = {}          # 工序 -> 失效条目列表
if "generated_pfmea_data" not in st.session_state:
    st.session_state.generated_pfmea_data = {}
if "selected_ai_scheme" not in st.session_state:
    st.session_state.selected_ai_scheme = {}

# ===================== 辅助函数 =====================
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
            payload = {
                "msgtype": "image",
                "image": {"base64": b64, "md5": md5}
            }
            if text_content and idx == 0:
                text_payload = {"msgtype": "text", "text": {"content": text_content}}
                requests.post(webhook_url, json=text_payload, timeout=10)
            response = requests.post(webhook_url, json=payload, timeout=10)
            if response.json().get("errcode") == 0:
                success_count += 1
        except:
            pass
    return success_count

def clean_history_limit(history, max_total=200, keep=100):
    if len(history) > max_total:
        return history[-keep:]
    return history

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

# ===================== 模块一：Excel 图片工具 =====================
def excel_image_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📸 Excel 图片工具")
    st.markdown("将多张图片按顺序插入 Excel 表格指定区域，支持新建或加载现有文件。")

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
            st.info(f"区域共 {total_cells} 个单元格（{rows} 行 × {cols} 列）")
        else:
            st.warning("单元格格式错误，示例：A1")
            total_cells = 0
    except:
        total_cells = 0

    # 图片上传
    uploaded_files = st.file_uploader(
        "从相册选择图片",
        type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True,
        key="img_upload"
    )
    if uploaded_files:
        st.session_state.uploaded_images = [(f.name, f.getvalue()) for f in uploaded_files]
        st.session_state.image_order = list(range(len(st.session_state.uploaded_images)))
        st.success(f"已上传 {len(st.session_state.uploaded_images)} 张图片")

    # 顺序调整：使用滑块重新排序
    if st.session_state.uploaded_images:
        st.subheader("图片顺序调整（拖动滑块改变位置）")
        # 显示预览表格（缩略图）
        if total_cells > 0:
            st.markdown("**实时预览（按当前顺序从左到右、从上到下填充）**")
            # 构建表格数据
            preview_data = []
            for r in range(rows):
                row_cells = []
                for c in range(cols):
                    idx = r * cols + c
                    if idx < len(st.session_state.uploaded_images):
                        img_idx = st.session_state.image_order[idx]
                        img_name = st.session_state.uploaded_images[img_idx][0][:8]
                        row_cells.append(f"{img_name}\n(#{img_idx+1})")
                    else:
                        row_cells.append("空")
                preview_data.append(row_cells)
            preview_df = pd.DataFrame(preview_data, columns=[f"{chr(65+i)}" for i in range(cols)])
            st.table(preview_df)

        # 滑块调整：为每张图片指定目标位置
        st.markdown("**调整每张图片的位置序号（1~{}）**".format(len(st.session_state.uploaded_images)))
        new_order = [0] * len(st.session_state.uploaded_images)
        for idx, (name, _) in enumerate(st.session_state.uploaded_images):
            col_a, col_b = st.columns([1, 3])
            with col_a:
                st.image(io.BytesIO(st.session_state.uploaded_images[idx][1]), width=60, caption=name)
            with col_b:
                # 滑块范围 1 到总图片数，默认当前顺序位置
                current_pos = st.session_state.image_order.index(idx) + 1
                new_pos = st.slider(f"位置", 1, len(st.session_state.uploaded_images), current_pos, key=f"pos_{idx}")
                new_order[new_pos-1] = idx
        # 应用新顺序
        if st.button("✅ 应用新顺序", width="stretch"):
            # 确保所有位置都被分配
            if all(p != 0 for p in new_order):
                st.session_state.image_order = new_order
                st.success("顺序已更新")
                st.rerun()
            else:
                st.error("位置分配有冲突，请确保每个位置都有唯一图片")

    # Excel 来源选择
    excel_source = st.radio("Excel 来源", ["新建空白工作簿", "上传现有 Excel 文件"])
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
    if st.button("🚀 生成并下载 Excel 文件", type="primary", width="stretch"):
        if not st.session_state.uploaded_images:
            st.error("请先上传图片")
        elif not start_cell or not end_cell or total_cells == 0:
            st.error("请填写正确的起始和结束单元格")
        else:
            try:
                # 解析单元格范围
                start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
                end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
                start_col = openpyxl.utils.column_index_from_string(start_match.group(1))
                start_row = int(start_match.group(2))
                end_col = openpyxl.utils.column_index_from_string(end_match.group(1))
                end_row = int(end_match.group(2))

                # 创建工作簿
                if existing_wb:
                    wb = existing_wb
                    ws = wb.active
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "图片表格"

                # 设置行高列宽
                for row in range(start_row, end_row+1):
                    ws.row_dimensions[row].height = 150
                for col in range(start_col, end_col+1):
                    ws.column_dimensions[get_column_letter(col)].width = 15

                # 插入图片
                idx = 0
                for r in range(start_row, end_row+1):
                    for c in range(start_col, end_col+1):
                        if idx >= len(st.session_state.uploaded_images):
                            break
                        img_idx = st.session_state.image_order[idx]
                        img_name, img_bytes = st.session_state.uploaded_images[img_idx]
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
                    width="stretch"
                )
                st.success("Excel 生成完成！")
            except Exception as e:
                st.error(f"生成失败: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 模块二：信息推送工具 =====================
def image_push_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📱 信息推送工具")
    st.markdown("填写检测信息，拍照/选图推送至企业微信群")

    WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=bdf3c7d5-a7fd-4d5a-92bb-4ab15a32e042"

    # 表单
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

    # 历史记录管理
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

        if st.button("📥 导出当前筛选记录为 Excel", width="stretch"):
            if not filtered_df.empty:
                excel_file = export_history_to_excel(filtered_df)
                st.download_button(
                    label="点击下载 Excel",
                    data=excel_file,
                    file_name=f"推送记录_{start_date}_{end_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )
            else:
                st.warning("没有符合条件的记录")
    else:
        st.info("暂无历史记录")
    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 模块三：PFMEA 智能生成工具 =====================
def pfmea_tool():
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_ENDPOINTS = [
        "https://api.doubao.com/v1/chat/completions",
        "https://api.doubaoai.com/v1/chat/completions",
        "https://open.doubao.com/v1/chat/completions"
    ]
    SYSTEM_NAME = "电池包/充电器PFMEA智能生成系统"
    STANDARD = "AIAG-VDA FMEA 第一版 | IATF16949:2016"

    # ---------- 本地标准库（每个工序至少3种失效模式）----------
    BATTERY_PROCESS_LIB = {
        "电芯来料检验": [
            {"失效模式": "电芯外观尺寸超差", "失效后果": "电芯无法装入模组壳体，导致装配中断", "失效原因": "来料尺寸公差不符合图纸要求，量具未定期校准", "预防措施": "制定电芯来料检验规范，每批次抽取样件全尺寸检测，量具定期校准并记录", "探测措施": "首件全尺寸检验，巡检按AQL抽样标准检测，超差件隔离标识", "严重度S": 6, "频度O": 3, "探测度D": 4, "AP等级": "中"},
            {"失效模式": "电芯电压/内阻异常", "失效后果": "模组充放电异常，循环寿命衰减过快，严重时引发热失控风险", "失效原因": "电芯生产过程工艺异常，来料存储环境温湿度不符合要求", "预防措施": "每批次电芯进行电压、内阻全检，存储环境温湿度24小时监控记录", "探测措施": "自动化检测设备100%全检，异常数据自动报警隔离，数据可追溯", "严重度S": 9, "频度O": 2, "探测度D": 2, "AP等级": "高"},
            {"失效模式": "电芯表面划伤/破损", "失效后果": "绝缘性能下降，可能引发短路，安全风险", "失效原因": "来料包装破损，搬运过程中磕碰，作业人员操作不当", "预防措施": "包装标准化，运输防护升级，培训作业人员轻拿轻放", "探测措施": "目视全检，不良品隔离，定期抽检包装防护效果", "严重度S": 7, "频度O": 3, "探测度D": 3, "AP等级": "高"},
        ],
        "模组堆叠装配": [
            {"失效模式": "电芯堆叠顺序错误、极性反向", "失效后果": "模组电路连接错误，充放电功能失效，严重时引发短路烧毁", "失效原因": "作业人员未按SOP操作，防错装置失效，首件检验未执行", "预防措施": "制定极性防错SOP，安装极性视觉防错装置，作业人员岗前培训考核", "探测措施": "首件极性全检，过程中视觉设备100%检测，异常自动停机报警", "严重度S": 9, "频度O": 2, "探测度D": 2, "AP等级": "高"},
            {"失效模式": "电芯之间间距不均", "失效后果": "散热不均，影响模组寿命，可能造成局部过热", "失效原因": "堆叠工装磨损，定位不准，来料尺寸偏差", "预防措施": "定期校准工装，首件确认，SPC监控堆叠精度", "探测措施": "激光测距抽检，过程能力分析", "严重度S": 5, "频度O": 3, "探测度D": 3, "AP等级": "中"},
            {"失效模式": "绝缘片漏装/错装", "失效后果": "短路风险，严重时起火，造成安全事故", "失效原因": "物料清单错误，作业疏忽，防错系统失效", "预防措施": "物料扫码防错，双人复核，装配前视觉检测绝缘片有无", "探测措施": "视觉系统100%检测绝缘片位置，异常自动报警", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
        ],
        "激光焊接": [
            {"失效模式": "焊接熔深不足", "失效后果": "连接强度不足，虚焊导致断路，产品功能失效", "失效原因": "激光功率不稳定，焦距偏移，保护气体流量不当", "预防措施": "每日焊接参数验证，设备定期维护，焊接前清洁板材", "探测措施": "焊接后拉力测试，在线监控焊接能量，SPC控制", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
            {"失效模式": "焊接飞溅", "失效后果": "污染其他部件，可能引起短路，影响外观", "失效原因": "保护气体流量不足，板材表面脏污，焊接参数不当", "预防措施": "清洁板材，优化焊接参数，定期更换保护气体", "探测措施": "目视检查，飞溅残留检测，定期抽检", "严重度S": 6, "频度O": 3, "探测度D": 3, "AP等级": "中"},
            {"失效模式": "焊接位置偏移", "失效后果": "焊接区域未对准，强度不足，导致虚焊", "失效原因": "定位夹具松动，视觉定位误差，来料尺寸偏差", "预防措施": "定期校准夹具，视觉定位自检，来料尺寸管控", "探测措施": "首件全检，过程SPC监控焊接位置，定期校准", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        ]
    }
    CHARGER_PROCESS_LIB = {
        "PCB来料检验": [
            {"失效模式": "PCB板尺寸超差", "失效后果": "PCB无法装入壳体，装配中断", "失效原因": "PCB生产制程偏差，来料检验规范未执行", "预防措施": "制定PCB来料检验规范，每批次首件全尺寸检测，量具定期校准", "探测措施": "首件全尺寸检验，巡检按AQL抽样，超差件隔离返工", "严重度S": 5, "频度O": 3, "探测度D": 4, "AP等级": "中"},
            {"失效模式": "铜箔起泡", "失效后果": "焊接可靠性下降，虚焊，产品功能失效", "失效原因": "PCB受潮，层压工艺不良，存储环境湿度超标", "预防措施": "来料烘烤，存储湿度控制，定期检查存储环境", "探测措施": "外观检查，切片分析，定期抽检", "严重度S": 6, "频度O": 2, "探测度D": 3, "AP等级": "中"},
            {"失效模式": "丝印错误", "失效后果": "元器件贴装错误，导致功能失效或返工", "失效原因": "PCB制板厂丝印工序失误，图纸版本错误", "预防措施": "IQC核对图纸，首件确认，建立丝印样板", "探测措施": "AOI检测丝印内容，首件全检", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        ],
        "SMT贴片焊接": [
            {"失效模式": "元器件贴装偏移", "失效后果": "焊接不良，功能失效，批量返工成本", "失效原因": "贴片机吸嘴磨损，程序坐标偏差，钢网定位不准", "预防措施": "定期校准设备，首件验证，设备维护计划", "探测措施": "AOI全检，SPI锡膏检测，异常报警", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
            {"失效模式": "立碑", "失效后果": "开路，功能失效，产品报废", "失效原因": "回流焊温度曲线不当，焊盘设计不合理，贴装偏移", "预防措施": "优化炉温曲线，PCB焊盘设计DFM评审，贴装精度管控", "探测措施": "AOI检测，X-ray抽查，定期监控炉温", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
            {"失效模式": "少锡/锡珠", "失效后果": "虚焊，短路风险，功能不稳定", "失效原因": "钢网堵塞，刮刀压力不当，锡膏回温不当", "预防措施": "钢网清洗周期，SPI监控，锡膏管理规范", "探测措施": "SPI全检，AOI复检，定期抽查", "严重度S": 6, "频度O": 3, "探测度D": 1, "AP等级": "中"},
        ]
    }

    # ---------- AI 生成函数 ----------
    def create_retry_session():
        session = requests.Session()
        retry = Retry(total=3, backoff_factor=1, status_forcelist=[429,500,502,503,504])
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        return session

    def generate_pfmea_ai(process_name, product_type, scheme_count=3):
        session = create_retry_session()
        headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
        prompt = f"""
        你是专业的汽车电子行业PFMEA工程师，精通AIAG-VDA FMEA标准和IATF16949质量管理体系。
        请针对【{process_name}】工序，生成{scheme_count}组完全不同、无重复内容的PFMEA方案，每组方案包含3-5条独立的PFMEA条目。
        严格遵守以下要求：
        1. 每组方案必须有明显差异化：分别从人、机、料、法、环不同维度切入，失效模式、失效后果、失效原因、预防/探测措施完全不同，禁止内容重复。
        2. 所有内容必须严格贴合{product_type}装配现场的实际作业场景。
        3. 严格遵循AIAG-VDA FMEA标准，失效链必须完整：失效模式→失效后果→失效原因→预防措施→探测措施。
        4. S/O/D评分严格符合AIAG-VDA评分标准：严重度S(1-10)、频度O(1-10)、探测度D(1-10)，AP等级为高/中/低。
        5. 必须返回严格的JSON格式，外层是一个数组，每个元素是一组方案，格式如下：
        [
            {{
                "方案名称": "方案1：人员操作维度管控",
                "pfmea_list": [
                    {{
                        "失效模式": "xxx",
                        "失效后果": "xxx",
                        "失效原因": "xxx",
                        "预防措施": "xxx",
                        "探测措施": "xxx",
                        "严重度S": x,
                        "频度O": x,
                        "探测度D": x,
                        "AP等级": "x"
                    }}
                ]
            }}
        ]
        6. 禁止返回任何JSON以外的内容，禁止注释、解释、markdown格式，确保JSON可直接解析。
        7. 每组方案的pfmea_list必须包含3-5条。
        """
        data = {
            "model": "doubao-pro-32k",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 4000
        }
        for endpoint in API_ENDPOINTS:
            try:
                response = session.post(endpoint, headers=headers, json=data, timeout=60)
                response.raise_for_status()
                content = response.json()["choices"][0]["message"]["content"]
                content = re.sub(r'^```json\s*|\s*```$', '', content.strip())
                return json.loads(content), None
            except Exception as e:
                continue
        return None, "所有API端点均失败"

    # ---------- 知识库管理函数 ----------
    def parse_pfmea_excel(file_bytes):
        """解析上传的PFMEA Excel文件，返回工序->失效条目字典"""
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
            # 智能列名映射
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
                    if any(c in str(col) for c in candidates):
                        col_mapping[target] = col
                        break
            if "工序" not in col_mapping:
                return {}
            knowledge = {}
            for _, row in df.iterrows():
                process = str(row[col_mapping["工序"]]).strip()
                if not process or process == "nan":
                    continue
                item = {}
                for target, col in col_mapping.items():
                    if target != "工序":
                        item[target] = row[col] if pd.notna(row[col]) else ""
                # 确保S/O/D为整数
                for score in ["严重度S", "频度O", "探测度D"]:
                    if score in item:
                        try:
                            item[score] = int(float(item[score]))
                        except:
                            item[score] = 5
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

    def merge_knowledge(base, new_items):
        """合并知识库，去重（基于失效模式+失效原因）"""
        if not new_items:
            return base
        existing_keys = set(f"{item.get('失效模式','')}_{item.get('失效原因','')}" for item in base)
        for item in new_items:
            key = f"{item.get('失效模式','')}_{item.get('失效原因','')}"
            if key not in existing_keys:
                base.append(item)
                existing_keys.add(key)
        return base

    # ---------- Excel导出函数（符合IATF16949审核格式）----------
    def export_pfmea_excel(pfmea_data, product_type):
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "PFMEA"
        # 标题
        ws.merge_cells('A1:K1')
        ws['A1'] = f"过程失效模式及后果分析（PFMEA） - {product_type}"
        ws['A1'].font = Font(name="微软雅黑", size=14, bold=True)
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        # 基础信息行
        ws.merge_cells('A2:K2')
        ws['A2'] = f"项目: {product_type} | 生成日期: {datetime.now().strftime('%Y-%m-%d')} | 版本: 1.0"
        ws['A2'].font = Font(name="微软雅黑", size=10)
        ws['A2'].alignment = Alignment(horizontal="center", vertical="center")
        # 表头
        headers = ["工序", "失效模式", "失效后果", "失效原因", "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"]
        header_font = Font(name="微软雅黑", bold=True, size=10)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=h)
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        # 数据
        row = 5
        for process, items in pfmea_data.items():
            for item in items:
                ws.cell(row=row, column=1, value=process).border = thin_border
                ws.cell(row=row, column=2, value=item.get("失效模式", "")).border = thin_border
                ws.cell(row=row, column=3, value=item.get("失效后果", "")).border = thin_border
                ws.cell(row=row, column=4, value=item.get("失效原因", "")).border = thin_border
                ws.cell(row=row, column=5, value=item.get("预防措施", "")).border = thin_border
                ws.cell(row=row, column=6, value=item.get("探测措施", "")).border = thin_border
                ws.cell(row=row, column=7, value=item.get("严重度S", "")).border = thin_border
                ws.cell(row=row, column=8, value=item.get("频度O", "")).border = thin_border
                ws.cell(row=row, column=9, value=item.get("探测度D", "")).border = thin_border
                ws.cell(row=row, column=10, value=item.get("AP等级", "")).border = thin_border
                # AP等级颜色
                ap = item.get("AP等级", "")
                if ap == "高":
                    ws.cell(row=row, column=10).fill = PatternFill(start_color="FF4D4F", end_color="FF4D4F", fill_type="solid")
                elif ap == "中":
                    ws.cell(row=row, column=10).fill = PatternFill(start_color="FAAD14", end_color="FAAD14", fill_type="solid")
                elif ap == "低":
                    ws.cell(row=row, column=10).fill = PatternFill(start_color="52C41A", end_color="52C41A", fill_type="solid")
                # 对齐
                for col in range(1, 11):
                    ws.cell(row=row, column=col).alignment = Alignment(horizontal="left" if col in [2,3,4,5,6] else "center", vertical="center", wrap_text=True)
                row += 1
        # 列宽
        col_widths = [18, 25, 30, 30, 35, 35, 8, 8, 8, 8]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64+i)].width = width
        # 冻结窗格
        ws.freeze_panes = 'A5'
        wb.save(output)
        output.seek(0)
        return output

    # ---------- 界面 ----------
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title(SYSTEM_NAME)
    st.caption(STANDARD)

    # 产品类型
    product_type = st.radio("产品类型", ["电池包", "充电器"])

    # 工序选择（多选+自定义）
    process_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
    all_process = list(process_lib.keys())
    if st.session_state.user_knowledge_base:
        all_process = list(set(all_process + list(st.session_state.user_knowledge_base.keys())))
    all_process.sort()

    col1, col2 = st.columns([3, 1])
    with col1:
        selected_processes = st.multiselect("选择工序（可多选）", all_process, default=all_process[:2] if all_process else [])
    with col2:
        custom_process = st.text_input("自定义工序名称")
        if st.button("➕ 添加自定义工序", width="stretch"):
            if custom_process and custom_process not in selected_processes:
                selected_processes.append(custom_process)
                st.success(f"已添加工序：{custom_process}")
                st.rerun()

    # 知识库管理区域
    with st.expander("📚 知识库管理", expanded=False):
        st.markdown("### 导入已有PFMEA文件（.xlsx）")
        uploaded_kb = st.file_uploader("选择文件", type=["xlsx"], key="kb_upload")
        if uploaded_kb:
            with st.spinner("正在解析..."):
                kb_data = parse_pfmea_excel(uploaded_kb.getvalue())
                if kb_data:
                    for proc, items in kb_data.items():
                        if proc in st.session_state.user_knowledge_base:
                            st.session_state.user_knowledge_base[proc] = merge_knowledge(st.session_state.user_knowledge_base[proc], items)
                        else:
                            st.session_state.user_knowledge_base[proc] = items
                    st.success(f"成功导入 {len(kb_data)} 个工序的知识库条目！")
                    st.rerun()

        st.markdown("### 当前知识库内容")
        if st.session_state.user_knowledge_base:
            for proc, items in st.session_state.user_knowledge_base.items():
                with st.expander(f"📦 工序：{proc}（共{len(items)}条）", expanded=False):
                    df_proc = pd.DataFrame(items)
                    df_proc = reset_df_index(df_proc)
                    edited_df = st.data_editor(df_proc, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"kb_edit_{proc}")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"💾 保存修改", key=f"kb_save_{proc}", width="stretch"):
                            st.session_state.user_knowledge_base[proc] = edited_df.to_dict("records")
                            st.success("已保存")
                            st.rerun()
                    with col2:
                        if st.button(f"🗑️ 删除此工序", key=f"kb_del_{proc}", width="stretch"):
                            del st.session_state.user_knowledge_base[proc]
                            st.success("已删除")
                            st.rerun()
            # 备份导出和恢复导入
            col1, col2 = st.columns(2)
            with col1:
                kb_json = json.dumps(st.session_state.user_knowledge_base, ensure_ascii=False, indent=2)
                st.download_button(
                    label="📤 导出知识库备份",
                    data=kb_json,
                    file_name=f"PFMEA知识库_{datetime.now().strftime('%Y%m%d%H%M%S')}.json",
                    mime="application/json",
                    width="stretch"
                )
            with col2:
                import_file = st.file_uploader("导入备份文件", type=["json"], key="kb_import")
                if import_file:
                    try:
                        import_data = json.load(import_file)
                        for proc, items in import_data.items():
                            if proc in st.session_state.user_knowledge_base:
                                st.session_state.user_knowledge_base[proc] = merge_knowledge(st.session_state.user_knowledge_base[proc], items)
                            else:
                                st.session_state.user_knowledge_base[proc] = items
                        st.success("导入成功")
                        st.rerun()
                    except Exception as e:
                        st.error(f"导入失败: {e}")
        else:
            st.info("知识库为空，请导入旧PFMEA文件或等待生成后自动积累。")

    # 生成模式
    gen_mode = st.radio("生成模式", ["本地标准库", "AI智能生成"], horizontal=True)
    if gen_mode == "AI智能生成":
        scheme_count = st.slider("生成方案数量", min_value=2, max_value=5, value=3, step=1)
        mix_knowledge = st.checkbox("混合知识库内容生成", value=False)

    # 生成按钮
    if st.button("🚀 生成PFMEA方案", type="primary", width="stretch") and selected_processes:
        st.session_state.generated_pfmea_data = {}
        st.session_state.selected_ai_scheme = {}
        for proc in selected_processes:
            if gen_mode == "本地标准库":
                # 合并标准库+知识库
                lib_items = process_lib.get(proc, [])
                user_items = st.session_state.user_knowledge_base.get(proc, [])
                combined = merge_knowledge(lib_items[:], user_items)  # 去重合并
                if not combined:
                    st.warning(f"工序【{proc}】无本地数据，请使用AI生成")
                    continue
                st.session_state.generated_pfmea_data[proc] = combined
            else:
                # AI生成多组方案
                with st.spinner(f"AI正在生成【{proc}】的{scheme_count}组方案..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if err:
                        st.error(f"{proc} AI生成失败: {err}")
                        continue
                    # 如果混合知识库，添加一个知识库方案
                    if mix_knowledge and proc in st.session_state.user_knowledge_base:
                        user_scheme = {
                            "方案名称": "知识库方案（本地已有）",
                            "pfmea_list": st.session_state.user_knowledge_base[proc]
                        }
                        schemes.append(user_scheme)
                    st.session_state.generated_pfmea_data[proc] = schemes
                    st.session_state.selected_ai_scheme[proc] = 0
        if st.session_state.generated_pfmea_data:
            st.success("生成完成！请选择方案")
            st.rerun()

    # 方案选择与预览
    if st.session_state.generated_pfmea_data:
        st.markdown("### 选择最终方案")
        final_data = {}
        for proc in selected_processes:
            if proc not in st.session_state.generated_pfmea_data:
                continue
            data = st.session_state.generated_pfmea_data[proc]
            if gen_mode == "本地标准库":
                # 本地库直接展示所有条目（可编辑）
                with st.expander(f"工序：{proc}", expanded=True):
                    df_proc = pd.DataFrame(data)
                    df_proc = reset_df_index(df_proc)
                    edited = st.data_editor(df_proc, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"edit_{proc}")
                    final_data[proc] = edited.to_dict("records")
            else:
                # AI模式，显示方案卡片
                schemes = data
                scheme_names = [s.get("方案名称", f"方案{i+1}") for i, s in enumerate(schemes)]
                selected_idx = st.radio(f"为【{proc}】选择方案", options=range(len(scheme_names)), format_func=lambda i: scheme_names[i], key=f"select_{proc}", index=st.session_state.selected_ai_scheme.get(proc, 0))
                st.session_state.selected_ai_scheme[proc] = selected_idx
                selected_scheme = schemes[selected_idx]
                st.markdown("**方案预览**")
                df_scheme = pd.DataFrame(selected_scheme["pfmea_list"])
                st.dataframe(df_scheme, use_container_width=True)
                # 允许编辑
                edited_df = st.data_editor(df_scheme, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"edit_ai_{proc}")
                final_data[proc] = edited_df.to_dict("records")

        # 确认生成最终Excel
        if st.button("✅ 确认使用当前方案，导出Excel", type="primary", width="stretch"):
            excel_file = export_pfmea_excel(final_data, product_type)
            st.download_button(
                label="📥 下载 PFMEA Excel 文件",
                data=excel_file,
                file_name=f"{product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 主界面 =====================
def main():
    if st.session_state.current_page == "home":
        st.markdown("<div style='text-align: center; padding: 2rem 0;'><h1>🛠️ 多功能智能工具集</h1><p>请选择要使用的工具模块</p></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        with col1:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("<div style='font-size: 3rem;'>📸</div>", unsafe_allow_html=True)
                st.subheader("Excel 图片工具")
                st.markdown("将多张图片按顺序插入 Excel 表格")
                if st.button("进入工具", key="btn_excel", width="stretch"):
                    st.session_state.current_page = "excel_image"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("<div style='font-size: 3rem;'>📱</div>", unsafe_allow_html=True)
                st.subheader("信息推送工具")
                st.markdown("拍照/选图推送至企业微信群")
                if st.button("进入工具", key="btn_push", width="stretch"):
                    st.session_state.current_page = "image_push"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col3:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.markdown("<div style='font-size: 3rem;'>⚡</div>", unsafe_allow_html=True)
                st.subheader("PFMEA 智能生成")
                st.markdown("符合 AIAG-VDA 标准的 FMEA 生成")
                if st.button("进入工具", key="btn_pfmea", width="stretch"):
                    st.session_state.current_page = "pfmea"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

    elif st.session_state.current_page == "excel_image":
        excel_image_tool()
        if st.button("🏠 返回首页", width="stretch"):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "image_push":
        image_push_tool()
        if st.button("🏠 返回首页", width="stretch"):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "pfmea":
        pfmea_tool()
        if st.button("🏠 返回首页", width="stretch"):
            st.session_state.current_page = "home"
            st.rerun()

if __name__ == "__main__":
    main()
