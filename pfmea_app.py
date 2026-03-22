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

# 自定义CSS（淡绿色主题）
st.markdown("""
<style>
    /* 全局背景 */
    .stApp {
        background-color: #f0f7e8;
    }
    /* 卡片样式 */
    .card {
        background-color: white;
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid #d4e2c1;
        transition: all 0.3s;
    }
    .card:hover {
        box-shadow: 0 8px 24px rgba(0,0,0,0.08);
    }
    /* 按钮样式 */
    .stButton button {
        background-color: #6f9e6f;
        color: white;
        border-radius: 10px;
        border: none;
        padding: 0.6rem 1.2rem;
        font-weight: 500;
        transition: 0.2s;
    }
    .stButton button:hover {
        background-color: #5a805a;
        color: white;
    }
    /* 输入框样式 */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 10px;
        border: 1px solid #cce0b5;
    }
    /* 标题样式 */
    h1, h2, h3 {
        color: #3a6b3a;
        font-weight: 600;
    }
    /* 表格样式 */
    .dataframe {
        border-radius: 12px;
        overflow: hidden;
        font-size: 14px;
    }
    /* 缩略图网格 */
    .preview-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(100px, 1fr));
        gap: 10px;
        margin-top: 10px;
    }
    .preview-cell {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 8px;
        text-align: center;
        background: #f9f9f9;
    }
    .preview-cell img {
        max-width: 80px;
        max-height: 80px;
        object-fit: contain;
    }
</style>
""", unsafe_allow_html=True)

# 初始化 session_state
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
        st.session_state.user_knowledge_base = {}      # 格式：{工序名称: [失效条目列表]}
    if "generated_pfmea_data" not in st.session_state:
        st.session_state.generated_pfmea_data = {}
    if "selected_ai_scheme" not in st.session_state:
        st.session_state.selected_ai_scheme = {}
    if "temp_kb_import" not in st.session_state:
        st.session_state.temp_kb_import = None

init_session()

# ===================== 通用辅助函数 =====================
def compress_image_to_limit(image_bytes, max_size_mb=2, max_side=1024):
    """压缩图片到指定大小以下"""
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
    """发送图片（企业微信一次一张，循环发送）"""
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

    # 预览布局按钮（带缩略图）
    if st.session_state.uploaded_images and total_cells > 0:
        if st.button("🔍 预览布局（显示缩略图）", use_container_width=True):
            st.session_state.show_preview = True
        else:
            st.session_state.show_preview = False

        if st.session_state.get("show_preview", False):
            st.subheader("📋 实时预览（缩略图 + 位置调整）")
            # 计算需要显示的网格
            preview_rows = rows
            preview_cols = cols
            # 创建一个容器展示网格
            for r in range(preview_rows):
                cols_layout = st.columns(preview_cols)
                for c in range(preview_cols):
                    idx = r * preview_cols + c
                    if idx < len(st.session_state.uploaded_images):
                        img_idx = st.session_state.image_order[idx]
                        img_name, img_bytes = st.session_state.uploaded_images[img_idx]
                        with cols_layout[c]:
                            # 显示缩略图
                            st.image(io.BytesIO(img_bytes), width=80, caption=img_name[:10])
                            # 提供位置输入框
                            new_pos = st.number_input(
                                f"位置", min_value=0, max_value=len(st.session_state.uploaded_images)-1,
                                value=idx, key=f"pos_{r}_{c}", step=1
                            )
                            if new_pos != idx:
                                # 重新排序
                                order = st.session_state.image_order
                                # 找到当前图片的索引
                                current_idx = order.index(img_idx)
                                # 移除并插入新位置
                                order.pop(current_idx)
                                order.insert(new_pos, img_idx)
                                st.session_state.image_order = order
                                st.rerun()
                    else:
                        with cols_layout[c]:
                            st.markdown("空")
            # 显示当前顺序摘要
            order_display = []
            for idx in st.session_state.image_order:
                order_display.append(st.session_state.uploaded_images[idx][0][:20])
            st.write("**当前顺序：** " + " → ".join(order_display))

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
    if st.button("🚀 生成并下载 Excel 文件", type="primary", use_container_width=True):
        if not st.session_state.uploaded_images:
            st.error("请先上传图片")
        elif not start_cell or not end_cell or total_cells == 0:
            st.error("请填写正确的起始和结束单元格")
        else:
            try:
                # 解析单元格
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
                    use_container_width=True
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

        # 导出
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

# ===================== 模块三：PFMEA 智能生成工具 =====================
# ------------------------- 本地标准库（每个工序至少3组方案，扩充工序）-------------------------
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
    ],
    "激光焊接": [
        {"失效模式": "焊接熔深不足", "失效后果": "连接强度不足，虚焊导致断路", "失效原因": "激光功率不稳定，焦距偏移", "预防措施": "每日焊接参数验证，设备定期维护", "探测措施": "焊接后拉力测试，在线监控", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "焊接飞溅", "失效后果": "污染其他部件，可能引起短路", "失效原因": "保护气体流量不足，板材表面脏污", "预防措施": "清洁板材，优化焊接参数", "探测措施": "目视检查，飞溅残留检测", "严重度S": 6, "频度O": 3, "探测度D": 3, "AP等级": "中"},
        {"失效模式": "焊接位置偏移", "失效后果": "焊接区域未对准，强度不足", "失效原因": "定位夹具松动，视觉定位误差", "预防措施": "定期校准夹具，视觉定位自检", "探测措施": "首件全检，过程SPC监控", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
    ],
    "BMS装配": [
        {"失效模式": "BMS板装配位置偏移", "失效后果": "信号采集异常，BMS通讯故障", "失效原因": "定位工装磨损，装配手法不当", "预防措施": "使用定位治具，首件确认", "探测措施": "视觉系统检测位置，电测验证", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "线束接插件插接不到位", "失效后果": "接触不良，信号中断", "失效原因": "作业人员未插到位，防错缺失", "预防措施": "安装插接防错装置，培训", "探测措施": "自动插拔力检测，功能测试", "严重度S": 8, "频度O": 2, "探测度D": 1, "AP等级": "高"},
        {"失效模式": "BMS固件烧录错误", "失效后果": "BMS无法正常工作，功能失效", "失效原因": "烧录程序版本错误，烧录工装接触不良", "预防措施": "扫码自动匹配程序，定期维护烧录座", "探测措施": "烧录后自检，功能测试", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
    ],
    "密封测试": [
        {"失效模式": "密封胶涂胶不均匀", "失效后果": "防水性能下降，IP等级不达标", "失效原因": "胶阀堵塞，轨迹参数偏差", "预防措施": "每日清洗胶阀，定期校准轨迹", "探测措施": "视觉检测胶宽，气密测试", "严重度S": 7, "频度O": 3, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "壳体螺丝紧固扭矩不足", "失效后果": "松动漏水，连接失效", "失效原因": "扭矩枪未校准，漏打螺丝", "预防措施": "扭矩枪每日点检，防错计数", "探测措施": "扭矩抽检，气密测试", "严重度S": 6, "频度O": 2, "探测度D": 2, "AP等级": "中"},
        {"失效模式": "气密测试泄漏", "失效后果": "防水失效，内部器件损坏", "失效原因": "密封圈破损，壳体变形", "预防措施": "来料密封圈检验，壳体尺寸监控", "探测措施": "气密测试仪100%检测，泄漏定位", "严重度S": 8, "频度O": 2, "探测度D": 1, "AP等级": "高"},
    ],
    "老化测试": [
        {"失效模式": "老化过程中通讯中断", "失效后果": "产品功能不稳定，客户投诉", "失效原因": "线束松动，BMS软件bug", "预防措施": "老化前插接确认，软件版本管控", "探测措施": "在线监控通讯状态，报警", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "充放电循环异常", "失效后果": "容量不达标，寿命短", "失效原因": "电芯一致性差，BMS保护参数错误", "预防措施": "电芯分选配组，BMS参数验证", "探测措施": "充放电设备监控，数据记录分析", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
        {"失效模式": "温度监控失效", "失效后果": "过温未保护，热失控风险", "失效原因": "温度传感器故障，线束接触不良", "预防措施": "传感器来料检验，插接防错", "探测措施": "老化过程温度曲线监控，报警", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
    ]
}
CHARGER_PROCESS_LIB = {
    "PCB来料检验": [
        {"失效模式": "PCB板尺寸超差", "失效后果": "PCB无法装入壳体", "失效原因": "PCB生产制程偏差", "预防措施": "制定PCB来料检验规范，首件全检", "探测措施": "首件全尺寸检验，巡检抽检", "严重度S": 5, "频度O": 3, "探测度D": 4, "AP等级": "中"},
        {"失效模式": "铜箔起泡", "失效后果": "焊接可靠性下降，虚焊", "失效原因": "PCB受潮，层压工艺不良", "预防措施": "来料烘烤，存储湿度控制", "探测措施": "外观检查，切片分析", "严重度S": 6, "频度O": 2, "探测度D": 3, "AP等级": "中"},
        {"失效模式": "丝印错误", "失效后果": "元器件贴装错误", "失效原因": "PCB制板厂丝印工序失误", "预防措施": "IQC核对图纸，首件确认", "探测措施": "AOI检测丝印内容", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
    ],
    "SMT贴片焊接": [
        {"失效模式": "元器件贴装偏移", "失效后果": "焊接不良，功能失效", "失效原因": "贴片机吸嘴磨损，程序坐标偏差", "预防措施": "定期校准设备，首件验证", "探测措施": "AOI全检，SPI锡膏检测", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "立碑", "失效后果": "开路，功能失效", "失效原因": "回流焊温度曲线不当，焊盘设计不合理", "预防措施": "优化炉温曲线，PCB焊盘设计DFM评审", "探测措施": "AOI检测，X-ray抽查", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "少锡/锡珠", "失效后果": "虚焊，短路风险", "失效原因": "钢网堵塞，刮刀压力不当", "预防措施": "钢网清洗周期，SPI监控", "探测措施": "SPI全检，AOI复检", "严重度S": 6, "频度O": 3, "探测度D": 1, "AP等级": "中"},
    ],
    "插件后焊": [
        {"失效模式": "插件极性反向", "失效后果": "电路功能异常，烧毁", "失效原因": "作业人员插反，防错缺失", "预防措施": "极性标识清晰，防错工装", "探测措施": "AOI检测，电测验证", "严重度S": 9, "频度O": 2, "探测度D": 1, "AP等级": "高"},
        {"失效模式": "焊点虚焊/连锡", "失效后果": "功能失效，短路", "失效原因": "烙铁温度不当，助焊剂残留", "预防措施": "定期校准烙铁，作业指导", "探测措施": "AOI检测，ICT测试", "严重度S": 7, "频度O": 3, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "元件漏插", "失效后果": "功能缺失", "失效原因": "物料漏放，作业疏忽", "预防措施": "物料清单核对，首件确认", "探测措施": "AOI检测，功能测试", "严重度S": 8, "频度O": 2, "探测度D": 1, "AP等级": "高"},
    ],
    "功能测试": [
        {"失效模式": "测试程序未正确加载", "失效后果": "测试结果误判", "失效原因": "程序版本错误，上传失败", "预防措施": "扫码自动匹配程序，版本管控", "探测措施": "自检程序验证", "严重度S": 6, "频度O": 2, "探测度D": 2, "AP等级": "中"},
        {"失效模式": "测试探针接触不良", "失效后果": "误判为不良品", "失效原因": "探针磨损，氧化", "预防措施": "定期更换探针，清洁", "探测措施": "标准板校验", "严重度S": 5, "频度O": 3, "探测度D": 3, "AP等级": "中"},
        {"失效模式": "测试参数设置错误", "失效后果": "不良品流出", "失效原因": "操作员误改参数", "预防措施": "权限管理，参数锁定", "探测措施": "首件测试验证", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
    ],
    "老化测试": [
        {"失效模式": "老化过程中无输出", "失效后果": "功能失效", "失效原因": "内部元器件损坏，焊接不良", "预防措施": "老化前功能测试，老化架连接确认", "探测措施": "在线监控输出，报警", "严重度S": 8, "频度O": 2, "探测度D": 2, "AP等级": "高"},
        {"失效模式": "老化温度过高", "失效后果": "器件寿命缩短", "失效原因": "散热风扇故障，环境温度高", "预防措施": "定期维护老化架，温度监控", "探测措施": "温度传感器实时监控，超温报警", "严重度S": 7, "频度O": 2, "探测度D": 1, "AP等级": "高"},
        {"失效模式": "老化时间不足", "失效后果": "早期失效未暴露", "失效原因": "人为提前下架", "预防措施": "自动计时，防错设计", "探测措施": "系统记录老化时间，超时报警", "严重度S": 6, "频度O": 2, "探测度D": 2, "AP等级": "中"},
    ]
}

# ------------------------- AI 生成函数（增强版，确保能连通）-------------------------
def create_retry_session():
    session = requests.Session()
    retry = Retry(total=3, backoff_factor=1, status_forcelist=[429,500,502,503,504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    return session

def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    """调用豆包API生成多组PFMEA方案，返回列表格式： [{"方案名称":..., "pfmea_list":[...]}, ...]"""
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    # 尝试多个端点
    API_ENDPOINTS = [
        "https://api.doubao.com/v1/chat/completions",
        "https://api.doubaoai.com/v1/chat/completions",
        "https://open.doubao.com/v1/chat/completions"
    ]
    # 尝试多个模型名称
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
    last_error = None
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
                    continue
                content = result["choices"][0]["message"]["content"]
                # 清理可能的 markdown 代码块
                content = re.sub(r'^```json\s*|\s*```$', '', content.strip())
                parsed = json.loads(content)
                # 验证格式
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

# ------------------------- 知识库管理函数 -------------------------
def parse_pfmea_excel(file_bytes):
    """解析上传的PFMEA Excel文件，返回工序->失效条目字典"""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        # 列名映射
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
            # 自动补充AP等级
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
    """合并知识库，根据失效模式+失效原因去重"""
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

# ------------------------- PFMEA 导出函数 -------------------------
def export_pfmea_excel(pfmea_data, product_type):
    """导出符合格式的Excel"""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "PFMEA"
    # 表头
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
            # AP等级标色
            ap = item.get("AP等级", "")
            fill = None
            if ap == "高":
                fill = PatternFill(start_color="FF4D4F", end_color="FF4D4F", fill_type="solid")
            elif ap == "中":
                fill = PatternFill(start_color="FAAD14", end_color="FAAD14", fill_type="solid")
            elif ap == "低":
                fill = PatternFill(start_color="52C41A", end_color="52C41A", fill_type="solid")
            if fill:
                ws.cell(row=row, column=10).fill = fill
            row += 1
    # 列宽
    col_widths = [18, 25, 30, 30, 35, 35, 8, 8, 8, 8]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[chr(64+i)].width = width
    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output

# ------------------------- PFMEA 主界面 -------------------------
def pfmea_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title("⚡ PFMEA 智能生成系统")
    st.caption("符合 AIAG-VDA FMEA 第一版 | IATF16949:2016")
    st.divider()

    # 产品类型
    product_type = st.radio("产品类型", ["电池包", "充电器"], horizontal=True)

    # ---------- 工序选择 ----------
    process_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
    all_process = list(process_lib.keys())
    # 合并知识库中的工序
    if st.session_state.user_knowledge_base:
        all_process = list(set(all_process + list(st.session_state.user_knowledge_base.keys())))
    all_process.sort()

    col1, col2 = st.columns([3, 1])
    with col1:
        selected_processes = st.multiselect("选择工序（可多选）", all_process, default=all_process[:2] if all_process else [])
    with col2:
        custom_process = st.text_input("自定义工序名称")
        if st.button("➕ 添加自定义工序", use_container_width=True):
            if custom_process and custom_process not in selected_processes:
                selected_processes.append(custom_process)
                st.success(f"已添加工序：{custom_process}")
                st.rerun()

    # ---------- 知识库管理 ----------
    with st.expander("📚 知识库管理（可导入/编辑/导出）"):
        # 导入旧PFMEA文件
        uploaded_kb_file = st.file_uploader("导入已有 PFMEA Excel 文件（.xlsx）", type=["xlsx"], key="kb_import")
        if uploaded_kb_file:
            with st.spinner("正在解析文件..."):
                kb_data = parse_pfmea_excel(uploaded_kb_file.getvalue())
                if kb_data:
                    st.session_state.user_knowledge_base = merge_knowledge(kb_data, st.session_state.user_knowledge_base)
                    st.success(f"成功导入 {len(kb_data)} 个工序的条目！")
                    st.rerun()
        # 显示知识库内容
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
            # 导出/导入备份
            col1, col2 = st.columns(2)
            with col1:
                kb_json = export_knowledge_to_json()
                st.download_button("📤 导出知识库备份（JSON）", data=kb_json, file_name=f"PFMEA知识库_{datetime.now().strftime('%Y%m%d')}.json", mime="application/json", use_container_width=True)
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
            st.info("暂无知识库内容，可通过上传旧PFMEA文件或AI生成后自动入库来积累。")

    # ---------- 生成设置 ----------
    st.subheader("生成设置")
    gen_mode = st.radio("生成模式", ["本地标准库（含知识库）", "AI智能生成（多方案）"], horizontal=True)
    scheme_count = 3
    mix_knowledge = False
    if gen_mode == "AI智能生成（多方案）":
        scheme_count = st.slider("AI生成方案数量", 2, 5, 3)
        mix_knowledge = st.checkbox("混合知识库内容作为独立方案", value=True, help="勾选后，知识库中该工序的内容将作为一个额外方案供选择")

    # 测试AI连接按钮
    if st.button("🔌 测试AI连接", use_container_width=True):
        with st.spinner("正在测试AI连接..."):
            test_result, err = generate_pfmea_ai("电芯来料检验", "电池包", 1)
            if test_result:
                st.success("AI连接正常，可以生成PFMEA方案！")
            else:
                st.error(f"AI连接失败: {err}")

    # ---------- 生成按钮 ----------
    if st.button("🚀 生成PFMEA方案", type="primary", use_container_width=True) and selected_processes:
        st.session_state.generated_pfmea_data = {}
        st.session_state.selected_ai_scheme = {}
        progress_bar = st.progress(0)
        for idx, proc in enumerate(selected_processes):
            progress_bar.progress((idx) / len(selected_processes))
            if gen_mode == "本地标准库（含知识库）":
                # 合并标准库和知识库
                lib_items = process_lib.get(proc, [])
                kb_items = st.session_state.user_knowledge_base.get(proc, [])
                combined = lib_items + kb_items
                if not combined:
                    st.warning(f"工序【{proc}】无本地数据，请使用AI生成或先导入知识库")
                    continue
                # 将所有条目作为一个方案
                st.session_state.generated_pfmea_data[proc] = [{"方案名称": "本地库+知识库方案", "pfmea_list": combined}]
            else:
                with st.spinner(f"AI正在生成【{proc}】的方案..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if err:
                        st.error(f"{proc} AI生成失败: {err}")
                        # 如果失败，可以尝试用本地库作为备用
                        lib_items = process_lib.get(proc, [])
                        if lib_items:
                            st.warning(f"已使用本地库作为备用方案")
                            st.session_state.generated_pfmea_data[proc] = [{"方案名称": "本地库备用方案", "pfmea_list": lib_items}]
                        else:
                            continue
                    else:
                        # 如果需要混合知识库
                        if mix_knowledge and proc in st.session_state.user_knowledge_base:
                            kb_scheme = {"方案名称": "📁 我的知识库方案", "pfmea_list": st.session_state.user_knowledge_base[proc]}
                            schemes.append(kb_scheme)
                        st.session_state.generated_pfmea_data[proc] = schemes
                        st.session_state.selected_ai_scheme[proc] = 0  # 默认选中第一个方案
        progress_bar.progress(1.0)
        if st.session_state.generated_pfmea_data:
            st.success("生成完成！请选择方案")
            st.rerun()

    # ---------- 方案选择与预览 ----------
    if st.session_state.generated_pfmea_data:
        st.subheader("选择最终方案")
        final_data = {}
        for proc in selected_processes:
            if proc not in st.session_state.generated_pfmea_data:
                continue
            data = st.session_state.generated_pfmea_data[proc]
            if len(data) == 1:
                # 只有一个方案，直接使用
                selected_scheme = data[0]
            else:
                # 多个方案，提供单选框
                scheme_names = [s["方案名称"] for s in data]
                selected_idx = st.radio(f"为【{proc}】选择方案", options=range(len(scheme_names)), format_func=lambda i: scheme_names[i], key=f"select_{proc}", index=st.session_state.selected_ai_scheme.get(proc, 0))
                st.session_state.selected_ai_scheme[proc] = selected_idx
                selected_scheme = data[selected_idx]
            st.markdown(f"**{selected_scheme['方案名称']}**")
            df_scheme = pd.DataFrame(selected_scheme["pfmea_list"])
            df_scheme = reset_df_index(df_scheme)
            # 可编辑预览
            edited_df = st.data_editor(df_scheme, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"edit_{proc}")
            final_data[proc] = edited_df.to_dict("records")
            # 可选：将用户编辑后的内容保存回知识库
            if st.button(f"💾 将当前方案存入知识库", key=f"save_kb_{proc}"):
                # 去重合并
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

        # 导出最终Excel
        if st.button("✅ 确认并导出 Excel 文件", type="primary", use_container_width=True):
            excel_file = export_pfmea_excel(final_data, product_type)
            st.download_button(
                label="📥 下载 PFMEA Excel 文件",
                data=excel_file,
                file_name=f"{product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
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
                st.image("https://img.icons8.com/fluency/96/000000/microsoft-excel-2019.png", width=60)
                st.subheader("Excel 图片工具")
                st.markdown("将多张图片按顺序插入 Excel 表格")
                if st.button("进入工具", key="btn_excel", use_container_width=True):
                    st.session_state.current_page = "excel_image"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.image("https://img.icons8.com/fluency/96/000000/wechat.png", width=60)
                st.subheader("信息推送工具")
                st.markdown("拍照/选图推送至企业微信群")
                if st.button("进入工具", key="btn_push", use_container_width=True):
                    st.session_state.current_page = "image_push"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with col3:
            with st.container():
                st.markdown("<div class='card' style='text-align: center;'>", unsafe_allow_html=True)
                st.image("https://img.icons8.com/fluency/96/000000/quality.png", width=60)
                st.subheader("PFMEA 智能生成")
                st.markdown("符合 AIAG-VDA 标准的 FMEA 生成")
                if st.button("进入工具", key="btn_pfmea", use_container_width=True):
                    st.session_state.current_page = "pfmea"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

    elif st.session_state.current_page == "excel_image":
        excel_image_tool()
        if st.button("🏠 返回首页", use_container_width=True):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "image_push":
        image_push_tool()
        if st.button("🏠 返回首页", use_container_width=True):
            st.session_state.current_page = "home"
            st.rerun()
    elif st.session_state.current_page == "pfmea":
        pfmea_tool()
        if st.button("🏠 返回首页", use_container_width=True):
            st.session_state.current_page = "home"
            st.rerun()

if __name__ == "__main__":
    main()
