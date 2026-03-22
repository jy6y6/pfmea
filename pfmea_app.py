import streamlit as st
import pandas as pd
import requests
import json
import io
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
from openpyxl.styles import Font, Alignment, Border, Side

# ===================== 全局配置 =====================
st.set_page_config(
    page_title="多功能智能工具集",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 清新绿色主题CSS
st.markdown("""
<style>
    .stApp {
        background-color: #f5f9f0;
    }
    .card {
        background-color: white;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: 1px solid #d9e8c5;
    }
    .stButton button {
        background-color: #7fb77e;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    .stButton button:hover {
        background-color: #6ca06b;
        color: white;
    }
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 8px;
        border: 1px solid #cbd5e0;
    }
    h1, h2, h3 {
        color: #3c6e3c;
    }
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
    }
    .image-card {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 8px;
        text-align: center;
        background: #fafafa;
    }
    .image-card img {
        max-height: 80px;
        margin-bottom: 8px;
    }
    .image-card button {
        margin-top: 4px;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

# 初始化 session_state
if "current_page" not in st.session_state:
    st.session_state.current_page = "home"
if "push_history" not in st.session_state:
    st.session_state.push_history = []
if "uploaded_images" not in st.session_state:
    st.session_state.uploaded_images = []          # [(name, bytes)]
if "image_order" not in st.session_state:
    st.session_state.image_order = []              # 索引顺序
if "user_knowledge_base" not in st.session_state:
    st.session_state.user_knowledge_base = {}      # 工序 -> list of dict
if "generated_pfmea_data" not in st.session_state:
    st.session_state.generated_pfmea_data = {}     # 工序 -> 最终选中的条目
if "ai_schemes_temp" not in st.session_state:
    st.session_state.ai_schemes_temp = {}          # 临时存储AI生成的方案
if "selected_ai_scheme_idx" not in st.session_state:
    st.session_state.selected_ai_scheme_idx = {}   # 工序 -> 选中的方案索引

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
    success = 0
    for i, img_bytes in enumerate(image_bytes_list):
        try:
            compressed = compress_image_to_limit(img_bytes)
            b64 = base64.b64encode(compressed).decode('utf-8')
            md5 = hashlib.md5(compressed).hexdigest()
            payload = {
                "msgtype": "image",
                "image": {"base64": b64, "md5": md5}
            }
            if text_content and i == 0:
                requests.post(webhook_url, json={"msgtype":"text","text":{"content":text_content}}, timeout=10)
            resp = requests.post(webhook_url, json=payload, timeout=10)
            if resp.json().get("errcode") == 0:
                success += 1
        except:
            pass
    return success

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

    # 可视化顺序调整（卡片+按钮）
    if st.session_state.uploaded_images:
        st.subheader("图片顺序调整")
        st.markdown("点击【上移】【下移】或【交换】调整图片位置。下方的预览表格会实时更新。")

        # 显示所有图片卡片，每个卡片带操作按钮
        num_images = len(st.session_state.uploaded_images)
        cols_per_row = min(4, num_images)
        for i in range(0, num_images, cols_per_row):
            cols = st.columns(cols_per_row)
            for j in range(cols_per_row):
                idx = i + j
                if idx >= num_images:
                    break
                with cols[j]:
                    img_name, img_bytes = st.session_state.uploaded_images[idx]
                    st.image(io.BytesIO(img_bytes), width=100)
                    st.caption(f"{idx+1}. {img_name[:12]}")
                    # 操作按钮
                    col_btn1, col_btn2, col_btn3 = st.columns(3)
                    with col_btn1:
                        if idx > 0 and st.button("⬆️", key=f"up_{idx}", use_container_width=True):
                            st.session_state.image_order[idx], st.session_state.image_order[idx-1] = \
                                st.session_state.image_order[idx-1], st.session_state.image_order[idx]
                            st.rerun()
                    with col_btn2:
                        if idx < num_images-1 and st.button("⬇️", key=f"down_{idx}", use_container_width=True):
                            st.session_state.image_order[idx], st.session_state.image_order[idx+1] = \
                                st.session_state.image_order[idx+1], st.session_state.image_order[idx]
                            st.rerun()
                    with col_btn3:
                        if st.button("🔄 交换", key=f"swap_{idx}", use_container_width=True):
                            # 弹出输入框选择要交换的图片编号
                            swap_target = st.number_input("交换至位置", min_value=1, max_value=num_images, value=idx+1, step=1, key=f"swap_input_{idx}", label_visibility="collapsed")
                            if swap_target != idx+1:
                                target_idx = swap_target - 1
                                st.session_state.image_order[idx], st.session_state.image_order[target_idx] = \
                                    st.session_state.image_order[target_idx], st.session_state.image_order[idx]
                                st.rerun()

        # 实时预览表格
        if total_cells > 0:
            st.subheader("实时预览（图片在表格中的位置）")
            preview_data = []
            for r in range(rows):
                row_cells = []
                for c in range(cols):
                    pos = r * cols + c
                    if pos < len(st.session_state.uploaded_images):
                        img_idx = st.session_state.image_order[pos]
                        img_name = st.session_state.uploaded_images[img_idx][0][:10]
                        row_cells.append(img_name)
                    else:
                        row_cells.append("空")
                preview_data.append(row_cells)
            preview_df = pd.DataFrame(preview_data, columns=[f"{chr(65+i)}" for i in range(cols)])
            st.table(preview_df)

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

                # 设置行高列宽
                for row in range(start_row, end_row+1):
                    ws.row_dimensions[row].height = 150
                for col in range(start_col, end_col+1):
                    ws.column_dimensions[get_column_letter(col)].width = 15

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

# ===================== 模块三：PFMEA 智能生成工具 =====================
def pfmea_tool():
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_ENDPOINTS = [
        "https://api.doubao.com/v1/chat/completions",
        "https://api.doubaoai.com/v1/chat/completions",
        "https://open.doubao.com/v1/chat/completions"
    ]

    # ---------- 本地标准库（每个工序至少3种方案）----------
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
        ]
    }

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
        你是专业的汽车电子行业PFMEA工程师，精通AIAG-VDA FMEA标准。
        针对【{process_name}】工序，生成{scheme_count}组不同的PFMEA方案。
        每组方案包含3-5条失效模式，必须涵盖不同维度（人、机、料、法、环）。
        返回严格的JSON格式：[{{"方案名称":"方案1：...","pfmea_list":[{{"失效模式":"...","失效后果":"...","失效原因":"...","预防措施":"...","探测措施":"...","严重度S":x,"频度O":x,"探测度D":x,"AP等级":"x"}}]}}]
        """
        data = {"model": "doubao-pro-32k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.7}
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

    def export_pfmea_excel(pfmea_data, product_type):
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "PFMEA汇总"

        # 设置表头样式
        header_font = Font(bold=True, size=11)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        headers = ["工序", "失效模式", "失效后果", "失效原因", "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.alignment = header_align
            cell.border = thin_border

        # 写入数据
        row = 2
        for process, items in pfmea_data.items():
            for item in items:
                ws.cell(row=row, column=1, value=process).border = thin_border
                ws.cell(row=row, column=2, value=item.get("失效模式","")).border = thin_border
                ws.cell(row=row, column=3, value=item.get("失效后果","")).border = thin_border
                ws.cell(row=row, column=4, value=item.get("失效原因","")).border = thin_border
                ws.cell(row=row, column=5, value=item.get("预防措施","")).border = thin_border
                ws.cell(row=row, column=6, value=item.get("探测措施","")).border = thin_border
                ws.cell(row=row, column=7, value=item.get("严重度S","")).border = thin_border
                ws.cell(row=row, column=8, value=item.get("频度O","")).border = thin_border
                ws.cell(row=row, column=9, value=item.get("探测度D","")).border = thin_border
                ws.cell(row=row, column=10, value=item.get("AP等级","")).border = thin_border
                row += 1

        # 设置列宽
        col_widths = [20, 30, 35, 35, 40, 40, 8, 8, 8, 8]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64+i)].width = width

        # 冻结表头
        ws.freeze_panes = "A2"

        wb.save(output)
        output.seek(0)
        return output

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

    # ---------- 界面 ----------
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title("⚡ 电池包/充电器PFMEA智能生成系统")
    st.caption("AIAG-VDA FMEA 第一版 | IATF16949:2016")

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
        if st.button("➕ 添加自定义工序", use_container_width=True):
            if custom_process and custom_process not in selected_processes:
                selected_processes.append(custom_process)
                st.success(f"已添加工序：{custom_process}")
                st.rerun()

    # ---------- 知识库管理 ----------
    with st.expander("📚 我的PFMEA知识库管理", expanded=False):
        st.markdown("支持上传旧PFMEA Excel文件，AI自动解析入库，也可在线编辑、删除、备份、恢复。")

        # 上传并解析入库
        uploaded_kb = st.file_uploader("上传旧PFMEA Excel文件（.xlsx）", type=["xlsx"], key="kb_upload")
        if uploaded_kb:
            with st.spinner("正在解析文件，请稍候..."):
                kb_data = parse_pfmea_excel(uploaded_kb.getvalue())
                if kb_data:
                    st.success(f"解析成功，共 {len(kb_data)} 个工序")
                    # 展示预览并确认入库
                    for proc, items in kb_data.items():
                        with st.expander(f"工序：{proc}（{len(items)}条）", expanded=False):
                            st.dataframe(pd.DataFrame(items), use_container_width=True)
                    if st.button("✅ 确认入库", key="confirm_kb_import"):
                        # 合并去重
                        for proc, items in kb_data.items():
                            if proc not in st.session_state.user_knowledge_base:
                                st.session_state.user_knowledge_base[proc] = []
                            existing_keys = {f"{i['失效模式']}_{i['失效原因']}" for i in st.session_state.user_knowledge_base[proc]}
                            for new_item in items:
                                key = f"{new_item.get('失效模式','')}_{new_item.get('失效原因','')}"
                                if key not in existing_keys:
                                    st.session_state.user_knowledge_base[proc].append(new_item)
                        st.success("知识库入库完成！")
                        st.rerun()

        # 显示知识库内容（可编辑、删除）
        if st.session_state.user_knowledge_base:
            for proc, items in list(st.session_state.user_knowledge_base.items()):
                with st.expander(f"📦 工序：{proc}（共{len(items)}条）", expanded=False):
                    df_items = pd.DataFrame(items)
                    df_items = reset_df_index(df_items)
                    edited_df = st.data_editor(df_items, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"kb_edit_{proc}")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"✅ 更新【{proc}】", key=f"update_{proc}"):
                            st.session_state.user_knowledge_base[proc] = edited_df.to_dict("records")
                            st.success("更新成功")
                            st.rerun()
                    with col2:
                        if st.button(f"🗑️ 删除【{proc}】", key=f"delete_{proc}"):
                            del st.session_state.user_knowledge_base[proc]
                            st.success("已删除")
                            st.rerun()
            # 备份和恢复
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📤 备份知识库", use_container_width=True):
                    kb_json = json.dumps(st.session_state.user_knowledge_base, ensure_ascii=False, indent=2)
                    st.download_button(
                        label="下载备份文件",
                        data=kb_json,
                        file_name=f"PFMEA知识库备份_{datetime.now().strftime('%Y%m%d%H%M%S')}.json",
                        mime="application/json",
                        use_container_width=True
                    )
            with col2:
                backup_file = st.file_uploader("恢复知识库（上传备份文件）", type=["json"], key="kb_restore")
                if backup_file:
                    try:
                        restored = json.load(backup_file)
                        if st.button("✅ 确认恢复", use_container_width=True):
                            st.session_state.user_knowledge_base = restored
                            st.success("恢复成功")
                            st.rerun()
                    except Exception as e:
                        st.error(f"恢复失败：{e}")

    # ---------- 生成模式 ----------
    gen_mode = st.radio("生成模式", ["本地标准库（自动合并知识库）", "AI智能生成"], horizontal=True)
    if gen_mode == "AI智能生成":
        scheme_count = st.slider("AI生成方案数量", 2, 5, 3)
        mix_knowledge = st.checkbox("混合我的知识库内容生成（AI会结合知识库生成方案）", value=False)
    else:
        scheme_count = 3
        mix_knowledge = False

    if st.button("🚀 生成PFMEA方案", type="primary", use_container_width=True) and selected_processes:
        st.session_state.generated_pfmea_data = {}
        st.session_state.ai_schemes_temp = {}
        st.session_state.selected_ai_scheme_idx = {}
        for proc in selected_processes:
            if gen_mode == "本地标准库（自动合并知识库）":
                lib_items = process_lib.get(proc, [])
                user_items = st.session_state.user_knowledge_base.get(proc, [])
                combined = lib_items + user_items
                if not combined:
                    st.warning(f"工序【{proc}】无本地数据，请使用AI生成")
                    continue
                st.session_state.generated_pfmea_data[proc] = combined
            else:
                with st.spinner(f"AI正在生成【{proc}】的方案..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if err:
                        st.error(f"{proc} AI生成失败: {err}")
                        continue
                    # 如果混合知识库，将知识库内容作为一个独立方案追加
                    if mix_knowledge and proc in st.session_state.user_knowledge_base:
                        kb_items = st.session_state.user_knowledge_base[proc]
                        if kb_items:
                            kb_scheme = {
                                "方案名称": "我的知识库方案",
                                "pfmea_list": kb_items
                            }
                            schemes.append(kb_scheme)
                    st.session_state.ai_schemes_temp[proc] = schemes
                    # 默认选中第一个方案
                    st.session_state.selected_ai_scheme_idx[proc] = 0
        if st.session_state.generated_pfmea_data or st.session_state.ai_schemes_temp:
            st.success("生成完成！请选择最终方案")
            st.rerun()

    # ---------- 方案选择与确认 ----------
    if st.session_state.generated_pfmea_data or st.session_state.ai_schemes_temp:
        st.markdown("### 选择最终方案")
        final_data = {}
        # 处理本地模式（直接展示编辑表格）
        for proc in selected_processes:
            if proc in st.session_state.generated_pfmea_data:
                with st.expander(f"工序：{proc}", expanded=True):
                    df_proc = pd.DataFrame(st.session_state.generated_pfmea_data[proc])
                    df_proc = reset_df_index(df_proc)
                    edited = st.data_editor(df_proc, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"final_edit_{proc}")
                    final_data[proc] = edited.to_dict("records")
        # 处理AI模式（单选方案）
        for proc in selected_processes:
            if proc in st.session_state.ai_schemes_temp:
                schemes = st.session_state.ai_schemes_temp[proc]
                scheme_names = [s["方案名称"] for s in schemes]
                selected_idx = st.radio(
                    f"为【{proc}】选择方案",
                    options=range(len(scheme_names)),
                    format_func=lambda i: scheme_names[i],
                    index=st.session_state.selected_ai_scheme_idx[proc],
                    key=f"final_select_{proc}"
                )
                st.session_state.selected_ai_scheme_idx[proc] = selected_idx
                selected_scheme = schemes[selected_idx]
                st.markdown("**方案预览**")
                df_scheme = pd.DataFrame(selected_scheme["pfmea_list"])
                st.dataframe(df_scheme, use_container_width=True)
                # 允许编辑
                edited_df = st.data_editor(df_scheme, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"final_edit_ai_{proc}")
                final_data[proc] = edited_df.to_dict("records")

        if st.button("✅ 确认使用当前方案，导出Excel", type="primary", use_container_width=True):
            if final_data:
                excel_file = export_pfmea_excel(final_data, product_type)
                st.download_button(
                    label="📥 下载 PFMEA Excel 文件",
                    data=excel_file,
                    file_name=f"{product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("文件已准备就绪，点击上方按钮下载。")
            else:
                st.warning("没有可导出的数据，请先选择方案。")

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
