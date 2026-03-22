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

# 🔥 全新淡绿色高级UI自定义CSS
st.markdown("""
<style>
/* 全局重置与基础配色 */
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
}
.stApp {
    background: linear-gradient(135deg, #F5FAF5 0%, #E6F4EA 100%);
    padding: 1rem;
}
/* 滚动条美化 */
::-webkit-scrollbar {
    width: 6px;
    height: 6px;
}
::-webkit-scrollbar-thumb {
    background: #A5D6A7;
    border-radius: 3px;
}
::-webkit-scrollbar-track {
    background: #F5FAF5;
}

/* 顶级卡片容器 */
.card {
    background: #FFFFFF;
    border-radius: 20px;
    padding: 28px;
    margin-bottom: 24px;
    box-shadow: 0 8px 24px rgba(76, 175, 80, 0.08);
    border: 1px solid #E6F4EA;
    transition: all 0.3s ease;
}
.card:hover {
    box-shadow: 0 12px 32px rgba(76, 175, 80, 0.12);
    transform: translateY(-2px);
}

/* 标题样式 */
h1 {
    color: #2E7D32;
    font-weight: 700;
    font-size: 2.2rem;
    margin-bottom: 12px;
}
h2, h3 {
    color: #388E3C;
    font-weight: 600;
}

/* 按钮样式 */
.stButton button {
    background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%);
    color: white;
    border-radius: 12px;
    border: none;
    padding: 0.8rem 1.5rem;
    font-weight: 600;
    font-size: 0.95rem;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(76, 175, 80, 0.2);
}
.stButton button:hover {
    background: linear-gradient(135deg, #388E3C 0%, #2E7D32 100%);
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(76, 175, 80, 0.3);
}
.stButton button:active {
    transform: translateY(0);
}

/* 输入框/选择框 */
.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"], .stRadio div {
    border-radius: 12px;
    border: 1px solid #C8E6C9;
    transition: all 0.2s ease;
    background: #FCFFFD;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #4CAF50;
    box-shadow: 0 0 0 2px rgba(76, 175, 80, 0.2);
}

/* 表格/编辑器 */
.dataframe, .stDataEditor {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid #E6F4EA;
}
.stDataEditor {
    background: #FCFFFD;
}

/* 拖动网格样式（模块一专用） */
.drag-grid {
    display: grid;
    gap: 12px;
    padding: 16px;
    background: #F5FAF5;
    border-radius: 16px;
    margin: 16px 0;
}
.drag-item {
    background: white;
    border: 2px solid #E6F4EA;
    border-radius: 12px;
    padding: 8px;
    text-align: center;
    cursor: grab;
    transition: all 0.2s ease;
    position: relative;
}
.drag-item:active {
    cursor: grabbing;
    transform: scale(0.95);
    border-color: #4CAF50;
}
.drag-item img {
    max-width: 80px;
    max-height: 80px;
    object-fit: contain;
    border-radius: 8px;
}
.drag-item .name {
    font-size: 12px;
    color: #555;
    margin-top: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
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
    if "temp_kb_import" not in st.session_state:
        st.session_state.temp_kb_import = None
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

# ===================== 模块一：Excel图片工具（长按拖动排序） =====================
def excel_image_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.header("📸 Excel 图片批量插入工具")
    st.markdown("长按图片0.4秒后拖动，即可自由调整插入顺序")

    # 单元格区域设置
    col1, col2 = st.columns(2)
    with col1:
        start_cell = st.text_input("起始单元格", "A1")
    with col2:
        end_cell = st.text_input("结束单元格", "C5")

    # 计算单元格数量
    total_cells = 0
    rows, cols = 0, 0
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
            st.info(f"✅ 区域共 {total_cells} 个单元格（{rows} 行 × {cols} 列）")
        else:
            st.warning("⚠️ 格式示例：A1")
    except:
        st.warning("⚠️ 单元格格式错误")

    # 图片上传
    uploaded_files = st.file_uploader(
        "选择图片", type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True, key="img_upload"
    )
    if uploaded_files:
        st.session_state.uploaded_images = [(f.name, f.getvalue()) for f in uploaded_files]
        if not st.session_state.image_order:
            st.session_state.image_order = list(range(len(st.session_state.uploaded_images)))
        st.success(f"✅ 已上传 {len(st.session_state.uploaded_images)} 张图片")

    # 🔥 长按拖动排序预览（核心优化）
    if st.session_state.uploaded_images and total_cells > 0:
        st.subheader("📋 长按拖动调整顺序")
        # 生成拖动网格HTML+JS
        drag_html = f"""
        <div class="drag-grid" style="grid-template-columns: repeat({cols}, 1fr);">
        """
        for idx in st.session_state.image_order:
            img_name, img_bytes = st.session_state.uploaded_images[idx]
            b64_img = base64.b64encode(img_bytes).decode()
            drag_html += f"""
            <div class="drag-item" data-idx="{idx}">
                <img src="data:image/png;base64,{b64_img}" />
                <div class="name">{img_name[:10]}</div>
            </div>
            """
        drag_html += "</div>"
        # 长按拖动JS逻辑
        drag_js = f"""
        <script>
        let longPressTimer;
        let isDragging = false;
        let draggedItem = null;
        const grid = document.querySelector('.drag-grid');
        const items = document.querySelectorAll('.drag-item');

        // 长按400ms触发拖动
        items.forEach(item => {{
            item.addEventListener('mousedown', startLongPress);
            item.addEventListener('mouseup', clearLongPress);
            item.addEventListener('mouseleave', clearLongPress);
            item.addEventListener('touchstart', startLongPress);
            item.addEventListener('touchend', clearLongPress);
        }});

        function startLongPress(e) {{
            longPressTimer = setTimeout(() => {{
                isDragging = true;
                draggedItem = this;
                this.style.opacity = '0.7';
            }}, 400);
        }}

        function clearLongPress() {{
            clearTimeout(longPressTimer);
            if (draggedItem) draggedItem.style.opacity = '1';
            isDragging = false;
            draggedItem = null;
        }}

        // 拖动交换位置
        grid.addEventListener('dragover', (e) => {{
            if (!isDragging || !draggedItem) return;
            e.preventDefault();
            const afterElement = getDragAfterElement(grid, e.clientX, e.clientY);
            if (afterElement) {{
                grid.insertBefore(draggedItem, afterElement);
            }} else {{
                grid.appendChild(draggedItem);
            }}
        }});

        // 拖动结束同步顺序到Streamlit
        grid.addEventListener('dragend', () => {{
            if (!isDragging) return;
            const newOrder = Array.from(grid.querySelectorAll('.drag-item'))
                .map(item => Number(item.dataset.idx));
            // 发送顺序到Python
            window.parent.postMessage({{
                type: 'update_order',
                order: newOrder
            }}, '*');
            clearLongPress();
        }});

        function getDragAfterElement(container, x, y) {{
            const elements = [...container.querySelectorAll('.drag-item:not([dragging])')];
            return elements.reduce((closest, child) => {{
                const box = child.getBoundingClientRect();
                const offset = y - box.top - box.height/2;
                if (offset < 0 && offset > closest.offset) {{
                    return {{ offset: offset, element: child }};
                }} else {{
                    return closest;
                }}
            }}, {{ offset: Number.NEGATIVE_INFINITY }}).element;
        }}
        </script>
        """
        # 渲染拖动组件
        st.markdown(drag_html + drag_js, unsafe_allow_html=True)
        # 接收JS传递的顺序
        try:
            from streamlit import runtime
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            ctx = get_script_run_ctx()
            if ctx:
                msg = st.experimental_get_query_params().get("msg", [None])[0]
                if msg and msg.startswith("order:"):
                    new_order = json.loads(msg.replace("order:", ""))
                    st.session_state.image_order = new_order
                    st.rerun()
        except:
            pass

        # 显示当前顺序
        order_names = [st.session_state.uploaded_images[idx][0][:15] for idx in st.session_state.image_order]
        st.write("**当前顺序：** " + " → ".join(order_names))

    # Excel来源选择
    excel_source = st.radio("Excel 来源", ["新建空白工作簿", "上传现有 Excel 文件"])
    existing_wb = None
    if excel_source == "上传现有 Excel 文件":
        existing_file = st.file_uploader("选择 Excel", type=["xlsx", "xlsm"])
        if existing_file:
            try:
                existing_wb = load_workbook(io.BytesIO(existing_file.read()))
                st.success("✅ 加载成功")
            except Exception as e:
                st.error(f"❌ 加载失败：{e}")

    # 生成Excel
    if st.button("🚀 生成并下载 Excel", type="primary", use_container_width=True):
        if not st.session_state.uploaded_images:
            st.error("❌ 请先上传图片")
        elif total_cells == 0:
            st.error("❌ 请填写正确单元格")
        else:
            try:
                start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
                end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
                start_col = openpyxl.utils.column_index_from_string(start_match.group(1))
                start_row = int(start_match.group(2))
                end_col = openpyxl.utils.column_index_from_string(end_match.group(1))
                end_row = int(end_match.group(2))

                # 创建工作簿
                wb = existing_wb if existing_wb else Workbook()
                ws = wb.active
                if not existing_wb:
                    ws.title = "图片表格"

                # 行高列宽
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

                # 下载
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
                st.error(f"❌ 生成失败：{e}")
    st.markdown("</div>", unsafe_allow_html=True)

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

# ===================== 模块三：PFMEA智能生成（AI连接修复） =====================
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
    ],
    "SMT贴片焊接": [
        {"失效模式": "元器件贴装偏移", "失效后果": "焊接不良，功能失效", "失效原因": "贴片机吸嘴磨损，程序坐标偏差", "预防措施": "定期校准设备，首件验证", "探测措施": "AOI全检，SPI锡膏检测", "严重度S": 7, "频度O": 2, "探测度D": 2, "AP等级": "高"},
    ]
}

# 🔥 修复豆包API连接（官方接口+正确参数）
def create_retry_session():
    session = requests.Session()
    retry = Retry(total=3, backoff_factor=0.5, status_forcelist=[429,500,502,503,504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    return session

def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    # 🔥 字节豆包官方API（修复连接失败）
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_ENDPOINT = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    MODEL = "doubao-1.5-pro"

    session = create_retry_session()
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    prompt = f"""你是专业AIAG-VDA PFMEA工程师，为【{process_name}】({product_type})生成{scheme_count}组不同PFMEA方案，每组3条，严格返回JSON，无其他文字：
    [{{"方案名称":"方案1：xxx","pfmea_list":[{{"失效模式":"","失效后果":"","失效原因":"","预防措施":"","探测措施":"","严重度S":int,"频度O":int,"探测度D":int,"AP等级":"高/中/低"}}]}}]"""

    try:
        data = {
            "model": MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 3000
        }
        response = session.post(API_ENDPOINT, headers=headers, json=data, timeout=60)
        response.raise_for_status()
        result = response.json()
        content = result["choices"][0]["message"]["content"].strip()
        content = re.sub(r'^```json|```$', '', content)
        parsed = json.loads(content)
        return parsed, None
    except Exception as e:
        return None, f"API错误：{str(e)}"

# 知识库/导出函数（保留优化）
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
                if any(c in str(col) for c in candidates):
                    col_mapping[target] = col
                    break
        if "工序" not in col_mapping:
            st.error("未找到工序列")
            return {}
        knowledge = {}
        for _, row in df.iterrows():
            process = str(row[col_mapping["工序"]]).strip()
            if process == "nan": continue
            item = {}
            for target, col in col_mapping.items():
                if target != "工序":
                    val = row[col] if pd.notna(row[col]) else ""
                    if target in ["严重度S","频度O","探测度D"]:
                        val = int(float(val)) if str(val).replace('.','').isdigit() else 5
                    item[target] = val
            if "AP等级" not in item:
                s,o,d = item.get("严重度S",5),item.get("频度O",3),item.get("探测度D",4)
                item["AP等级"] = "高" if s>=9 else "低" if s<=4 and o<=6 else "中"
            if process not in knowledge:
                knowledge[process] = []
            knowledge[process].append(item)
        return knowledge
    except Exception as e:
        st.error(f"解析失败：{e}")
        return {}

def merge_knowledge(knowledge_dict, existing_kb):
    for proc, items in knowledge_dict.items():
        if proc not in existing_kb: existing_kb[proc] = []
        keys = {f"{i['失效模式']}_{i['失效原因']}" for i in existing_kb[proc]}
        for item in items:
            key = f"{item['失效模式']}_{item['失效原因']}"
            if key not in keys:
                existing_kb[proc].append(item)
    return existing_kb

def export_pfmea_excel(pfmea_data, product_type):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "PFMEA"
    headers = ["工序","失效模式","失效后果","失效原因","预防措施","探测措施","严重度S","频度O","探测度D","AP等级"]
    for col, h in enumerate(headers,1):
        cell = ws.cell(row=1,column=col,value=h)
        cell.font = Font(bold=True, name="微软雅黑")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    row = 2
    for process, items in pfmea_data.items():
        for item in items:
            ws.cell(row=row,column=1,value=process)
            ws.cell(row=row,column=2,value=item.get("失效模式",""))
            ws.cell(row=row,column=3,value=item.get("失效后果",""))
            ws.cell(row=row,column=4,value=item.get("失效原因",""))
            ws.cell(row=row,column=5,value=item.get("预防措施",""))
            ws.cell(row=row,column=6,value=item.get("探测措施",""))
            ws.cell(row=row,column=7,value=item.get("严重度S",""))
            ws.cell(row=row,column=8,value=item.get("频度O",""))
            ws.cell(row=row,column=9,value=item.get("探测度D",""))
            ws.cell(row=row,column=10,value=item.get("AP等级",""))
            fill = None
            if item.get("AP等级")=="高": fill=PatternFill("FF4D4F", fill_type="solid")
            elif item.get("AP等级")=="中": fill=PatternFill("FAAD14", fill_type="solid")
            elif item.get("AP等级")=="低": fill=PatternFill("52C41A", fill_type="solid")
            if fill: ws.cell(row=row,column=10).fill = fill
            row +=1
    wb.save(output)
    output.seek(0)
    return output

def pfmea_tool():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.title("⚡ PFMEA 智能生成系统")
    st.caption("符合 AIAG-VDA FMEA 标准 | IATF16949:2016")
    st.divider()

    product_type = st.radio("产品类型", ["电池包", "充电器"], horizontal=True)
    process_lib = BATTERY_PROCESS_LIB if product_type=="电池包" else CHARGER_PROCESS_LIB
    all_process = list(process_lib.keys())
    if st.session_state.user_knowledge_base:
        all_process = list(set(all_process + list(st.session_state.user_knowledge_base.keys())))
    all_process.sort()

    col1, col2 = st.columns([3,1])
    with col1:
        selected_processes = st.multiselect("选择工序", all_process, default=all_process[:2] if all_process else [])
    with col2:
        custom_process = st.text_input("自定义工序")
        if st.button("➕ 添加", use_container_width=True):
            if custom_process and custom_process not in selected_processes:
                selected_processes.append(custom_process)
                st.rerun()

    with st.expander("📚 知识库管理"):
        uploaded_kb_file = st.file_uploader("导入PFMEA Excel", type=["xlsx"])
        if uploaded_kb_file:
            kb_data = parse_pfmea_excel(uploaded_kb_file.getvalue())
            st.session_state.user_knowledge_base = merge_knowledge(kb_data, st.session_state.user_knowledge_base)
            st.success(f"导入成功")
            st.rerun()

    st.subheader("生成设置")
    gen_mode = st.radio("生成模式", ["本地标准库+知识库", "AI智能生成"], horizontal=True)
    scheme_count = 3
    mix_knowledge = False
    if gen_mode == "AI智能生成":
        scheme_count = st.slider("AI方案数", 2,5,3)
        mix_knowledge = st.checkbox("混合知识库", True)

    if st.button("🔌 测试AI连接", use_container_width=True):
        res, err = generate_pfmea_ai("电芯来料检验", "电池包", 1)
        if res: st.success("✅ AI连接正常")
        else: st.error(f"❌ {err}")

    if st.button("🚀 生成PFMEA方案", type="primary", use_container_width=True) and selected_processes:
        st.session_state.generated_pfmea_data = {}
        for idx, proc in enumerate(selected_processes):
            if gen_mode == "本地标准库+知识库":
                lib_items = process_lib.get(proc, [])
                kb_items = st.session_state.user_knowledge_base.get(proc, [])
                st.session_state.generated_pfmea_data[proc] = [{"方案名称":"本地+知识库","pfmea_list":lib_items+kb_items}]
            else:
                schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                if err:
                    st.error(f"{proc}生成失败：{err}")
                    continue
                if mix_knowledge and proc in st.session_state.user_knowledge_base:
                    schemes.append({"方案名称":"📁 知识库方案","pfmea_list":st.session_state.user_knowledge_base[proc]})
                st.session_state.generated_pfmea_data[proc] = schemes
        st.success("✅ 生成完成")
        st.rerun()

    if st.session_state.generated_pfmea_data:
        st.subheader("选择最终方案")
        final_data = {}
        for proc in selected_processes:
            if proc not in st.session_state.generated_pfmea_data: continue
            data = st.session_state.generated_pfmea_data[proc]
            if len(data)>1:
                idx = st.radio(f"【{proc}】选方案", range(len(data)), format_func=lambda i:data[i]["方案名称"])
                sel = data[idx]
            else:
                sel = data[0]
            df = pd.DataFrame(sel["pfmea_list"])
            edited = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            final_data[proc] = edited.to_dict("records")
            st.divider()

        if st.button("✅ 导出Excel", type="primary", use_container_width=True):
            file = export_pfmea_excel(final_data, product_type)
            st.download_button("📥 下载PFMEA", file, f"{product_type}_PFMEA.xlsx", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ===================== 主界面 =====================
def main():
    if st.session_state.current_page == "home":
        st.markdown("<div style='text-align:center; padding:2rem 0;'><h1>🛠️ 多功能智能工具集</h1><p>淡绿色轻奢版 | 高效实用</p></div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;'>", unsafe_allow_html=True)
                st.subheader("Excel 图片工具")
                st.markdown("长按拖动排序 | 批量插入")
                if st.button("进入工具", key="btn_excel", use_container_width=True):
                    st.session_state.current_page = "excel_image"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
        with col2:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;'>", unsafe_allow_html=True)
                st.subheader("信息推送工具")
                st.markdown("企业微信推送 | 历史记录")
                if st.button("进入工具", key="btn_push", use_container_width=True):
                    st.session_state.current_page = "image_push"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
        with col3:
            with st.container():
                st.markdown("<div class='card' style='text-align:center;'>", unsafe_allow_html=True)
                st.subheader("PFMEA 智能生成")
                st.markdown("AI快速生成 | 知识库管理")
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
