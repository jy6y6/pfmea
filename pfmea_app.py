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
from datetime import datetime
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

# 初始化 session_state
if "current_page" not in st.session_state:
    st.session_state.current_page = "home"
if "push_history" not in st.session_state:
    st.session_state.push_history = []          # 推送历史记录
if "image_order" not in st.session_state:
    st.session_state.image_order = []           # 模块一图片顺序
if "uploaded_images" not in st.session_state:
    st.session_state.uploaded_images = []       # 模块一上传的图片文件

# ===================== 辅助函数 =====================
def compress_image_to_limit(image_bytes, max_size_mb=2, max_side=1024):
    """压缩图片到指定大小以下（返回压缩后的 bytes）"""
    img = PILImage.open(io.BytesIO(image_bytes))
    # 转换为 RGB（避免 PNG 透明通道问题）
    if img.mode in ('RGBA', 'LA', 'P'):
        rgb = PILImage.new('RGB', img.size, (255, 255, 255))
        rgb.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
        img = rgb
    # 按比例缩放
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

def send_to_wechat_robot(image_bytes, webhook_url):
    """发送图片到企业微信群机器人（需先压缩）"""
    try:
        # 压缩图片至 2MB 以内
        compressed = compress_image_to_limit(image_bytes)
        # 计算 base64 和 md5
        b64 = base64.b64encode(compressed).decode('utf-8')
        md5 = hashlib.md5(compressed).hexdigest()
        payload = {
            "msgtype": "image",
            "image": {
                "base64": b64,
                "md5": md5
            }
        }
        response = requests.post(webhook_url, json=payload, timeout=10)
        result = response.json()
        if result.get("errcode") == 0:
            return True, "推送成功"
        else:
            return False, f"企业微信返回错误: {result}"
    except Exception as e:
        return False, str(e)

def clean_history_limit(history, max_total=200, keep=100):
    """历史记录清理：超过 max_total 则保留 keep 条"""
    if len(history) > max_total:
        return history[-keep:]
    return history

def reset_df_index(df):
    """重置 DataFrame 索引为 RangeIndex，避免 data_editor 警告"""
    if not df.empty and not isinstance(df.index, pd.RangeIndex):
        df = df.reset_index(drop=True)
    return df

# ===================== 模块一：Excel 图片工具 =====================
def excel_image_tool():
    st.header("📸 Excel 图片工具")
    st.markdown("将多张图片按顺序插入 Excel 表格指定区域，支持新建或加载现有文件。")

    # 1. 图片上传（支持多选）
    uploaded_files = st.file_uploader(
        "选择图片（支持 jpg/png/bmp，多选）",
        type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True,
        key="img_upload"
    )
    if uploaded_files:
        # 更新 session_state 中的图片列表（保持顺序）
        current_files = [(f.name, f.getvalue()) for f in uploaded_files]
        st.session_state.uploaded_images = current_files
        st.session_state.image_order = list(range(len(current_files)))
        st.success(f"已上传 {len(current_files)} 张图片")

    # 2. 顺序调整（手动排序）
    if st.session_state.uploaded_images:
        st.subheader("调整图片顺序（点击上下移动）")
        for idx, (name, _) in enumerate(st.session_state.uploaded_images):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.image(io.BytesIO(st.session_state.uploaded_images[idx][1]), width=80, caption=name)
            with col2:
                if idx > 0 and st.button("⬆️ 上移", key=f"up_{idx}", width="stretch"):
                    st.session_state.image_order[idx], st.session_state.image_order[idx-1] = \
                        st.session_state.image_order[idx-1], st.session_state.image_order[idx]
                    st.rerun()
            with col3:
                if idx < len(st.session_state.image_order)-1 and st.button("⬇️ 下移", key=f"down_{idx}", width="stretch"):
                    st.session_state.image_order[idx], st.session_state.image_order[idx+1] = \
                        st.session_state.image_order[idx+1], st.session_state.image_order[idx]
                    st.rerun()

    # 3. 单元格区域设置
    col1, col2 = st.columns(2)
    with col1:
        start_cell = st.text_input("起始单元格 (如 A1)", "A1")
    with col2:
        end_cell = st.text_input("结束单元格 (如 C5)", "C5")

    # 4. 选择 Excel 源（新建或上传现有）
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

    # 5. 预览布局（模拟网格）
    if st.session_state.uploaded_images and start_cell and end_cell:
        try:
            start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
            end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
            if not start_match or not end_match:
                st.error("单元格格式错误，示例：A1")
            else:
                start_col_letter = start_match.group(1)
                start_row = int(start_match.group(2))
                end_col_letter = end_match.group(1)
                end_row = int(end_match.group(2))
                start_col = openpyxl.utils.column_index_from_string(start_col_letter)
                end_col = openpyxl.utils.column_index_from_string(end_col_letter)
                rows = end_row - start_row + 1
                cols = end_col - start_col + 1
                total_cells = rows * cols
                if total_cells < len(st.session_state.uploaded_images):
                    st.warning(f"区域共有 {total_cells} 个单元格，但图片数量为 {len(st.session_state.uploaded_images)}，多余图片将不会被插入")
                # 预览网格
                st.subheader("📋 预览布局（按当前顺序从左到右、从上到下填充）")
                preview_html = "<table style='border-collapse: collapse;'>"
                for r in range(rows):
                    preview_html += "<tr>"
                    for c in range(cols):
                        idx = r * cols + c
                        if idx < len(st.session_state.uploaded_images):
                            img_idx = st.session_state.image_order[idx]
                            img_name = st.session_state.uploaded_images[img_idx][0][:10]
                            preview_html += f"<td style='border:1px solid #ddd; padding:8px; text-align:center;'><div style='width:80px; height:80px; background:#f0f0f0;'><small>{img_name}</small></div>顶替"
                        else:
                            preview_html += "<td style='border:1px solid #ddd; padding:8px; text-align:center;'>空顶替"
                    preview_html += "²"
                preview_html += "∧"
                st.markdown(preview_html, unsafe_allow_html=True)
        except Exception as e:
            st.error(f"解析单元格出错: {e}")

    # 6. 生成并下载 Excel
    if st.button("🚀 生成 Excel 并下载", type="primary", width="stretch"):
        if not st.session_state.uploaded_images:
            st.error("请先上传图片")
        elif not start_cell or not end_cell:
            st.error("请填写起始和结束单元格")
        else:
            try:
                # 解析单元格范围
                start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
                end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())
                if not start_match or not end_match:
                    st.error("单元格格式错误")
                    return
                start_col_letter = start_match.group(1)
                start_row = int(start_match.group(2))
                end_col_letter = end_match.group(1)
                end_row = int(end_match.group(2))
                start_col = openpyxl.utils.column_index_from_string(start_col_letter)
                end_col = openpyxl.utils.column_index_from_string(end_col_letter)

                # 创建工作簿
                if existing_wb:
                    wb = existing_wb
                    ws = wb.active
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "图片表格"

                # 设置行高列宽（每个单元格 150 像素高，列宽约 15 字符）
                CELL_HEIGHT = 150
                CELL_WIDTH = 15
                for row in range(start_row, end_row + 1):
                    ws.row_dimensions[row].height = CELL_HEIGHT
                for col in range(start_col, end_col + 1):
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = CELL_WIDTH

                # 按顺序插入图片
                idx = 0
                for r in range(start_row, end_row + 1):
                    for c in range(start_col, end_col + 1):
                        if idx >= len(st.session_state.uploaded_images):
                            break
                        img_idx = st.session_state.image_order[idx]
                        img_name, img_bytes = st.session_state.uploaded_images[img_idx]
                        try:
                            # 打开图片并缩放至单元格大小（保持比例）
                            pil_img = PILImage.open(io.BytesIO(img_bytes))
                            # 计算缩放比例（单元格内边距 140px）
                            max_w = 140
                            max_h = 140
                            ratio = min(max_w / pil_img.width, max_h / pil_img.height)
                            new_w = int(pil_img.width * ratio)
                            new_h = int(pil_img.height * ratio)
                            resized = pil_img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                            temp_buf = io.BytesIO()
                            resized.save(temp_buf, format='PNG')
                            temp_buf.seek(0)

                            xl_img = XLImage(temp_buf)
                            xl_img.width = new_w
                            xl_img.height = new_h

                            # 直接锚定到单元格（左上角）
                            cell_coord = f"{get_column_letter(c)}{r}"
                            ws.add_image(xl_img, cell_coord)
                        except Exception as e:
                            st.warning(f"插入图片 {img_name} 失败: {e}")
                        idx += 1
                    if idx >= len(st.session_state.uploaded_images):
                        break

                # 保存到内存
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

# ===================== 模块二：信息推送工具 =====================
def image_push_tool():
    st.header("📸 信息推送工具")
    st.markdown("拍照或选择图片，压缩后推送到企业微信群")

    # 企业微信机器人地址（固定）
    WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=bdf3c7d5-a7fd-4d5a-92bb-4ab15a32e042"

    # 图片获取方式
    col_cam, col_album = st.columns(2)
    with col_cam:
        camera_image = st.camera_input("📷 拍照上传", key="camera")
    with col_album:
        album_image = st.file_uploader("🖼️ 从相册选择", type=["jpg", "jpeg", "png", "bmp"], key="album")

    image_data = None
    if camera_image:
        image_data = camera_image.getvalue()
    elif album_image:
        image_data = album_image.getvalue()

    if image_data:
        # 显示原图大小
        st.image(io.BytesIO(image_data), caption="待推送图片", width=200)
        orig_size = len(image_data) / 1024
        st.info(f"原图大小: {orig_size:.2f} KB")

        if st.button("📤 推送至企业微信", type="primary", width="stretch"):
            with st.spinner("正在压缩并推送..."):
                success, msg = send_to_wechat_robot(image_data, WEBHOOK_URL)
                if success:
                    st.success("推送成功！")
                    st.session_state.push_history.append({
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "status": "成功",
                        "size": orig_size
                    })
                else:
                    st.error(f"推送失败: {msg}")
                    st.session_state.push_history.append({
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "status": "失败",
                        "size": orig_size,
                        "error": msg
                    })
                # 清理历史记录（超过200条则保留100条）
                st.session_state.push_history = clean_history_limit(
                    st.session_state.push_history, max_total=200, keep=100
                )
                st.rerun()

    # 显示推送历史
    if st.session_state.push_history:
        with st.expander("📜 推送历史记录", expanded=False):
            df_history = pd.DataFrame(st.session_state.push_history)
            # 重置索引避免警告
            df_history = reset_df_index(df_history)
            st.dataframe(df_history, use_container_width=True)
            if st.button("🧹 清空历史记录", width="stretch"):
                st.session_state.push_history = []
                st.rerun()
    else:
        st.info("暂无推送记录")

# ===================== 模块三：PFMEA 智能生成工具 =====================
def pfmea_tool():
    # 核心配置
    API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
    API_ENDPOINTS = [
        "https://api.doubao.com/v1/chat/completions",
        "https://api.doubaoai.com/v1/chat/completions",
        "https://open.doubao.com/v1/chat/completions"
    ]
    SYSTEM_NAME = "电池包/充电器PFMEA智能生成系统"
    STANDARD = "AIAG-VDA FMEA 第一版 | IATF16949:2016"

    # 本地标准库（简化示例，实际可扩展）
    BATTERY_PROCESS_LIB = {
        "电芯来料检验": [
            {
                "失效模式": "电芯外观尺寸超差",
                "失效后果": "电芯无法装入模组壳体，导致装配中断",
                "失效原因": "来料尺寸公差不符合图纸要求",
                "预防措施": "制定电芯来料检验规范，量具定期校准",
                "探测措施": "首件全尺寸检验，巡检按AQL抽样",
                "严重度S": 6,
                "频度O": 3,
                "探测度D": 4,
                "AP等级": "中"
            }
        ],
        "模组堆叠装配": [
            {
                "失效模式": "电芯堆叠顺序错误",
                "失效后果": "模组电路连接错误，严重时短路",
                "失效原因": "作业人员未按SOP操作",
                "预防措施": "安装极性视觉防错装置",
                "探测措施": "视觉设备100%检测",
                "严重度S": 9,
                "频度O": 2,
                "探测度D": 2,
                "AP等级": "高"
            }
        ]
    }
    CHARGER_PROCESS_LIB = {
        "PCB来料检验": [
            {
                "失效模式": "PCB板尺寸超差",
                "失效后果": "PCB无法装入壳体",
                "失效原因": "PCB生产制程偏差",
                "预防措施": "制定PCB来料检验规范",
                "探测措施": "首件全尺寸检验",
                "严重度S": 5,
                "频度O": 3,
                "探测度D": 4,
                "AP等级": "中"
            }
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
        返回严格的JSON格式：[{{"方案名称":"...","pfmea_list":[{{"失效模式":"...","失效后果":"...","失效原因":"...","预防措施":"...","探测措施":"...","严重度S":x,"频度O":x,"探测度D":x,"AP等级":"x"}}]}}]
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
        headers = ["工序", "失效模式", "失效后果", "失效原因", "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row = 2
        for process, items in pfmea_data.items():
            for item in items:
                ws.cell(row=row, column=1, value=process)
                ws.cell(row=row, column=2, value=item.get("失效模式",""))
                ws.cell(row=row, column=3, value=item.get("失效后果",""))
                ws.cell(row=row, column=4, value=item.get("失效原因",""))
                ws.cell(row=row, column=5, value=item.get("预防措施",""))
                ws.cell(row=row, column=6, value=item.get("探测措施",""))
                ws.cell(row=row, column=7, value=item.get("严重度S",""))
                ws.cell(row=row, column=8, value=item.get("频度O",""))
                ws.cell(row=row, column=9, value=item.get("探测度D",""))
                ws.cell(row=row, column=10, value=item.get("AP等级",""))
                row += 1
        for col in range(1, 11):
            ws.column_dimensions[chr(64+col)].width = 20
        wb.save(output)
        output.seek(0)
        return output

    st.title(SYSTEM_NAME)
    st.caption(STANDARD)

    # 产品类型和工序选择
    product_type = st.radio("产品类型", ["电池包", "充电器"])
    process_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
    selected_processes = st.multiselect("选择工序", list(process_lib.keys()), default=list(process_lib.keys())[:1])

    # 生成模式
    gen_mode = st.radio("生成模式", ["本地标准库", "AI智能生成"])
    scheme_count = 3
    if gen_mode == "AI智能生成":
        scheme_count = st.slider("方案数量", 2, 5, 3)

    if st.button("🚀 生成PFMEA", type="primary", width="stretch") and selected_processes:
        pfmea_data = {}
        for proc in selected_processes:
            if gen_mode == "本地标准库":
                pfmea_data[proc] = process_lib.get(proc, [])
            else:
                with st.spinner(f"AI生成 {proc} ..."):
                    schemes, err = generate_pfmea_ai(proc, product_type, scheme_count)
                    if err:
                        st.error(f"{proc} 生成失败: {err}")
                        pfmea_data[proc] = []
                    else:
                        pfmea_data[proc] = schemes[0]["pfmea_list"] if schemes else []
        if pfmea_data:
            st.success("生成完成")
            # 可编辑预览
            all_rows = []
            for p, items in pfmea_data.items():
                for item in items:
                    all_rows.append({"工序": p, **item})
            if all_rows:
                df = pd.DataFrame(all_rows)
                df = reset_df_index(df)  # 避免 data_editor 索引警告
                edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True)
                # 重新组装
                updated = {}
                for _, row in edited.iterrows():
                    p = row["工序"]
                    if p not in updated:
                        updated[p] = []
                    item = {k:v for k,v in row.items() if k != "工序"}
                    updated[p].append(item)
                pfmea_data = updated
                # 导出
                excel_file = export_pfmea_excel(pfmea_data, product_type)
                st.download_button(
                    label="📥 下载PFMEA Excel",
                    data=excel_file,
                    file_name=f"{product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )

# ===================== 主界面 =====================
def main():
    # 首页显示三个卡片按钮
    if st.session_state.current_page == "home":
        st.title("🛠️ 多功能智能工具集")
        st.markdown("请选择要使用的工具模块：")
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("📸 Excel 图片工具\n\n将多张图片按顺序插入 Excel 表格", width="stretch"):
                st.session_state.current_page = "excel_image"
                st.rerun()

        with col2:
            if st.button("📱 信息推送工具\n\n拍照/选图推送至企业微信群", width="stretch"):
                st.session_state.current_page = "image_push"
                st.rerun()

        with col3:
            if st.button("⚡ PFMEA 智能生成工具\n\n符合 AIAG-VDA 标准的 FMEA 生成", width="stretch"):
                st.session_state.current_page = "pfmea"
                st.rerun()

    # 模块页面
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
