# --------------- 依赖安装命令 ---------------
# pip install streamlit openpyxl requests pandas
# 本地Ollama模式需额外安装：https://ollama.com/ ，拉取模型：ollama pull qwen2:7b
# --------------------------------------------

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
import json
from datetime import datetime

# -------------------------- 全局配置 --------------------------
# 页面配置（手机端自适应）
st.set_page_config(
    page_title="PFMEA智能生成工具",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 手机端自适应CSS
st.markdown("""
<style>
@media (max-width: 768px) {
    .row-widget.stButton {
        width: 100% !important;
    }
    .stDataFrame {
        font-size: 12px !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        flex-direction: column !important;
    }
    .stTextInput, .stTextArea, .stSelectbox {
        width: 100% !important;
    }
}
/* AP等级标色 */
.ap-high {
    background-color: #ff4d4f;
    color: white;
    font-weight: bold;
}
.ap-medium {
    background-color: #faad14;
    color: white;
    font-weight: bold;
}
.ap-low {
    background-color: #52c41a;
    color: white;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# AIAG-VDA AP等级判定矩阵（标准）
def calculate_ap(severity, occurrence, detection):
    """根据S/O/D自动计算AP行动优先级"""
    if not (1<=severity<=10 and 1<=occurrence<=10 and 1<=detection<=10):
        return "无效评分"
    # 高优先级H：严重度≥9，或严重度7-8且频度≥4，或严重度7-8且探测度≥8
    if severity >= 9:
        return "H"
    elif 7 <= severity <= 8:
        if occurrence >= 4 or detection >= 8:
            return "H"
        else:
            return "M"
    # 中优先级M：严重度5-6，或严重度1-4且频度≥7
    elif 5 <= severity <= 6:
        if occurrence >= 7 or detection >= 9:
            return "H"
        else:
            return "M"
    elif 1 <= severity <= 4:
        if occurrence >= 7:
            return "M"
        else:
            return "L"
    # 低优先级L
    else:
        return "L"

# 预设行业工序库
PRESET_PROCESSES = {
    "电池包PACK装配": [
        "来料检验（电芯/结构件/BMS）",
        "电芯分选配组",
        "电芯堆叠与固定",
        "母线/极耳激光焊接",
        "模组装配与固定",
        "BMS板装配与接线",
        "高压线束装配",
        "壳体密封与上盖组装",
        "绝缘耐压测试",
        "充放电循环测试",
        "成品外观检验",
        "成品包装入库"
    ],
    "充电器装配": [
        "PCBA SMT贴片",
        "插件后焊",
        "PCBA功能测试",
        "外壳注塑与预处理",
        "PCB与外壳组装",
        "成品老化测试",
        "耐压绝缘测试",
        "输出性能测试",
        "成品外观检验",
        "成品包装入库"
    ]
}

# PFMEA标准列名
PFMEA_COLUMNS = [
    "工序编号", "工序名称", "过程功能", "过程要求",
    "失效模式", "失效影响", "严重度S", "失效起因/机理",
    "频度O", "预防控制措施", "探测控制措施", "探测度D",
    "AP等级", "优化措施", "责任人", "完成期限"
]

# -------------------------- AI生成功能 --------------------------
# 本地Ollama AI生成
def generate_pfmea_ollama(process_name, product_type):
    prompt = f"""
    你是专业的新能源行业FMEA工程师，针对{product_type}的【{process_name}】工序，严格按照AIAG-VDA PFMEA标准，生成完整的PFMEA内容。
    要求：
    1. 严格遵循「过程功能→过程要求→失效模式→失效影响→失效起因→控制措施」逻辑链
    2. 严重度S、频度O、探测度D按1-10分评分，符合汽车行业标准，安全相关项严重度≥8
    3. 失效模式、控制措施必须贴合{product_type}装配的实际生产场景，不能通用化
    4. 输出格式为JSON数组，每个失效项为一个对象，key必须包含：
    过程功能、过程要求、失效模式、失效影响、严重度S、失效起因/机理、频度O、预防控制措施、探测控制措施、探测度D
    5. 每个工序生成3-5个核心失效项，无需多余解释，直接输出JSON
    """
    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": "qwen2:7b",
                "prompt": prompt,
                "stream": False,
                "temperature": 0.3
            },
            timeout=60
        )
        result = response.json()
        return json.loads(result["response"])
    except Exception as e:
        st.error(f"本地AI生成失败：{str(e)}，请检查Ollama是否启动、模型是否已拉取")
        return []

# 豆包API AI生成（用户自填密钥）
def generate_pfmea_doubao(process_name, product_type, api_key):
    prompt = f"""
    你是专业的新能源行业FMEA工程师，针对{product_type}的【{process_name}】工序，严格按照AIAG-VDA PFMEA标准，生成完整的PFMEA内容。
    要求：
    1. 严格遵循「过程功能→过程要求→失效模式→失效影响→失效起因→控制措施」逻辑链
    2. 严重度S、频度O、探测度D按1-10分评分，符合汽车行业标准，安全相关项严重度≥8
    3. 失效模式、控制措施必须贴合{product_type}装配的实际生产场景，不能通用化
    4. 输出格式为JSON数组，每个失效项为一个对象，key必须包含：
    过程功能、过程要求、失效模式、失效影响、严重度S、失效起因/机理、频度O、预防控制措施、探测控制措施、探测度D
    5. 每个工序生成3-5个核心失效项，无需多余解释，直接输出JSON
    """
    try:
        url = "https://api.doubao.com/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        data = {
            "model": "ep-20240805194357-jzrql",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "stream": False
        }
        response = requests.post(url, headers=headers, json=data, timeout=60)
        result = response.json()
        content = result["choices"][0]["message"]["content"]
        return json.loads(content)
    except Exception as e:
        st.error(f"豆包API生成失败：{str(e)}，请检查API密钥是否正确")
        return []

# -------------------------- Excel导出功能 --------------------------
def export_to_excel(pfmea_df, base_info):
    """生成Excel2016兼容的PFMEA文件，带标准格式"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFMEA"

    # 定义样式
    title_font = Font(name="微软雅黑", size=14, bold=True)
    header_font = Font(name="微软雅黑", size=10, bold=True)
    content_font = Font(name="微软雅黑", size=9)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    # AP等级填充色
    fill_h = PatternFill(start_color="FF4D4F", end_color="FF4D4F", fill_type="solid")
    fill_m = PatternFill(start_color="FAAD14", end_color="FAAD14", fill_type="solid")
    fill_l = PatternFill(start_color="52C41A", end_color="52C41A", fill_type="solid")

    # 写入基础信息
    ws.merge_cells('A1:Q1')
    ws['A1'] = f"过程失效模式及后果分析（PFMEA）"
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align

    base_info_rows = [
        ["项目名称", base_info["project_name"], "产品型号", base_info["product_type"], "版本号", base_info["version"], "生成日期", datetime.now().strftime("%Y-%m-%d")],
        ["责任部门", base_info["department"], "团队成员", base_info["team_member"], "", "", "", ""]
    ]
    for i, row in enumerate(base_info_rows):
        for j, value in enumerate(row):
            cell = ws.cell(row=i+2, column=j+1, value=value)
            cell.font = header_font if j%2==0 else content_font
            cell.alignment = center_align
            cell.border = thin_border

    # 写入表头
    header_row = 5
    for col_idx, col_name in enumerate(PFMEA_COLUMNS):
        cell = ws.cell(row=header_row, column=col_idx+1, value=col_name)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # 写入内容
    for row_idx, row_data in pfmea_df.iterrows():
        current_row = header_row + 1 + row_idx
        for col_idx, col_name in enumerate(PFMEA_COLUMNS):
            cell_value = row_data.get(col_name, "")
            cell = ws.cell(row=current_row, column=col_idx+1, value=cell_value)
            cell.font = content_font
            cell.alignment = left_align if col_idx in [2,3,4,5,7,9,10,14] else center_align
            cell.border = thin_border
            # AP等级标色
            if col_name == "AP等级":
                if cell_value == "H":
                    cell.fill = fill_h
                elif cell_value == "M":
                    cell.fill = fill_m
                elif cell_value == "L":
                    cell.fill = fill_l

    # 设置列宽
    col_widths = [8, 18, 20, 18, 20, 22, 8, 22, 8, 22, 22, 8, 8, 22, 12, 12]
    for i, width in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i+1)].width = width

    # 冻结表头
    ws.freeze_panes = f"A{header_row+1}"

    # 保存文件
    file_name = f"PFMEA_{base_info['project_name']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    wb.save(file_name)
    return file_name

# -------------------------- 主页面逻辑 --------------------------
def main():
    st.title("📋 电池包&充电器PFMEA智能生成工具")
    st.caption("严格对齐AIAG-VDA最新标准 | 手机端自适应 | 永久免费无过期")

    # 1. 基础信息配置
    with st.expander("📝 PFMEA基础信息配置", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            project_name = st.text_input("项目名称", value="电池包PACK总成装配项目")
            product_type = st.selectbox("产品类型", options=["电池包PACK", "充电器", "电池包+充电器"])
            version = st.text_input("版本号", value="V1.0")
        with col2:
            department = st.text_input("责任部门", value="工艺部/IE")
            team_member = st.text_input("团队成员", value="")
            generate_date = st.date_input("生成日期", value=datetime.now())

    # 2. AI模式选择
    with st.expander("🤖 AI生成模式配置", expanded=False):
        ai_mode = st.radio("AI模式", options=["本地Ollama模式（完全免费无限制）", "豆包API模式（精准生成）"], horizontal=True)
        api_key = ""
        if ai_mode == "豆包API模式（精准生成）":
            api_key = st.text_input("豆包API密钥", type="password", help="自行在豆包开放平台申请，密钥仅保存在本地，不会上传")

    # 3. 工序选择与生成
    st.subheader("🔧 工序选择与PFMEA生成")
    col_process1, col_process2 = st.columns(2)
    with col_process1:
        # 预设工序选择
        preset_category = st.selectbox("选择预设工序库", options=PRESET_PROCESSES.keys())
        selected_preset_processes = st.multiselect("一键选择预设工序", options=PRESET_PROCESSES[preset_category])
    with col_process2:
        # 自定义工序输入
        custom_process = st.text_input("自定义工序输入（多个工序用英文逗号分隔）", placeholder="例如：激光打标,螺丝紧固")
        custom_process_list = [p.strip() for p in custom_process.split(",") if p.strip()]

    # 合并所有选中的工序
    all_selected_processes = list(set(selected_preset_processes + custom_process_list))
    st.write(f"已选中工序：{all_selected_processes if all_selected_processes else '无'}")

    # 生成按钮
    generate_btn = st.button("🚀 一键生成选中工序PFMEA", use_container_width=True)

    # 4. PFMEA数据管理
    # 初始化session_state
    if "pfmea_data" not in st.session_state:
        st.session_state.pfmea_data = pd.DataFrame(columns=PFMEA_COLUMNS)

    # 执行生成逻辑
    if generate_btn and all_selected_processes:
        with st.spinner("正在生成PFMEA内容，请稍候..."):
            new_data = []
            process_no = len(st.session_state.pfmea_data) + 1
            for process in all_selected_processes:
                st.write(f"正在生成【{process}】工序内容...")
                # 调用AI生成
                if ai_mode == "本地Ollama模式（完全免费无限制）":
                    pfmea_items = generate_pfmea_ollama(process, product_type)
                else:
                    if not api_key:
                        st.error("请填写豆包API密钥")
                        break
                    pfmea_items = generate_pfmea_doubao(process, product_type, api_key)
                # 整理数据
                for item in pfmea_items:
                    s = item.get("严重度S", 5)
                    o = item.get("频度O", 3)
                    d = item.get("探测度D", 4)
                    ap = calculate_ap(int(s), int(o), int(d))
                    new_data.append({
                        "工序编号": f"OP{str(process_no).zfill(2)}",
                        "工序名称": process,
                        "过程功能": item.get("过程功能", ""),
                        "过程要求": item.get("过程要求", ""),
                        "失效模式": item.get("失效模式", ""),
                        "失效影响": item.get("失效影响", ""),
                        "严重度S": int(s),
                        "失效起因/机理": item.get("失效起因/机理", ""),
                        "频度O": int(o),
                        "预防控制措施": item.get("预防控制措施", ""),
                        "探测控制措施": item.get("探测控制措施", ""),
                        "探测度D": int(d),
                        "AP等级": ap,
                        "优化措施": "",
                        "责任人": "",
                        "完成期限": ""
                    })
                process_no += 1
            # 追加到现有数据
            if new_data:
                new_df = pd.DataFrame(new_data)
                st.session_state.pfmea_data = pd.concat([st.session_state.pfmea_data, new_df], ignore_index=True)
                st.success("PFMEA生成完成！")

    # 5. 数据编辑与操作
    st.subheader("📋 PFMEA内容编辑与管理")
    if not st.session_state.pfmea_data.empty:
        # 可编辑表格
        edited_df = st.data_editor(
            st.session_state.pfmea_data,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "AP等级": st.column_config.Column(width="small"),
                "严重度S": st.column_config.NumberColumn(width="small", min_value=1, max_value=10),
                "频度O": st.column_config.NumberColumn(width="small", min_value=1, max_value=10),
                "探测度D": st.column_config.NumberColumn(width="small", min_value=1, max_value=10),
            }
        )
        # 更新数据
        st.session_state.pfmea_data = edited_df

        # 批量操作按钮
        col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
        with col_btn1:
            if st.button("🔄 重新计算AP等级", use_container_width=True):
                for idx, row in st.session_state.pfmea_data.iterrows():
                    s = row["严重度S"]
                    o = row["频度O"]
                    d = row["探测度D"]
                    st.session_state.pfmea_data.at[idx, "AP等级"] = calculate_ap(int(s), int(o), int(d))
                st.success("AP等级重新计算完成！")
                st.rerun()
        with col_btn2:
            if st.button("🗑️ 清空选中行", use_container_width=True):
                st.session_state.pfmea_data = edited_df[~edited_df.index.isin(st.session_state.get("selected_rows", []))]
                st.success("已清空选中行！")
                st.rerun()
        with col_btn3:
            if st.button("📋 保存为模板", use_container_width=True):
                st.session_state["template_data"] = st.session_state.pfmea_data.copy()
                st.success("模板保存成功！")
        with col_btn4:
            if st.button("🔄 加载模板", use_container_width=True):
                if "template_data" in st.session_state:
                    st.session_state.pfmea_data = st.session_state["template_data"].copy()
                    st.success("模板加载成功！")
                    st.rerun()
                else:
                    st.error("暂无保存的模板")

        # 6. 导出Excel
        st.subheader("📤 导出Excel文件")
        export_btn = st.button("📥 导出Excel2016兼容文件", use_container_width=True, type="primary")
        if export_btn:
            base_info = {
                "project_name": project_name,
                "product_type": product_type,
                "version": version,
                "department": department,
                "team_member": team_member
            }
            with st.spinner("正在生成Excel文件，请稍候..."):
                file_name = export_to_excel(st.session_state.pfmea_data, base_info)
                with open(file_name, "rb") as f:
                    st.download_button(
                        label="✅ 点击下载PFMEA Excel文件",
                        data=f,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
    else:
        st.info("暂无PFMEA数据，请先选择工序并生成内容")

if __name__ == "__main__":
    main()
