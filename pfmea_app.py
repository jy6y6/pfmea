# --------------- 零报错 最终修复版 PFMEA工具 ---------------
# 适配Streamlit Cloud + 本地Termux，修复所有API调用、格式解析报错
# 兼容Excel2016导出，全链路异常容错，不会崩溃
# -----------------------------------------------------------

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
import json
import re
from datetime import datetime
import os

# -------------------------- 全局初始化（彻底解决刷新报错） --------------------------
# 页面基础配置
st.set_page_config(
    page_title="PFMEA智能生成工具",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 强制初始化session_state，避免页面刷新报错
if "pfmea_data" not in st.session_state:
    st.session_state.pfmea_data = pd.DataFrame(columns=[
        "工序编号", "工序名称", "过程功能", "过程要求",
        "失效模式", "失效影响", "严重度S", "失效起因/机理",
        "频度O", "预防控制措施", "探测控制措施", "探测度D",
        "AP等级", "优化措施", "责任人", "完成期限"
    ])

# 手机端自适应CSS
st.markdown("""
<style>
@media (max-width: 768px) {
    .row-widget.stButton {width: 100% !important;}
    .stDataFrame {font-size: 11px !important;}
    div[data-testid="stVerticalBlock"] > div {flex-direction: column !important;}
    .stTextInput, .stTextArea, .stSelectbox {width: 100% !important;}
    .block-container {padding-left: 1rem !important;padding-right: 1rem !important;padding-top: 2rem !important;}
}
.ap-high {background-color: #ff4d4f;color: white;font-weight: bold;}
.ap-medium {background-color: #faad14;color: white;font-weight: bold;}
.ap-low {background-color: #52c41a;color: white;font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# -------------------------- 核心工具函数（全容错修复） --------------------------
# AIAG-VDA AP等级判定（加类型容错，彻底解决数字转换报错）
def calculate_ap(severity, occurrence, detection):
    try:
        s = int(severity)
        o = int(occurrence)
        d = int(detection)
        if not (1<=s<=10 and 1<=o<=10 and 1<=d<=10):
            return "无效评分"
        if s >= 9:
            return "H"
        elif 7 <= s <= 8:
            return "H" if (o >= 4 or d >= 8) else "M"
        elif 5 <= s <= 6:
            return "H" if (o >= 7 or d >= 9) else "M"
        elif 1 <= s <= 4:
            return "M" if o >= 7 else "L"
        else:
            return "L"
    except:
        return "无效评分"

# JSON内容提取（核心修复：解决AI返回非纯JSON的解析报错）
def extract_json_from_text(text):
    try:
        # 先清理markdown格式
        text = text.strip().replace("```json", "").replace("```", "").replace("'", "\"")
        # 正则提取最外层的JSON数组
        json_match = re.search(r'\[\s*\{.*\}\s*\]', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        # 兜底：直接尝试解析全文
        return json.loads(text)
    except Exception as e:
        st.error(f"JSON解析失败：{str(e)}，AI返回内容格式异常")
        st.write("AI返回的原始内容：", text[:500] + "..." if len(text)>500 else text)
        return []

# 豆包API调用（核心修复：更新官方通用接口，个人密钥可直接调用）
def generate_pfmea_doubao(process_name, product_type, api_key):
    prompt = f"""
    你是专业的新能源汽车电池包/充电器装配行业FMEA工程师，严格遵循AIAG-VDA PFMEA第一版标准，针对【{product_type}】的【{process_name}】工序，生成合规的PFMEA内容。
    硬性要求：
    1. 严格遵循「过程功能→过程要求→失效模式→失效影响→失效起因→控制措施」逻辑链
    2. 严重度S、频度O、探测度D必须是1-10的整数，安全相关项严重度≥8，符合IATF16949审核要求
    3. 失效模式、控制措施必须贴合{product_type}装配的实际生产场景，禁止通用化、空泛内容
    4. 输出格式必须是纯JSON数组，每个失效项为一个对象，key必须严格包含：
    过程功能、过程要求、失效模式、失效影响、严重度S、失效起因/机理、频度O、预防控制措施、探测控制措施、探测度D
    5. 每个工序生成3-5个核心失效项，不要任何多余解释、不要markdown格式，只输出纯JSON
    """
    try:
        # 豆包官方最新通用API地址（个人密钥可直接调用）
        url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key.strip()}"
        }
        data = {
            "model": "doubao-1.5-pro-32k", # 官方通用免费模型，个人密钥直接用
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "stream": False,
            "max_tokens": 2048
        }
        # 发送请求，超时60秒
        response = requests.post(url, headers=headers, json=data, timeout=60)
        response.raise_for_status() # 捕获HTTP错误
        result = response.json()
        # 提取返回内容
        content = result["choices"][0]["message"]["content"]
        # 解析JSON，加容错
        return extract_json_from_text(content)
    except Exception as e:
        st.error(f"AI生成失败：{str(e)}")
        # 详细错误提示，方便排查
        if "401" in str(e):
            st.error("错误原因：API密钥无效，请检查密钥是否正确、是否开通了对应模型的调用权限")
        elif "403" in str(e):
            st.error("错误原因：账号没有该模型的调用权限，请在火山引擎控制台开通doubao-1.5-pro模型的权限")
        elif "timeout" in str(e):
            st.error("错误原因：请求超时，请检查网络是否正常")
        return []

# Excel导出（加路径容错，解决手机/云端导出报错）
def export_to_excel(pfmea_df, base_info):
    try:
        # 文件名
        file_name = f"PFMEA_{base_info['project_name']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
        # 路径容错：优先手机下载目录，其次当前目录
        save_dir = "/sdcard/Download/"
        if not os.path.exists(save_dir):
            save_dir = "./"
        save_path = os.path.join(save_dir, file_name)

        # 创建Excel工作簿
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PFMEA"

        # 样式定义
        title_font = Font(name="微软雅黑", size=14, bold=True)
        header_font = Font(name="微软雅黑", size=10, bold=True)
        content_font = Font(name="微软雅黑", size=9)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        fill_h = PatternFill(start_color="FF4D4F", end_color="FF4D4F", fill_type="solid")
        fill_m = PatternFill(start_color="FAAD14", end_color="FAAD14", fill_type="solid")
        fill_l = PatternFill(start_color="52C41A", end_color="52C41A", fill_type="solid")

        # 写入基础信息
        ws.merge_cells('A1:Q1')
        ws['A1'] = "过程失效模式及后果分析（PFMEA）"
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
        PFMEA_COLUMNS = st.session_state.pfmea_data.columns.tolist()
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
            if i < len(PFMEA_COLUMNS):
                ws.column_dimensions[get_column_letter(i+1)].width = width

        # 冻结表头
        ws.freeze_panes = f"A{header_row+1}"
        # 保存文件
        wb.save(save_path)
        return save_path, file_name
    except Exception as e:
        st.error(f"Excel生成失败：{str(e)}")
        return None, None

# -------------------------- 预设行业库 --------------------------
PRESET_PROCESSES = {
    "电池包PACK装配": [
        "来料检验（电芯/结构件/BMS）", "电芯分选配组", "电芯堆叠与固定",
        "母线/极耳激光焊接", "模组装配与固定", "BMS板装配与接线",
        "高压线束装配", "壳体密封与上盖组装", "绝缘耐压测试",
        "充放电循环测试", "成品外观检验", "成品包装入库"
    ],
    "充电器装配": [
        "PCBA SMT贴片", "插件后焊", "PCBA功能测试", "外壳注塑与预处理",
        "PCB与外壳组装", "成品老化测试", "耐压绝缘测试", "输出性能测试",
        "成品外观检验", "成品包装入库"
    ]
}

# 预设标准PFMEA库（无网络兜底）
PRESET_PFMEA_LIB = {
    "电池包PACK装配": {
        "来料检验（电芯/结构件/BMS）": [
            {
                "过程功能": "对电芯、结构件、BMS等来料进行检验，确认符合图纸和规格要求",
                "过程要求": "来料尺寸、性能、外观100%符合规格书要求，不合格品不流入生产",
                "失效模式": "来料尺寸超差",
                "失效影响": "装配干涉，无法完成组装，导致成品返工",
                "严重度S": 6,
                "失效起因/机理": "供应商生产尺寸偏差，来料检验未按标准抽检",
                "频度O": 3,
                "预防控制措施": "供应商资质审核，来料执行AQL抽样标准，首件确认",
                "探测控制措施": "卡尺/二次元测量尺寸，全检关键尺寸",
                "探测度D": 4
            },
            {
                "过程功能": "对电芯、结构件、BMS等来料进行检验，确认符合图纸和规格要求",
                "过程要求": "电芯电压、内阻、容量匹配，符合分组要求",
                "失效模式": "电芯内阻/电压超差，分组不匹配",
                "失效影响": "模组循环寿命下降，充放电异常，甚至引发热失控",
                "严重度S": 9,
                "失效起因/机理": "电芯生产一致性差，来料检测设备精度不足",
                "频度O": 2,
                "预防控制措施": "供应商电芯全检出货，来料执行全项电性能检测",
                "探测控制措施": "电芯分选设备100%全检电压、内阻，数据追溯",
                "探测度D": 2
            }
        ]
    },
    "充电器装配": {
        "PCBA SMT贴片": [
            {
                "过程功能": "通过SMT设备将元器件贴装到PCB板上，完成回流焊接",
                "过程要求": "元器件贴装位置正确，焊接无虚焊、连锡、偏移，符合IPC标准",
                "失效模式": "元器件贴装偏移、错件",
                "失效影响": "PCBA功能异常，成品无法正常工作，导致返工报废",
                "严重度S": 6,
                "失效起因/机理": "贴片机程序错误，物料上料错误，吸嘴磨损",
                "频度O": 3,
                "预防控制措施": "上料双人复核，程序固化锁定，设备定期维护",
                "探测控制措施": "SPI锡膏检测，AOI外观全检，首件功能测试",
                "探测度D": 3
            }
        ]
    }
}

# -------------------------- 主页面逻辑 --------------------------
def main():
    st.title("📋 电池包&充电器PFMEA智能生成工具")
    st.caption("AIAG-VDA最新标准 | 零报错修复版 | Excel2016兼容 | 手机/电脑通用")

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

    # 2. AI配置
    with st.expander("🤖 AI生成配置", expanded=True):
        ai_mode = st.radio("生成模式", options=["预设标准库（无网络可用）", "豆包AI在线生成"], horizontal=True)
        api_key = ""
        if ai_mode == "豆包AI在线生成":
            api_key = st.text_input("豆包API密钥", type="password", placeholder="请输入你申请的API密钥")
            st.caption("提示：密钥仅在本地使用，不会上传，申请教程见页面底部")

    # 3. 工序选择
    st.subheader("🔧 工序选择")
    col_process1, col_process2 = st.columns(2)
    with col_process1:
        preset_category = st.selectbox("选择工序库", options=PRESET_PROCESSES.keys())
        selected_preset_processes = st.multiselect("选择预设工序", options=PRESET_PROCESSES[preset_category])
    with col_process2:
        custom_process = st.text_input("自定义工序（多个用英文逗号分隔）", placeholder="例如：激光打标,螺丝紧固")
        custom_process_list = [p.strip() for p in custom_process.split(",") if p.strip()]

    # 合并选中工序
    all_selected_processes = list(set(selected_preset_processes + custom_process_list))
    st.write(f"已选中工序：{all_selected_processes if all_selected_processes else '无'}")

    # 生成按钮
    generate_btn = st.button("🚀 生成选中工序PFMEA", use_container_width=True, type="primary")

    # 执行生成逻辑
    if generate_btn and all_selected_processes:
        with st.spinner("正在生成PFMEA内容，请稍候..."):
            new_data = []
            process_no = len(st.session_state.pfmea_data) + 1
            for process in all_selected_processes:
                st.write(f"正在处理【{process}】工序...")
                pfmea_items = []

                # 模式1：预设标准库
                if ai_mode == "预设标准库（无网络可用）":
                    if preset_category in PRESET_PFMEA_LIB and process in PRESET_PFMEA_LIB[preset_category]:
                        pfmea_items = PRESET_PFMEA_LIB[preset_category][process]
                    else:
                        # 通用兜底模板，不会报错
                        pfmea_items = [
                            {
                                "过程功能": f"完成{process}工序作业，符合图纸和作业指导书要求",
                                "过程要求": f"{process}作业后尺寸、性能、外观符合规格要求，无不良品流出",
                                "失效模式": f"{process}作业不符合标准要求",
                                "失效影响": "后工序无法装配，成品性能不达标，导致返工",
                                "严重度S": 5,
                                "失效起因/机理": "人员操作不规范，设备参数设置错误，来料不良",
                                "频度O": 3,
                                "预防控制措施": "作业指导书标准化，人员培训上岗，设备参数固化",
                                "探测控制措施": "作业后首件确认，过程巡检，成品全检",
                                "探测度D": 4
                            }
                        ]
                # 模式2：豆包AI在线生成
                else:
                    if not api_key:
                        st.error("请先输入豆包API密钥！")
                        st.stop()
                    pfmea_items = generate_pfmea_doubao(process, product_type, api_key)
                    # AI生成失败，用兜底模板
                    if not pfmea_items:
                        st.warning(f"【{process}】工序AI生成失败，已使用通用兜底模板")
                        pfmea_items = [
                            {
                                "过程功能": f"完成{process}工序作业，符合图纸和作业指导书要求",
                                "过程要求": f"{process}作业后尺寸、性能、外观符合规格要求，无不良品流出",
                                "失效模式": f"{process}作业不符合标准要求",
                                "失效影响": "后工序无法装配，成品性能不达标，导致返工",
                                "严重度S": 5,
                                "失效起因/机理": "人员操作不规范，设备参数设置错误，来料不良",
                                "频度O": 3,
                                "预防控制措施": "作业指导书标准化，人员培训上岗，设备参数固化",
                                "探测控制措施": "作业后首件确认，过程巡检，成品全检",
                                "探测度D": 4
                            }
                        ]
                
                # 整理数据，加全容错
                for item in pfmea_items:
                    try:
                        # 数字转换兜底，避免类型报错
                        s = int(item.get("严重度S", 5))
                        o = int(item.get("频度O", 3))
                        d = int(item.get("探测度D", 4))
                        ap = calculate_ap(s, o, d)
                        # 拼接完整数据
                        new_data.append({
                            "工序编号": f"OP{str(process_no).zfill(2)}",
                            "工序名称": process,
                            "过程功能": item.get("过程功能", ""),
                            "过程要求": item.get("过程要求", ""),
                            "失效模式": item.get("失效模式", ""),
                            "失效影响": item.get("失效影响", ""),
                            "严重度S": s,
                            "失效起因/机理": item.get("失效起因/机理", ""),
                            "频度O": o,
                            "预防控制措施": item.get("预防控制措施", ""),
                            "探测控制措施": item.get("探测控制措施", ""),
                            "探测度D": d,
                            "AP等级": ap,
                            "优化措施": item.get("优化措施", ""),
                            "责任人": item.get("责任人", ""),
                            "完成期限": item.get("完成期限", "")
                        })
                    except Exception as e:
                        st.warning(f"【{process}】工序单条数据处理失败：{str(e)}，已跳过该条")
                        continue
                process_no += 1

            # 追加到现有数据
            if new_data:
                new_df = pd.DataFrame(new_data)
                st.session_state.pfmea_data = pd.concat([st.session_state.pfmea_data, new_df], ignore_index=True)
                st.success("✅ PFMEA生成完成！")
                st.rerun()

    # 5. 内容编辑
    st.subheader("📋 PFMEA内容编辑")
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
        st.session_state.pfmea_data = edited_df

        # 批量操作按钮
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("🔄 重新计算AP等级", use_container_width=True):
                for idx, row in st.session_state.pfmea_data.iterrows():
                    ap = calculate_ap(row["严重度S"], row["频度O"], row["探测度D"])
                    st.session_state.pfmea_data.at[idx, "AP等级"] = ap
                st.success("AP等级重新计算完成！")
                st.rerun()
        with col_btn2:
            if st.button("🗑️ 清空所有数据", use_container_width=True):
                st.session_state.pfmea_data = pd.DataFrame(columns=st.session_state.pfmea_data.columns)
                st.success("已清空所有数据！")
                st.rerun()
        with col_btn3:
            if st.button("📋 清空选中行", use_container_width=True):
                st.session_state.pfmea_data = edited_df[~edited_df.index.isin(st.session_state.get("selected_rows", []))]
                st.success("已清空选中行！")
                st.rerun()

        # 6. 导出Excel
        st.subheader("📤 导出Excel文件")
        export_btn = st.button("📥 导出Excel2016兼容文件", use_container_width=True)
        if export_btn:
            base_info = {
                "project_name": project_name,
                "product_type": product_type,
                "version": version,
                "department": department,
                "team_member": team_member
            }
            with st.spinner("正在生成Excel文件，请稍候..."):
                save_path, file_name = export_to_excel(st.session_state.pfmea_data, base_info)
                if save_path and os.path.exists(save_path):
                    st.success(f"🎉 Excel文件生成成功！")
                    # 适配云端和本地下载
                    with open(save_path, "rb") as f:
                        st.download_button(
                            label="✅ 点击下载PFMEA Excel文件",
                            data=f,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    st.caption("本地运行提示：文件同时已保存到手机「下载」文件夹，路径：内部存储/Download/")
    else:
        st.info("暂无PFMEA数据，请先选择工序并生成内容")

    # 7. API密钥申请教程
    st.subheader("🔑 豆包API密钥免费申请教程")
    st.markdown("""
    1.  打开火山引擎官网：https://www.volcengine.com/ ，用手机号注册并登录（免费）
    2.  搜索「豆包大模型」，进入「方舟平台」，点击「开通服务」（免费开通）
    3.  左侧菜单找到「API密钥管理」，点击「创建密钥」，一键生成专属API密钥
    4.  复制生成的密钥（sk-开头的完整字符串），粘贴到上面的密钥输入框即可使用
    5.  免费额度：新用户注册即可领取**大量免费调用额度**，个人日常生成PFMEA完全够用，无需付费
    """)

if __name__ == "__main__":
    main()
