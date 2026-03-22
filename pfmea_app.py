import streamlit as st
import pandas as pd
import requests
import json
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime

# ===================== 核心配置（已内置完成，无需修改）=====================
# 内置API密钥，开箱即用
API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
API_ENDPOINT = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"  # 修复域名解析问题
SYSTEM_NAME = "电池包/充电器PFMEA智能生成系统"
STANDARD = "AIAG-VDA FMEA 第一版 | IATF16949:2016"

# ===================== 淡绿色主题配置 =====================
st.set_page_config(
    page_title=SYSTEM_NAME,
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS：淡绿色风格
st.markdown("""
<style>
    /* 全局背景 */
    .stApp {
        background-color: #f0f9f4;
    }
    /* 侧边栏背景 */
    .css-1d391kg {
        background-color: #e6f7ef;
    }
    /* 按钮样式 */
    .stButton>button {
        background-color: #52c41a;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 20px;
        font-weight: 500;
    }
    .stButton>button:hover {
        background-color: #73d13d;
        color: white;
    }
    /* 单选框/复选框选中颜色 */
    .stRadio>div>div>div>div>div {
        background-color: #52c41a;
    }
    .stCheckbox>div>div>div>div>div {
        background-color: #52c41a;
    }
    /* 滑块颜色 */
    .stSlider>div>div>div>div>div {
        background-color: #52c41a;
    }
    /* 标题颜色 */
    h1, h2, h3 {
        color: #23856d;
    }
    /* 成功提示框 */
    .stAlert {
        background-color: #f6ffed;
        border: 1px solid #b7eb8f;
        color: #23856d;
    }
    /* 警告提示框 */
    .stAlert.warning {
        background-color: #fffbe6;
        border: 1px solid #ffe58f;
        color: #d48806;
    }
</style>
""", unsafe_allow_html=True)

# ===================== 1. 全工序专业本地标准库（符合审核要求）=====================
BATTERY_PROCESS_LIB = {
    "电芯来料检验": [
        {
            "失效模式": "电芯外观尺寸超差",
            "失效后果": "电芯无法装入模组壳体，导致装配中断，产生返工成本",
            "失效原因": "来料尺寸公差不符合图纸要求，检验量具未定期校准",
            "预防措施": "制定电芯来料检验规范，每批次抽取样件全尺寸检测，量具定期校准并记录",
            "探测措施": "首件全尺寸检验，巡检按AQL抽样标准检测，超差件隔离标识",
            "严重度S": 6,
            "频度O": 3,
            "探测度D": 4,
            "AP等级": "中"
        },
        {
            "失效模式": "电芯电压/内阻异常",
            "失效后果": "模组充放电异常，循环寿命衰减过快，严重时引发热失控风险",
            "失效原因": "电芯生产过程工艺异常，来料存储环境温湿度不符合要求",
            "预防措施": "每批次电芯进行电压、内阻全检，存储环境温湿度24小时监控记录",
            "探测措施": "自动化检测设备100%全检，异常数据自动报警隔离，数据可追溯",
            "严重度S": 9,
            "频度O": 2,
            "探测度D": 2,
            "AP等级": "高"
        }
    ],
    "模组堆叠装配": [
        {
            "失效模式": "电芯堆叠顺序错误、极性反向",
            "失效后果": "模组电路连接错误，充放电功能失效，严重时引发短路烧毁",
            "失效原因": "作业人员未按SOP操作，防错装置失效，首件检验未执行",
            "预防措施": "制定极性防错SOP，安装极性视觉防错装置，作业人员岗前培训考核",
            "探测措施": "首件极性全检，过程中视觉设备100%检测，异常自动停机报警",
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
            "失效模式": "PCB板尺寸、孔位超差",
            "失效后果": "PCB无法装入壳体，元器件安装错位，装配中断",
            "失效原因": "PCB生产制程偏差，来料检验规范未执行，量具未校准",
            "预防措施": "制定PCB来料检验规范，每批次首件全尺寸检测，量具定期校准",
            "探测措施": "首件全尺寸检验，巡检按AQL抽样，超差件隔离返工",
            "严重度S": 5,
            "频度O": 3,
            "探测度D": 4,
            "AP等级": "中"
        }
    ],
    "SMT贴片焊接": [
        {
            "失效模式": "元器件贴装偏移、错件、漏件",
            "失效后果": "电路功能失效，产品无法正常工作，批量返工成本",
            "失效原因": "贴片机程序错误，元器件料盘上错，吸嘴磨损定位偏差",
            "预防措施": "贴片机程序首件验证，上料双人复核，设备定期维护保养",
            "探测措施": "首件全项核对，SPI锡膏检测，AOI光学100%检测，异常报警",
            "严重度S": 7,
            "频度O": 2,
            "探测度D": 2,
            "AP等级": "高"
        }
    ]
}

# ===================== 2. 核心工具函数 =====================
def init_session_state():
    if "user_knowledge_base" not in st.session_state:
        st.session_state.user_knowledge_base = {}
    if "batch_pfmea_data" not in st.session_state:
        st.session_state.batch_pfmea_data = {}
    if "ai_schemes" not in st.session_state:
        st.session_state.ai_schemes = {}
    if "custom_process" not in st.session_state:
        st.session_state.custom_process = ""

def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    prompt = f"""
    你是专业的汽车电子行业PFMEA工程师，精通AIAG-VDA FMEA标准和IATF16949质量管理体系要求，专注于{product_type}装配制造场景。
    请针对【{process_name}】工序，生成{scheme_count}组完全不同、无重复内容的PFMEA方案，每组方案包含3-5条独立的PFMEA条目。
    严格遵守以下要求：
    1. 每组方案必须有明显差异化：分别从人、机、料、法、环、测不同维度切入，失效模式、失效后果、失效原因、预防/探测措施完全不同，禁止内容重复
    2. 所有内容必须严格贴合{product_type}装配现场的实际作业场景，禁止通用化、理论化内容
    3. 严格遵循AIAG-VDA FMEA标准，失效链必须完整：失效模式→失效后果→失效原因→预防措施→探测措施
    4. S/O/D评分严格符合AIAG-VDA评分标准：严重度S(1-10)、频度O(1-10)、探测度D(1-10)
    5. AP等级严格按S/O/D评分判定：高/中/低三个等级
    6. 必须返回严格的JSON格式，外层是一个数组，每个元素是一组方案，格式如下：
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
    7. 禁止返回任何JSON以外的内容，禁止注释、解释、 markdown格式，确保JSON可直接解析
    """
    data = {
        "model": "ep-20250218",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "top_p": 0.9,
        "max_tokens": 4000
    }
    try:
        response = requests.post(API_ENDPOINT, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        result = response.json()
        ai_content = result["choices"][0]["message"]["content"]
        ai_content = ai_content.strip()
        if ai_content.startswith("```json"):
            ai_content = ai_content[7:]
        if ai_content.endswith("```"):
            ai_content = ai_content[:-3]
        ai_content = ai_content.strip()
        schemes = json.loads(ai_content)
        return schemes, None
    except Exception as e:
        error_msg = f"AI生成失败：{str(e)}，已自动切换为本地专业标准库内容"
        local_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
        local_content = local_lib.get(process_name, [])
        fallback_schemes = [
            {
                "方案名称": "本地标准库方案（AI生成失败兜底）",
                "pfmea_list": local_content
            }
        ]
        return fallback_schemes, error_msg

def parse_pfmea_knowledge(file_content, file_name):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    prompt = f"""
    你是专业的PFMEA工程师，精通IATF16949和AIAG-VDA FMEA标准。
    请分析用户上传的PFMEA文件内容，提取所有符合标准的PFMEA条目，完成以下处理：
    1. 按工序名称分类，提取每个工序下的所有PFMEA条目
    2. 每个条目必须整理成标准格式，包含以下字段：
       失效模式、失效后果、失效原因、预防措施、探测措施、严重度S、频度O、探测度D、AP等级
    3. 过滤无效、重复、不符合标准的内容，补充缺失的S/O/D评分和AP等级
    4. 确保所有内容贴合电池包/充电器装配场景
    5. 返回严格的JSON格式，外层是对象，key为工序名称，value为该工序下的PFMEA条目数组
    待解析的PFMEA文件内容：
    {file_content}
    """
    data = {
        "model": "ep-20250218",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
        "max_tokens": 8000
    }
    try:
        response = requests.post(API_ENDPOINT, headers=headers, json=data, timeout=60)
        response.raise_for_status()
        result = response.json()
        ai_content = result["choices"][0]["message"]["content"]
        ai_content = ai_content.strip()
        if ai_content.startswith("```json"):
            ai_content = ai_content[7:]
        if ai_content.endswith("```"):
            ai_content = ai_content[:-3]
        ai_content = ai_content.strip()
        knowledge_data = json.loads(ai_content)
        return knowledge_data, None
    except Exception as e:
        return None, f"知识库解析失败：{str(e)}"

def export_pfmea_excel(batch_pfmea_data, product_type):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "PFMEA汇总"
    title_font = Font(name="微软雅黑", bold=True, size=14)
    header_font = Font(name="微软雅黑", bold=True, size=10)
    content_font = Font(name="微软雅黑", size=10)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    ws.merge_cells("A1:K1")
    ws["A1"] = f"{product_type} 全工序PFMEA汇总表"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A2:K2")
    ws["A2"] = f"符合标准：{STANDARD} | 生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A2"].font = Font(name="微软雅黑", size=10)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    headers = [
        "序号", "工序名称", "失效模式", "失效后果", "失效原因",
        "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = alignment
        cell.border = border
    row_num = 5
    seq = 1
    for process_name, pfmea_list in batch_pfmea_data.items():
        for item in pfmea_list:
            ws.cell(row=row_num, column=1, value=seq).font = content_font
            ws.cell(row=row_num, column=2, value=process_name).font = content_font
            ws.cell(row=row_num, column=3, value=item["失效模式"]).font = content_font
            ws.cell(row=row_num, column=4, value=item["失效后果"]).font = content_font
            ws.cell(row=row_num, column=5, value=item["失效原因"]).font = content_font
            ws.cell(row=row_num, column=6, value=item["预防措施"]).font = content_font
            ws.cell(row=row_num, column=7, value=item["探测措施"]).font = content_font
            ws.cell(row=row_num, column=8, value=item["严重度S"]).font = content_font
            ws.cell(row=row_num, column=9, value=item["频度O"]).font = content_font
            ws.cell(row=row_num, column=10, value=item["探测度D"]).font = content_font
            ws.cell(row=row_num, column=11, value=item["AP等级"]).font = content_font
            for col_num in range(1, 12):
                cell = ws.cell(row=row_num, column=col_num)
                cell.alignment = alignment
                cell.border = border
            row_num += 1
            seq += 1
    column_widths = [6, 20, 25, 30, 30, 35, 35, 8, 8, 8, 8]
    for col_num, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64+col_num)].width = width
    ws.freeze_panes = "A5"
    wb.save(output)
    output.seek(0)
    return output

# ===================== 3. 界面布局与主逻辑 =====================
def main():
    init_session_state()
    st.sidebar.title("⚡ " + SYSTEM_NAME)
    st.sidebar.markdown(f"**符合标准：** {STANDARD}")
    st.sidebar.divider()
    menu = st.sidebar.radio("功能导航", ["批量PFMEA生成", "我的知识库管理"])
    st.sidebar.divider()
    st.sidebar.markdown("**内置密钥已配置，开箱即用**")
    st.sidebar.markdown("**兼容：本地Termux | Streamlit云端 | Excel2016**")
    if menu == "批量PFMEA生成":
        st.title("🌿 批量PFMEA智能生成")
        st.divider()
        st.subheader("第一步：基础参数设置")
        col1, col2 = st.columns(2)
        with col1:
            product_type = st.radio("产品类型", ["电池包", "充电器"], index=0)
        with col2:
            generate_mode = st.radio("生成模式", ["本地专业标准库", "豆包AI智能生成"], index=1)
        st.divider()
        st.subheader("第二步：选择/自定义工序（可多选）")
        process_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
        all_process = list(process_lib.keys())
        if st.session_state.user_knowledge_base:
            user_process = list(st.session_state.user_knowledge_base.keys())
            all_process = list(set(all_process + user_process))
            all_process.sort()
        selected_processes = st.multiselect("选择要生成的工序", all_process, default=all_process[:2])
        st.markdown("**或自定义新工序**")
        custom_process = st.text_input("输入自定义工序名称", placeholder="例如：电池包气密性测试")
        if custom_process and custom_process not in selected_processes:
            if st.button("➕ 添加自定义工序"):
                selected_processes.append(custom_process)
                st.success(f"已添加自定义工序：{custom_process}")
        st.divider()
        scheme_count = 3
        mix_user_knowledge = False
        if generate_mode == "豆包AI智能生成":
            st.subheader("第三步：AI生成参数设置")
            col3, col4 = st.columns(2)
            with col3:
                scheme_count = st.slider("AI生成方案数量", min_value=2, max_value=5, value=3, step=1)
            with col4:
                mix_user_knowledge = st.checkbox("混合我的知识库内容生成", value=False)
        st.divider()
        if st.button("🚀 开始批量生成PFMEA", type="primary", use_container_width=True):
            st.session_state.batch_pfmea_data = {}
            st.session_state.ai_schemes = {}
            with st.spinner("正在批量生成PFMEA内容，请稍候..."):
                for process_name in selected_processes:
                    if generate_mode == "本地专业标准库":
                        standard_content = process_lib.get(process_name, [])
                        user_content = st.session_state.user_knowledge_base.get(process_name, [])
                        final_content = standard_content + user_content
                        st.session_state.batch_pfmea_data[process_name] = final_content
                    else:
                        schemes, error_msg = generate_pfmea_ai(process_name, product_type, scheme_count)
                        if error_msg:
                            st.warning(f"{process_name}：{error_msg}")
                        if mix_user_knowledge and process_name in st.session_state.user_knowledge_base:
                            user_content = st.session_state.user_knowledge_base[process_name]
                            schemes.append({
                                "方案名称": "我的知识库方案",
                                "pfmea_list": user_content
                            })
                        st.session_state.ai_schemes[process_name] = schemes
                st.success("✅ 批量生成完成！请选择各工序方案后导出")
        st.divider()
        if generate_mode == "豆包AI智能生成" and st.session_state.ai_schemes:
            st.subheader("第三步：各工序方案选择")
            for process_name, schemes in st.session_state.ai_schemes.items():
                with st.expander(f"📦 工序：{process_name}", expanded=False):
                    scheme_names = [scheme["方案名称"] for scheme in schemes]
                    selected_scheme_name = st.radio(f"选择{process_name}的PFMEA方案", scheme_names, index=0, key=f"scheme_{process_name}")
                    selected_scheme = next(scheme for scheme in schemes if scheme["方案名称"] == selected_scheme_name)
                    st.dataframe(pd.DataFrame(selected_scheme["pfmea_list"]), use_container_width=True)
                    st.session_state.batch_pfmea_data[process_name] = selected_scheme["pfmea_list"]
        st.divider()
        if st.session_state.batch_pfmea_data:
            st.subheader("第四步：全工序PFMEA预览与导出")
            all_data = []
            for process_name, pfmea_list in st.session_state.batch_pfmea_data.items():
                for item in pfmea_list:
                    item["工序名称"] = process_name
                    all_data.append(item)
            df = pd.DataFrame(all_data)
            edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            st.divider()
            excel_file = export_pfmea_excel(st.session_state.batch_pfmea_data, product_type)
            st.download_button(
                label="📥 下载全工序PFMEA汇总Excel",
                data=excel_file,
                file_name=f"{product_type}_全工序PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
    elif menu == "我的知识库管理":
        st.title("📚 我的PFMEA知识库管理")
        st.markdown("支持上传您现场编写的旧PFMEA文件，AI自动分析筛选入库，生成时可直接调用")
        st.divider()
        st.subheader("上传旧PFMEA文件入库")
        uploaded_file = st.file_uploader("请上传PFMEA Excel文件（.xlsx/.xls格式）", type=["xlsx", "xls"])
        if uploaded_file:
            with st.spinner("正在读取并解析文件内容，请稍候..."):
                try:
                    df_list = pd.read_excel(uploaded_file, sheet_name=None)
                    file_content = ""
                    for sheet_name, df in df_list.items():
                        file_content += f"===== 工作表：{sheet_name} =====\n"
                        file_content += df.to_string(index=False)
                        file_content += "\n\n"
                    knowledge_data, error_msg = parse_pfmea_knowledge(file_content, uploaded_file.name)
                    if error_msg:
                        st.error(error_msg)
                    else:
                        st.success("✅ 文件解析完成！请确认要入库的内容")
                        with st.expander("📄 解析结果预览", expanded=True):
                            for process_name, pfmea_list in knowledge_data.items():
                                st.markdown(f"**工序名称：{process_name}**")
                                st.dataframe(pd.DataFrame(pfmea_list), use_container_width=True)
                                st.divider()
                        if st.button("✅ 确认入库", type="primary", use_container_width=True):
                            for process_name, pfmea_list in knowledge_data.items():
                                if process_name in st.session_state.user_knowledge_base:
                                    existing_list = st.session_state.user_knowledge_base[process_name]
                                    existing_keys = set(f"{item['失效模式']}_{item['失效原因']}" for item in existing_list)
                                    for item in pfmea_list:
                                        item_key = f"{item['失效模式']}_{item['失效原因']}"
                                        if item_key not in existing_keys:
                                            existing_list.append(item)
                                    st.session_state.user_knowledge_base[process_name] = existing_list
                                else:
                                    st.session_state.user_knowledge_base[process_name] = pfmea_list
                            st.success("✅ 知识库入库完成！生成PFMEA时可直接调用")
                            st.rerun()
                except Exception as e:
                    st.error(f"文件读取失败：{str(e)}，请检查文件格式是否正确")
        st.divider()
        st.subheader("我的知识库内容")
        user_kb = st.session_state.user_knowledge_base
        if not user_kb:
            st.info("您的知识库暂无内容，请上传PFMEA文件入库")
        else:
            for process_name, pfmea_list in user_kb.items():
                with st.expander(f"📦 工序：{process_name}（共{len(pfmea_list)}条PFMEA）", expanded=False):
                    edited_df = st.data_editor(
                        pd.DataFrame(pfmea_list),
                        use_container_width=True,
                        num_rows="dynamic",
                        key=f"edit_{process_name}"
                    )
                    col_update, col_delete = st.columns(2)
                    with col_update:
                        if st.button("✅ 更新内容", key=f"update_{process_name}", use_container_width=True):
                            st.session_state.user_knowledge_base[process_name] = edited_df.to_dict("records")
                            st.success("内容更新成功！")
                            st.rerun()
                    with col_delete:
                        if st.button("🗑️ 删除此工序", key=f"delete_{process_name}", use_container_width=True, type="secondary"):
                            del st.session_state.user_knowledge_base[process_name]
                            st.success("工序删除成功！")
                            st.rerun()
                st.divider()
            st.subheader("知识库备份与恢复")
            col_export, col_import = st.columns(2)
            with col_export:
                kb_json = json.dumps(user_kb, ensure_ascii=False, indent=2)
                st.download_button(
                    label="📤 导出知识库备份文件",
                    data=kb_json,
                    file_name=f"PFMEA知识库备份_{datetime.now().strftime('%Y%m%d%H%M%S')}.json",
                    mime="application/json",
                    use_container_width=True
                )
            with col_import:
                import_file = st.file_uploader("导入知识库备份文件", type=["json"], label_visibility="collapsed")
                if import_file:
                    try:
                        import_data = json.load(import_file)
                        if st.button("✅ 确认导入", use_container_width=True):
                            for process_name, pfmea_list in import_data.items():
                                if process_name in st.session_state.user_knowledge_base:
                                    existing_list = st.session_state.user_knowledge_base[process_name]
                                    existing_keys = set(f"{item['失效模式']}_{item['失效原因']}" for item in existing_list)
                                    for item in pfmea_list:
                                        item_key = f"{item['失效模式']}_{item['失效原因']}"
                                        if item_key not in existing_keys:
                                            existing_list.append(item)
                                    st.session_state.user_knowledge_base[process_name] = existing_list
                                else:
                                    st.session_state.user_knowledge_base[process_name] = pfmea_list
                            st.success("✅ 知识库导入完成！")
                            st.rerun()
                    except Exception as e:
                        st.error(f"导入失败：{str(e)}，请检查备份文件格式是否正确")

if __name__ == "__main__":
    main()
