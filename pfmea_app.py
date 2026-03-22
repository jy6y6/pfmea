import streamlit as st
import pandas as pd
import requests
import json
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from datetime import datetime
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

# ===================== 核心配置（已内置完成，无需修改）=====================
# 内置API密钥，开箱即用
API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
API_ENDPOINTS = [
    "https://api.doubao.com/v1/chat/completions",
    "https://api.doubaoai.com/v1/chat/completions",
    "https://open.doubao.com/v1/chat/completions"
]
# 系统基础配置
SYSTEM_NAME = "电池包/充电器PFMEA智能生成系统"
STANDARD = "AIAG-VDA FMEA 第一版 | IATF16949:2016"

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
    if "generated_pfmea_data" not in st.session_state:
        st.session_state.generated_pfmea_data = {}
    if "ai_schemes" not in st.session_state:
        st.session_state.ai_schemes = {}
    if "current_product" not in st.session_state:
        st.session_state.current_product = "电池包"
    if "upload_status" not in st.session_state:
        st.session_state.upload_status = None
    if "custom_process" not in st.session_state:
        st.session_state.custom_process = ""

# 网络请求增强（解决域名解析失败）
def create_retry_session():
    session = requests.Session()
    retry = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["POST"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

def generate_pfmea_ai(process_name, product_type, scheme_count=3):
    session = create_retry_session()
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
    4. S/O/D评分严格符合AIAG-VDA评分标准：严重度S(1-10)、频度O(1-10)、探测度D(1-10)，AP等级为高/中/低
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
    6. 禁止返回任何JSON以外的内容，禁止注释、解释、markdown格式，确保JSON可直接解析
    7. 禁止使用本地标准库中的重复内容，必须生成全新的、差异化的PFMEA条目
    """
    
    data = {
        "model": "doubao-pro-32k",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "top_p": 0.9,
        "max_tokens": 4000
    }
    
    last_error = None
    for endpoint in API_ENDPOINTS:
        try:
            response = session.post(endpoint, headers=headers, json=data, timeout=60)
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
            last_error = f"Endpoint {endpoint} 失败: {str(e)}"
            continue
    return None, f"所有API端点均失败，最后错误: {last_error}"

def parse_pfmea_knowledge(file_content, file_name):
    session = create_retry_session()
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
    3. 过滤无效、重复、不符合标准的内容，补充缺失的S/O/D评分和AP等级，确保符合AIAG-VDA标准
    4. 确保所有内容贴合电池包/充电器装配场景，可直接用于PFMEA生成
    5. 返回严格的JSON格式，外层是对象，key为工序名称，value为该工序下的PFMEA条目数组
    待解析的PFMEA文件内容：
    {file_content}
    """
    
    data = {
        "model": "doubao-pro-32k",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
        "max_tokens": 8000
    }
    
    last_error = None
    for endpoint in API_ENDPOINTS:
        try:
            response = session.post(endpoint, headers=headers, json=data, timeout=120)
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
            last_error = f"Endpoint {endpoint} 失败: {str(e)}"
            continue
    return None, f"所有API端点均失败，最后错误: {last_error}"

def export_pfmea_excel(pfmea_data, product_type):
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
    ws["A1"] = f"{product_type} 全工序PFMEA汇总"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("A2:K2")
    ws["A2"] = f"符合标准：{STANDARD} | 生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A2"].font = Font(name="微软雅黑", size=10)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    
    headers = [
        "序号", "过程步骤/工序", "失效模式", "失效后果", "失效原因",
        "预防措施", "探测措施", "严重度S", "频度O", "探测度D", "AP等级"
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = alignment
        cell.border = border
    
    row_num = 5
    for process_name, items in pfmea_data.items():
        for item in items:
            ws.cell(row=row_num, column=1, value=row_num-4).font = content_font
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
    
    column_widths = [6, 20, 25, 30, 30, 35, 35, 8, 8, 8, 8]
    for col_num, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64+col_num)].width = width
    
    ws.freeze_panes = "A5"
    wb.save(output)
    output.seek(0)
    return output

# ===================== 3. 界面布局与主逻辑 =====================
def main():
    st.set_page_config(
        page_title=SYSTEM_NAME,
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    init_session_state()
    
    st.sidebar.title(SYSTEM_NAME)
    st.sidebar.markdown(f"**符合标准：** {STANDARD}")
    st.sidebar.divider()
    menu = st.sidebar.radio("功能导航", ["批量PFMEA生成", "我的知识库管理"])
    st.sidebar.divider()
    st.sidebar.markdown("**内置密钥已配置，开箱即用**")
    st.sidebar.markdown("**兼容：本地Termux | Streamlit云端 | Excel2016**")
    
    if menu == "批量PFMEA生成":
        st.title("⚡ 电池包/充电器PFMEA批量生成系统")
        st.divider()
        
        st.subheader("第一步：基础参数设置")
        col1, col2 = st.columns(2)
        with col1:
            product_type = st.radio("产品类型", ["电池包", "充电器"], index=0)
        with col2:
            generate_mode = st.radio("生成模式", ["本地专业标准库", "豆包AI智能生成"], index=1)
        
        st.divider()
        st.subheader("第二步：选择/自定义工序（支持多选）")
        process_lib = BATTERY_PROCESS_LIB if product_type == "电池包" else CHARGER_PROCESS_LIB
        all_process = list(process_lib.keys())
        if st.session_state.user_knowledge_base:
            user_process = list(st.session_state.user_knowledge_base.keys())
            all_process = list(set(all_process + user_process))
            all_process.sort()
        
        col3, col4 = st.columns([3, 1])
        with col3:
            selected_processes = st.multiselect("选择工序（可多选）", all_process, default=all_process[:2])
        with col4:
            st.text_input("自定义工序名称", key="custom_process")
            if st.button("➕ 添加自定义工序") and st.session_state.custom_process:
                if st.session_state.custom_process not in selected_processes:
                    selected_processes.append(st.session_state.custom_process)
                    st.success(f"已添加自定义工序：{st.session_state.custom_process}")
                    st.session_state.custom_process = ""
                    st.rerun()
        
        st.divider()
        st.subheader("第三步：AI生成设置（仅AI模式生效）")
        scheme_count = 3
        mix_user_knowledge = False
        if generate_mode == "豆包AI智能生成":
            col5, col6 = st.columns(2)
            with col5:
                scheme_count = st.slider("AI生成方案数量", min_value=2, max_value=5, value=3, step=1)
            with col6:
                mix_user_knowledge = st.checkbox("混合我的知识库内容生成", value=False)
        
        st.divider()
        generate_btn = st.button(f"🚀 批量生成 {len(selected_processes)} 个工序的PFMEA", type="primary", use_container_width=True)
        
        if generate_btn and selected_processes:
            st.session_state.generated_pfmea_data = {}
            st.session_state.ai_schemes = {}
            st.session_state.current_product = product_type
            
            with st.spinner(f"正在批量生成 {len(selected_processes)} 个工序的PFMEA，请稍候..."):
                if generate_mode == "本地专业标准库":
                    for process_name in selected_processes:
                        standard_content = process_lib.get(process_name, [])
                        user_content = st.session_state.user_knowledge_base.get(process_name, [])
                        st.session_state.generated_pfmea_data[process_name] = standard_content + user_content
                    st.success(f"✅ 本地标准库PFMEA批量生成完成！共 {len(selected_processes)} 个工序")
                
                else:
                    failed_processes = []
                    for process_name in selected_processes:
                        st.info(f"正在生成工序：{process_name}")
                        schemes, error_msg = generate_pfmea_ai(process_name, product_type, scheme_count)
                        if error_msg:
                            failed_processes.append(f"{process_name}：{error_msg}")
                            st.session_state.generated_pfmea_data[process_name] = []
                        else:
                            if mix_user_knowledge and process_name in st.session_state.user_knowledge_base:
                                user_content = st.session_state.user_knowledge_base[process_name]
                                schemes.append({
                                    "方案名称": "我的知识库方案",
                                    "pfmea_list": user_content
                                })
                            st.session_state.ai_schemes[process_name] = schemes
                            st.session_state.generated_pfmea_data[process_name] = schemes[0]["pfmea_list"]
                    
                    if failed_processes:
                        st.error("⚠️ 以下工序AI生成失败：\n" + "\n".join(failed_processes))
                    else:
                        st.success(f"✅ AI多方案批量生成完成！共 {len(selected_processes)} 个工序")
        
        st.divider()
        if st.session_state.generated_pfmea_data:
            st.subheader("第四步：PFMEA Excel预览与调整")
            current_product = st.session_state.current_product
            pfmea_data = st.session_state.generated_pfmea_data
            
            if generate_mode == "豆包AI智能生成" and st.session_state.ai_schemes:
                st.info("AI模式：可切换不同工序的生成方案")
                process_to_switch = st.selectbox("选择要切换方案的工序", list(st.session_state.ai_schemes.keys()))
                schemes = st.session_state.ai_schemes[process_to_switch]
                scheme_names = [scheme["方案名称"] for scheme in schemes]
                selected_scheme_name = st.radio(f"为【{process_to_switch}】选择方案", scheme_names, index=0)
                selected_scheme = next(scheme for scheme in schemes if scheme["方案名称"] == selected_scheme_name)
                if st.button("✅ 确认切换此方案", key=f"switch_{process_to_switch}"):
                    st.session_state.generated_pfmea_data[process_to_switch] = selected_scheme["pfmea_list"]
                    st.success(f"✅ 已切换【{process_to_switch}】至方案：{selected_scheme_name}")
                    st.rerun()
            
            st.divider()
            st.subheader("全工序PFMEA预览（可在线编辑）")
            all_data = []
            for process_name, items in pfmea_data.items():
                for item in items:
                    all_data.append({"工序": process_name, **item})
            if all_data:
                df = pd.DataFrame(all_data)
                edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
                
                updated_pfmea_data = {}
                for _, row in edited_df.iterrows():
                    process_name = row["工序"]
                    if process_name not in updated_pfmea_data:
                        updated_pfmea_data[process_name] = []
                    item = {k: v for k, v in row.items() if k != "工序"}
                    updated_pfmea_data[process_name].append(item)
                st.session_state.generated_pfmea_data = updated_pfmea_data
            
            st.divider()
            st.subheader("第五步：导出全工序Excel文件")
            if pfmea_data:
                excel_file = export_pfmea_excel(pfmea_data, current_product)
                st.download_button(
                    label="📥 下载全工序PFMEA汇总Excel",
                    data=excel_file,
                    file_name=f"{current_product}_全工序PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.warning("暂无PFMEA数据可导出")
    
    elif menu == "我的知识库管理":
        st.title("📚 我的PFMEA知识库管理")
        st.markdown("支持上传您现场编写的旧PFMEA Excel文件，AI自动分析筛选入库，生成时可直接调用")
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
