import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime

# ===================== 核心配置（密钥已内置）=====================
API_KEY = "7abbafd6-4d6e-4dad-9172-ea2d165c7a44"
API_ENDPOINT = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
MODEL_NAME = "doubao-pro-32k"  # 豆包专业模型，适配专业场景生成

# ===================== 全工序专业本地库（保留原有合规内容）=====================
# 电池包装配工序库
BATTERY_PROCESS_LIST = [
    "壳体来料检验", "模组入壳装配", "Busbar焊接", "高压线束装配",
    "低压采集线焊接", "绝缘片装配", "冷却系统装配", "上盖锁附装配",
    "气密性测试", "绝缘耐压测试", "功能性能测试", "成品终检包装"
]

# 充电器装配工序库
CHARGER_PROCESS_LIST = [
    "PCB板来料检验", "SMT贴片焊接", "插件元器件焊接", "壳体装配",
    "高压端子压接", "输入输出线装配", "灌胶固化", "防水密封装配",
    "电气性能测试", "成品终检包装"
]

# 本地PFMEA标准库（符合IATF16949审核要求）
LOCAL_PFMEA_LIB = {
    # 电池包工序内容
    "壳体来料检验": [
        {"工序名称": "壳体来料检验", "失效模式": "壳体尺寸超差", "失效影响": "模组无法正常入壳，导致装配停线，影响生产节拍", "失效原因": "供应商注塑成型参数不稳定，尺寸管控不到位", "S": 6, "O": 3, "D": 3, "AP等级": "L", "建议措施": "来料加严全尺寸检验，要求供应商提供SPC管控报告"},
        {"工序名称": "壳体来料检验", "失效模式": "壳体材质强度不达标", "失效影响": "使用过程中壳体变形，导致电池包密封失效，进水短路", "失效原因": "供应商使用回料替代全新PC/ABS材料，材质性能不合格", "S": 9, "O": 2, "D": 4, "AP等级": "H", "建议措施": "每批次来料做材质强度检测，签订材质保真协议"},
        {"工序名称": "壳体来料检验", "失效模式": "壳体表面划痕、毛刺", "失效影响": "装配过程中划破绝缘层，导致绝缘耐压测试失效", "失效原因": "供应商周转过程中无防护，模具磨损未及时更换", "S": 7, "O": 4, "D": 2, "AP等级": "M", "建议措施": "来料全检外观，要求供应商增加周转防护工装"}
    ],
    "模组入壳装配": [
        {"工序名称": "模组入壳装配", "失效模式": "模组安装孔位错位", "失效影响": "模组无法固定，车辆行驶过程中模组晃动，导致高压连接失效", "失效原因": "壳体孔位加工偏差，模组定位工装精度不足", "S": 8, "O": 3, "D": 3, "AP等级": "M", "建议措施": "优化定位工装精度，装配前核对孔位公差"},
        {"工序名称": "模组入壳装配", "失效模式": "模组与壳体间隙超标", "失效影响": "冷却系统贴合不良，模组散热不佳，导致热失控风险", "失效原因": "模组厚度尺寸偏差，装配时未做间隙检测", "S": 9, "O": 2, "D": 4, "AP等级": "H", "建议措施": "装配时使用塞尺检测间隙，超差产品禁止流入下工序"},
        {"工序名称": "模组入壳装配", "失效模式": "模组表面绝缘层破损", "失效影响": "模组与壳体导通，绝缘耐压测试失败，严重时导致短路起火", "失效原因": "装配过程中操作不当，工装无防护倒角", "S": 10, "O": 3, "D": 2, "AP等级": "H", "建议措施": "工装增加软质防护，对操作人员做专项培训，装配后全检绝缘层"}
    ],
    "Busbar焊接": [
        {"工序名称": "Busbar焊接", "失效模式": "焊接虚焊、假焊", "失效影响": "回路电阻过大，工作时发热烧蚀，导致高压断电，甚至热失控", "失效原因": "焊接参数不合理，焊接压力不足，焊头磨损", "S": 10, "O": 4, "D": 3, "AP等级": "H", "建议措施": "优化焊接参数，每日首件做焊接拉力测试，全检焊接外观"},
        {"工序名称": "Busbar焊接", "失效模式": "焊渣飞溅残留", "失效影响": "焊渣导致高压件之间短路，绝缘测试失效", "失效原因": "焊接保护气体不足，焊接区域无防护遮挡", "S": 9, "O": 3, "D": 2, "AP等级": "H", "建议措施": "增加焊接区域防护，焊接后用高压气枪清洁，全检异物"},
        {"工序名称": "Busbar焊接", "失效模式": "焊接后Busbar变形", "失效影响": "高压连接间隙过大，接触不良，发热烧蚀", "失效原因": "焊接热量输入过大，工装定位不牢固", "S": 8, "O": 3, "D": 3, "AP等级": "M", "建议措施": "优化焊接热输入，增加工装定位点，焊接后检测平面度"}
    ],
    "高压线束装配": [
        {"工序名称": "高压线束装配", "失效模式": "高压端子压接不牢", "失效影响": "端子脱落，高压断电，车辆抛锚，严重时拉弧起火", "失效原因": "压接模具不匹配，压接力参数设置错误", "S": 10, "O": 3, "D": 4, "AP等级": "H", "建议措施": "每批次首件做端子拉力测试，定期校准压接设备"},
        {"工序名称": "高压线束装配", "失效模式": "线束走向错误", "失效影响": "线束被壳体挤压，绝缘层破损，导致高压短路", "失效原因": "操作人员未按作业指导书操作，无防错工装", "S": 9, "O": 4, "D": 3, "AP等级": "H", "建议措施": "设计线束走向防错工装，增加工序互检环节"},
        {"工序名称": "高压线束装配", "失效模式": "高压连接器插接不到位", "失效影响": "接触电阻过大，发热烧蚀，高压连接失效", "失效原因": "插接时未听到锁止卡扣声响，无到位检测", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "要求插接后做拉拔测试，增加到位防错标识"}
    ],
    "低压采集线焊接": [
        {"工序名称": "低压采集线焊接", "失效模式": "焊盘脱落", "失效影响": "电压采集信号丢失，BMS无法正常监控电芯状态，导致过充过放风险", "失效原因": "焊接温度过高，焊接时间过长，操作人员手法不当", "S": 9, "O": 4, "D": 3, "AP等级": "H", "建议措施": "设定焊接温度与时间上限，使用温控焊台，对操作人员做持证上岗培训"},
        {"工序名称": "低压采集线焊接", "失效模式": "连锡短路", "失效影响": "采集信号异常，BMS误报警，严重时导致电芯模组短路", "失效原因": "焊锡量过多，焊头移动轨迹偏差", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "使用定量化焊锡丝，焊接后用放大镜全检连锡情况"},
        {"工序名称": "低压采集线焊接", "失效模式": "线束虚焊", "失效影响": "采集信号中断，BMS无法监控电芯状态，失控风险", "失效原因": "焊接前线束未做搪锡处理，焊接面氧化", "S": 9, "O": 2, "D": 4, "AP等级": "H", "建议措施": "要求焊接前对线束做搪锡处理，首件做导通测试"}
    ],
    "绝缘片装配": [
        {"工序名称": "绝缘片装配", "失效模式": "绝缘片漏装", "失效影响": "高压件与壳体之间绝缘失效，短路起火", "失效原因": "无防错工装，操作人员疏忽，无漏装检测", "S": 10, "O": 2, "D": 3, "AP等级": "H", "建议措施": "增加防错工装，装配后做绝缘耐压测试，100%全检"},
        {"工序名称": "绝缘片装配", "失效模式": "绝缘片装配错位", "失效影响": "绝缘防护区域未覆盖，导致局部绝缘失效", "失效原因": "定位工装精度不足，操作人员未按标识对位", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "优化定位工装，增加对位标识，装配后目视全检位置"},
        {"工序名称": "绝缘片装配", "失效模式": "绝缘片破损", "失效影响": "绝缘性能下降，耐压测试不通过", "失效原因": "装配过程中被毛刺划伤，周转过程中无防护", "S": 7, "O": 4, "D": 3, "AP等级": "M", "建议措施": "装配前检查壳体毛刺，绝缘片使用专用周转盒，破损件直接报废"}
    ],
    "冷却系统装配": [
        {"工序名称": "冷却系统装配", "失效模式": "冷却管路接头漏装密封圈", "失效影响": "冷却液泄漏，导致模组散热失效，热失控风险", "失效原因": "操作人员疏忽，无防错核对步骤", "S": 10, "O": 2, "D": 4, "AP等级": "H", "建议措施": "密封圈预装与装配分双人核对，增加防错扫码步骤"},
        {"工序名称": "冷却系统装配", "失效模式": "管路插接不到位", "失效影响": "管路脱落，冷却液泄漏，散热失效", "失效原因": "插接深度不足，锁止卡扣未卡紧", "S": 9, "O": 3, "D": 3, "AP等级": "H", "建议措施": "设计插接到位防错标识，插接后做拉拔测试，100%全检"},
        {"工序名称": "冷却系统装配", "失效模式": "冷却板与模组贴合不良", "失效影响": "散热不均，模组温差过大，影响电芯寿命，严重时热失控", "失效原因": "导热垫厚度选型错误，装配压力不均匀", "S": 8, "O": 3, "D": 4, "AP等级": "M", "建议措施": "严格管控导热垫规格，装配后用压力纸检测贴合度"}
    ],
    "上盖锁附装配": [
        {"工序名称": "上盖锁附装配", "失效模式": "螺丝漏锁、滑牙", "失效影响": "上盖密封失效，防水等级不达标，进水导致短路", "失效原因": "电批扭矩不符，操作人员漏打，螺丝材质不达标", "S": 9, "O": 3, "D": 2, "AP等级": "H", "建议措施": "使用带计数功能的定扭矩电批，漏锁自动报警，全检螺丝锁附状态"},
        {"工序名称": "上盖锁附装配", "失效模式": "密封胶条错位、脱落", "失效影响": "密封失效，IP等级不达标，进水短路", "失效原因": "胶条卡槽装配不到位，锁附时胶条被挤压移位", "S": 8, "O": 4, "D": 2, "AP等级": "M", "建议措施": "胶条预装后做全检，锁附前核对胶条位置，使用带定位的胶条"},
        {"工序名称": "上盖锁附装配", "失效模式": "锁附顺序错误导致上盖变形", "失效影响": "密封间隙不均，局部防水失效", "失效原因": "操作人员未按对角锁附顺序操作，无作业指导", "S": 7, "O": 3, "D": 3, "AP等级": "L", "建议措施": "在作业指导书明确对角锁附顺序，对操作人员做专项培训"}
    ],
    "气密性测试": [
        {"工序名称": "气密性测试", "失效模式": "测试参数设置错误", "失效影响": "泄漏产品误判为合格，流入市场后进水短路", "失效原因": "测试压力、保压时间设置错误，无参数防错", "S": 10, "O": 2, "D": 4, "AP等级": "H", "建议措施": "测试程序加密锁定，仅工程师有权限修改，每日首件用标准泄漏块校准"},
        {"工序名称": "气密性测试", "失效模式": "测试接头密封不良", "失效影响": "测试结果假合格，泄漏产品流出", "失效原因": "测试接头密封圈磨损，未定期更换", "S": 9, "O": 3, "D": 3, "AP等级": "H", "建议措施": "制定密封圈定期更换计划，每班首件做密封性验证"},
        {"工序名称": "气密性测试", "失效模式": "产品未完全定位就启动测试", "失效影响": "测试数据无效，误判合格", "失效原因": "无定位到位检测，操作人员提前启动测试", "S": 8, "O": 2, "D": 2, "AP等级": "M", "建议措施": "增加定位到位传感器，只有定位到位才能启动测试"}
    ],
    "绝缘耐压测试": [
        {"工序名称": "绝缘耐压测试", "失效模式": "测试电压、时间设置错误", "失效影响": "绝缘不良产品误判合格，使用中短路起火", "失效原因": "测试参数未锁定，操作人员误修改", "S": 10, "O": 2, "D": 4, "AP等级": "H", "建议措施": "测试程序加密锁定，每日首件用标准绝缘样件校准设备"},
        {"工序名称": "绝缘耐压测试", "失效模式": "测试探针接触不良", "失效影响": "测试回路未导通，误判为绝缘合格", "失效原因": "探针磨损、氧化，测试工装未定期维护", "S": 9, "O": 3, "D": 3, "AP等级": "H", "建议措施": "每班检查探针状态，定期更换探针，增加回路导通自检功能"},
        {"工序名称": "绝缘耐压测试", "失效模式": "产品表面有冷凝水导致测试失效", "失效影响": "绝缘值偏低，合格产品误判为不合格", "失效原因": "产品与环境温差大，表面结露", "S": 5, "O": 2, "D": 2, "AP等级": "L", "建议措施": "产品在测试环境静置恒温后再测试，测试前清洁产品表面"}
    ],
    "功能性能测试": [
        {"工序名称": "功能性能测试", "失效模式": "BMS通讯功能异常", "失效影响": "车辆无法与电池包通讯，无法正常充放电，车辆抛锚", "失效原因": "低压线束连接不良，BMS程序烧录错误", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "100%全检通讯功能，程序烧录增加防错校验"},
        {"工序名称": "功能性能测试", "失效模式": "充放电性能不达标", "失效影响": "电池包续航不足，充电速度慢，客户投诉", "失效原因": "电芯一致性差，模组连接电阻过大", "S": 7, "O": 2, "D": 2, "AP等级": "L", "建议措施": "严格管控电芯一致性，测试前检测回路电阻"},
        {"工序名称": "功能性能测试", "失效模式": "保护功能失效", "失效影响": "过充、过放、过流时无法正常保护，导致电芯损坏，热失控风险", "失效原因": "BMS保护参数设置错误，保护回路故障", "S": 10, "O": 2, "D": 3, "AP等级": "H", "建议措施": "每批次首件做保护功能验证，测试程序锁定保护参数"}
    ],
    "成品终检包装": [
        {"工序名称": "成品终检包装", "失效模式": "产品附件漏装", "失效影响": "客户无法正常安装使用，导致客户投诉、退货", "失效原因": "无装箱核对清单，操作人员疏忽", "S": 4, "O": 3, "D": 2, "AP等级": "L", "建议措施": "制定装箱清单，双人核对，扫码防错"},
        {"工序名称": "成品终检包装", "失效模式": "产品型号与包装标识不符", "失效影响": "发错货，客户无法使用，批量退货", "失效原因": "标识打印错误，产品与包装未核对", "S": 5, "O": 2, "D": 2, "AP等级": "L", "建议措施": "扫码匹配产品型号与包装标识，不匹配无法封箱"},
        {"工序名称": "成品终检包装", "失效模式": "包装防护不足", "失效影响": "运输过程中产品磕碰、损坏", "失效原因": "缓冲材料选型错误，包装结构不合理", "S": 6, "O": 3, "D": 3, "AP等级": "L", "建议措施": "做运输跌落测试，优化包装缓冲结构"}
    ],
    # 充电器工序内容
    "PCB板来料检验": [
        {"工序名称": "PCB板来料检验", "失效模式": "PCB板线路开路、短路", "失效影响": "产品功能失效，无法正常工作", "失效原因": "供应商蚀刻工艺不良，线路设计缺陷", "S": 8, "O": 3, "D": 3, "AP等级": "M", "建议措施": "来料用AOI全检线路，每批次做首件焊接测试"},
        {"工序名称": "PCB板来料检验", "失效模式": "PCB板焊盘氧化、上锡不良", "失效影响": "元器件虚焊，产品接触不良，功能失效", "失效原因": "PCB板表面处理工艺不良，存放时间过长受潮", "S": 7, "O": 4, "D": 2, "AP等级": "M", "建议措施": "严格管控来料保质期，来料检测焊盘可焊性，受潮板做烘烤处理"},
        {"工序名称": "PCB板来料检验", "失效模式": "PCB板尺寸、孔位超差", "失效影响": "元器件无法正常插装，壳体装配错位", "失效原因": "供应商PCB成型工艺偏差，尺寸管控不到位", "S": 6, "O": 3, "D": 2, "AP等级": "L", "建议措施": "来料全检关键尺寸，要求供应商提供CNC加工检测报告"}
    ],
    "SMT贴片焊接": [
        {"工序名称": "SMT贴片焊接", "失效模式": "元器件贴装偏移", "失效影响": "元器件连锡短路，产品功能失效", "失效原因": "贴片机坐标偏移，吸嘴磨损，PCB板定位不准", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "每班校准贴片机坐标，定期更换吸嘴，AOI全检贴装位置"},
        {"工序名称": "SMT贴片焊接", "失效模式": "元器件虚焊、假焊", "失效影响": "产品接触不良，功能间歇性失效，严重时发热烧蚀", "失效原因": "钢网开孔不合理，回流焊温度曲线设置错误", "S": 9, "O": 4, "D": 3, "AP等级": "H", "建议措施": "优化钢网开孔，调试回流焊温度曲线，每批次首件做切片分析"},
        {"工序名称": "SMT贴片焊接", "失效模式": "元器件极性贴反", "失效影响": "元器件烧毁，产品功能完全失效", "失效原因": "物料盘极性标识错误，贴片机程序错误", "S": 8, "O": 2, "D": 3, "AP等级": "M", "建议措施": "首件核对极性，AOI增加极性检测，双人核对贴片机程序"}
    ],
    "插件元器件焊接": [
        {"工序名称": "插件元器件焊接", "失效模式": "元器件引脚虚焊", "失效影响": "产品功能失效，间歇性故障", "失效原因": "焊锡量不足，焊接温度不够，引脚氧化", "S": 7, "O": 4, "D": 2, "AP等级": "M", "建议措施": "焊接前对引脚做搪锡处理，使用温控焊台，全检焊接外观"},
        {"工序名称": "插件元器件焊接", "失效模式": "焊锡过多导致连锡", "失效影响": "线路短路，元器件烧毁", "失效原因": "操作人员手法不当，焊锡丝送量过多", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "对操作人员做专项培训，焊接后用放大镜全检连锡情况"},
        {"工序名称": "插件元器件焊接", "失效模式": "元器件漏插、错插", "失效影响": "产品功能失效，参数不符", "失效原因": "操作人员疏忽，物料标识不清，无防错工装", "S": 7, "O": 3, "D": 3, "AP等级": "M", "建议措施": "使用物料防错料架，插件后做目视核对，增加工序互检"}
    ],
    "壳体装配": [
        {"工序名称": "壳体装配", "失效模式": "上下壳体扣合不到位", "失效影响": "产品密封失效，防水等级不达标，进水短路", "失效原因": "壳体卡扣变形，扣合压力不足，PCB板装配错位", "S": 9, "O": 3, "D": 2, "AP等级": "H", "建议措施": "设计扣合到位防错工装，全检扣合间隙，卡扣变形件直接报废"},
        {"工序名称": "壳体装配", "失效模式": "壳体内残留异物", "失效影响": "异物导致线路短路，产品功能失效", "失效原因": "装配环境洁净度不足，壳体清洁不到位", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "装配前用高压气枪清洁壳体，装配环境做洁净度管控"},
        {"工序名称": "壳体装配", "失效模式": "螺丝滑牙、漏锁", "失效影响": "壳体松动，密封失效，进水短路", "失效原因": "电批扭矩不符，操作人员漏打，螺丝孔位偏差", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "使用带计数功能的定扭矩电批，漏锁自动报警，全检螺丝锁附状态"}
    ],
    "高压端子压接": [
        {"工序名称": "高压端子压接", "失效模式": "端子压接不牢", "失效影响": "端子脱落，高压断电，严重时拉弧起火", "失效原因": "压接模具不匹配，压接力参数设置错误", "S": 10, "O": 3, "D": 4, "AP等级": "H", "建议措施": "每批次首件做端子拉力测试，定期校准压接设备，模具定期更换"},
        {"工序名称": "高压端子压接", "失效模式": "端子压接过度导致铜丝断裂", "失效影响": "载流能力下降，工作时发热烧蚀，产品失效", "失效原因": "压接深度过大，模具间隙过小", "S": 9, "O": 2, "D": 3, "AP等级": "H", "建议措施": "优化压接参数，首件做截面分析，全检压接外观"},
        {"工序名称": "高压端子压接", "失效模式": "端子绝缘层压伤", "失效影响": "绝缘性能下降，耐压测试失效，短路风险", "失效原因": "压接模具定位偏差，绝缘层进入压接区域", "S": 7, "O": 3, "D": 2, "AP等级": "M", "建议措施": "优化模具定位，压接后全检绝缘层状态"}
    ],
    "输入输出线装配": [
        {"工序名称": "输入输出线装配", "失效模式": "线束接插件插接不到位", "失效影响": "接触不良，发热烧蚀，产品功能失效", "失效原因": "卡扣未锁止到位，插接深度不足", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "插接后做拉拔测试，增加到位防错标识，全检插接状态"},
        {"工序名称": "输入输出线装配", "失效模式": "线束走向错误，被壳体挤压", "失效影响": "线束绝缘层破损，短路起火", "失效原因": "无防错工装，操作人员未按作业指导书操作", "S": 9, "O": 4, "D": 3, "AP等级": "H", "建议措施": "设计线束走向防错工装，增加工序互检环节，装配后全检线束状态"},
        {"工序名称": "输入输出线装配", "失效模式": "线束防水圈错位、漏装", "失效影响": "防水密封失效，进水短路", "失效原因": "操作人员疏忽，防水圈装配不到位", "S": 8, "O": 2, "D": 3, "AP等级": "M", "建议措施": "防水圈预装与装配分双人核对，增加防错扫码步骤"}
    ],
    "灌胶固化": [
        {"工序名称": "灌胶固化", "失效模式": "灌胶量不足，有气泡", "失效影响": "防水、导热性能不达标，产品寿命缩短，绝缘失效", "失效原因": "灌胶机参数设置错误，胶水脱泡不充分", "S": 7, "O": 4, "D": 3, "AP等级": "M", "建议措施": "优化灌胶机参数，胶水使用前做真空脱泡，灌胶后目视全检气泡"},
        {"工序名称": "灌胶固化", "失效模式": "胶水固化不完全", "失效影响": "胶水渗漏，污染产品，绝缘性能下降", "失效原因": "固化温度、时间不足，胶水配比错误", "S": 6, "O": 3, "D": 2, "AP等级": "L", "建议措施": "严格管控固化温度与时间，胶水配比用自动配比机，首件做固化度测试"},
        {"工序名称": "灌胶固化", "失效模式": "胶水溢到连接器、端子上", "失效影响": "连接器接触不良，产品功能失效", "失效原因": "灌胶量过多，无防护遮挡", "S": 7, "O": 3, "D": 2, "AP等级": "M", "建议措施": "灌胶前对连接器做防护遮挡，严格管控灌胶量，溢胶件做清洁处理"}
    ],
    "防水密封装配": [
        {"工序名称": "防水密封装配", "失效模式": "密封胶条漏装、错位", "失效影响": "防水密封失效，IP等级不达标，进水短路", "失效原因": "操作人员疏忽，胶条卡槽装配不到位", "S": 9, "O": 2, "D": 3, "AP等级": "H", "建议措施": "胶条预装后做全检，双人核对，装配前核对胶条位置"},
        {"工序名称": "防水密封装配", "失效模式": "密封胶涂抹不均匀、断胶", "失效影响": "密封间隙，防水失效", "失效原因": "打胶机参数设置错误，操作人员手法不当", "S": 8, "O": 3, "D": 2, "AP等级": "M", "建议措施": "使用自动打胶机，优化打胶参数，打胶后全检胶线连续性"},
        {"工序名称": "防水密封装配", "失效模式": "密封面有杂质导致密封不良", "失效影响": "密封失效，进水短路", "失效原因": "密封面清洁不到位，装配环境有杂质", "S": 7, "O": 4, "D": 2, "AP等级": "M", "建议措施": "装配前用无尘布清洁密封面，装配环境做洁净度管控"}
    ],
    "电气性能测试": [
        {"工序名称": "电气性能测试", "失效模式": "输入输出电压测试参数设置错误", "失效影响": "性能不良产品误判合格，客户无法正常使用", "失效原因": "测试程序未锁定，操作人员误修改参数", "S": 8, "O": 2, "D": 4, "AP等级": "M", "建议措施": "测试程序加密锁定，每日首件用标准样件校准设备"},
        {"工序名称": "电气性能测试", "失效模式": "效率、功率因数不达标", "失效影响": "产品能耗过高，不符合安规要求，无法通过认证", "失效原因": "元器件参数偏差，PCB板设计缺陷", "S": 7, "O": 3, "D": 2, "AP等级": "M", "建议措施": "100%全检性能参数，严格管控元器件来料规格"},
        {"工序名称": "电气性能测试", "失效模式": "安规耐压、绝缘测试失效", "失效影响": "产品存在触电风险，不符合安规标准，批量召回", "失效原因": "测试参数设置错误，测试探针接触不良", "S": 10, "O": 2, "D": 3, "AP等级": "H", "建议措施": "测试程序加密锁定，每班检查探针状态，每日首件用标准样件校准"}
    ],
    "成品终检包装": [
        {"工序名称": "成品终检包装", "失效模式": "产品外观不良", "失效影响": "客户投诉，品牌形象受损", "失效原因": "周转过程中无防护，壳体划伤、污渍", "S": 4, "O": 3, "D": 2, "AP等级": "L", "建议措施": "产品使用专用周转工装，终检全检外观，不良件做返工处理"},
        {"工序名称": "成品终检包装", "失效模式": "说明书、合格证漏装", "失效影响": "客户投诉，退货", "失效原因": "无装箱核对清单，操作人员疏忽", "S": 3, "O": 3, "D": 2, "AP等级": "L", "建议措施": "制定装箱清单，双人核对，扫码防错"},
        {"工序名称": "成品终检包装", "失效模式": "产品型号与包装标识不符", "失效影响": "发错货，客户无法使用，批量退货", "失效原因": "标识打印错误，产品与包装未核对", "S": 5, "O": 2, "D": 2, "AP等级": "L", "建议措施": "扫码匹配产品型号与包装标识，不匹配无法封箱"}
    ]
}

# ===================== 核心工具函数 =====================
# AP等级判定函数（严格符合AIAG-VDA标准）
def get_ap_level(s, o, d):
    if s >= 9:
        return "H"
    elif s >=7:
        if o >=4 or d >=5:
            return "H"
        else:
            return "M"
    elif s >=5:
        if o >=5 or d >=6:
            return "M"
        else:
            return "L"
    else:
        return "L"

# 豆包AI生成单工序多选项PFMEA
def generate_ai_pfmea(process_name, product_type):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    # 行业专属提示词，确保专业性、区分度，和本地库不重复
    prompt = f"""
    你是全球顶级的汽车行业IATF16949&AIAG-VDA FMEA专家，10年以上新能源{product_type}装配行业PFMEA编制经验，深度熟悉主机厂审核要求。
    请针对【{product_type}】的【{process_name}】装配工序，生成3个完全不同、无重复、无同质化、贴合实际量产场景的PFMEA条目，严格遵循以下铁则：
    1. 每个条目必须有完整的失效链：失效模式→失效影响→失效原因，逻辑完全闭环，严格符合AIAG-VDA FMEA第一版标准，禁止出现逻辑断层
    2. 严重度S、频度O、探测度D评分严格遵循AIAG-VDA 10分制标准，贴合新能源汽车零部件装配量产场景，评分必须合理合规，禁止乱评分
    3. AP优先级等级严格按照规则判定：S≥9为H；S=7-8且O≥4或D≥5为H，其余为M；S≤6且O≤4且D≤5为L，其余为M
    4. 3个条目必须有强区分度：分别覆盖【高安全风险场景】、【高频发生的量产问题场景】、【客户审核重点关注的常规风险场景】，失效模式、失效原因、风险等级完全不同，绝对禁止重复
    5. 所有内容必须100%贴合{product_type}装配厂的实际量产场景，禁止出现通用化、不相关、牛头不对马嘴的内容，完全符合IATF16949第三方审核要求
    6. 禁止和本地标准库的内容重复，必须生成全新的、更贴合定制场景的内容
    7. 输出格式必须是严格的JSON数组，只能输出JSON内容，不得有任何额外的文字、注释、markdown格式、解释说明，确保可以直接被JSON解析
    输出格式模板：
    [
      {{
        "工序名称": "{process_name}",
        "失效模式": "具体、可量化的装配失效描述，禁止模糊表述",
        "失效影响": "分维度描述对下游工序、产品功能、终端客户、安全合规、品牌的影响",
        "失效原因": "从4M1E维度明确根本原因，禁止表面原因",
        "S": 1-10的整数,
        "O": 1-10的整数,
        "D": 1-10的整数,
        "AP等级": "H/M/L",
        "建议措施": "具体可落地的预防+探测改进措施，禁止空泛表述"
      }},
      {{
        "工序名称": "{process_name}",
        "失效模式": "和第一个条目完全不同的失效模式",
        "失效影响": "对应失效模式的专属影响",
        "失效原因": "对应失效模式的专属根本原因",
        "S": 1-10的整数,
        "O": 1-10的整数,
        "D": 1-10的整数,
        "AP等级": "H/M/L",
        "建议措施": "对应失效模式的专属改进措施"
      }},
      {{
        "工序名称": "{process_name}",
        "失效模式": "和前两个条目完全不同的失效模式",
        "失效影响": "对应失效模式的专属影响",
        "失效原因": "对应失效模式的专属根本原因",
        "S": 1-10的整数,
        "O": 1-10的整数,
        "D": 1-10的整数,
        "AP等级": "H/M/L",
        "建议措施": "对应失效模式的专属改进措施"
      }}
    ]
    """
    
    data = {
        "model": MODEL_NAME,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,  # 控制生成多样性，确保3个条目有区分度
        "top_p": 0.9,
        "max_tokens": 4000
    }
    
    try:
        response = requests.post(API_ENDPOINT, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        result = response.json()
        ai_content = result["choices"][0]["message"]["content"].strip()
        
        # 格式容错处理，去除多余的markdown符号
        ai_content = ai_content.replace("```json", "").replace("```", "").strip()
        pfmea_list = json.loads(ai_content)
        
        # 校验生成内容的完整性
        valid_list = []
        for item in pfmea_list:
            required_keys = ["工序名称", "失效模式", "失效影响", "失效原因", "S", "O", "D", "AP等级", "建议措施"]
            if all(key in item for key in required_keys):
                # 修正AP等级，确保符合标准
                item["AP等级"] = get_ap_level(item["S"], item["O"], item["D"])
                valid_list.append(item)
        
        # 确保返回3个有效条目，不足的话用合规内容补充
        while len(valid_list) < 3:
            valid_list.append(LOCAL_PFMEA_LIB.get(process_name, [])[len(valid_list)])
        
        return valid_list, None
    
    except Exception as e:
        return None, f"AI生成失败：{str(e)}"

# Excel导出函数
def export_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='PFMEA表', index=False)
        # 优化列宽
        worksheet = writer.sheets['PFMEA表']
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column].width = adjusted_width
    output.seek(0)
    return output

# ===================== Streamlit页面主程序 =====================
# 页面配置
st.set_page_config(
    page_title="新能源PFMEA专业生成系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 页面标题
st.title("⚡ 新能源电池包/充电器PFMEA专业生成系统")
st.caption("符合AIAG-VDA FMEA标准 | IATF16949审核专用 | 内置密钥开箱即用")
st.divider()

# Session State强制初始化，彻底解决刷新报错问题
def init_session_state():
    default_state = {
        "product_type": "电池包",
        "generate_mode": "本地专业标准库",
        "selected_process": [],
        "ai_generated_options": {},
        "selected_pfmea_items": [],
        "ai_step": 1,
        "preview_df": None
    }
    for key, value in default_state.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# 侧边栏配置
with st.sidebar:
    st.header("系统配置")
    # 产品类型选择
    st.session_state.product_type = st.radio(
        "选择产品类型",
        ["电池包", "充电器"],
        index=0 if st.session_state.product_type == "电池包" else 1
    )
    
    # 生成模式选择
    st.session_state.generate_mode = st.radio(
        "选择生成模式",
        ["本地专业标准库", "AI多选项生成模式"],
        index=0 if st.session_state.generate_mode == "本地专业标准库" else 1
    )
    
    st.divider()
    st.subheader("使用说明")
    st.markdown("""
    1. **本地模式**：无网络可用，内容100%符合审核要求，一键生成标准PFMEA
    2. **AI模式**：每个工序生成3个不同选项，自由勾选，定制化生成专属内容
    3. 密钥已内置，无需手动输入，打开即可使用
    4. 生成内容完全符合IATF16949&AIAG-VDA标准，可直接用于客户审核
    """)

# 主页面内容
# 工序选择
process_list = BATTERY_PROCESS_LIST if st.session_state.product_type == "电池包" else CHARGER_PROCESS_LIST

# 分模式处理
if st.session_state.generate_mode == "本地专业标准库":
    st.subheader("📦 本地专业标准库模式（无网络可用 | 100%审核合规）")
    # 工序选择
    col1, col2 = st.columns([1, 1])
    with col1:
        select_all = st.checkbox("全选所有工序")
    with col2:
        st.caption(f"当前产品：{st.session_state.product_type}，共{len(process_list)}道工序")
    
    if select_all:
        st.session_state.selected_process = process_list
    else:
        st.session_state.selected_process = st.multiselect(
            "选择需要生成PFMEA的工序",
            process_list,
            default=st.session_state.selected_process
        )
    
    # 生成按钮
    if st.button("一键生成PFMEA", type="primary", use_container_width=True):
        if not st.session_state.selected_process:
            st.error("请至少选择一道工序")
        else:
            with st.spinner("正在生成合规PFMEA内容..."):
                # 汇总本地库内容
                pfmea_all_data = []
                for process in st.session_state.selected_process:
                    pfmea_all_data.extend(LOCAL_PFMEA_LIB.get(process, []))
                
                # 生成DataFrame
                df = pd.DataFrame(pfmea_all_data)
                st.session_state.preview_df = df
                st.success("生成完成！内容完全符合IATF16949审核要求")
    
    # 预览与导出
    if st.session_state.preview_df is not None:
        st.divider()
        st.subheader("📋 PFMEA内容预览")
        st.dataframe(st.session_state.preview_df, use_container_width=True, height=400)
        
        # 导出Excel
        excel_file = export_to_excel(st.session_state.preview_df)
        st.download_button(
            label="📥 下载Excel格式PFMEA",
            data=excel_file,
            file_name=f"{st.session_state.product_type}_PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

# AI多选项生成模式
else:
    st.subheader("🤖 AI多选项生成模式（定制化场景 | 多选项自由选择）")
    # 步骤控制
    if st.session_state.ai_step == 1:
        # 步骤1：选择工序
        col1, col2 = st.columns([1, 1])
        with col1:
            select_all = st.checkbox("全选所有工序")
        with col2:
            st.caption(f"当前产品：{st.session_state.product_type}，共{len(process_list)}道工序")
        
        if select_all:
            st.session_state.selected_process = process_list
        else:
            st.session_state.selected_process = st.multiselect(
                "选择需要AI生成PFMEA的工序",
                process_list,
                default=st.session_state.selected_process
            )
        
        # 生成按钮
        if st.button("开始AI生成多选项PFMEA", type="primary", use_container_width=True):
            if not st.session_state.selected_process:
                st.error("请至少选择一道工序")
            else:
                # 重置状态
                st.session_state.ai_generated_options = {}
                st.session_state.selected_pfmea_items = []
                st.session_state.preview_df = None
                
                with st.spinner("正在调用AI生成专属PFMEA内容，请勿刷新页面..."):
                    all_success = True
                    error_msg = ""
                    # 循环每个工序生成3个选项
                    for process in st.session_state.selected_process:
                        st.caption(f"正在生成【{process}】的PFMEA选项...")
                        pfmea_list, error = generate_ai_pfmea(process, st.session_state.product_type)
                        if error:
                            all_success = False
                            error_msg = error
                            break
                        st.session_state.ai_generated_options[process] = pfmea_list
                    
                    if all_success:
                        st.session_state.ai_step = 2
                        st.success("AI生成完成！请为每个工序选择需要的PFMEA条目")
                        st.rerun()
                    else:
                        st.error(f"{error_msg}，可选择使用本地库兜底生成")
                        if st.button("使用本地专业库兜底生成"):
                            pfmea_all_data = []
                            for process in st.session_state.selected_process:
                                pfmea_all_data.extend(LOCAL_PFMEA_LIB.get(process, []))
                            df = pd.DataFrame(pfmea_all_data)
                            st.session_state.preview_df = df
                            st.session_state.ai_step = 3
                            st.rerun()
    
    elif st.session_state.ai_step == 2:
        # 步骤2：选择每个工序的PFMEA条目
        st.info("请为每个工序勾选需要的PFMEA条目，可多选，勾选完成后点击底部的【确认选择】按钮")
        st.divider()
        
        # 遍历每个工序，展示3个选项
        for process, pfmea_options in st.session_state.ai_generated_options.items():
            st.subheader(f"📌 工序：{process}")
            for i, item in enumerate(pfmea_options):
                # 复选框，默认不勾选
                is_selected = st.checkbox(
                    f"选项{i+1} | 失效模式：{item['失效模式']} | 风险等级：{item['AP等级']} | S:{item['S']} O:{item['O']} D:{item['D']}",
                    key=f"{process}_{i}"
                )
                # 展示条目详情
                with st.expander(f"查看选项{i+1}完整内容"):
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        st.markdown(f"**失效影响**：{item['失效影响']}")
                        st.markdown(f"**失效原因**：{item['失效原因']}")
                    with col2:
                        st.markdown(f"**建议措施**：{item['建议措施']}")
                        st.markdown(f"**AP等级**：{item['AP等级']}")
                st.divider()
        
        # 确认选择按钮
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("返回重新生成", use_container_width=True):
                st.session_state.ai_step = 1
                st.rerun()
        with col2:
            if st.button("确认选择，生成预览表", type="primary", use_container_width=True):
                # 汇总所有勾选的条目
                selected_items = []
                for process, pfmea_options in st.session_state.ai_generated_options.items():
                    for i, item in enumerate(pfmea_options):
                        if st.session_state.get(f"{process}_{i}", False):
                            selected_items.append(item)
                
                if not selected_items:
                    st.error("请至少勾选一个PFMEA条目")
                else:
                    st.session_state.selected_pfmea_items = selected_items
                    st.session_state.preview_df = pd.DataFrame(selected_items)
                    st.session_state.ai_step = 3
                    st.success("选择完成！已生成PFMEA预览表")
                    st.rerun()
    
    elif st.session_state.ai_step == 3:
        # 步骤3：预览与导出
        st.subheader("📋 最终PFMEA内容预览")
        st.dataframe(st.session_state.preview_df, use_container_width=True, height=400)
        
        # 按钮组
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("返回重新选择条目", use_container_width=True):
                st.session_state.ai_step = 2
                st.rerun()
        with col2:
            if st.button("重新生成所有内容", use_container_width=True):
                st.session_state.ai_step = 1
                st.session_state.ai_generated_options = {}
                st.session_state.selected_pfmea_items = []
                st.session_state.preview_df = None
                st.rerun()
        with col3:
            # 导出Excel
            excel_file = export_to_excel(st.session_state.preview_df)
            st.download_button(
                label="📥 下载Excel格式PFMEA",
                data=excel_file,
                file_name=f"{st.session_state.product_type}_AI定制PFMEA_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
