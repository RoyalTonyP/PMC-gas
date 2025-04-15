import json
import os

import xlwt
from langchain.prompts import PromptTemplate
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI
from langchain_unstructured import UnstructuredLoader

# 获取环境变量中的 API 密钥
api_key = os.getenv("ZHIPU_API_KEY")

# 初始化大语言模型
llm = ChatOpenAI(
    model="glm-4-flash",
    openai_api_key=api_key,
    openai_api_base="https://open.bigmodel.cn/api/paas/v4/"
)

# 定义提示模板
template = """Question: {question}

Answer: 按照要求回答这个问题"""
prompt = PromptTemplate(
    template=template,
    input_variables=["question"]
)

# 使用 RunnableSequence 替代 LLMChain
runnable = prompt | llm

# 数据集所在的基础目录
base_dir = "D:/PyCharm/Py_Projects/gaspaper/datasets/"
# 获取基础目录下的所有文件
files = [os.path.join(base_dir, file) for file in os.listdir(base_dir)]

# 创建一个新的 Excel 工作簿
book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
# 写入表头
sheet.write(0, 0, '政策文件名')
sheet.write(0, 1, '资源勘探与开发')
sheet.write(0, 2, '管网建设与运营')
sheet.write(0, 3, '储气调峰')
sheet.write(0, 4, '消费引导')
sheet.write(0, 5, '市场准入资质')
sheet.write(0, 6, '安全环保')
sheet.write(0, 7, '价格调控')
sheet.write(0, 8, '税收')
sheet.write(0, 9, '监管')

row = 1

# 遍历文件列表
for file in files:
    print(f"Processing file: {file}")
    sheet.write(row, 0, file)
    # 使用 UnstructuredLoader 加载文件
    loader = UnstructuredLoader(file)
    documents = loader.load()

    # 对文档进行文本分割
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=10)
    split_docs = text_splitter.split_documents(documents)
    split_texts = [doc.page_content for doc in split_docs]

    # 处理资源勘探与开发内容
    gas_exploration_template = '''请对给定的政策文本进行分析，判断其是否与天然气相关。若不相关，输出空列表[]；
    若相关，从文本里提取和天然气资源勘探与开发有关的内容，示例包括但不限于：
    1. 资源勘探政策：例如明确勘探资质条件、规定区块获取方式等。
    2. 资源勘探区域：像划定允许或限制勘探的具体区域。
    3. 资源勘探标准：如确定地震勘探、钻井等方面的技术规范。
    4. 资源开发环保：涵盖废弃物处理、生态恢复等环保要求。
    5. 资源开发收益：包括税收减免、财政补贴等相关政策。
    6. 资源勘探规则：例如招标流程、准入门槛等规定。
    提取内容时要以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未找到相关内容则输出空列表[]。请从这段文本中提取相关信息：{text}'''

    gas_exploration_results = []
    for text in split_texts:
        prompt_text = gas_exploration_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_exploration_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_exploration_results = list(set(gas_exploration_results))
    gas_exploration_str = str(gas_exploration_results)
    print(gas_exploration_str)
    sheet.write(row, 1, gas_exploration_str)

    # 处理管网建设与运营内容
    gas_pipeline_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气管网建设与运营相关的内容，示例包括但不限于：
    1. 规划布局：如管网建设的总体路线规划、区域连接方案、资金支持政策等。
    2. 管输政策：包括运输价格定价原则、第三方准入规则、运输服务标准等。
    3. 技术标准：涉及管道材质要求、耐压等级规范、接口标准等。
    4. 市场管理：涵盖运营企业资质要求、安全监管规定、互联互通协调机制等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_pipeline_results = []
    for text in split_texts:
        prompt_text = gas_pipeline_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_pipeline_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_pipeline_results = list(set(gas_pipeline_results))
    gas_pipeline_str = str(gas_pipeline_results)
    print(gas_pipeline_str)
    sheet.write(row, 2, gas_pipeline_str)

    # 处理储气调峰内容
    gas_peak_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气储气调峰相关的内容，示例包括但不限于：
    1. 储气设施管理：如储气库建设责任主体划分、运营管理规范、安全监测标准等。
    2. 储气服务与市场：包括储气能力租赁规则、交易机制、第三方准入标准等。
    3. 储气能力与调峰：涉及储气库容量指标、可调节气量要求、调峰调度机制等。
    4. 奖励与政策：包含建设补贴、税收优惠、专项奖励等激励措施。
    5. 项目管理：涵盖储气设施建设审批流程、工程质量监管要求等。
    6. 冬季保供：如极端天气应急储气预案、冬季供气保障机制等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_peak_results = []
    for text in split_texts:
        prompt_text = gas_peak_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_peak_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_peak_results = list(set(gas_peak_results))
    gas_peak_str = str(gas_peak_results)
    print(gas_peak_str)
    sheet.write(row, 3, gas_peak_str)

    # 处理消费引导内容
    gas_consumption_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气消费引导相关的内容，示例包括但不限于：
    1. 激励政策：如用气积分奖励、消费阶梯优惠、环保认证奖励等。
    2. 宣传活动：包括用气知识普及活动、环保优势推广、用户教育计划等。
    3. 补贴政策：涉及居民用气补贴、工商业用气折扣、设备改造补贴等。
    4. 税收优惠：例如企业所得税减免、增值税抵扣、固定资产加速折旧等。
    5. 金融支持：包含低息贷款、融资担保、绿色债券支持等。
    6. 能源项目：如分布式能源补贴、加气站建设资助、天然气车辆购置补贴等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_consumption_results = []
    for text in split_texts:
        prompt_text = gas_consumption_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_consumption_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_consumption_results = list(set(gas_consumption_results))
    gas_consumption_str = str(gas_consumption_results)
    print(gas_consumption_str)
    sheet.write(row, 4, gas_consumption_str)

    # 处理市场准入内容
    gas_market_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气市场准入资质相关的内容，示例包括但不限于：
    1. 矿业权管理：如勘探开发资质条件、区块获取方式（招标/竞拍）、外资准入限制等。
    2. 基础设施开放：包括管网第三方准入规则、储气设施共享机制、基础设施公平接入标准等。
    3. 天然气交易资质：涉及企业经营许可区域、技术能力门槛、环保合规要求等。
    4. 市场体系：涵盖天然气定价机制、市场主体培育政策、交易平台建设规范等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_market_results = []
    for text in split_texts:
        prompt_text = gas_market_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_market_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_market_results = list(set(gas_market_results))
    gas_market_str = str(gas_market_results)
    print(gas_market_str)
    sheet.write(row, 5, gas_market_str)

    # 处理安全环保内容
    gas_environment_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气安全环保相关的内容，示例包括但不限于：
    1. 污染物管理标准：如废气/废水/噪音排放限值、甲烷泄漏控制指标等。
    2. 环保奖励与支持：包含清洁生产补贴、低碳技术研发资助、替代能源奖励等。
    3. 环保设施建设与维护：涉及脱硫脱硝设备规范、污染处理设施运行标准等。
    4. 污染物去除与控制：例如净化处理环节的脱硫脱碳要求、燃烧设备氮氧化物控制等。
    5. 生态保护与绿色措施：包括施工期土地复垦要求、生物多样性保护方案等。
    6. 绿色技术与应用：如碳捕捉技术推广、智能监测系统部署、新能源替代技术支持等。
    7. 安全监管：涵盖全流程安全评估标准、应急预案管理规范、隐患排查制度等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_environment_results = []
    for text in split_texts:
        prompt_text = gas_environment_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_environment_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_environment_results = list(set(gas_environment_results))
    gas_environment_str = str(gas_environment_results)
    print(gas_environment_str)
    sheet.write(row, 6, gas_environment_str)

    # 处理价格调控
    gas_price_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气价格调控相关的内容，示例包括但不限于：
    1. 价格管理：如基准价格设定、价格上下限规定、跨区域价格协调机制等。
    2. 定价机制：包括与国际能源价格挂钩的联动机制、阶梯气价制度、季节差价政策等。
    3. 价格监管：涉及成本监审办法、价格监测预警体系、违规定价处罚措施等。
    4. 市场定价：涵盖竞争性环节价格形成机制、天然气交易中心建设规范等。
    5. 特殊领域定价：例如居民用气保障价格、重点工业用户优惠价格等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_price_results = []
    for text in split_texts:
        prompt_text = gas_price_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_price_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_price_results = list(set(gas_price_results))
    gas_price_str = str(gas_price_results)
    print(gas_price_str)
    sheet.write(row, 7, gas_price_str)

    # 处理税收内容
    gas_tax_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气税收相关的内容，示例包括但不限于：
    1. 税收政策：如资源税税率设定、进口关税优惠、销售印花税标准等。
    2. 增值税政策：涉及简易计税办法、进项税额抵扣规则、跨境交易增值税处理等。
    3. 资源税政策：例如开采企业资源税征收方式（从价/从量）、减免税适用条件等。
    4. 企业支持政策：包括企业所得税减免、消费税补贴、车船税优惠等激励措施。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_tax_results = []
    for text in split_texts:
        prompt_text = gas_tax_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_tax_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_tax_results = list(set(gas_tax_results))
    gas_tax_str = str(gas_tax_results)
    print(gas_tax_str)
    sheet.write(row, 8, gas_tax_str)

    # 处理监管内容
    gas_supervision_template = '''请分析给定政策文本是否与天然气相关。若不相关，输出空列表[]；
    若相关，请从文本中提取与天然气监管相关的内容，示例包括但不限于：
    1. 行业监管：如监管主体职责分工、质量标准制定（成分/热值）、设施安全规范（检查周期/维护标准）等。
    2. 市场监管：包括市场行为监管机制（反垄断/反不正当竞争）、价格监管措施（审核程序/违规处罚）等。
    3. 项目审批：涉及勘探开发/基础设施建设的审批流程（规划/环评/施工许可）、审查要点与时限要求等。
    4. 信息披露：例如生产/销售/成本数据定期公开机制、企业信用信息公示规则等。
    提取内容需以列表形式呈现，且只保留政策文本中原文内容，如["内容1", "内容2",...]，若未涉及相关内容则输出空列表[]。请从这段文本中提取：{text}'''

    gas_supervision_results = []
    for text in split_texts:
        prompt_text = gas_supervision_template.format(text=text)
        print(f"Starting to invoke LLM for content in {file}...")
        res = runnable.invoke(prompt_text)
        print(f"LLM response received for content in {file}: {res}")
        try:
            gas_supervision_results.extend(json.loads(res.content))
        except json.JSONDecodeError:
            print(f"Error decoding JSON for results in {file}: {res.content}")
    gas_supervision_results = list(set(gas_supervision_results))
    gas_supervision_str = str(gas_supervision_results)
    print(gas_supervision_str)
    sheet.write(row, 9, gas_supervision_str)
    row += 1

# 保存 Excel 文件
book.save("D:/PyCharm/Py_Projects/gaspaper/results/extraction_results.xls")