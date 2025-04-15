import json
import os
import re

import xlrd
import xlwt
from langchain.prompts import PromptTemplate
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI
from langchain_unstructured import UnstructuredLoader
from tqdm import tqdm

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

wb = xlrd.open_workbook("D:/PMC-gas/results/extraction_results.xls")
sheet = wb.sheet_by_index(0)
rows = sheet.nrows

base_dir = "D:/PMC-gas/datasets"
# 获取当前目录下的所有文件
files = [os.path.join(base_dir, file) for file in os.listdir(base_dir)]
gas_exploration_list = []
gas_pipeline_list = []
gas_peak_list = []
gas_consumption_list = []
gas_market_list = []
gas_environment_list = []
gas_price_list = []
gas_tax_list = []
gas_supervision_list = []

# 遍历文件列表，输出文件名

book = xlwt.Workbook(encoding='utf-8')
sheet2 = book.add_sheet('sheet1', cell_overwrite_ok=True)
sheet2.write(0, 0, '政策文件名')
sheet2.write(0, 1, '资源勘探与开发')
sheet2.write(0, 2, '管网建设与运营')
sheet2.write(0, 3, '储气调峰')
sheet2.write(0, 4, '消费引导')
sheet2.write(0, 5, '市场准入')
sheet2.write(0, 6, '安全环保')
sheet2.write(0, 7, '价格调控')
sheet2.write(0, 8, '税收')
sheet2.write(0, 9, '监管')

row_w = 1
for file in tqdm(files):
    sheet2.write(row_w, 0, file)
    loader = UnstructuredLoader(file)
    row = 1
    for i in range(1, rows):
        if sheet.cell(i, 0).value == file:
            row = i
    cell_value = sheet.cell(row, 1).value
    try:
        # 清理单元格内容
        cell_value = cell_value.strip()
        if cell_value:
            gas_price = json.loads(cell_value)
            gas_market = json.loads(cell_value)
            gas_exploration = json.loads(cell_value)
            gas_pipeline = json.loads(cell_value)
            gas_environment = json.loads(cell_value)
            gas_consumption = json.loads(cell_value)
            gas_peak = json.loads(cell_value)
            gas_tax = json.loads(cell_value)
            gas_supervision = json.loads(cell_value)
        else:
            print(f"Cell in row {row}, column 1 is empty. Skipping JSON decoding.")
            gas_exploration = []
            gas_pipeline = []
            gas_peak = []
            gas_consumption = []
            gas_market = []
            gas_environment = []
            gas_price = []
            gas_tax = []
            gas_supervision = []
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON in row {row}, column 1: {e}, original content: {cell_value}")
        gas_exploration = []
        gas_pipeline = []
        gas_peak = []
        gas_consumption = []
        gas_market = []
        gas_environment = []
        gas_price = []
        gas_tax = []
        gas_supervision = []

    documents = loader.load()
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=10)
    split_docs = text_splitter.split_documents(documents)
    gas_exploration_temp = []
    gas_pipeline_temp = []
    gas_peak_temp = []
    gas_consumption_temp = []
    gas_market_temp = []
    gas_environment_temp = []
    gas_price_temp = []
    gas_tax_temp = []
    gas_supervision_temp = []

    for doc in split_docs:
        prompt1 = ("对于给定的天然气资源勘探与开发政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，资源勘探政策：指制定资源勘探领域的制度性政策，如勘探许可证发放、勘探权招标机制等。例如建立油气勘探区块竞争出让制度，明确勘探作业者资质要求。"
                   "第二类，资源勘探区域：涉及资源勘探空间范围的规划政策。如划分国家规划矿区与商业勘探区，设定自然保护区内勘探限制。"
                   "第三类，资源勘探标准：规定资源勘探技术与安全标准的政策。例如制定油气储量评估规范，设定勘探设备环保排放标准。"
                   "第四类，资源开发环保：针对资源开发过程的环境保护政策。如要求开发企业提交生态恢复方案，实施开发过程环境监测制度。"
                   "第五类，资源开发收益：关于资源开发收益分配的调控政策。例如建立油气资源税动态调整机制，制定区块出让收益分成办法。"
                   "第六类，资源勘探规则：规范资源勘探行为的管理政策。如明确勘探数据汇交义务，制定勘探区块退出机制。"
                   "回复格式为字典形式，即{资源勘探政策:1或0, 资源勘探区域:1或0, 资源勘探标准:1或0, 资源开发环保:1或0, 资源开发收益:1或0, 资源勘探规则:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt1 += doc.page_content
        res1 = runnable.invoke(prompt1)
        print(res1.content)
        # 将属性名添加双引号
        res1_fixed = re.sub(r'(\w+):', r'"\1":', res1.content)
        try:
            gas_exploration_temp.append(json.loads(res1_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res1: {e}, original content: {res1.content}, fixed content: {res1_fixed}")

        prompt2 = ("对于给定的天然气管网建设与运营政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，规划布局：指制定管网建设空间布局与统筹规划的政策。如编制国家主干管网建设规划，推动区域管网互联互通。"
                   "第二类，管输政策：涉及天然气管道运输服务的价格、合同及质量监管政策。例如制定管输费定价规则，建立运输服务质量评价体系。"
                   "第三类，技术标准：规定管网建设与运营技术规范的政策。如设定管道材料安全标准，制定智能管网监测系统技术要求。"
                   "第四类，市场管理：关于管网运营市场准入与竞争秩序的管理政策。例如实施管网运营企业资质许可制度，建立跨区域管道运输公平开放机制。"
                   "回复格式为字典形式，即{规划布局:1或0, 管输政策:1或0, 技术标准:1或0, 市场管理:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt2 += doc.page_content
        res2 = runnable.invoke(prompt2)
        print(res2.content)
        # 将属性名添加双引号
        res2_fixed = re.sub(r'(\w+):', r'"\1":', res2.content)
        try:
            gas_pipeline_temp.append(json.loads(res2_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res2: {e}, original content: {res2.content}, fixed content: {res2_fixed}")

        prompt3 = ("对于给定的天然气储气调峰政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，储气设施管理：指制定储气设施建设、运营及安全维护的管理政策。如规定储气库设计标准，建立设施定期检修制度。"
                   "第二类，储气服务与市场：涉及储气服务定价、交易及市场化运作的政策。例如实施储气服务竞价机制，建立储气容量租赁市场规则。"
                   "第三类，储气能力与调峰：关于储气能力建设目标及调峰技术应用的政策。如设定地方政府储气能力达标指标，推广智能化调峰技术应用。"
                   "第四类，奖励与政策：针对储气调峰行为的激励政策。如对超额完成储气任务企业给予税收优惠，建立调峰贡献度奖励机制。"
                   "第五类，项目管理：规范储气调峰项目全周期管理的政策。例如制定储气项目审批流程，建立项目进度跟踪监管体系。"
                   "第六类，冬季保供：专门针对冬季天然气保供的专项政策。如启动储气设施冬季应急调度预案，优先保障民生用气调峰需求。"
                   "回复格式为字典形式，即{储气设施管理:1或0, 储气服务与市场:1或0, 储气能力与调峰:1或0, 奖励与政策:1或0, 项目管理:1或0, 冬季保供:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt3 += doc.page_content
        res3 = runnable.invoke(prompt3)
        print(res3.content)
        # 将属性名添加双引号
        res3_fixed = re.sub(r'(\w+):', r'"\1":', res3.content)
        try:
            gas_peak_temp.append(json.loads(res3_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res3: {e}, original content: {res3.content}, fixed content: {res3_fixed}")

        prompt4 = ("对于给定的天然气消费引导政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，激励政策：指通过非物质奖励手段引导消费行为的政策。如建立绿色消费积分制度，对节能产品购买者给予荣誉表彰。"
                   "第二类，宣传活动：涉及消费理念传播与公众教育的政策。例如开展天然气高效利用科普巡展，制作家庭节能宣传手册。"
                   "第三类，补贴政策：针对特定消费行为的直接资金补贴政策。如对居民更换节能灶具给予现金补贴，对工商业用户使用清洁能源设备提供购置补贴。"
                   "第四类，税收优惠：通过税收减免引导消费方向的政策。例如对天然气汽车减免车船税，对分布式能源项目实施增值税优惠。"
                   "第五类，金融支持：提供融资便利或利率优惠的消费引导政策。如推出零首付清洁能源消费贷款，对节能改造项目给予贴息支持。"
                   "第六类，能源项目：通过实施具体工程促进消费升级的政策。例如推广天然气分布式能源项目，建设社区综合供能服务站。"
                   "回复格式为字典形式，即{激励政策:1或0, 宣传活动:1或0, 补贴政策:1或0, 税收优惠:1或0, 金融支持:1或0, 能源项目:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt4 += doc.page_content
        res4 = runnable.invoke(prompt4)
        print(res4.content)
        # 将属性名添加双引号
        res4_fixed = re.sub(r'(\w+):', r'"\1":', res4.content)
        try:
            gas_consumption_temp.append(json.loads(res4_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res4: {e}, original content: {res4.content}, fixed content: {res4_fixed}")

        prompt5 = ("对于给定的天然气市场准入资质政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，矿业权管理：指规范矿产资源开发权利的审批与管理政策。如实施油气矿业权竞争性出让制度，建立矿业权退出机制。"
                   "第二类，基础设施开放：涉及能源基础设施公平接入与共享的政策。例如制定管网设施向第三方公平开放规则，推动储气设施租赁市场化。"
                   "第三类，天然气交易资质：规定天然气市场主体交易资格的政策。如建立天然气交易中心会员准入标准，实施液化天然气（LNG）接收站运营资质许可。"
                   "第四类，市场体系：关于市场规则与监管框架的基础性政策。例如构建全国统一的天然气市场体系，制定市场主体信用监管制度。"
                   "回复格式为字典形式，即{矿业权管理:1或0, 基础设施开放:1或0, 天然气交易资质:1或0, 市场体系:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt5 += doc.page_content
        res5 = runnable.invoke(prompt5)
        print(res5.content)
        # 将属性名添加双引号
        res5_fixed = re.sub(r'(\w+):', r'"\1":', res5.content)
        try:
            gas_market_temp.append(json.loads(res5_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res5: {e}, original content: {res5.content}, fixed content: {res5_fixed}")

        prompt6 = ("对于给定的天然气安全环保政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，污染物管理标准：指制定污染物排放限值与监测规范的政策。如设定工业废气氮氧化物排放限值，建立企业污染排放在线监测系统。"
                   "第二类，环保奖励与支持：涉及环保行为激励与资金扶持的政策。例如对清洁生产企业给予税收减免，设立环保技术研发专项补贴。"
                   "第三类，环保设施建设与维护：规定污染防治设施建设与运维的政策。如要求燃煤电厂配套建设脱硫脱硝装置，建立污水处理设施定期检修制度。"
                   "第四类，污染物去除与控制：针对污染治理技术应用的政策。如推广烟气超低排放改造技术，实施挥发性有机物（VOCs）综合治理方案。"
                   "第五类，生态保护与绿色措施：关于生态修复与低碳发展的政策。例如建立矿山生态环境恢复补偿机制，推行政府绿色采购制度。"
                   "第六类，绿色技术与应用：促进环保技术研发与推广的政策。如支持碳捕捉与封存（CCUS）技术示范项目，鼓励新能源汽车充电基础设施建设。"
                   "第七类，安全监管：涉及安全生产监督与应急管理的政策。如制定危化品运输安全标准，建立环境事故应急预案演练制度。"
                   "回复格式为字典形式，即{污染物管理标准:1或0, 环保奖励与支持:1或0, 环保设施建设与维护:1或0, 污染物去除与控制:1或0, 生态保护与绿色措施:1或0, 绿色技术与应用:1或0, 安全监管:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt6 += doc.page_content
        res6 = runnable.invoke(prompt6)
        print(res6.content)
        # 将属性名添加双引号
        res6_fixed = re.sub(r'(\w+):', r'"\1":', res6.content)
        try:
            gas_environment_temp.append(json.loads(res6_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res6: {e}, original content: {res6.content}, fixed content: {res6_fixed}")

        prompt7 = ("对于给定的价格调控政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，价格管理：指统筹协调价格调控目标与政策实施的管理性政策。如制定年度价格总水平调控计划，建立跨部门价格协调机制。"
                   "第二类，定价机制：涉及价格形成方式与调整规则的政策。例如实施天然气井口价成本加成定价机制，建立居民用气价格联动调整公式。"
                   "第三类，价格监管：关于价格行为监督与违规查处的政策。如开展天然气价格专项检查，建立价格违法行为举报奖励制度。"
                   "第四类，市场定价：通过市场竞争形成价格的政策。如推行液化天然气（LNG）现货拍卖交易，取消非居民用气政府定价。"
                   "第五类，特殊领域定价：针对特定行业或群体的价格调控政策。例如制定化肥生产用气优惠价格，建立学校食堂用气价格备案制度。"
                   "回复格式为字典形式，即{价格管理:1或0, 定价机制:1或0, 价格监管:1或0, 市场定价:1或0, 特殊领域定价:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt7 += doc.page_content
        res7 = runnable.invoke(prompt7)
        print(res7.content)
        # 将属性名添加双引号
        res7_fixed = re.sub(r'(\w+):', r'"\1":', res7.content)
        try:
            gas_price_temp.append(json.loads(res7_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res7: {e}, original content: {res7.content}, fixed content: {res7_fixed}")

        prompt8 = ("对于给定的税收政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，税收政策：指制定税收制度框架与管理规则的综合性政策。如修订《税收征收管理法》，建立现代税收制度体系。"
                   "第二类，增值税政策：涉及增值税税率、抵扣及优惠的专项政策。例如将制造业增值税率从13%降至11%，扩大进项税额抵扣范围。"
                   "第三类，资源税政策：针对自然资源开发利用的税收调控政策。如对高耗能行业实施资源税从价计征改革，对页岩气开采企业减征资源税。"
                   "第四类，企业支持政策：通过税收手段支持企业发展的政策。如对小微企业实施所得税减半征收，对研发费用加计扣除比例提高至120%。"
                   "回复格式为字典形式，即{税收政策:1或0, 增值税政策:1或0, 资源税政策:1或0, 企业支持政策:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt8 += doc.page_content
        res8 = runnable.invoke(prompt8)
        print(res8.content)
        # 将属性名添加双引号
        res8_fixed = re.sub(r'(\w+):', r'"\1":', res8.content)
        try:
            gas_tax_temp.append(json.loads(res8_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res8: {e}, original content: {res8.content}, fixed content: {res8_fixed}")

        prompt9 = ("对于给定的监管政策列表，请判断这些政策所属类型。类型包含："
                   "第一类，行业监管：指针对特定行业制定的专业性监管政策。如建立能源行业碳排放监测体系，实施银行业资本充足率监管标准。"
                   "第二类，市场监管：涉及市场秩序维护与公平竞争的政策。例如开展反不正当竞争执法行动，建立价格垄断行为举报机制。"
                   "第三类，项目审批：规定项目准入条件与审批流程的政策。如实施固定资产投资项目节能审查制度，建立跨境并购安全审查机制。"
                   "第四类，信息披露：关于市场主体信息公开的管理政策。例如要求上市公司披露环境信息，制定政府投资项目进展公示规则。"
                   "回复格式为字典形式，即{行业监管:1或0, 市场监管:1或0, 项目审批:1或0, 信息披露:1或0}，"
                   "其中，1表示列表中存在该类政策元素，0表示不存在，列表中一个元素只属一类且必属其中一类，不会出现所有键值都为0的情况。"
                   "若列表元素较多，可能存在多个1；若仅一个元素，则仅有一个1，其余为0。"
                   "请勿改变字典键名，仅回复该字典，不要添加其他内容或换行符。文本如下：")
        prompt9 += doc.page_content
        res9 = runnable.invoke(prompt9)
        print(res9.content)
        # 将属性名添加双引号
        res9_fixed = re.sub(r'(\w+):', r'"\1":', res9.content)
        try:
            gas_supervision_temp.append(json.loads(res9_fixed))
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON in res9: {e}, original content: {res9.content}, fixed content: {res9_fixed}")

    print('===========================================')
    # ==========汇总===========
    if gas_exploration_temp:
        result1 = {
            "资源勘探政策": 0, "资源勘探区域": 0, "资源勘探标准": 0, "资源开发环保": 0, "资源开发收益": 0, "资源勘探规则": 0
        }
        for key in result1.keys():
            result1[key] = 1 if sum(
                [gas_exploration_temp[index].get(key, 0) for index in range(len(gas_exploration_temp))]) > 0 \
                else 0
        print(result1)
        gas_exploration.append(result1)
        sheet2.write(row_w, 1, str(result1))

    if gas_pipeline_temp:
        result2 = {
            "规划布局": 0, "管输政策": 0, "技术标准": 0, "市场管理": 0
        }
        for key in result2.keys():
            result2[key] = 1 if sum(
                [gas_pipeline_temp[index].get(key, 0) for index in range(len(gas_pipeline_temp))]) > 0 \
                else 0
        print(result2)
        gas_pipeline.append(result2)
        sheet2.write(row_w, 2, str(result2))

    if gas_peak_temp:
        result3 = {
            "储气设施管理": 0, "储气服务与市场": 0, "储气能力与调峰": 0, "奖励与政策": 0, "项目管理": 0, "冬季保供": 0
        }
        for key in result3.keys():
            result3[key] = 1 if sum(
                [gas_peak_temp[index].get(key, 0) for index in range(len(gas_peak_temp))]) > 0 \
                else 0
        print(result3)
        gas_peak.append(result3)
        sheet2.write(row_w, 3, str(result3))

    if gas_consumption_temp:
        result4 = {
            "激励政策": 0, "宣传活动": 0, "补贴政策": 0, "税收优惠": 0, "金融支持": 0, "能源项目": 0
        }
        for key in result4.keys():
            result4[key] = 1 if sum(
                [gas_consumption_temp[index].get(key, 0) for index in range(len(gas_consumption_temp))]) > 0 \
                else 0
        print(result4)
        gas_consumption.append(result4)
        sheet2.write(row_w, 4, str(result4))

    if gas_market_temp:
        result5 = {
            "矿业权管理": 0, "基础设施开放": 0, "天然气交易资质": 0, "市场体系": 0
        }
        for key in result5.keys():
            result5[key] = 1 if sum(
                [gas_market_temp[index].get(key, 0) for index in range(len(gas_market_temp))]) > 0 \
                else 0
        print(result5)
        gas_market.append(result5)
        sheet2.write(row_w, 5, str(result5))

    if gas_environment_temp:
        result6 = {
            "污染物管理标准": 0, "环保奖励与支持": 0, "环保设施建设与维护": 0, "污染物去除与控制": 0, "生态保护与绿色措施": 0, "绿色技术与应用": 0, "安全监管": 0
        }
        for key in result6.keys():
            result6[key] = 1 if sum(
                [gas_environment_temp[index].get(key, 0) for index in range(len(gas_environment_temp))]) > 0 \
                else 0
        print(result6)
        gas_environment.append(result6)
        sheet2.write(row_w, 6, str(result6))

    if gas_price_temp:
        result7 = {
            "价格管理": 0, "定价机制": 0, "价格监管": 0, "市场定价": 0, "特殊领域定价": 0
        }
        for key in result7.keys():
            result7[key] = 1 if sum([gas_price_temp[index].get(key, 0) for index in range(len(gas_price_temp))]) > 0 \
                else 0
        print(result7)
        gas_price.append(result7)
        sheet2.write(row_w, 7, str(result7))

    if gas_tax_temp:
        result8 = {
            "税收政策": 0, "增值税政策": 0, "资源税政策": 0, "企业支持政策": 0
        }
        for key in result8.keys():
            result8[key] = 1 if sum([gas_tax_temp[index].get(key, 0) for index in range(len(gas_tax_temp))]) > 0 \
                else 0
        print(result8)
        gas_tax.append(result8)
        sheet2.write(row_w, 8, str(result8))

    if gas_supervision_temp:
        result9 = {
            "行业监管": 0, "市场监管": 0, "项目审批": 0, "信息披露": 0
        }
        for key in result9.keys():
            result9[key] = 1 if sum(
                [gas_supervision_temp[index].get(key, 0) for index in range(len(gas_supervision_temp))]) > 0 \
                else 0
        print(result9)
        gas_supervision.append(result9)
        sheet2.write(row_w, 9, str(result9))

    row_w += 1
book.save("D:/PMC-gas/results/sub_variables_scores_results.xls")
