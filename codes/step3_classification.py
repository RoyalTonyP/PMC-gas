import json
import os
import re
import xlrd
import xlwt
from langchain.prompts import PromptTemplate
from langchain_openai import ChatOpenAI
from langchain.schema.runnable import RunnableSequence

# 主属性关键词映射
keyword_map = {
    '资源勘探与开发': ['勘探', '开发', '资源', '气源', '开采'],
    '管网建设与运营': ['管网', '管道', '输配', '运营', '建设'],
    '储气调峰': ['储气', '调峰', '储备', '应急', '能力'],
    '消费引导': ['消费', '用气', '用户', '需求', '引导'],
    '市场准入资质': ['准入', '资质', '许可', '审批', '企业'],
    '安全环保': ['安全', '环保', '生态', '排放', '风险'],
    '价格调控': ['价格', '收费', '成本', '补贴', '税收'],
    '税收': ['税收', '税率', '优惠', '减免', '财政'],
    '监管': ['监管', '监督', '规范', '执法', '查处']
}


# 初始化大语言模型
def initialize_llm():
    api_key = os.getenv("ZHIPU_API_KEY")
    return ChatOpenAI(
        model="glm-4-flash",
        openai_api_key=api_key,
        openai_api_base="https://open.bigmodel.cn/api/paas/v4/"
    )


# 数据预处理
def preprocess_data(file_path):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)
    rows = sheet.nrows

    categories = {
        '资源勘探与开发': [],
        '管网建设与运营': [],
        '储气调峰': [],
        '消费引导': [],
        '市场准入资质': [],
        '安全环保': [],
        '价格调控': [],
        '税收': [],
        '监管': []
    }

    for i in range(1, rows):
        for category in categories:
            column = list(categories.keys()).index(category) + 1
            content = sheet.cell(i, column).value
            cleaned_content = re.sub(r'\s+', ' ', content).strip()
            items = json.loads(cleaned_content.replace('\'', '\"').replace('，', ','))
            categories[category] += items

    return categories


# 生成聚类提示
def generate_prompt(category_name, items):
    return (f'将以下列表按"{category_name}"相关内容进行语义聚类，形成不超过6个核心子类别，'
            f'每个子类别名称必须包含"{category_name}"领域专有名词（如{", ".join(keyword_map[category_name])}），'
            f'且不与其他主属性子类别重复（如"储气设施建设"只能属于储气调峰类），'
            f'名称需准确反映政策核心内容，禁止使用"类别X"等虚拟命名和通用词汇（如"管理"、"机制"等）。'
            f'只输出形成的核心子类别名，用列表形式回复，即["类别1", "类别2", ...,"类别6"]，'
            f'每个类别不超过7个字，不要用省略号。列表：{items}')


# 执行聚类
def perform_clustering(llm, category_name, items):
    prompt = PromptTemplate.from_template("{question}")
    runnable = prompt | llm

    prompt_text = generate_prompt(category_name, items)
    res = runnable.invoke({"question": prompt_text})
    print(f"聚类提示: {prompt_text}")
    print(f"聚类响应: {res.content}")

    return res.content


# 跨类别去重
def deduplicate_categories(all_categories):
    seen = set()
    for category in all_categories:
        main_attr = category[0]
        try:
            sub_attrs = json.loads(category[1])
        except json.JSONDecodeError:
            print(f"JSON 解析错误: 主属性 {main_attr} 的子属性 {category[1]} 不是有效的 JSON 字符串")
            category[1] = []
            continue
        unique_subs = []
        for sub in sub_attrs:
            if sub not in seen:
                seen.add(sub)
                unique_subs.append(sub)
        category[1] = json.dumps(unique_subs, ensure_ascii=False)
    return all_categories


# 跨类别去重和验证
def validate_categories(all_categories):
    seen = set()
    valid_categories = []
    for main_attr, sub_attrs in all_categories:
        try:
            # 处理可能的列表类型
            if isinstance(sub_attrs, list):
                subs = sub_attrs
            else:
                subs = json.loads(sub_attrs)
            unique_subs = []
            for sub in subs:
                # 验证领域专有名词和格式
                if (any(word in sub for word in keyword_map[main_attr]) and
                        not re.search(r'类别\d+', sub)):  # and  禁止虚拟类别
                    # not re.search(r'管理|机制|体系', sub)):  # 禁止通用词汇
                    if sub not in seen:
                        seen.add(sub)
                        unique_subs.append(sub)
            valid_categories.append([main_attr, json.dumps(unique_subs, ensure_ascii=False)])
        except json.JSONDecodeError:
            valid_categories.append([main_attr, '聚类失败'])
    return valid_categories


# 主程序
def main():
    llm = initialize_llm()
    categories = preprocess_data("D:/PyCharm/Py_Projects/gaspaper/results/extraction_results.xls")

    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('分类结果', cell_overwrite_ok=True)
    sheet.write(0, 0, '主变量')
    sheet.write(0, 1, '聚类结果')

    all_category_data = []
    row = 1
    for category in categories:
        items = list(set(categories[category]))  # 初始去重
        final_response = perform_clustering(llm, category, items)
        all_category_data.append([category, final_response])
        row += 1

    # 跨类别去重
    all_category_data = deduplicate_categories(all_category_data)

    # 验证并补充类别
    all_category_data = validate_categories(all_category_data)

    # 写入Excel
    row = 1
    for category_data in all_category_data:
        main_attr, sub_attrs = category_data
        sheet.write(row, 0, main_attr)
        sheet.write(row, 1, sub_attrs if sub_attrs else '聚类失败')
        row += 1

    book.save("D:/PyCharm/Py_Projects/gaspaper/results/classification_results.xls")


if __name__ == "__main__":
    main()
