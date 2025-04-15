import os
import docx
import pandas as pd
from langchain_core.messages import HumanMessage
from langchain_openai import ChatOpenAI

# 设置 API 信息
api_key = os.getenv("ZHIPU_API_KEY")
llm = ChatOpenAI(
    model="glm-4-flash",
    openai_api_key=api_key,
    openai_api_base="https://open.bigmodel.cn/api/paas/v4/"
)

# 定义文件夹路径和输出文件路径
folder_path = r"D:/PMC-gas/datasets"
output_file = r"D:/PMC-gas/results/mian_variable_result.xlsx"

# 用于存储所有文档的文本
all_texts = []

# 遍历文件夹中的所有 Word 文档
for filename in os.listdir(folder_path):
    if filename.endswith('.docx'):
        file_path = os.path.join(folder_path, filename)
        doc = docx.Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        text = '\n'.join(full_text)
        all_texts.append(text)

# 合并所有文本
combined_text = '\n'.join(all_texts)

# 设计提示词以提取高频词
prompt = f"请从以下文本中提取与天然气相关的高频词：{combined_text}，“天然气”或者公司名等专有名词不提取，只提取天然气相关政策措施的高频词。"

# 调用智谱清言 API 提取高频词
response = llm.invoke([HumanMessage(content=prompt)])
high_frequency_words = response.content

# 设计提示词以对高频词进行聚类
cluster_prompt = f"请将以下与天然气相关的高频词进行聚类，分成与天然气相关的九类：{high_frequency_words}，类别尽量不要重叠。"

# 调用智谱清言 API 进行聚类
cluster_response = llm.invoke([HumanMessage(content=cluster_prompt)])
clustered_words = cluster_response.content

# 处理聚类结果，区分出类别和高频词
category_lines = []
current_category = None
current_words = []
for line in clustered_words.strip().split('\n'):
    line = line.strip()
    if line.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.')):
        if current_category is not None:
            category_lines.append((current_category, ', '.join(current_words)))
        current_category = line.split('. ')[1].strip().replace('**', '')
        current_words = []
    elif line.startswith('-'):
        current_words.append(line.replace('- ', '').strip())

# 添加最后一个类别
if current_category is not None:
    category_lines.append((current_category, ', '.join(current_words)))

# 保存结果到 Excel 文件
data = {
    '类别': [category for category, _ in category_lines],
    '高频词': [words for _, words in category_lines]
}
df = pd.DataFrame(data)
df.to_excel(output_file, index=False)

print(f"高频词和类别结果已保存到 {output_file}")
