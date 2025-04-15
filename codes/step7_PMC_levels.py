import openpyxl

# 加载工作簿
workbook = openpyxl.load_workbook("D:/PMC-gas/results/main_variable_scores_&_PMC_index.xlsx")
sheet = workbook.active

# 初始化总和
sum = 0

# 遍历前10行和前15列，计算每行最大值的总和
for row in sheet.iter_rows(min_row=2, max_col=15, max_row=10):
    # 过滤掉 None 值
    valid_values = [cell.value for cell in row if cell.value is not None]
    if valid_values:
        sum += max(valid_values[1:])

# 获取 PMC 指数
PMC_index = [sheet['B11:O11'][0][i].value for i in range(len(sheet['B11:O11'][0]))]
# 过滤掉 PMC_index 中的 None 值
PMC_index = [PMC for PMC in PMC_index if PMC is not None]

# 计算区间
poor_interval = "[{},{})".format(0, "{:.2f}".format(4 / 9 * sum))
acceptable_interval = "[{},{})".format("{:.2f}".format(4 / 9 * sum), "{:.2f}".format(2 / 3 * sum))
excellent_interval = "[{},{})".format("{:.2f}".format(2 / 3 * sum), "{:.2f}".format(8 / 9 * sum))
perfect_interval = "[{},{}]".format("{:.2f}".format(8 / 9 * sum), "{:.2f}".format(sum))

# 创建新的工作簿
from openpyxl import Workbook
wb = Workbook()
ws = wb.active

# 写入表头
ws.cell(row=1, column=1, value="政策一致性级别")
ws.cell(row=1, column=2, value="poor")
ws.cell(row=1, column=3, value="acceptable")
ws.cell(row=1, column=4, value="excellent")
ws.cell(row=1, column=5, value="perfect")

# 写入 PMC 指数区间
ws.cell(row=2, column=1, value="PMC指数")
ws.cell(row=2, column=2, value=poor_interval)
ws.cell(row=2, column=3, value=acceptable_interval)
ws.cell(row=2, column=4, value=excellent_interval)
ws.cell(row=2, column=5, value=perfect_interval)

# 写入政策数量
ws.cell(row=3, column=1, value="政策数量")
ws.cell(row=3, column=2, value=len([PMC for PMC in PMC_index if 0 <= PMC < 4 / 9 * sum]))
ws.cell(row=3, column=3, value=len([PMC for PMC in PMC_index if 4 / 9 * sum <= PMC < 2 / 3 * sum]))
ws.cell(row=3, column=4, value=len([PMC for PMC in PMC_index if 2 / 3 * sum <= PMC < 8 / 9 * sum]))
ws.cell(row=3, column=5, value=len([PMC for PMC in PMC_index if 8 / 9 * sum <= PMC <= sum]))

# 保存工作簿
wb.save("D:/PMC-gas/results/Evaluation_criteria_of_the_PMC.xlsx")
