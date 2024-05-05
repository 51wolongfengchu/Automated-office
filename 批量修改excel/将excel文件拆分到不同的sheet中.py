import pandas as pd

# 从 Excel 文件中读取数据
df = pd.read_excel("三年级总成绩.xlsx")

# 创建一个 ExcelWriter 对象，并指定输出文件名为 "三年级总成绩单——1.xlsx"
writer = pd.ExcelWriter("三年级总成绩单——1.xlsx")

# 将整个数据框保存到 "总成绩" 工作表中，不包含索引列
df.to_excel(writer, sheet_name="总成绩", index=False)

# 遍历班级列的唯一值
for i in df['班级'].unique():
    # 将每个班级的数据保存到以班级名命名的工作表中，不包含索引列
    df[df["班级"] == i].to_excel(writer, sheet_name=i, index=False)

# 保存并关闭 Excel 文件
writer._save()

# 处理完成后打印消息
print('已完成')
