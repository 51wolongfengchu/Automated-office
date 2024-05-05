import pandas as pd

# 读取 Excel 文件的所有工作表名称
sheet_names = pd.ExcelFile('三年级分班成绩单.xlsx').sheet_names

# 创建一个空的 DataFrame，用于保存所有班级的成绩数据
df_all = pd.DataFrame()

# 遍历每个工作表
for i in sheet_names:
    # 从 Excel 文件中读取当前工作表的数据
    df = pd.read_excel('三年级分班成绩单.xlsx', sheet_name=i)

    # 将当前工作表的数据追加到 df_all 中
    df_all = df_all._append(df)

# 将合并后的数据保存到新的 Excel 文件中，不包含索引列，并命名为“总成绩”
df_all.to_excel('三年级总成绩单.xlsx', index=False, sheet_name='总成绩')

print('已完成')