import glob

import pandas as pd

# 创建一个空的 DataFrame，用于保存所有班级的成绩数据
df_all = pd.DataFrame()

# 使用 glob.glob 获取符合条件的文件路径
for i in glob.glob('三年*班.xlsx', recursive=True):
    # 从 Excel 文件中读取当前工作表的数据
    df = pd.read_excel(i)

    # 将当前工作表的数据追加到 df_all 中
    df_all = df_all._append(df)

# 将合并后的数据保存到新的 Excel 文件中，不包含索引列
df_all.to_excel('wo shi sha bi.xlsx', index=False)

print('已完成')