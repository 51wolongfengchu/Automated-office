import pandas as pd

df = pd.read_excel("三年级总成绩.xlsx")

# 遍历班级列的唯一值
for i in df["班级"].unique():
    # 将该班级的数据筛选出来，并保存到以班级名称命名的 Excel 文件中
    df[df["班级"] == i].to_excel(f"{i}.xlsx", index=False)

print('已完成')
