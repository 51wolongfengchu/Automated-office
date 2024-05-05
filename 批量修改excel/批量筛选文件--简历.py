import os
import shutil
import pdfplumber

#用os模块，读取并筛选出符合条件的简历文件，获取每个文件的相对的路径
file_list = os.listdir("./简历")
new_file_list = []
for index, file in enumerate(file_list):
    if file.split(".")[-1] == "pdf":
        new_file_list.append(file_list[index])
full_path_list = ["./简历/" + i for i in new_file_list]
des_path = "./简历/简历筛选_SQL"
#读取每个简历的文件，并提取文字的内容，将其转化为字符串
for full_path in full_path_list:
    string = ""
    with pdfplumber.open(full_path) as pdf:
        pages_list = pdf.pages
        for page in pages_list:
            text = page.extract_text()
            string += text
        pdf.close()
        #判断sql是否在字符串string中
        if "sql" in string.lower():
            #判断目标的中间价是否存在，若不存在，则需要创建一个目标的文件夹
            if os.path.exists(des_path):
                print("指定文件夹存在")
            else:
                os.mkdir(des_path)
                #关键词在，则简历移动到目标的文件中
            shutil.move(full_path,des_path)

print('have finished')