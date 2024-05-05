from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

#打开文件的绝对路径
doc = Document("5_4_1.docx")
#遍历每个段落
for paragraph in doc.paragraphs:
    #遍历段落中的每个文字块
    for run in paragraph.runs:
        #标题1的设置
        if "Heading 1" in paragraph.style.name:
            #磅数
            run.font.size = Pt(22)
            #中文字体设置为’等线’
            run._element.rPr.rFonts.set(qn("w:eastAsia"),"等线")
            #英文字体设置为‘corbel'
            run.font.name = "Corbel"
            #标题2的设置
        elif "Heading 2" in paragraph.style.name:
            #标题2设置磅数为16
            run.font.size = Pt(16)
            #中文字体为’宋体
            run._element.rPr.rFonts.set(qn("w:eastAsia"),"宋体")
            #英文字体设置为"Times New Roman"
            run.font.name = "Times New Roman"
        #正文
        else:
            #字体颜色设置为红色
            run.font.color.rgb = RGBColor(234,22,72)
            #是否加粗（是）
            run.font.bold = True
            #是否斜体（是）
            run.font.italic = True
            #是否加下划线（否）
            run.font.underline = False
            #保存文件的名字
            doc.save("5_4_1_字体样式.docx")

print('have finished')