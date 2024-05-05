from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Cm
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

#文件的名字
doc = Document('作文.docx')
#遍历每一个段落
for p in doc.paragraphs:
    # 修改每个段落的行间距，首行缩进，段前和段后
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    #首行缩进，1=4个字符（空格为2）
    p.paragraph_format.first_line_indent = Cm(1)
    #段前
    p.paragraph_format.space_before = Pt(30)
    #段后
    p.paragraph_format.space_after = Pt(20)
    #两端对齐
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
#修改后的文件名字
doc.save('修改之后的论文`.docx')

print('have finished')