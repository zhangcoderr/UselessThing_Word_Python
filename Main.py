import docx
import re
from docx import Document
from docx.enum.text import WD_UNDERLINE
path = r"C:\Users\Administrator\Desktop\Test\test.docx"
doc=Document(path)

start=False
for p in doc.paragraphs:
    if('（封面）' in p.text):
        start=True
    if(start):
        if('投标文件内容' in p.text):
            start=False
        #print(p.text)
        for i in range(len(p.runs)):
            r=p.runs[i]
            lastText=None
            nextText=None
            if (i != 0):
                lastText=p.runs[i-1].text
            if(i!=len(p.runs)-1):
                nextText=p.runs[i+1].text

            if(lastText!=None and '建设单位' in lastText):
                #print(r.underline)
                #print(r.text)


                #输入数据
                if((r.underline==WD_UNDERLINE.THICK or r.underline==True) and len(r.text)>3):
                    input=True
                    for char in r.text:
                        if char!=' ':
                            input=False
                    if(input):
                        r.clear()

                        r.add_text(' fasdfsd;fk ')
                        #r.text='afsdfa'
                        print('ok')



doc.save(r"C:\Users\Administrator\Desktop\Test\testResult.docx")