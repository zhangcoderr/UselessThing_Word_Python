# UselessThing_Word_Python
根据上下文自动填充word文档中格式为下划线的文本，争取做到：
1、输入关键字查找，输入正则匹配，如  大米__元，  上下文 '我们的大米2元，十分便宜' 输入查找格式为 '大米'、'元' ，正则：'大米.*?元'
2、数据通过TXT文本读取,简化代码


6.14：更新：代码中if ((r.underline == WD_UNDERLINE.THICK or r.underline == True) and len(r.text) > 3): 的 len(r.text) > 3) 可单独拿出来做对输入文字中原文本的替换、文本筛选
if xx in r.text
