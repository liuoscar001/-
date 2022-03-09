'''
1、使用python-docx读取文件
2、使用jieba进行分词
3、统计大于出现2次的词语频率
4、使用grid排版窗体
5、输出

V1.0 程序主体完成
V1.1 加入结果排序功能
V1.2 调整界面布局，使用grid布局工具，关键词使用【】使其更加美观合理，易看
'''
import re
from tkinter import Button, Tk,filedialog
from tkinter.constants import END, W
from docx import Document
import tkinter as tk
import jieba
import tkinter.messagebox

def showMsg(): 
    
    tkinter.messagebox.showinfo(title='提示', message='请选择word后缀格式为.docx的文档！！！')
    
def count_words():
    words=[]
    all_words=[]
    counts={}
    
    #选择文件
    filename = filedialog.askopenfilename()
    #打开word文档
    files=Document(filename)
    add_ent.insert(END,filename)
    #使用jieba分词
    for x in files.paragraphs:
        words.append(x.text)
    for i in words:
        all_words.extend(jieba.lcut(i))
    #创建词频字典
    for word in all_words:
        if len(word)==1:
            continue
        else:
            counts[word]=counts.get(word,0)+1
    #字典转换为列表        
    items = list(counts.items())
    #按照词频出现次数进行排序
    items.sort(key=lambda x: x[1], reverse=True)
    #输入出现大于一次的词语
    for item in items:
        word, count = item
        if count==1:
            continue
        else:
            res_text.insert(END,f'【{word}】 共出现：{count}次\n\n')
    return
  
def delete():
    #清除地址框内容
    add_ent.delete('0',END)
    #清除结果框内容
    res_text.delete('0.0',END)
    
if __name__ == '__main__':
    
    #创建窗体
    fm_main=tk.Tk()
    fm_main.title('词频统计V1.5 for Smily  请选择word后缀格式为.docx的文档！！！')
    
    #创建地址输入框
    add_ent=tk.Entry(fm_main)
    add_ent.grid(row=0,column=0,ipadx=100,ipady=5,sticky=W,padx=5,pady=10)
   
    #选择文件按钮
    choice_but=tk.Button(fm_main,text='选择文件',command=count_words)
    choice_but.grid(row=0,column=1,ipady=5)
    
    #清除内容按钮
    del_but=tk.Button(fm_main,text='清空',command=delete)
    del_but.grid(row=0,column=2,ipady=5)
    
    #创建结果输出窗
    res_text=tk.Text(fm_main)
    res_text.grid(row=2,columnspan=10,ipadx=10,ipady=30)
    
    fm_main.mainloop()