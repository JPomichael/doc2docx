#引入模块
import PySimpleGUI as sg
#创建元素
#在layout中输入sg.T("第一个文本元素")之后就可以为他搞属性了！
layout=[
sg.T("第一个文本元素")
]
#创建窗口
window=sg.Window('python图形化开发教程',layout)
#重复检测事件+重复刷新窗口
while True:
    event,values=window.read()    
    # print(event,values)
    #按x退出循环
    if event==None:  
        break
#关闭窗口
window.close()
