import os
from win32com import client as wc
import win32api
import time
import glob
from datetime import datetime

#  注意：目录的格式必须写成双反斜杠
path = ""
files = []
errs = []

def writeFile(_txt):
    txt= open(os.path.join(os.getcwd(),'log.txt'),'a',encoding='UTF-8')   # 创建文件，权限为写入
    _txt=datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]+" "+_txt
    print(_txt)
    txt.write(_txt +'\n')

def takeErrs(_f):
    if _f not in errs:
        errs.append(_f)

def convertDocx(_files):
    n=0;
    for file in _files[:]:
    # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件）
        if not file.endswith('.docx') and file.endswith('.doc') and not os.path.basename(file).startswith('~$'):                             
                new_file_path = os.path.splitext(file)[0] + ".docx"                
                try:
                    word = wc.Dispatch("Word.Application")
                    # 后台运行，不显示，不警告
                    word.Visible = 0
                    word.DisplayAlerts = 0
                    #print("正在转换：[{0}/{1}] ".format(n+1,len(_files))+ file+" => "+new_file_path)
                    writeFile("正在转换：[{}/{}] ".format(n+1,len(_files))+ file+" => "+new_file_path)
                    # 打开文件
                    doc = word.Documents.Open(file)
                    time.sleep(0.5)
                    # 将文件另存为.docx
                    doc.SaveAs("{}x".format(file), 12)  # 12表示docx格式
                    # 删除原doc文件
                    # os.remove(files[0])
                    # 在files数组中删除第一个文件地址（已处理的文件地址）
                    del _files[_files.index(file)]
                except:
                    writeFile("\033[1;31m 转换失败：{}\033[0m".format(file))
                    takeErrs(file)
                finally:                                                            
                    try:
                        doc.Close()
                        word.Quit()
                    except:
                        writeFile("\033[1;31m 未知错误：{}\033[0m".format(file))
                        takeErrs(file)                  
                    time.sleep(0.5)
        else:
            writeFile("\033[1;38m 跳过转换：[{}/{}]\033[0m".format(n+1,len(_files))+ file)
        n=n+1;               
    writeFile("队列完成")    

def exec():
    path=input("请输入目录：").replace(r'\\',r'')
    while not os.path.isdir(path) or not os.path.exists(path):
        writeFile("\033[1;31m 目录不合法：{}\033[0m".format(path))
        path=input("请输入目录：").replace(r'\\',r'')
    writeFile("指定目录：{}".format(path))
    
    path = path+'/**/*.doc'
    for file in glob.glob(path,recursive = True):
        if not file.endswith('.docx') and file.endswith('.doc') and not os.path.basename(file).startswith('~$'): 
            files.append(file)
    writeFile("检索完成：{}".format(len(files)))
    
    if len(files)>0:
        convertDocx(files)
    err=0
    while len(errs)>0 and err<=5:
        err=err+1
        print("失败队列{}：{}".format(err,len(errs)))
        convertDocx(errs)
    exec()

# 程序入口
if __name__ == "__main__":
    exec()