import os
from win32com import client as wc
import win32api
import time
import glob

#  注意：目录的格式必须写成双反斜杠
path = ""
files = []
errs = []

def writeFile(_txt):
    txt= open(os.path.join(os.getcwd(),'log.txt'),'a',encoding='UTF-8')   # 创建文件，权限为写入
    print(_txt)
    txt.write(_txt +'\n')


def takeErrs(_f):
    if _f not in errs:
        errs.append(_f)

# def readInfo(_dir):
#     for entry in os.scandir(_dir):
#         if entry.is_file() and not entry.endswith('.docx') and entry.endswith('.doc') and not os.path.basename(entry).startswith('~$'):
#             print('')
#         else:
#             readInfo(entry.path)    
#     name = os.listdir(_dir)         # os.listdir方法返回一个列表对象
#     return name

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
                    writeFile("正在转换：[{0}/{1}] ".format(n+1,len(_files))+ file+" => "+new_file_path)
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
                    writeFile("转换失败：" + file)
                    takeErrs(file)
                finally:                                                            
                    try:
                        doc.Close()
                        word.Quit()
                    except:
                        writeFile("未知错误：" + file)
                        takeErrs(file)                  
                    time.sleep(0.5)
        else:
            writeFile("跳过转换：[{0}/{1}]".format(n+1,len(_files))+ file)
        n=n+1;               
    writeFile("队列完成")    

# 程序入口
if __name__ == "__main__":
    path=input("请输入目录：").replace(r'\\',r'')
    while not os.path.isdir(path) or not os.path.exists(path):
        writeFile("目录不合法：{0}".format(path))
        path=input("请输入目录：").replace(r'\\',r'')
    path = path+'/**/*.doc'
    for file in glob.glob(path,recursive = True):
        if not file.endswith('.docx') and file.endswith('.doc') and not os.path.basename(file).startswith('~$'): 
            files.append(file)
    # for root,dirs,names in os.walk(path):
    #     for d in dirs:
    #         _dir=os.path.join(root,d)
    #         #print(_dir)
    #         _dir2=_dir.replace(os.path.dirname(_dir)+'\\','')
    #         if os.path.exists(_dir) and os.path.isdir(_dir) and _dir2!='err':
    #             fileList = readInfo(_dir)       # 读取文件夹下所有的文件名，返回一个列表
    #             for i in fileList:
    #                 rowInfo = os.path.join(_dir,i)
    #                 #rowInfo=os.path.abspath(_dir+'\\'+i)
    #                 #print(rowInfo)
    #                 files.append(rowInfo)
    #     for f in names:
    #         _f=os.path.join(root,f)
    #         files.append(_f)
    writeFile("检索完成：{}".format(len(files)))
    
    if len(files)>0:
        convertDocx(files)
    while len(errs)>0:
        print("失败队列：{}".format(len(errs)))
        convertDocx(errs)