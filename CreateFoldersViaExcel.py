#coding=UTF-8
import xlrd
import os
import shutil
import sys
sys.path.append("../your/target/path/")

def read_xls(path):
    xl = xlrd.open_workbook(path)
    sheet = xl.sheets()[eval(input("请输入需要提取数据的sheet序号"))-1] # 0表示读取第一个工作表sheet
    datas = []
    l_space = []
    for i in range(0,sheet.nrows):
        datas=datas+['']
        l_space=l_space+[' ']
    cols=list(map(int, input("请输入需要提取数据的列序号,如果有多列，用空格隔开").split()))
    for i in cols:
        data=sheet.col_values(i-1)
        if i != cols[-1]:
            data=list(map(lambda x,y: x+y, data,l_space))
        datas=list(map(lambda x,y: x+y, datas,data))
    return datas
# xlrd为第三方包，可以通过用pip下载，具体操作：打开运行，输入cmd→在cmd中输入pip install xlrd，enter →等待安装完成即可。在后续若存在需要使用的第三方包，都可以通过这种方式下载和安装。
# 传入参数为path，path为excel所在路径。
# 传入的path需如下定义：path= r’ D:\excel.xlsx’或path= r’ D:\excel.xls’
# col_values(i)表示按照第i+1列中的所有单元格遍历读取
# return data ：返回的data是一个列表，列表每个元素是对应位置所有元素组成的字符串。

def buildfile(echkeyfile):
    if os.path.exists(echkeyfile):
            #创建前先判断是否存在文件夹，if存在则删除
            shutil.rmtree(echkeyfile)
            os.makedirs(echkeyfile)
    else:
            os.makedirs(echkeyfile)#else则创建语句
    return echkeyfile#返回创建路径
#传入的参数是需要创建文件夹的路径，比如我想在D盘下创建一个名字为newfile的文件夹，则传入参数为r’ D:\newfile’。同样，返回的参数也是r’ D:\newfile’

path=input("请输入Excel文件路径")
root=input("请输入创建的文件夹目标路径")
for i in read_xls(path):
    buildfile(root+'/'+i)

print("文件夹创建成功！")
