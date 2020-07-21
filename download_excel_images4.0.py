import xlrd
import requests
from io import BytesIO
from PIL import Image
import time
import os
# 更新说明
# V1.0:可下载火瞳主版本平台人员信息列表导出的excel的照片 date:20191226
# V2.0:新增人员信息收集平台导出的excel（两种模式，企业和学校） date:20200327
# V3.0:适配一个人员信息存在多个照片链接的情况 date:20200609
# V3.1:打印日志 date:20200622
# V4.0:新增【新火瞳平台人员信息导出的列表人脸照片下载】
# 作者：tellmorning@outlook.com

flag = 0
print("*****************************************************************************************************************")
input("*\t欢迎使用此程序，此程序仅用于火瞳产品根据excel中的url下载生成图片保存使用，输入enter开始使用\t\t\t*")

logFile = open('log.log', 'w')
# 选择excel类型
while flag == 0:
    excel_type = input("*\t输入1：老火瞳主版本人员信息excel的照片下载。姓名第1列，照片第11列，从第3行开始有数据\t\t\t\t*"
                       "\n*\t输入2：人员信息收集平台企业excel照片下载。姓名第3列，照片第9列，从第2行开始有数据\t\t\t*"
                       "\n*\t输入3：人员信息收集平台学校excel照片下载。姓名第3列，照片第8列，从第2行开始有数据\t\t\t*"
                       "\n*\t输入4：新火瞳人员信息收集平台企业excel照片下载。姓名第1列，照片第13列，从第3行开始有数据\t\t\t"
                       "\n*\t请输入后按enter确认：\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t*")
    if excel_type == "1":
        name_col = 0
        image_col = 10
        data_row = 2
        flag = 1
    elif excel_type == "2":
        name_col = 2
        image_col = 8
        data_row = 1
        flag = 1
    elif excel_type == "3":
        name_col = 2
        image_col = 7
        data_row = 1
        flag = 1
    elif excel_type == "4":
        name_col = 0
        image_col = 12
        data_row = 2
        flag = 1
    else:
        print("*\t输入有误，请重新输入\n请输入后按enter确认：\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t*")
    # 自动获取目录下的excel文件
p = os.listdir()
for k in range(0, len(p)):
    if 'xlsx' in p[k]:
        excel_file_name = p[k]

# 读取获取到的excel文件
try:
    book = xlrd.open_workbook(excel_file_name)
except Exception as e:
    logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + '：' + str(e) + '\n')
    input("请确认excel文件是否存在！")
    exit(0)

logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + '：' + '读取excel文件成功！' + '\n')
# 读取获取到的excel文件的sheet
try:
    sheet = book.sheet_by_index(0)
except Exception as e:
    logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + '：' + str(e) + '\n')
    input("请确认excel文件内容！")
    exit(0)
logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + '：' + '读取sheet页文件成功！' + '\n')
row = sheet.nrows
if row == 0:
    input("请确认excel文件内容！")
    exit(0)
succeed = 0
namelist = ""
except_data_num = 0
# num = (len(name)-len(name.replace('朱兴', "")))//len('朱兴')
for i in range(data_row, row):  # 第data_row行为有效数据
    a = sheet.cell(i, image_col)  # 第image_col+1列为照片URL
    b = sheet.cell(i, name_col)  # 第name_col+1列为姓名
    b = str(b)
    b = b.replace('text:', "")
    b = b.replace('\'', "")
    namelist = namelist + '_' + b
    num = (len(namelist) - len(namelist.replace(b, ""))) // len(b) - 1
    if num > 0:
        b = b + str(num)
    filename = b + ".jpg"  # 文件命名为序号+姓名+.jpg
    a = str(a)
    a = a.replace('text:', "")
    a = a.replace('\'', "")
    alist = a.split(',')
    kflag = 0
    for urlV in alist:
        try:
            r = requests.get(urlV)
            f = BytesIO(r.content)
            img = Image.open(f)
            if kflag > 0:
                filename = str(kflag) + filename
            kflag = kflag+1
            img.save(filename, 'jpeg')
        except Exception as e:
            logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                          + '：' + '链接' + str(urlV) + str(e) + '\n')
            print(filename.replace('.jpg', '') + "下载失败，请确认excel中的" + b + "的图片链接是否过期！")
            except_data_num = except_data_num+1
            continue
        f = BytesIO(r.content)
        img = Image.open(f)
        img.save(filename, 'jpeg')
        succeed = succeed + 1
        logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) +
                      '：' + filename.replace('.jpg', '') + "的照片生成成功" + '\n')
        print(filename.replace('.jpg', '') + "的照片生成成功")
print("总共" + str(succeed+except_data_num) + " 张照片 " + str(succeed) + " 张成功下载")
logFile.write(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + '：' + '任务结束，' + str(succeed) +
              '张照片生成成功！' + '\n')
logFile.close()
input("照片已在此应用软件同目录生成，输入enter退出软件,谢谢使用")
