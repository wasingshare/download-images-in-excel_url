import xlrd
import requests
from io import BytesIO
from PIL import Image
import os


p=os.listdir()
for k in range(0,len(p)):
    if 'xlsx' in p[k]:
        excel_file_name=p[k]

try:
     book = xlrd.open_workbook(excel_file_name)
except:
     input("请确认excel文件是否存在！")
     exit(0)

    
try:
     sheet = book.sheet_by_index(0)
except:
     input("请确认excel文件内容！")
     exit(0)

 
row=sheet.nrows
if row==0:
     input("请确认excel文件内容！")
     exit(0)

for i in range(2,row-1):
     a=sheet.cell(i,10)##第十一列为照片URL
     b=sheet.cell(i,0)##第一列为姓名
     b=str(b)
     b=b.replace('text:',"")
     b=b.replace('\'',"")
     filename=str(i)+b+".jpg"##文件命名为序号+姓名+.jpg
     a=str(a)
     a=a.replace('text:',"")
     a=a.replace('\'',"")
     try:
         r = requests.get(a)
         f = BytesIO(r.content)
         img = Image.open(f)
         img.save(filename,'jpeg')
     except:
         print("下载失败，请确认excel中的"+b+"的图片链接是否过期！")
         continue
     f = BytesIO(r.content)
     img = Image.open(f)
     img.save(filename,'jpeg')
     print(b+"的照片生成成功")

input("Prease enter")
