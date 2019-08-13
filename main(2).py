import re
import xlsxwriter

#读取文本数据并存储
file=open("China_Daily2019-07-14_1-200.TXT","r",encoding="utf-8-sig")
data=file.read()
file.close()

#数据按行分割，并去除空字符串
lines=data.split("\n")
lines=list(filter(None,lines))

#初始化，用于存放文档数据
info_all=[]

#每个文档开始的正则匹配
num_pattern = re.compile(r'\s*\d+\sof\s\d+\sDOCUMENTS')
#时间匹配
date_pattern=re.compile(r'\s*\w+\s\d{1,2},\s\d{4}\s\w+')
#HEADLINE匹配
headline_pattern=re.compile(r'HEADLINE: .+')
#BYLINE匹配
byline_pattern=re.compile(r'BYLINE: .+')
#BODY匹配
body_pattern=re.compile(r'BODY:')
#endline匹配
end_pattern=re.compile(r'\(China\sDaily.+\)')

have_match=0#判断目前匹配进行到哪一程度的布尔值
for line in lines:
    num_match = num_pattern.match(line)
    if num_match:#发现了新的number数据
        if have_match==1:#若已经进行了一次数据匹配，则把读取到的数据存到info_all中
            info_all.append([number,date,headline,byline,body,end])  
        #新的number数据的数据存放初始化
        number=re.findall(r"\s*(.+?)DOCUMENTS",line)[0]
        date=""
        headline=""
        byline=""
        body=""
        end=""
        have_match=0

    if have_match==0:
        #开始逐行匹配除body,end以外的基本信息
        date_match=date_pattern.match(line)
        if date_match:
            date=re.findall(r"\s*(.+?)\s\D+",line)[0]
            continue
        headline_match=headline_pattern.match(line)
        if headline_match:
            headline=re.findall(r"HEADLINE:\s(.+)",line)[0]
            continue
        byline_match=byline_pattern.match(line)
        if byline_match:
            byline=re.findall(r"BYLINE:\s(.+)",line)[0]
            continue
        #开始准备存储body信息
        body_match=body_pattern.match(line)
        if body_match:
            have_match=1
            continue
    else:
        #存储body信息，并时刻观察有无end出现
        end_match=end_pattern.match(line)
        if end_match:
            end=re.findall(r"\((.+?)\)",line)[0]
        else:
            body+=' '
            body+=line
#把最后一次的number数据加上
info_all.append([number,date,headline,byline,body,end])

#开始准备存入Excel表格
workbook=xlsxwriter.Workbook('output.xlsx')
worksheet=workbook.add_worksheet()
for head in enumerate(["number","date","HEADLINE","BYLINE","BODY","END"]):
    worksheet.write(0,head[0],head[1])
for info in enumerate(info_all):
    for col in enumerate(info[1]):
        worksheet.write(info[0]+1,col[0],col[1])
workbook.close()