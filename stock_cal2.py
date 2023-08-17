import xml.etree.ElementTree as ET
import requests
import time
from openpyxl import Workbook
def xml_to_dict(element):
    result={}
    for child in element:
        if len(child)==0:
            result[child.tag]=child.text
        else:
            result[child.tag]=xml_to_dict(child)
    return result
def returnStrDayList(startyear,startmonth,endyear,endmonth,day="01"):
    result=[]
    startyear,startmonth=int(startyear),int(startmonth)
    endyear,endmonth=int(endyear),int(endmonth)
    if startyear==endyear:
        for month in range(startmonth,endmonth):
            month=str(month)
            if len(month)==1:
                month="0"+month
            result.append(str(startyear)+month+day)
        return result
    for year in range(startyear,endyear+1):
        if year==startyear:
            for month in range(startmonth,13):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+day)
        elif year==startyear:
            for month in range(1,endmonth+1):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+day)
        else:
            for month in range(1,13):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+day)
    return result 
def fillsheet(sheet,data,row):
    for column, value in enumerate(data,1):
        sheet.cell(row=row,column=column,value=value)      
tree=ET.parse("setting2.xml")
root=tree.getroot()
data=xml_to_dict(root)
fields=["date","profit margin","closing price","starting price","highest point","lowest point","transaction amount","symphony of transaction"]
print(data)
startyear,startmonth=data["startyear"],data["startmonth"]
endyear,endmonth=data["endyear"],data["endmonth"]
yearlist=returnStrDayList(startyear,startmonth,endyear,endmonth)
wb=Workbook()
sheet=wb.active
sheet.title="Excel Everday records"
row=1
fillsheet(sheet,fields,row)
row=row+1
for month in yearlist:
    rq=requests.get(data["url"],params={
        "date":month,
        "stockNo":data["stockNo"],
        "response":"json"
    })          
    print(rq)
    if str(rq)!="<Response [200]>":
        break
    jsondata=rq.json()
    dailypriceslist=jsondata.get("data",[])
    for dailyprices in dailypriceslist:
        fillsheet(sheet,dailyprices,row)
        row=row+1
    time.sleep(3)
name=data["excelname"]
wb.save(name+".xlsx")
