import pandas as pd
import openpyxl
plan_csv=pd.read_excel("1.6排期.xlsx")
line1=list(plan_csv["班别"])
style1=list(plan_csv["款号"])
start_date1=list(plan_csv["开始出产"])
finish_date1=list(plan_csv["完成生产"])
line=["Line"]
style=["style"]
start_date=["Start date"]
finish_date=["Finish date"]
line_simple=[]
c=-1
i=int
temp=""
count=len(list(plan_csv["班别"]))-1
for i in range(count):
  if i==0:
    i=i+1
  elif i>0:
    if style1[i]!=style1[i-1] and style1[i]!="-":
      line.append(line1[i])
      start_date.append(start_date1[i])
      style.append(style1[i])
      finish_date.append(finish_date1[i])
                
                   
x=len(line)
y=len(style)
nline=[]
nstyle=[]
nstart_date=[]
nfinish_date=[]
i=0
j=0
for i in range(0,x):
  if not start_date[i]==finish_date[i]:
    nline.append(line[i])
    nstyle.append(style[i])
    nstart_date.append(start_date[i])
    nfinish_date.append(finish_date[i])
    i=i+1

new_data1=pd.DataFrame({'line': nline,'style': nstyle,'start date':nstart_date,'finish date':nfinish_date})  
#new_data=percentile_list = pd.DataFrame({'line': line,'style': style,'start date':start_date,'finish date':finish_date})
#print(df_new)
with pd.ExcelWriter("1.6.xlsx") as writer:
    new_data1.to_excel(writer, sheet_name="Sheet1")  
