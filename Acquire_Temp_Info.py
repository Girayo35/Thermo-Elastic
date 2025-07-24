import numpy as np
import pandas as pd
import openpyxl
import math
import random
Alpabet="BCDEFGHIJKLMN"

wb=openpyxl.load_workbook("All_Node_Information.xlsx")
ws=wb.active
sheet=wb["Sheet2"]
xlsx = pd.ExcelFile('All_Node_Information.xlsx')
df = xlsx.parse('Sheet1')
dt = xlsx.parse('Sheet2')
time_steps_Solstice=[129120,133200,137760,146400,155040,163680,170160,176400,180960,189600,198240,206880,210000]
time_steps_EQNX=[86400,89520,95040,103680,112320,120960,127440,131520,138240,146880,155520,164160,172800]
Panels=[]
for i in range(2,49):
    Panels.append(ws["A"+str(i)].value)
Node_List=[]

for i in range(0,len(Panels)):
    j=0
    col = df[Panels[i]]
    a=[]
    while j < len(col) and not pd.isna(col.iloc[j]):
        a.append(df[Panels[i]][j])
        j+=1
    Node_List.append(a)
 

Main_Temp_List=[]

for i in range(0,len(time_steps_EQNX)):
    Temp_List=[]
    for k in range (0,len(Node_List)):
    #for k in range (0,len(List)):
        List=Node_List[k]
        xlsx = pd.ExcelFile('Thermo_Elastic_Data.xlsx')
        df = xlsx.parse("EQNXB_"+str(time_steps_EQNX[i]))
        df.columns = [" "," ",'Node', 'Temperature']
        b=[]
        for j in range(0,len(List)):
            value = df.loc[df['Node'] == List[j], 'Temperature'].values[0]
            b.append(value)
        Temp_List.append(b)
    Main_Temp_List.append(Temp_List)
Average_Temp_List=[]

print(Main_Temp_List)

for i in range(0,len(Main_Temp_List)):
    List=Main_Temp_List[i]
    b=[]
    for j in range (0,len(List)):
        List_b=List[j]
        b.append(sum(List_b)/len(List_b))
    Average_Temp_List.append(b)

#print(Average_Temp_List,"Average_Temp_List")

for i in range(0,len(Average_Temp_List)):
    print(i,"i")
    for j in range (0,len(Average_Temp_List[i])):
        sheet[str(Alpabet[i])+str(j+2)]=str(Average_Temp_List[i][j])
        print(j,"j")
        print(str(Alpabet[i]))
wb.save("All_Node_Information_EQNX.xlsx")

