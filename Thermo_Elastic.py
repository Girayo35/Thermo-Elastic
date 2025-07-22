import numpy as np
import pandas as pd
import openpyxl
import math
import random

wb=openpyxl.load_workbook("Thermo_Elastic_Data.xlsx")
ws=wb.active
print ("here")
file_name=["_ASG1_GLOBAL_MODEL_assyfem1_sim1-DSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-JSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-SMEQBOL_nastran_"]
time_steps=["t129120","t133200","t137760","t146400","t155040","t163680","t170160","t176400","t180960","t189600","t198240","t206880","t215520"]
i=0
for i in range (0,1):
    j=0
    for j in range (0,1):
        df = pd.read_csv(file_name[i]+time_steps[j]+".dat", delimiter='\t')
        #wb = openpyxl.load_workbook(file_name[i]+time_steps[j]+".dat")
        #ws = wb.active
        #print(ws["A"+str(1)])
        ws=df
        ws.save("Thermo_Elastic_Data.xlsx")

