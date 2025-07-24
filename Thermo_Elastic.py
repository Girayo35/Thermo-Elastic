import numpy as np
import pandas as pd
import openpyxl
import math
import random

wb=openpyxl.load_workbook("Thermo_Elastic_Data.xlsx")
ws=wb.active
print("here")
file_name=["_ASG1_GLOBAL_MODEL_assyfem1_sim1-DSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-JSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_"]
time_steps=["t129120","t133200","t137760","t146400","t155040","t163680","t170160","t176400","t180960","t189600","t198240","t206880","t210000"]
time_steps_1=["t86400","t89520","t95040","t103680","t112320","t120960","t127440","t131520","t138240","t146880","t155520","t164160","t172800"]
i=0
with pd.ExcelWriter('Thermo_Elastic_Data.xlsx', engine='xlsxwriter') as writer:
    for i in range (0,len(file_name)):
        j=0
        if file_name[i]!="_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_":
            for j in range (0,len(time_steps)):
                df = pd.read_csv(file_name[i]+time_steps[j]+".dat", sep=r'\s+',header=None)
                df.to_excel(writer, sheet_name=str(file_name[i][33:38])+"_"+str(time_steps[j][1:8]), index=False, header=False)
                j+=1
        else:
            for j in range (0,len(time_steps_1)):
                df1 = pd.read_csv(file_name[i]+time_steps_1[j]+".dat", sep=r'\s+',header=None)
                df1.to_excel(writer, sheet_name=str(file_name[i][34:39])+"_"+str(time_steps_1[j][1:8]), index=False, header=False)
                j+=1