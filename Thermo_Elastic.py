import numpy as np
import pandas as pd
import openpyxl
import math
import random

wb=openpyxl.load_workbook("Thermo_Elastic_Data.xlsx")
ws=wb.active
print ("here")
file_name=["_ASG1_GLOBAL_MODEL_assyfem1_sim1-DSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim2-JSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_"]
time_steps=["t129120","t133200","t137760","t146400","t155040","t163680","t170160","t176400","t180960","t189600","t198240","t206880","t215520"]
time_steps_1=["t172800","t175920","t181440","t190080","t198720","t207360","t213840","t217920","t224640","t233280","t241920","t250560","t259200"]
i=0
with pd.ExcelWriter('Thermo_Elastic_Data.xlsx', engine='xlsxwriter') as writer:
    for i in range (0,len(file_name)):
        j=0
        if file_name[i]!="_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_":
            for j in range (0,len(time_steps)):
                df = pd.read_csv(file_name[i]+time_steps[j]+".dat", sep='\s+',header=None)
                df.to_excel(writer, sheet_name=str(file_name[i][33:38])+"_"+str(time_steps[j][1:8]), index=False, header=False)
                j+=1
        else:
            for j in range (0,len(time_steps_1)):
                df1 = pd.read_csv(file_name[i]+time_steps_1[j]+".dat", sep='\s+',header=None)
                df1.to_excel(writer, sheet_name=str(file_name[i][34:39])+"_"+str(time_steps_1[j][1:8]), index=False, header=False)
                j+=1