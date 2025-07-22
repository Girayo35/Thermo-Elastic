import numpy as np
import pandas as pd
import openpyxl
import re
import random

wb=openpyxl.load_workbook("Thermo_Elastic_Data.xlsx")
ws=wb.active
print ("here")
file_name=["_ASG1_GLOBAL_MODEL_assyfem1_sim1-DSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim2-JSEOL_nastran_","_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_"]
time_steps=["t129120","t133200","t137760","t146400","t155040","t163680","t170160","t176400","t180960","t189600","t198240","t206880","t215520"]
time_steps_1=["t172800","t175920","t181440","t190080","t198720","t207360","t213840","t217920","t224640","t233280","t241920","t250560","t259200"]
i=0
used_names = set()

def make_unique(name, used):
    base = name[:31]
    while base in used:
        base += '_'
        base = base[:31]
    used.add(base)
    return base

with pd.ExcelWriter('Thermo_Elastic_Data.xlsx', engine='xlsxwriter') as writer:
    for i in range(len(file_name)):
        if file_name[i] != "_ASG1_GLOBAL_MODEL_assyfem1_sim1-EQNXBOL_nastran_":
            for j in range(len(time_steps)):
                try:
                    df = pd.read_csv(file_name[i] + time_steps[j] + ".dat", sep='\s+', header=None)
                    if not df.empty:
                        name = f"{file_name[i][33:38]}_{time_steps[j][1:8]}"
                        name = re.sub(r'[:\\/?*\[\]]', '_', name)
                        name = make_unique(name, used_names)
                        df.to_excel(writer, sheet_name=name, index=False, header=False)
                except Exception as e:
                    print(f"Error with {file_name[i]}{time_steps[j]}: {e}")
        else:
            for j in range(len(time_steps_1)):
                try:
                    df = pd.read_csv(file_name[i] + time_steps_1[j] + ".dat", sep='\s+', header=None)
                    if not df.empty:
                        name = f"{file_name[i][33:38]}_{time_steps_1[j][1:8]}"
                        name = re.sub(r'[:\\/?*\[\]]', '_', name)
                        name = make_unique(name, used_names)
                        df.to_excel(writer, sheet_name=name, index=False, header=False)
                except Exception as e:
                    print(f"Error with {file_name[i]}{time_steps_1[j]}: {e}")