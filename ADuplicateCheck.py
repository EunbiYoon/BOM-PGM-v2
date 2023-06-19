import pandas as pd
import openpyxl

today_date="0623"
DR_final=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/'+str(today_date)+'/DR/final_table.xlsx')
FL_final=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/'+str(today_date)+'/FL/final_table.xlsx')
TL_final=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/'+str(today_date)+'/TL/final_table.xlsx')

DR_result=DR_final["Seq"]
FL_result=FL_final["Seq"]
TL_result=TL_final["Seq"]

DR_result=DR_result.dropna()
DR_result=DR_result.sort_values()
DR_result.reset_index(drop=True, inplace=True)

FL_result=FL_result.dropna()
FL_result=FL_result.sort_values()
FL_result.reset_index(drop=True, inplace=True)

TL_result=TL_result.dropna()
TL_result=TL_result.sort_values()
TL_result.reset_index(drop=True, inplace=True)



# DR - exceptional - 120
DR_exc=120
DR_result1=DR_result[:DR_exc]
DR_result2=DR_result[DR_exc:]
DR_result2.reset_index(inplace=True, drop=True)

# TL - expectioanl - 66
TL_exc=66
TL_result1=TL_result[:TL_exc]
TL_result2=TL_result[TL_exc:]
TL_result2.reset_index(inplace=True, drop=True)

# compare index 
DR_result1=pd.DataFrame(DR_result1)
DR_result1.reset_index(inplace=True)
DR_result2=pd.DataFrame(DR_result2)
DR_result2.reset_index(inplace=True)

FL_result=pd.DataFrame(FL_result)
FL_result.reset_index(inplace=True)

TL_result1=pd.DataFrame(TL_result1)
TL_result1.reset_index(inplace=True)
TL_result2=pd.DataFrame(TL_result2)
TL_result2.reset_index(inplace=True)


# minus 
DR_error1=pd.DataFrame(DR_result1["Seq"]-DR_result1["index"])
DR_error2=pd.DataFrame(DR_result2["Seq"]-DR_result2["index"])

FL_error=pd.DataFrame(FL_result["Seq"]-FL_result["index"])

TL_error1=pd.DataFrame(TL_result1["Seq"]-TL_result1["index"])
TL_error2=pd.DataFrame(TL_result2["Seq"]-TL_result2["index"])



# error check
data=DR_error1
condition=0
message="DR below "+str(DR_exc)+" has error"
for i in range(len(data)):
    error=data.at[i,0]
    if error!=condition:
        print(message)

data=DR_error2
condition=DR_exc+1
message="DR above "+str(DR_exc)+" has error"
for i in range(len(data)):
    error=data.at[i,0]
    if error!=condition:
        print(message)

data=FL_error
condition=0
message="FL has error"
for i in range(len(data)):
    error=data.at[i,0]
    if error!=condition:
        print(message)

data=TL_error1
condition=0
message="TL below "+str(TL_exc)+" has error"
for i in range(len(data)):
    error=data.at[i,0]
    if error!=condition:
        print(message)

data=TL_error2
condition=TL_exc+1
message="TL above "+str(TL_exc)+" has error"
for i in range(len(data)):
    error=data.at[i,0]
    if error!=condition:
        print(message)
