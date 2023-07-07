import pandas as pd
from json2html import *

# today date
today_date="0707"
model_name="DR"

# model -> sheet name matching
if model_name=="FL":
    model_sheet="F3P2CYUBW.ABWEUUS"
elif model_name=="DR":
    model_sheet="RV13D1AMAZU.ABWEUUS"
elif model_name=="TL":
    model_sheet="T1889EFHUW.ABWEUUS"

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/"+today_date+"/result_"+today_date+".xlsx", sheet_name=model_sheet+'_result')
read_excel.index=read_excel['Unnamed: 1']
read_excel=read_excel.drop(['Unnamed: 1','Unnamed: 0'],axis=1)
read_excel=read_excel.T
extract_data=read_excel[["PAC Net - BOM Net","Price Change","Substitute"]]
extract_data["PO + Substitute"]=extract_data["Price Change"]+extract_data["Substitute"]

#column_list
column_list=list(extract_data.index)

#index_list
value1=list(extract_data["PAC Net - BOM Net"].round(1))
value2=list(extract_data["Price Change"].round(1))
value3=list(extract_data["Substitute"].round(1))
value4=list(extract_data["PO + Substitute"].round(1))

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print("")
print(column_json)

# data,index json file format
data_json=str({"PAC Net - BOM Net":value1,
               "Price Change":value2, 
               "Substitute":value3,
               "PO + Substitute":value4}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print(","+data_json)
print(" ")



