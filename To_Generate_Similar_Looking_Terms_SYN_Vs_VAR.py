import pandas as pd
Lexicon_Report=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
Variant_Report=r'D:\AIAA\AIAA Variant Report _ August_31_2020.xlsx'
save_path = r'D:\AIAA\SYN_Vs_VAR.xlsx'

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Creating input files required"))
df = pd.read_excel(Lexicon_Report)
df1 = pd.read_excel(Variant_Report)

MT_List=df[df['Is main term']==1]

SYN_List=df[df['Is main term']==0]

ACR_List=df[df['isAcronym (in Lexicon)']==1]

VAR_List=df1.iloc[:,[1,0,2]]

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Created input files required"))

Match_Percentage=75

SYN_List_row,MT_List_col = SYN_List.shape
VAR_List_row,VAR_List_col = VAR_List.shape

#SYN_List_row=25

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with SYN_Vs_VAR"))

SYN=()
SYN_MT=()
VAR=()
VAR_SYN_MT=()
VAR_MT=()
SEQ_RATIO=()
import difflib
for i in range(0,SYN_List_row):
    a=SYN_List.iat[i,1]
    for j in range(0,VAR_List_row):
        if  (SYN_List.iat[i,1] != VAR_List.iat[j,1]) and (SYN_List.iat[i,3] != VAR_List.iat[j,1]) and (SYN_List.iat[i,3] != VAR_List.iat[j,2]):
            b=VAR_List.iat[j,0]
            seq = difflib.SequenceMatcher(None,a,b)
            d = seq.ratio()*100
            if float(d) >= Match_Percentage:
                SYN+=(SYN_List.iat[i,1],)
                SYN_MT+=(SYN_List.iat[i,3],)
                VAR+=(VAR_List.iat[j,0],)
                VAR_SYN_MT+=(VAR_List.iat[j,1],)
                SEQ_RATIO+=(d,)
                for k in range(0,len(df)):
                    if VAR_List.iat[j,1] == df.iat[k,1]:
                        VAR_MT+=(df.iat[k,3],)
                        break;
    PER=(i+1)/SYN_List_row*100
    print("\nSYN_Vs_VAR", PER,"% Completed")
SYN_Vs_VAR = pd.DataFrame(list(zip(SYN, SYN_MT, VAR, VAR_SYN_MT,VAR_MT,SEQ_RATIO)), columns =['SYN', 'SYN_MT', 'VAR', 'VAR_SYN_MT','VAR_MT','SEQ_RATIO']) 
#SYN_Vs_VAR=SYN_Vs_VAR.sort_values(by='SEQ_RATIO', ascending=False)
SYN_Vs_VAR=SYN_Vs_VAR.round({"SEQ_RATIO":2})
#print(SYN_Vs_VAR)
SYN_Vs_VAR.to_excel (save_path, index = None, header=True)

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with SYN_Vs_VAR"))
