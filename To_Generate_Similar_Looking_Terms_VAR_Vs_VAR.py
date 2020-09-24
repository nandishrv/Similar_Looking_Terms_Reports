import pandas as pd
Lexicon_Report=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
Variant_Report=r'D:\AIAA\AIAA Variant Report _ August_31_2020.xlsx'
save_path = r'D:\AIAA\VAR_Vs_VAR.xlsx'


import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Creating input files required"))
df = pd.read_excel(Variant_Report)
df1 =pd.read_excel(Lexicon_Report)

VAR_List=df.iloc[:,[1,0,2]]

Match_Percentage=75

VAR_List_row,VAR_List_col = VAR_List.shape

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with VAR_Vs_VAR"))
VAR1=()
VAR1_SYN_MT=()
VAR1_MT=()

VAR2=()
VAR2_SYN_MT=()
VAR2_MT=()

SEQ_RATIO=()
import difflib
#for i in range(0,10):
for i in range(0,VAR_List_row-1):
    a=VAR_List.iat[i,0]
    for j in range(i+1,VAR_List_row):
        if VAR_List.iat[i,1] != VAR_List.iat[j,1]:
            b=VAR_List.iat[j,0]
            seq = difflib.SequenceMatcher(None,a,b)
            d = seq.ratio()*100
            if float(d) >= Match_Percentage:
                VAR1+=(VAR_List.iat[i,0],)
                VAR1_SYN_MT+=(VAR_List.iat[i,1],)
                VAR2+=(VAR_List.iat[j,0],)
                VAR2_SYN_MT+=(VAR_List.iat[j,1],)

                for k in range(0,len(df1)):
                    if VAR_List.iat[i,1] == df1.iat[k,1]:
                        VAR1_MT+=(df1.iat[k,3],)
                    
                    if VAR_List.iat[j,1] == df1.iat[k,1]:
                        VAR2_MT+=(df1.iat[k,3],)
                        
                SEQ_RATIO+=(d,)
    PER=(i+1)/VAR_List_row*100
    print("\nVAR_Vs_VAR", PER,"% Completed")
VAR_Vs_VAR = pd.DataFrame(list(zip(VAR1, VAR1_SYN_MT, VAR1_MT,VAR2, VAR2_SYN_MT, VAR2_MT,SEQ_RATIO)), columns =['VAR1', 'VAR1_SYN_MT', 'VAR1_MT','VAR2', 'VAR2_SYN_MT','VAR2_MT', 'SEQ_RATIO']) 
#VAR_Vs_VAR=VAR_Vs_VAR.sort_values(by='SEQ_RATIO', ascending=False)
VAR_Vs_VAR=VAR_Vs_VAR.round({"SEQ_RATIO":2})
#print(VAR_Vs_VAR)
VAR_Vs_VAR.to_excel (save_path, index = None, header=True)
import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with VAR_Vs_VAR"))
