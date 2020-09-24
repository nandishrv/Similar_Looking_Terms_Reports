import pandas as pd
Lexicon_Report_Link=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
save_path = r'D:\AIAA\SYN_Vs_SYN.xlsx'

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Creating input files required"))
df = pd.read_excel(Lexicon_Report_Link)

MT_List=df[df['Is main term']==1]

SYN_List=df[df['Is main term']==0]

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Created input files required"))

Match_Percentage=75

SYN_List_row,SYN_List_col = SYN_List.shape

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with SYN_Vs_SYN"))
SYN1=()
SYN1_MT=()
SYN2=()
SYN2_MT=()
SEQ_RATIO=()
import difflib
#for i in range(0,10):
for i in range(0,SYN_List_row-1):
	a=SYN_List.iat[i,1]
	for j in range(i+1,SYN_List_row):
		b=SYN_List.iat[j,1]
		seq = difflib.SequenceMatcher(None,a,b)
		d = seq.ratio()*100
		if (float(d) >= Match_Percentage) and (SYN_List.iat[i,3]!=SYN_List.iat[j,3]):
			SYN1+=(SYN_List.iat[i,1],)
			SYN1_MT+=(SYN_List.iat[i,3],)
			SYN2+=(SYN_List.iat[j,1],)
			SYN2_MT+=(SYN_List.iat[j,3],)
			SEQ_RATIO+=(d,)
	PER=(i+1)/SYN_List_row*100
	print("\nSYN_Vs_SYN", PER,"% Completed")
SYN_Vs_SYN = pd.DataFrame(list(zip(SYN1, SYN1_MT, SYN2, SYN2_MT,SEQ_RATIO)), columns =['SYN1', 'SYN1_MT','SYN2', 'SYN2_MT','SEQ_RATIO']) 
#SYN_Vs_SYN=SYN_Vs_SYN.sort_values(by='SEQ_RATIO', ascending=False)
SYN_Vs_SYN=SYN_Vs_SYN.round({"SEQ_RATIO":2})
#print(SYN_Vs_SYN)
SYN_Vs_SYN.to_excel (save_path, index = None, header=True)
import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with SYN_Vs_SYN"))
