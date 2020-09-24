import pandas as pd
Lexicon_Report=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
save_path=r'D:\AIAA\MT_Vs_SYN.xlsx'

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Creating input files required"))
df = pd.read_excel(Lexicon_Report)

MT_List=df[df['Is main term']==1]

SYN_List=df[df['Is main term']==0]

ACR_List=df[df['isAcronym (in Lexicon)']==1]

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Created input files required"))

Match_Percentage=75

MT_List_row,MT_List_col = MT_List.shape
SYN_List_row,SYN_List_col = SYN_List.shape

#MT_List_row=25

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with MT_Vs_SYN"))

MT=()
SYN=()
SYN_MT=()
SEQ_RATIO=()
import difflib
for i in range(0,MT_List_row):
	a=MT_List.iat[i,1]
	for j in range(0,SYN_List_row):
		b=SYN_List.iat[j,1]
		seq = difflib.SequenceMatcher(None,a,b)
		d = seq.ratio()*100
		if float(d) >= Match_Percentage:
			if  MT_List.iat[i,1] != SYN_List.iat[j,3]: 
				MT+=(MT_List.iat[i,1],)
				SYN+=(SYN_List.iat[j,1],)
				SYN_MT+=(SYN_List.iat[j,3],)
				SEQ_RATIO+=(d,)
	PER=(i+1)/MT_List_row*100
	print("\nMT_Vs_SYN", PER,"% Completed")
MT_Vs_SYN = pd.DataFrame(list(zip(MT, SYN, SYN_MT,SEQ_RATIO)), columns =['MT', 'SYN', 'SYN_MT','SEQ_RATIO']) 
#MT_Vs_SYN=MT_Vs_SYN.sort_values(by='SEQ_RATIO', ascending=False)
MT_Vs_SYN=MT_Vs_SYN.round({"SEQ_RATIO":2})
#print(MT_Vs_SYN)
MT_Vs_SYN.to_excel (save_path, index = None, header=True)

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with MT_Vs_SYN"))
