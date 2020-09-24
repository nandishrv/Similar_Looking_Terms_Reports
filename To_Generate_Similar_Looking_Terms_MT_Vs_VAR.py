import pandas as pd
Lexicon_Report=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
Variant_Report=r'D:\AIAA\AIAA Variant Report _ August_31_2020.xlsx'
save_path = r'D:\AIAA\MT_Vs_VAR.xlsx'

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

MT_List_row,MT_List_col = MT_List.shape
VAR_List_row,VAR_List_col = VAR_List.shape

#MT_List_row=25

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with MT_Vs_VAR"))

MT=()
VAR=()
VAR_SYN_MT=()
VAR_MT=()
SEQ_RATIO=()
import difflib
for i in range(0,MT_List_row):
	a=MT_List.iat[i,1]
	for j in range(0,VAR_List_row):
		b=VAR_List.iat[j,0]
		seq = difflib.SequenceMatcher(None,a,b)
		d = seq.ratio()*100
		if (float(d) >= Match_Percentage) and (d!=100):
			if (MT_List.iat[i,1] != VAR_List.iat[j,1]):
				MT+=(MT_List.iat[i,1],)
				VAR+=(VAR_List.iat[j,0],)
				VAR_SYN_MT+=(VAR_List.iat[j,1],)
				SEQ_RATIO+=(d,)
				for k in range(0,len(df)):
					if VAR_List.iat[j,1] == df.iat[k,1]:
						VAR_MT+=(df.iat[k,3],)
						break;
	PER=(i+1)/MT_List_row*100
	print("\nMT_Vs_VAR", PER,"% Completed")
MT_Vs_VAR = pd.DataFrame(list(zip(MT, VAR, VAR_SYN_MT,VAR_MT,SEQ_RATIO)), columns =['MT', 'VAR', 'VAR_SYN_MT', 'VAR_MT','SEQ_RATIO']) 
#MT_Vs_VAR=MT_Vs_VAR.sort_values(by='SEQ_RATIO', ascending=False)
MT_Vs_VAR=MT_Vs_VAR.round({"SEQ_RATIO":2})
#print(MT_Vs_VAR)
MT_Vs_VAR.to_excel (save_path, index = None, header=True)

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with MT_Vs_VAR"))
