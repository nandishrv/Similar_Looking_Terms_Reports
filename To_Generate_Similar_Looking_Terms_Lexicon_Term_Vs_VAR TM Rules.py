import pandas as pd
Lexicon_Report=r'D:\AIAA\AIAA_Lexicon Report_August_31_2020.xlsx'
Variant_Report=r'D:\AIAA\AIAA Variant Report _ August_31_2020.xlsx'
save_path = r'D:\AIAA\Lexicon_Term_Vs_VAR_TM_Rule.xlsx'

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
Lexicon_Term_List =df
Lexicon_Term_List_row,Lexicon_Term_List_col = Lexicon_Term_List.shape
VAR_List_row,VAR_List_col = VAR_List.shape

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with Lexicon_Term_Vs_VAR"))

Lexicon_Term=()
Lexicon_Term_MT=()
Lexicon_Term_TM=()
VAR=()
VAR_SYN_MT=()
VAR_MT=()
VAR_TM=()
SEQ_RATIO=()

import difflib
for i in range(0,Lexicon_Term_List_row):
	a=Lexicon_Term_List.iat[i,1]
	for j in range(0,VAR_List_row):
		b=VAR_List.iat[j,0]
		seq = difflib.SequenceMatcher(None,a,b)
		d = seq.ratio()*100
		if float(d) >= Match_Percentage:
			if  Lexicon_Term_List.iat[i,13] != VAR_List.iat[j,2]:
				Lexicon_Term+=(Lexicon_Term_List.iat[i,1],)
				Lexicon_Term_MT+=(Lexicon_Term_List.iat[i,3],)
				Lexicon_Term_TM+=(Lexicon_Term_List.iat[i,13],)
				VAR+=(VAR_List.iat[j,0],)
				VAR_SYN_MT+=(VAR_List.iat[j,1],)
				VAR_TM+=(VAR_List.iat[j,2],)
				SEQ_RATIO+=(d,)
				for k in range(0,Lexicon_Term_List_row):
					if VAR_List.iat[j,1] == Lexicon_Term_List.iat[k,1]:
						VAR_MT+=(Lexicon_Term_List.iat[k,3],)
						break;
						
	PER=(i+1)/Lexicon_Term_List_row*100
	
	print("\nLexicon_Term_Vs_VAR_TM_Rule", PER,"% Completed")

Lexicon_Term_Vs_VAR_TM_Rule = pd.DataFrame(list(zip(Lexicon_Term, Lexicon_Term_MT, Lexicon_Term_TM, VAR,VAR_SYN_MT,VAR_MT,VAR_TM,SEQ_RATIO)), columns =['Lexicon_Term', 'Lexicon_Term_MT', 'Lexicon_Term_TM', 'VAR','VAR_SYN_MT','VAR_MT','VAR_TM','SEQ_RATIO']) 
Lexicon_Term_Vs_VAR_TM_Rule=Lexicon_Term_Vs_VAR_TM_Rule.sort_values(by='SEQ_RATIO', ascending=False)
Lexicon_Term_Vs_VAR_TM_Rule=Lexicon_Term_Vs_VAR_TM_Rule.round({"SEQ_RATIO":2})
#print(Lexicon_Term_Vs_VAR_TM_Rule )
Lexicon_Term_Vs_VAR_TM_Rule.to_excel (save_path, index = None, header=True)

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with Lexicon_Term_Vs_VAR_TM_Rule"))
