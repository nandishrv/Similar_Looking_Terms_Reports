import pandas as pd
import sys
import csv
import glob
import os

# get data file names
path =r'D:\AIAA\DepthWise'
save_path=r'D:\AIAA\MT_Vs_MT.xlsx'

filenames = glob.glob(path + "/*.xlsx")
#print(filenames)
files_xls = [f for f in filenames if f[-4:] == 'xlsx']

a=()
for i in range(0,len(files_xls)):
	base=os.path.basename(files_xls[i])
	os.path.splitext(base)
	a+=(os.path.splitext(base)[0],)
#print(a)


df2 = pd.DataFrame()
i=0
for f in filenames:
	data = pd.read_excel(f, header=None)
	data=data.replace(data.iat[0,0], a[i] )
	i+=1
	df2 = df2.append(data)
	
df2_row, df2_col=df2.shape

NT=()
BT=()
for i in range(0,df2_row):
	for j in range(df2_col-1,-1,-1):
		if str(df2.iat[i,j])!=str(df2.iat[0,1] ) : 
			NT+=(df2.iat[i,j],)
			BT+=(df2.iat[i,j-1],)
			break	

BT_NT = pd.DataFrame(list(zip(NT, BT)), columns =['NT', 'BT']) 
BT_NT = BT_NT.sort_values('NT')
BT_NT = BT_NT.drop_duplicates() 

#########################################################

Match_Percentage=75

BT_NT_row,BT_NT_col=BT_NT.shape

import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generating list of similar looking terms with MT_Vs_MT"))


MT1=()
MT2=()
MT1_BT=()
MT2_BT=()
SEQ_RATIO=()
import difflib
#for i in range(0,100):
for i in range(0,BT_NT_row-1):
	a=BT_NT.iat[i,0]
	for j in range(i+1,BT_NT_row):
		b=BT_NT.iat[j,0]
		seq = difflib.SequenceMatcher(None,a,b)
		d = seq.ratio()*100
		if (float(d) >= Match_Percentage) and (d!=100):
			if (BT_NT.iat[i,1] != BT_NT.iat[j,1]) and (BT_NT.iat[i,0] != BT_NT.iat[j,1]) and (BT_NT.iat[i,1] != BT_NT.iat[j,0]):
				MT1+=(BT_NT.iat[i,0],)
				MT1_BT+=(BT_NT.iat[i,1],)
				MT2+=(BT_NT.iat[j,0],)
				MT2_BT+=(BT_NT.iat[j,1],)
				SEQ_RATIO+=(d,)		

	PER=(i+1)/BT_NT_row*100
	print("\nMT_Vs_MT", PER,"% Completed")
MT_Vs_MT = pd.DataFrame(list(zip(MT1,MT1_BT, MT2, MT2_BT, SEQ_RATIO)), columns =['MT1','MT1_BT','MT2','MT2_BT','SEQ_RATIO']) 
#MT_Vs_MT=MT_Vs_MT.sort_values(by='SEQ_RATIO', ascending=False)
MT_Vs_MT=MT_Vs_MT.round({"SEQ_RATIO":2})
#print(MT_Vs_MT)
MT_Vs_MT.to_excel (save_path, index = None, header=True)
import datetime
now = datetime.datetime.now()
print (now.strftime("\n%H:%M Generated list of similar looking terms with MT_Vs_MT"))
