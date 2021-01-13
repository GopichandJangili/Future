import pandas as pd,docx,os,xlrd
import itertools
from tqdm import tqdm
doc = docx.Document(os.path.join(os.getcwd(),'temp','word.docx'))
fullText=[para.text for para in doc.paragraphs]
df=pd.DataFrame(fullText,columns=['paragraphs'])

tqdm.pandas()

df['flag']=df['paragraphs'].progress_apply(lambda x:True if ('<' in x and '>' in x) else False)

df=df.loc[df['flag']==True].reset_index()

###input the entire logic in a pandas apply function

xls = xlrd.open_workbook(os.path.join(os.getcwd(),'Roche','Bhas-Data_FasterFiling.xlsx'), on_demand=True)  

     
a=[pd.read_excel(os.path.join(os.getcwd(),'Roche','Bhas-Data_FasterFiling.xlsx'),sheet_name=i) for i in xls.sheet_names()]     


col_list=list(itertools.chain(*[i.columns.tolist() for i in a]))

cols=[('<'+i.split('>')[0]+'>','{'+i.split('>')[0]+'}') for i in df.loc[0,'paragraphs'].split('<') if i.split('>')[0] in col_list ]

df['newparagraphs']=df['paragraphs']

for i in cols:
 df.loc[0,'newparagraphs']=df.loc[0,'newparagraphs'].replace(i[0],i[1])



rep=dict([(i.split('}')[0],[x.loc[0,i.split('}')[0]] for x in a if i.split('}')[0] in x.columns.tolist()][0]) for i in df.loc[0,'newparagraphs'].split('{') if '}' in i])



df.loc[0,'newparagraphs']=df.loc[0,'newparagraphs'].format(**rep)


####

