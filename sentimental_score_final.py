#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[1]:


import snownlp
import numpy as np
import pandas as pd
import jieba
import re
from snownlp import SnowNLP


#######read LM and NTUSD dict#########

LM_P=pd.read_excel('LM词典+NTUSD词典.xlsx',header=None,index_col=0,sheet_name='LM_Positive（CNRDS整理）')
LM_P=LM_P.reset_index().to_dict()
LM_P={value:key for key, value in LM_P[0].items()}
LM_P_key= {v : k for k, v in LM_P.items()}
LM_N=pd.read_excel('LM词典+NTUSD词典.xlsx',index_col=0,header=None,sheet_name='LM_Negative（CNRDS整理）')
LM_N=LM_N.reset_index().to_dict()
LM_N={value:key for key, value in LM_N[0].items()}
LM_N_key= {v : k for k, v in LM_N.items()}
NTUSD_P=pd.read_excel('LM词典+NTUSD词典.xlsx',index_col=0,header=None,sheet_name='NTUSD_Positive')
NTUSD_P=NTUSD_P.reset_index().to_dict()
NTUSD_P={value:key for key, value in NTUSD_P[0].items()}
NTUSD_P_key= {v : k for k, v in NTUSD_P.items()}
NTUSD_N=pd.read_excel('LM词典+NTUSD词典.xlsx',index_col=0,header=None,sheet_name='NTUSD_Negative')
NTUSD_N=NTUSD_N.reset_index().to_dict()
NTUSD_N={value:key for key, value in NTUSD_N[0].items()}
NTUSD_N_key= {v : k for k, v in NTUSD_N.items()}


# In[2]:


######## read different excel ##########

####### enter excel name here ##########
#excel='MDA_br'
excel='report_br'

####### enter the volume of the dataset here (percentage) #########
percentage=95
#percentage=99

####### use the whole sentence or several word around '一带一路' ##########
model='whole'
#model='substring'

######enter forward and backward words around '一带一路'##############
######if model is whole please still keep number there, it won't affect the result ########
number=10
######enter top k hot word you want #########
k=80


# In[ ]:


######## code start from the last block #########
       
        
        


# In[3]:


######To find all substring in one string#########

def find_all(str1,key):
    lstKey = [] #定义空列表用于存储多个指定字符的索引
    lengthKey = 0
    #字符串中存在指定字符串的个数
    countStr = str1.count(key)

    #利用获取的countStr进行判断
    if countStr < 1:
        return -1
    elif countStr == 1: #当字符串中只有一个指定字符时，直接通过find()方法即可解决
        indexKey = str1.find(key)
        return indexKey
    else: #当字符串中存在多个指定字符的处理方法
        #第一个指定字符的处理方法
        indexKey = str1.find(key)
        lstKey.append(indexKey) #将第一个索引加入到lstKey列表中
        #其余指定字符的处理方法
        while countStr > 1:
            #将前一个指定字符之后的字符串截取下来
            str_new = str1[indexKey+1:len(str1)+1]
            #获取截取后的字符串中前一个指定字符的索引值
            indexKey_new = str_new.find(key)
            #后一个索引值=前一个索引值+1+indexkey_new
            indexKey = indexKey+1 +indexKey_new
            #将后发现的索引值加入lstKey列表中
            lstKey.append(indexKey)
            countStr -= 1
        #print('查找的关键字的索引为',lstKey)
        return lstKey


# In[4]:





def find_yidaiyilu(string1):
    if type(string1)==float:
        return 'None'
    if string1[-1]!='。' or '，':
        string1=string1+'。'
        
    
    pattern=re.compile('[，。、,.;:][^，。]*一带一路[^，。、,.:;]*[，。,.:;]')
    search = pattern.search(string1)
    if search!=None:
        print(search.group(0))
        return search.group(0)
    else:
        return 'None'
    
######### if want to use snownlp for sentimental score #########   
def sent_find(string1):
    if type(string1)==float:
        return 'None'
    
    t = SnowNLP(string1)
    list1=[]
    for sen in t.sentences:
        list1.append(sen+'。')
   
    for i in list1:
        if type(i)==float:
            continue
        if i.find('一带一路')!=-1:
            return (i,len(i))
        
        
    return 'None'

########### find a few words before or after the substring ###########
def cut_ydyl(string1,num):
    if type(string1)==float:
        return ['None']
    lst=find_all(string1,'一带一路')
    if type(lst)==int:
        if lst==-1:
            return ['None']
        else:
            if lst-num<0:
                return [string1[:lst+num]]
            else:
                return [string1[lst-num:lst+num]]
    list1=[]
    n=0
    if type(lst)==list:
        for n in range(0,len(lst)):
            if n==len(lst)-1:
                continue
            elif lst[n]+num>=lst[n+1]:
                    lst[n+1]=-1
        lst=[n for n in lst if n>=0]
        
        for n in range(0,len(lst)):
                
            if lst[n]-num<0:
                list1.append(string1[:lst[n]+num+len('一带一路')])
            else:
                list1.append(string1[lst[n]-num:lst[n]+num+len('一带一路')])
                
    return list1
 

            
    
    
def get_key (dict, value):
    return [k for k, v in dict.items() if v == value]         
def search(dict, value):
    return dict.get(value)
def search_pos_neg(value):
    
    if search(LM_P,value)!=None or search(NTUSD_P,value)!=None:
        return 1
    if search(NTUSD_N,value)!=None or search(LM_N,value)!=None:
        return 0

                
        
####### calculate sentimental score LM_TONE NTUSD_TONE##########    
def sentimental(i,tone):
    list1=i.split('/')
    pscore=0
    nscore=0
    total=len(list1)
    list2=[]
    list3=[]
    list_p=[]
    list_n=[]
    if tone=='LM_TONE1':

        for a in list1:
            if search(LM_P,a)!=None:
                pscore+=1
                list2.append(a)
                list_p.append((a,1))
            if search(LM_N,a)!=None:
                nscore += 1
                list2.append(a)
                list_n.append((a,1))
        if pscore + nscore == 0:
            LM_TONE1=0
            return (pscore,nscore,LM_TONE1,list2,list_p,list_n)
        else:
            LM_TONE1=(pscore-nscore)/total
            return (pscore,nscore,LM_TONE1,list2,list_p,list_n)
    if tone=="LM_TONE2":
        for a in list1:
            if search(LM_P,a)!=None:
                pscore+=1
            if search(LM_N,a)!=None:
                nscore += 1
        if pscore + nscore == 0:
            LM_TONE2=0
            return (pscore,nscore,LM_TONE2)
        else:
            LM_TONE2=(pscore-nscore)/(pscore+nscore)
            return (pscore,nscore,LM_TONE2)
    if tone=='NTUSD_TONE':
        for a in list1:
            if search(NTUSD_P,a)!=None:
                pscore+=1
                list3.append(a)
                list_p.append((a,1))
            if search(NTUSD_N,a)!=None:
                nscore += 1
                list3.append(a)
                list_n.append((a,1))
        if pscore+nscore==0:
            NTUSD_TONE=0
            return (pscore,nscore,NTUSD_TONE,list3,list_p,list_n)
        else:
            NTUSD_TONE=(pscore-nscore)/(pscore+nscore)
            return (pscore,nscore,NTUSD_TONE,list3,list_p,list_n)
def seperate(tuple1,n):
    print(type(tuple1))
    x=tuple1[n]
    return x

def token(string):
    text = jieba.cut(string.strip())
    return '/'.join(text)


# In[5]:


#######if use find 10 words forward and 10 words backward########
def substring_model(df,num):
    list1=[]

    df['string']=df['内容'].apply(lambda x:cut_ydyl(x,num))
    df=df.drop(columns=['内容'])


    #df=df.reindex(df.index.repeat(df['string'].str.len()))
    #df=df.assign(B=np.concatenate(df.string.values))
    df=df.reindex(df.index.repeat(df['string'].str.len())).assign(string=np.concatenate(df.string.values))


    ######if use sent_find#########
    for x in df['string']:
        if x is not str:
            x=str(x)
            text = token(x.strip())
            list1.append(text)
    df_cut=pd.DataFrame(list1,columns=['cut'])
    df=pd.concat([df.reset_index(),df_cut],axis=1)
    df['len']=df['cut'].apply(lambda x:len(x))
    df=df.drop(columns=['index'])
    return df
#######if use whole sentence############
def whole_model(df):
    list1=[]
    df['string']=df['内容']
   

    for x in df['string']:
        if x is not str:
            x=str(x)
            text = token(x.strip())
            list1.append(text)
    df_cut=pd.DataFrame(list1,columns=['cut'])
    df=pd.concat([df.reset_index(),df_cut],axis=1)
    df['len']=df['cut'].apply(lambda x:len(x))
    df=df.drop(columns=['index'])
    return df
    


# In[6]:


######## sort to find the top n hot word ##########

import collections

from string import digits

def counter(text,n):
    
   
    
    while '' in text:
        text.remove('')
    freq=collections.Counter(text)
    frequency={k: v for k, v in freq.items() if not k.isdigit()}
    if len(frequency)>=n:
        
        return sorted(frequency.items(),key=lambda kv:(kv[1],kv[0]),reverse=True)[:n]
    else:
        return sorted(frequency.items(),key=lambda kv:(kv[1],kv[0]),reverse=True)


# In[7]:


def append_zero(list1,num):
    while len(list1)<num:
        list1.append(0)
    return list1

########calculate by firm year sentimental score ###########
def by_firm_year(df_fy):
    

    df_fy['NTUSD_TONE']=df_fy['cut'].apply(lambda x:sentimental(x,'NTUSD_TONE'))
    df_fy['Number of Positive terms NTUSD']=df_fy['NTUSD_TONE'].apply(lambda x:x[0])
    df_fy['Number of Negative terms NTUSD']=df_fy['NTUSD_TONE'].apply(lambda x:x[1])
    df_fy['NTUSD_total']=df_fy['NTUSD_TONE'].apply(lambda x:x[3])
    df_fy['NTUSD_positive']=df_fy['NTUSD_TONE'].apply(lambda x:x[4])
    df_fy['NTUSD_negative']=df_fy['NTUSD_TONE'].apply(lambda x:x[5])
    df_fy['NTUSD_TONE']=df_fy['NTUSD_TONE'].apply(lambda x:x[2])


    df_fy['LM_TONE1']=df_fy['cut'].apply(lambda x:sentimental(x,'LM_TONE1'))
    df_fy['LM_TONE2']=df_fy['cut'].apply(lambda x:sentimental(x,'LM_TONE2'))
    df_fy['Number of Positive terms LM_TONE']=df_fy['LM_TONE1'].apply(lambda x:x[0])
    df_fy['Number of Negative terms LM_TONE']=df_fy['LM_TONE1'].apply(lambda x:x[1])
    df_fy['LM_total']=df_fy['LM_TONE1'].apply(lambda x:x[3])
    df_fy['LM_positive']=df_fy['LM_TONE1'].apply(lambda x:x[4])
    df_fy['LM_negative']=df_fy['LM_TONE1'].apply(lambda x:x[5])
    df_fy['LM_TONE1']=df_fy['LM_TONE1'].apply(lambda x:x[2])
    df_fy['LM_TONE2']=df_fy['LM_TONE2'].apply(lambda x:x[2])
    
    df_fy=df_fy.drop(columns=['cut'])
    return df_fy

def dumlist(i):
    str=''
    for x in i:
        str+=x[0]
        #str+='/''
    return str

####### out put negative words and the orginal sentence #######
def original_sentence_negative_words(df,excel,model):
    df_n=by_firm_year(df)
    df_n['NTUSD_empty']=df_n["NTUSD_negative"].apply(lambda x: len(x)!=0)
    df_n['LM_empty']=df_n['LM_negative'].apply(lambda x:len(x)!=0)
    df_ntusd=df_n[df_n['NTUSD_empty']]

    df_l=pd.DataFrame(df_ntusd['NTUSD_negative'].tolist(), index=df_ntusd.index)
    df_l=df_l[0].apply(lambda x :x[0])
    df_ntusd=pd.concat([df_ntusd,df_l],axis=1)
    df_ntusd=df_ntusd.groupby([0]).agg(lambda x:'#'.join(x))['string']
   
    df_lm=df_n[df_n['LM_empty']]
    df_lm=pd.DataFrame(df_lm['LM_negative'].tolist(), index=df_lm.index)
    df_lm=df_lm[0].apply(lambda x :x[0])
    df_n=pd.concat([df_n,df_lm],axis=1)
    df_n=df_n.groupby([0]).agg(lambda x:'#'.join(x))['string']
    
    if model=='substring':
        df_ntusd=df_ntusd.to_frame().reset_index().to_excel(excel+'_negative_orgin_ntusd.xls')
        df_lm=df_n.to_frame().reset_index().to_excel(excel+'_negative_origin_lm.xls')
    if model=='whole':
        df_ntusd=df_ntusd.to_frame().reset_index().to_csv(excel+'_negative_orgin_ntusd.csv')
        df_lm=df_n.to_frame().reset_index().to_csv(excel+'_negative_origin_lm.csv')
def original_sentence_positive_words(df,excel,model):
    df_n=by_firm_year(df)
    df_n['NTUSD_empty']=df_n["NTUSD_positive"].apply(lambda x: len(x)!=0)
    df_n['LM_empty']=df_n['LM_positive'].apply(lambda x:len(x)!=0)
    df_ntusd=df_n[df_n['NTUSD_empty']]

    df_l=pd.DataFrame(df_ntusd['NTUSD_positive'].tolist(), index=df_ntusd.index)
    df_l=df_l[0].apply(lambda x :x[0])
    df_ntusd=pd.concat([df_ntusd,df_l],axis=1)
    df_ntusd=df_ntusd.groupby([0]).agg(lambda x:'#'.join(x))['string']
   
    df_lm=df_n[df_n['LM_empty']]
    df_lm=pd.DataFrame(df_lm['LM_positive'].tolist(), index=df_lm.index)
    df_lm=df_lm[0].apply(lambda x :x[0])
    df_n=pd.concat([df_n,df_lm],axis=1)
    df_n=df_n.groupby([0]).agg(lambda x:'#'.join(x))['string']
    
    if model=='substring':
        df_ntusd=df_ntusd.to_frame().reset_index().to_excel(excel+'_positive_orgin_ntusd.xls')
        df_lm=df_n.to_frame().reset_index().to_excel(excel+'_positive_origin_lm.xls')
    if model=='whole':
        df_ntusd=df_ntusd.to_frame().reset_index().to_csv(excel+'_positive_orgin_ntusd.csv')
        df_lm=df_n.to_frame().reset_index().to_csv(excel+'_positive_origin_lm.csv')
    
######## out put sentimental score 'report' && 'describe' && 'top n hot word' for excel ########
def quantile_fy(num,df,excel,k):
    df_fy_95=df[df['len']<=df['len'].quantile(num/100)]
    df_fy_95=by_firm_year(df_fy_95)
    df_fy_95.to_excel(excel+str(num)+'_report'+'.xls')
    df_fy_95.describe().to_excel(excel+str(num)+'_report'+'_describe.xls')
    df_all=df_fy_95['NTUSD_total'].sum()
    cut=counter(df_all,k)
    NTUSD_total=df_fy_95['Number of Positive terms NTUSD'].sum()+df_fy_95['Number of Negative terms NTUSD'].sum()
    LM_total=df_fy_95['Number of Positive terms LM_TONE'].sum()+df_fy_95['Number of Negative terms LM_TONE'].sum()
    
    
    #cut_df=pd.DataFrame(columns=['NTUSD_Negative_word','NTUSD_times','LM_Negative_word','LM_times'])
    cut_df=pd.DataFrame()
    cut_df['NTUSD_word']=pd.Series(append_zero([x[0] for x in cut],k))
    cut_df['NTUSD_negative/postive']=cut_df['NTUSD_word'].apply(lambda x: search_pos_neg(x))
    cut_df['NTUSD_times']=pd.Series(append_zero([x[1] for x in cut],k))
    cut_df['NTUSD_percentage']=cut_df['NTUSD_times']/NTUSD_total
    df_all=df_fy_95['LM_total'].sum()
    cut=counter(df_all,k)
    cut_df['LM_word']=pd.Series(append_zero([x[0] for x in cut],k))
    cut_df['LM_negative/postive']=cut_df['LM_word'].apply(lambda x: search_pos_neg(x))
    cut_df['LM_times']=pd.Series(append_zero([x[1] for x in cut],k))
    cut_df['LM_percentage']=cut_df['LM_times']/LM_total
    
    
    cut_df.to_excel(excel+str(num)+'_report'+'_top'+str(k)+'.xls')
    return(df_fy_95,df_fy_95.describe(),cut_df)


# In[8]:


def sentimental_score(excel,percentage,model,number,k):
    df=pd.read_excel(excel+'.xlsx',index_col=0,decode='unicode').dropna().reset_index()
    if model=='substring':
        df=substring_model(df,number)
        original_sentence_negative_words(df,excel,model)
        original_sentence_positive_words(df,excel,model)
        quantile_fy(percentage,df,excel,k)
    if model=='whole':
        df=whole_model(df)
        original_sentence_negative_words(df,excel,model)
        original_sentence_positive_words(df,excel,model)
        quantile_fy(percentage,df,excel,k)

sentimental_score(excel,percentage,model,number,k)

