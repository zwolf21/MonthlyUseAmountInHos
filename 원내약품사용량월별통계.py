
# coding: utf-8

# In[1]:

import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import numpy as np
from pandas import DataFrame, Series
import os, re


# In[2]:

OUTPUT_EXCEL = '월별원내약품사용현황.xlsx'


# In[3]:

# 데이타셋 준비
data_source_dir = '사용량월별통계/원내'
dfs = []
for fname in os.listdir(data_source_dir):
    fn, ext = os.path.splitext(fname)
    if ext in ['.xls', '.xlsx']:
        df = pd.read_excel(os.path.join(data_source_dir, fname))
        df['사용(개시)년월'] = fn
        dfs.append(df)
use_amount_df = pd.concat(dfs, ignore_index=True)


# In[4]:

drug_standard_df = pd.read_json('drug.json').T

drug_info_df = pd.read_excel('약품정보.xls')

use_amount_df = pd.merge(drug_info_df, use_amount_df[['사용량', '약품코드', '사용(개시)년월']], on='약품코드', how='left')

use_amount_df = pd.merge(use_amount_df, drug_standard_df[['보험코드', '제품명', '판매사', '성분/함량']], left_on='EDI코드', right_on='보험코드', how='left')

use_amount_df['제품명'] = use_amount_df['제품명'].fillna(use_amount_df['약품명(한글)'])

use_amount_df['사용개시년월'] = use_amount_df['수가시작일자'].map(lambda x: str(x)[0:4]+'-'+str(x)[4:6])

use_amount_df['사용(개시)년월'] = use_amount_df['사용(개시)년월'].fillna(use_amount_df['사용개시년월'])
use_amount_df['성분명'] = use_amount_df['성분명'].fillna(use_amount_df['성분/함량'])

use_amount_df['원내/원외 처방구분'] = use_amount_df['원내/원외 처방구분'].map({1: '원외', 2: '원외/원내', 3: '원내'})
use_amount_df['약품법적구분'] = use_amount_df['약품법적구분'].map({0: '일반', 1: '마약', 2: '향정약', 3: '독약', 4: '한방약', 5: '고가약'})


# In[5]:

def get_last(s):
    try:
        return max(s)
    except:
        return s


# In[6]:

months = use_amount_df['사용(개시)년월'].unique()
months = sorted(months.tolist(), reverse=1)
use_amount_df['최후사용월'] = use_amount_df.groupby(['제품명'])['사용(개시)년월'].transform(get_last)
use_amount_df['최근미사용월수'] = use_amount_df['최후사용월'].map(lambda x: months.index(x) if x in months else -1)


# In[7]:

use_amount_in_df = use_amount_df[use_amount_df['원내/원외 처방구분'] != '원외']


# In[8]:

use_amount_in_df['사용량'] = use_amount_in_df['사용량'].fillna('오픈후미사용')


# In[9]:

pat = '(\(([^\d].*?)\)+\s*)|퇴장방지\s*|생산원가보전,*\s*|사용장려(비\s*\d+원|및|비용지급,*\s*)'
use_amount_in_df = use_amount_in_df.rename(columns={'제품명': '약품명(드럭인포)', '약품명(한글)': '약품명(원내)'})
use_amount_in_df['약품명(드럭인포)'] = use_amount_in_df['약품명(드럭인포)'].str.replace(pat, '')


# In[10]:

pvt = use_amount_in_df.pivot_table(index = ['EDI코드','약품명(드럭인포)', '성분명','약품코드','약품명(원내)','효능코드명','규격단위', '최근미사용월수'], columns=['사용(개시)년월'], values=['사용량'], aggfunc=sum)


# In[11]:

pvt.to_excel(OUTPUT_EXCEL)
os.startfile(OUTPUT_EXCEL)


# In[ ]:



