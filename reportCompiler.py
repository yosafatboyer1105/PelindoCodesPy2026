import pandas as pd
import numpy as np
from openpyxl import load_workbook
import datetime

def dateFilter(df, mm1, mm2, dd1, dd2):
    clm = ['dd','mm','yyyy']
    for i,j in enumerate(clm):
        df[j] = df['Date'].apply(lambda x:x.split('/')[i]).astype(int)
    
    df1 = df.loc[(df.mm == mm1) & (df.dd > dd1)]
    df2 = df.loc[(df.mm == mm2) & (df.dd < dd2)]
    
    df = pd.concat([df1,df2]).drop(columns=clm)
    
    return df
    

def jheTranslator(ina, eng, colNameIna, colNameEng, colNum):
    try:
        for a,b in zip(colNameIna, colNameEng):
            ina.rename(columns={a:b}, inplace=True)
    except Exception as e:
        print("ERROR COLUMN RENAME PROCESS:", e)
    
    inaVal = ['Cetak', 'Transfer File', 'Transfer Email', 'Salin', 'Lembar Sistem', 'Nama File']
    engVal = ['Print', 'File Transfer', 'Email Transfer', 'Copy', 'System Sheet', 'File Name']
    
    try:
        for c,d in zip(inaVal, engVal):
            ina.replace({colNameEng[colNum]:{c:d}}, inplace=True)
    except Exception as e:
        print("ERROR JOB TYPE TRANSLATION PROCESS:", e)
        
    return ina

def paperClr(df,c):
    paper_size = [
    (df[c[24]] > 0) | (df[c[24+7]] > 0),
    (df[c[25]] > 0) | (df[c[25+7]] > 0),
    (df[c[26]] > 0) | (df[c[26+7]] > 0),
    (df[c[27]] > 0) | (df[c[27+7]] > 0),
    (df[c[28]] > 0) | (df[c[28+7]] > 0),
    (df[c[29]] > 0) | (df[c[29+7]] > 0),
    (df[c[30]] > 0) | (df[c[30+7]] > 0),
    (
        (df[c[24]] == 0)&
        (df[c[25]] == 0)&
        (df[c[26]] == 0)&
        (df[c[27]] == 0)&
        (df[c[28]] == 0)&
        (df[c[29]] == 0)&
        (df[c[30]] == 0)&
        (df[c[24+7]] == 0)&
        (df[c[25+7]] == 0)&
        (df[c[26+7]] == 0)&
        (df[c[27+7]] == 0)&
        (df[c[28+7]] == 0)&
        (df[c[29+7]] == 0)&
        (df[c[30+7]] == 0)
        )
    ]
    paper_desc = [
        'A4', 'JIS B4', 'A3', 'Letter', 'Legal', 'Ledger', 'Other', 'No Printing'
    ]

    df['Paper Size'] = np.select(paper_size, paper_desc)

    mesin = df.loc[df['Paper Size'] != 'No Printing']
    return mesin

def userDeptMapping(df, user, colUser, colData):
    users = dict(user[[colUser[0],colUser[2]]].values)
    depts = dict(user[[colUser[0],colUser[3]]].values)
    df['Nama'] = df[colData[12]].map(users)
    df['Dept'] = df[colData[12]].map(depts)
    df[colData[12]].fillna('No User Detected', inplace=True)
    return df

def filler(df, c, floor):
    df[c[12]] = df[c[12]].astype(str)
    df['Nama'] = df['Nama'].fillna(df[c[13]])
    df['Dept'] = df['Dept'].fillna(str(floor))
    df['Location'] = floor
    df[c[15]] = df[c[15]].fillna('Copy')
    df_prep = df[[
        c[12],
        'Nama',
        'Dept',
        'Location',
        c[15],
        c[7],
        'Paper Size',
        c[91],
        c[92]
    ]]
    return df_prep

def fillerTrans(compiledFx, compiledFf):
    cx = list(compiledFx.columns)
    cf = list(compiledFf.columns)
    for old, new in zip(cx, cf):
        compiledFx.rename(columns={old:new}, inplace=True)
    return compiledFx

def mergeData(d1,d2,d3=None,d4=None,d5=None,d6=None,d7=None,d8=None,d9=None,d10=None,d11=None):
    if (d3 is None) and (d4 is None) and (d5 is None) and (d6 is None) and (d7 is None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1, d2])
        
    elif (d3 is not None) and (d4 is None) and (d5 is None) and (d6 is None) and (d7 is None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is None) and (d6 is None) and (d7 is None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4])
    
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is None) and (d7 is None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is not None) and (d8 is None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6,d7])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is not None) and (d8 is not None) and (d9 is None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6,d7,d8])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is not None) and (d8 is not None) and (d9 is not None) and (d10 is None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6,d7,d8,d9])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is not None) and (d8 is not None) and (d9 is not None) and (d10 is not None) and (d11 is None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6,d7,d8,d9,d10])
        
    elif (d3 is not None) and (d4 is not None) and (d5 is not None) and (d6 is not None) and (d7 is not None) and (d8 is not None) and (d9 is not None) and (d10 is not None) and (d11 is not None):
        merged = pd.concat([d1,d2,d3,d4,d5,d6,d7,d8,d9,d10,d11])
        
    m = list(merged.columns)
    
    
    return merged