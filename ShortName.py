import pandas as pd
import numpy as np

def tratarShortNumber(df,siteNameCollumn):# 'Tecnologia','Station'
  df2 = df.copy()
  #df2['SHORT'] = df2[siteNameCollumn]
  df2.loc[(df2[siteNameCollumn].str[:2] == '18'),'SHORT'] = df2[siteNameCollumn].str[4:4+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'NL'),'SHORT'] = df2[siteNameCollumn].str[2:2+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '4G'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:4] == '4GSP'),'SHORT'] = df2[siteNameCollumn].str[4:4+7]
  df2.loc[(df2[siteNameCollumn].str[:2] == '4T'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '5T'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'ST'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '5G'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '4C'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '4S'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '3S'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:3] == '4D-'),'SHORT'] = df2[siteNameCollumn].str[3:3+9]
  df2.loc[(df2[siteNameCollumn].str[:2] == '7N'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == '4R'),'SHORT'] = df2[siteNameCollumn].str[3:3+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'SL'),'SHORT'] = df2[siteNameCollumn].str[2:2+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'SI'),'SHORT'] = df2[siteNameCollumn].str[0:0+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'SP'),'SHORT'] = df2[siteNameCollumn].str[0:0+6]
  df2.loc[(df2[siteNameCollumn].str[:2] == 'SY'),'SHORT'] = df2[siteNameCollumn].str[0:0+6]
  df2.loc[(df2[siteNameCollumn].str[-3:] == '_SP'),'SHORT'] = df2[siteNameCollumn]
  df2.loc[(df2[siteNameCollumn].str[:3] == '5D-'),'SHORT'] = df2[siteNameCollumn].str[3:3+9]



  return df2

def Short2Gto3G(df,Tecnologia,Station):
  df3 = df.copy()
  df2G4G5G = df.loc[(df[Tecnologia].isin(['2G','4G','5G'])) & (df['NGNIS'].isin(['In Service'])) ]
  
  KeepList_df4G2G = [Station,'SHORT']
  locationBase_df4G2G = list(df2G4G5G.columns)
  DellListdf4G2G = list(set(locationBase_df4G2G)^set(KeepList_df4G2G))
  df2G4G5G = df2G4G5G.drop(DellListdf4G2G,1)
  df2G4G5G = df2G4G5G.drop_duplicates()
  df2G4G5G = df2G4G5G.reset_index(drop=True)
  df2G4G5G.rename(columns={'SHORT':'SHORT2'},inplace=True)


  df3 = pd.merge(df3,df2G4G5G, how='left',left_on=[Station],right_on=[Station])
  df3.loc[(df3[Tecnologia] == '3G') & (~df3['SHORT2'].isna()),['SHORT']] = df3['SHORT2']
  df3 = df3.drop(['SHORT2'],1)





  

  '''
  df3 = df.copy()
  df3 = df3.reset_index(drop=True)
  df4G2G = df3.loc[(df3[Tecnologia] == '2G') | (df3[Tecnologia] == '4G')]

  KeepList_df4G2G = [Station,'SHORT']
  locationBase_df4G2G = list(df4G2G.columns)
  DellListdf4G2G = list(set(locationBase_df4G2G)^set(KeepList_df4G2G))
  df4G2G = df4G2G.drop(DellListdf4G2G,1)
  df4G2G = df4G2G.drop_duplicates()
  df4G2G = df4G2G.reset_index(drop=True)

  indexI = 0
  for i in df3[Station]:
      indexJ = 0
      for j in df4G2G[Station]:
          if i == j:
              df3.at[indexI,'SHORT'] = df4G2G.at[indexJ,'SHORT']
          indexJ += 1
      indexI += 1    
  '''
  return df3
   

def tratarShortNumber2(df,siteNameCollumn,Tecnologia,Station):
  df2 = df.copy()
  df2 = tratarShortNumber(df,siteNameCollumn) 

  df2.loc[(df2[Tecnologia] == '2G'),['SHORT']] = df2[siteNameCollumn]
  #todo incluir shot 3G puro
  df2.loc[(df2[Tecnologia] == '3G'),['SHORT']] = df2[siteNameCollumn]

  df2 = Short2Gto3G(df2,Tecnologia,Station)
    



  return df2
