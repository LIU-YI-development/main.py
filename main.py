
import pandas as pd
import numpy as np

path = "C:\\Users\\yliu\\Work Folders\\Desktop\\mainWork_XFab\\working_folder\\2021-09-16\\source_forder\\Control Plan Monitoring WET FP 101718 5.xlsx"
path2 = "C:\\Users\\yliu\\Work Folders\\Desktop\\mainWork_XFab\\working_folder\\2021-09-16\\111.xlsx"
path3 = "C:\\Users\\yliu\\Work Folders\\Desktop\\mainWork_XFab\\working_folder\\2021-09-16\\out_20210916.xlsx"

# first step

df = pd.read_excel(path)
df.to_excel(path2, index = False)

# second step

df = pd.read_excel(path2)

oldColumnNames = []
for col in df.columns[:26]:
    oldColumnNames.append(col)

newColumnNames = []
for col in df.iloc[0][:26]:
    newColumnNames.append(col)

i = len(newColumnNames) - 1
while i > 0:
    df.rename(columns={oldColumnNames[i]:newColumnNames[i]}, inplace=True)
    i = i-1

# third step
df1 = df.iloc[:, 0:25]

df1.replace(r'\s+', np.nan, regex=True)

df2 = df1.fillna(method='ffill')
# print(df.columns)
df2[['edge exclusion','reaction plan']]=df[['edge exclusion', 'reaction plan']]
df2.iloc[0,26]='reaction plan'
df2.iloc[46:48,26]=''
df2.iloc[64:65,26]=''
df2.iloc[98:99,26]=''


df2['reaction plan'] = df2['reaction plan'].fillna(method='ffill')
df2 = df2.iloc[1: , :]

df2.to_excel(path3, index = False)

# print(c.loc[0:3,'control item'])






