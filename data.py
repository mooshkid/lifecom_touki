import pandas as pd
import glob, os

# list of all excel files in directory 
os.chdir(r"\\192.168.1.194\lifecom_share\011_営業\01_本社\06_権利MD送付\スクリプト\02_excel")
excelFiles = glob.glob("*.xlsx")
# print(excelFiles)

# empty list
dataAll = []

for i in excelFiles:
    excelPath = i

    # create dataframe
    df = pd.read_excel(excelPath)
    # print(df)

    # cell data
    title = df.values[2][0]
    loc = df.values[6][1]
    name = df.values[6][2]

    # remove extra text
    title = str(title).replace('    所有者一覧表 （土地)','')

    # new data frame
    data = [i, title, loc, name]
    # append to dataAll dataframe
    dataAll.append(data)

# print(dataAll)

# convert the final dataframe to csv
df = pd.DataFrame(dataAll)
df.to_csv('../03_data/data.csv', header=['file', '設置住所', '所有者住所', '名前'], encoding='utf-8-sig', index=False)
print(df)