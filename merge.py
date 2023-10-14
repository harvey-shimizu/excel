import pandas as pd
import sys
import json
from tabulate import tabulate

def excel_to_dict(excel_file):
    # Read the Excel file
    df = pd.read_excel(excel_file, engine='openpyxl')

    # Convert the dataframe to dict format
    dict_data = df.to_dict(orient='records')
    return dict_data

def read_japanese_csv(file_path, encoding="utf-8"):
    try:
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path, encoding=encoding, header=[1])
        return df.to_dict(orient='records')
    except UnicodeDecodeError:
        print(f"Unable to read the file with {encoding} encoding. Please specify the correct encoding.")
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python master_data tData.csv")
        sys.exit(1)

    master = excel_to_dict(sys.argv[1])
    tdata = read_japanese_csv(sys.argv[2])
    if master is not None and tdata is not None:
        #for m,t in zip(master, tdata):
            #print(m['氏名(戸籍)'], ', ', t['児童生徒氏名'])
        count = 1
        #print("連番","\t,",'管理番号','\t,','学年',"\t,",'組',"\t,",'出席番号','\t,','学級名','\t,','master','\t,','tData' )
        data = []
        match = 0
        for m in master:
            mname = m['氏名(戸籍)']
            match = 0
            for t in tdata:
                tname = t['児童生徒氏名']
                if(mname == tname):
                    #print(count,"\t,",t['管理番号'],'\t,',m['学年'],"\t,",m['組'],"\t,",m['出席番号'],'\t,',t['学級名'],'\t,',mname,'\t,',tname )
                    data.append([str(count),t['管理番号'],m['学年'],m['組'],m['出席番号'],t['学級名'],mname,tname,m['個人識別コード']])
                    count += 1
                    match = 1
                    m['match'] = 1
                    break
            if(match == 0):
                m['match'] = 0

        print(tabulate(data, headers=['連番','管理番号(t)','学年(m)','組(m)','出席番号(m)','学級名(m)','master(m)','tData(t)','個人識別コード'], stralign='center', numalign='center'))
        #print(tabulate(data, headers='keys'))

        print('masterに登録されているが、tDataの名前とマッチでしなかった生徒')
        data = []
        count = 1
        for m in master:
            if(m['match'] == 0):
                data.append([str(count),None,m['学年'],m['組'],m['出席番号'],None,m['氏名(戸籍)'],None])
                count += 1
        print(tabulate(data, headers=['連番','管理番号(t)','学年(m)','組(m)','出席番号(m)','学級名(m)','master(m)','tData(t)'], stralign='center', numalign='center'))

        print('tDataに登録されているが、masterの名前とマッチでしなかった生徒')
        data = []
        match = 0
        count = 1
        for t in tdata:
            tname = t['児童生徒氏名']
            match = 0
            for m in master:
                mname = m['氏名(戸籍)']
                if(mname == tname):
                    #data.append([str(count),t['管理番号'],m['学年'],m['組'],m['出席番号'],t['学級名'],mname,tname])
                    match = 1
                    t['match'] = 1
                    break
            if(match == 0):
                t['match'] = 0

        data = []
        count = 1
        for t in tdata:
            if(t['match'] == 0):
                data.append([str(count),t['管理番号'],t['学年'],t['組'],t['出席番号'],t['学級名'],None,t['児童生徒氏名']])
                count += 1
        print(tabulate(data, headers=['連番','管理番号(t)','学年(m)','組(m)','出席番号(m)','学級名(m)','master(m)','tData(t)'], stralign='center', numalign='center'))

