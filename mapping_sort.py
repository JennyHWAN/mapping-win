import os

# from collections import defaultdict
from datetime import datetime
import pandas as pd



def process_excel_sorted():
    files = os.listdir('.\\input')
    # files = os.listdir(filename)
    path = os.getcwd() + '\\input'
    temp = []
    # path = os.getcwd() + filename
    for filename in files:
        # if filename.endswith('.xlsx') and len(files) != 1:
        #     file = pd.read_excel(os.path.join(path, 'test.xlsx'))
        #     file1 = pd.read_excel(os.path.join(path, 'test1.xlsx'))
        if filename.endswith('.xlsx'):
            temp.append(filename)
    if len(temp) == 2:
        file = pd.read_excel(os.path.join(path, 'Table_B.xlsx'))
        file1 = pd.read_excel(os.path.join(path, 'Table_A.xlsx')) 
        
        df = file[['控制点编号','TSP mapping','控制点描述（2023）']]
        df1 = df['TSP mapping'].str.split('\n').tolist()
        df_control = file[['控制点编号','控制点描述（2023）']]
        df_control_map = {}
        for i in range(len(df_control)):
            df_control_map[df_control.iloc[i, 0]] = df_control.iloc[i, 1]
        file1 = file1[['标准编号','标准描述']].dropna()
        df_discrip = file['控制点描述（2023）']
        dic = {}
        dic_control = {}
        standard_disc = {}
        for i in range(len(file1)):
            standard_disc[file1.iloc[i, 0]] = file1.iloc[i, 1]
        for i in range(len(df1)):
            for key in df1[i]:
                if key not in dic:
                    dic[key] = [df_control.iloc[i, 0]]
                else:
                    dic[key].append(df_control.iloc[i, 0])
        res_dic = {'标准编号':[], '标准描述':[], '控制编号':[], '控制点描述（2023）':[]}
        for key in dic:
            for i in range(len(dic[key])):
                res_dic['标准编号'].append(key)
                res_dic['标准描述'].append(standard_disc[key])
                res_dic['控制编号'].append(dic[key][i])
                res_dic['控制点描述（2023）'].append(df_control_map[dic[key][i]])

        res = pd.DataFrame(res_dic)
        res = res.set_index(['标准编号','标准描述','控制编号'])
        # res.to_excel('./output/result.xlsx')
        res1 = res.sort_values(by=['标准编号'], ascending=[True])
        filename = f'result_sorted_{str(datetime.now().strftime("%Y%m%d-%H%M%S"))}.xlsx'
        res1.to_excel(f'.\\output\\{filename}') 
    # return res