import os
from datetime import datetime
import pandas as pd



def process_excel_nodes():
    files = os.listdir('.\\input')
    path = os.getcwd() + '\\input'
    temp = []
    for filename in files:
        if filename.endswith('.xlsx'):
            temp.append(filename)
    if len(temp) == 1:
        file = pd.read_excel(os.path.join(path, 'Table_B.xlsx'))

        df = file[['控制点编号','TSP mapping','控制点描述（2023）']]
        df1 = df['TSP mapping'].str.split('\n').tolist()
        df_control = file[['控制点编号','控制点描述（2023）']]
        df_control_map = {}
        for i in range(len(df_control)):
            df_control_map[df_control.iloc[i, 0]] = df_control.iloc[i, 1]
        dic = {}
        for i in range(len(df1)):
            for key in df1[i]:
                if key not in dic:
                    dic[key] = [df_control.iloc[i, 0]]
                else:
                    dic[key].append(df_control.iloc[i, 0])
        res_dic = {'标准编号':[], '控制编号':[], '控制点描述（2023）':[]}
        for key in dic:
            for i in range(len(dic[key])):
                res_dic['标准编号'].append(key)
                res_dic['控制编号'].append(dic[key][i])
                res_dic['控制点描述（2023）'].append(df_control_map[dic[key][i]])

        res = pd.DataFrame(res_dic)
        res = res.set_index(['标准编号','控制编号'])
        path = os.getcwd() + '\output'
        filename = f'result_nodes_{str(datetime.now().strftime("%Y%m%d-%H%M%S"))}.xlsx'
        res.to_excel(f'{path}\\{filename}')