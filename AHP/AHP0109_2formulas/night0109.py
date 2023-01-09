import numpy as np
import pandas as pd
from openpyxl import load_workbook

#pro_standard = [1/7,1/5,1/3,1,3,5,7] #三个分数对应7种判断矩阵结果，C32+1
pro_standard = [0.181818181818182,0.2,0.222222222222222,0.25,0.285714285714286,0.333333333333333,0.4,0.5,0.666666666666667,1,1.5,2,2.5,3,3.5,4,4.5,5,5.5
]
pro_list = [1/5,1/4,1/3,2/5,1/2,3/5,2/3,3/4,4/5,1,5/4,4/3,3/2,5/3,3,4,5]
dict_pro = dict(zip(pro_list, pro_standard))


df_1 = pd.read_excel("matrix.xlsx", sheet_name="Sheet2")
df_1 = df_1.iloc[0:5,1:5]
print(df_1)
A1 = df_1.values

df_2 = pd.read_excel("matrix.xlsx", sheet_name="Sheet1")
print(df_2)
A2 = df_2.values

df_3 = pd.read_excel("matrix.xlsx", sheet_name="Sheet3")
df_3 = df_3.iloc[1:17,2:17]
#print(df_3)

df_4 = pd.read_excel("matrix.xlsx", sheet_name="Sheet4")
df_4 = df_4.iloc[1:13,2:13]


def data_process(df):
    A = df.values
    data_matrix = A.reshape(1,-1)
    rough_data = np.array(data_matrix)
    rough_list = rough_data.tolist()
    rough_list = sum(rough_list,[])
    print(rough_list)
    rough_set = set(rough_list)
    print("去重所得集合：", rough_set)

    pro_list = list(rough_set)
    pro_list.sort()
    print("去重所得列表：", *pro_list, sep=" ")

    pro_df = df.replace(dict_pro)
    print("初始结果如下：\n", df)
    print("调整得分如下:\n", pro_df)

    A = pro_df.values
    return pro_df, A



def AHP(A) -> np.array:
#def AHP(A,weight):
    # 平均随机一致性指标。
    RI_dict = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49, 
    11:1.52, 12:1.54, 13:1.56, 14:1.58, 15:1.59795, 16:1.60459, 17:1.61526, 18:1.62132, 19:1.63019, 20:1.63743,
}
    n = len(A)
    for i in range(1, n):
        for k in range(i):
            A[i][k] = 1 / A[k][i]
    A = np.array(A)
    w, v = np.linalg.eig(A)
    lambda_max = np.max(abs(w))
    index = list(w).index(abs(w).max())
    CI = (lambda_max - n) / (n - 1)
    RI = RI_dict[n]
    CR = CI / RI

    if CR < 0.1:
        print("随机一致性指标为{}，判断矩阵具有满意的一致性。".format(CR))
        x = v[:, index].sum(axis=0)  # 对列向量求和，对于第一列求和
        y = v[:, index] / x  # 第一列进行归一化处理
        print("最大的特征值为：", lambda_max)
        weight = abs(y)
        print("对应的特征向量为：", weight.round(3))
        return weight
        
    else:
        print("随机一致性指标为{}，判断矩阵不具有满意的一致性。".format(CR))
        return None


if __name__ == '__main__':

    A3 = data_process(df_3)[1]
    pro_df_3 = data_process(df_3)[0]

    A4 = data_process(df_4)[1]
    pro_df_4 = data_process(df_4)[0]

    m1 = AHP(A1)
    m1 = pd.DataFrame(m1)

    m2 = AHP(A2)
    m2 = pd.DataFrame(m2)

    m3 = AHP(A3)
    m3 = pd.DataFrame(m3)

    m4 = AHP(A4)
    m4 = pd.DataFrame(m4)



    outputfile = pd.ExcelWriter('NIGHT_st19.xlsx')

    m1.to_excel(outputfile,sheet_name = "result1")
    m2.to_excel(outputfile,sheet_name = "result2")
    pro_df_3.to_excel(outputfile,sheet_name = "process_data_3")
    m3.to_excel(outputfile,sheet_name = "night_result3")
    pro_df_4.to_excel(outputfile,sheet_name = "process_data_4")
    m4.to_excel(outputfile,sheet_name = "night_result4")
    
    # outputfile.save() FUTURE WARNING
    outputfile.close()


