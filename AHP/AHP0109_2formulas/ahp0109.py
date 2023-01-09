import numpy as np
import pandas as pd
from openpyxl import load_workbook


df_1 = pd.read_excel("matrix.xlsx", sheet_name="Sheet2")
df_1 = df_1.iloc[0:5,1:5]
print(df_1)
A1 = df_1.values

df_2 = pd.read_excel("matrix.xlsx", sheet_name="Sheet1")
print(df_2)
A2 = df_2.values

'''
#展平嵌套列表
def flatten(L):
    return sum(([x] if not isinstance(x, list) else flatten(x) for x in L), [])
'''
pro_standard = [1/7,1/5,1/3,1,3,5,7] #三个分数对应7种判断矩阵结果，C32+1


df_3 = pd.read_excel("matrix.xlsx", sheet_name="Sheet3")
df_3 = df_3.iloc[1:17,2:17]
#print(df_3)
A3 = df_3.values
data_3 = A3.reshape(1,-1)
rough_3 = np.array(data_3)
rough_3_list = rough_3.tolist()
rough_3_list = sum(rough_3_list,[]) #展平嵌套列表:SUM(嵌套list,[])
print(rough_3_list)
rough_3_set = set(rough_3_list)
print("去重所得集合：", rough_3_set)

pro_3_list = list(rough_3_set)
pro_3_list.sort()
print("去重所得列表：", pro_3_list)

dict_3 = dict(zip(pro_3_list,pro_standard))
print("df3对应字典：", dict_3)

pro_df_3 = df_3.replace(dict_3)
print("初始结果如下：\n", df_3)
print("调整得分如下:\n", pro_df_3)
A3 = pro_df_3.values

def rough_process():
    


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

    m1 = AHP(A1)
    m1 = pd.DataFrame(m1)

    m2 = AHP(A2)
    m2 = pd.DataFrame(m2)

    m3 = AHP(A3)
    m3 = pd.DataFrame(m3)

    outputfile = pd.ExcelWriter('OUTPUT0109.xlsx')

    m1.to_excel(outputfile,sheet_name = "result1")
    m2.to_excel(outputfile,sheet_name = "result2")
    pro_df_3.to_excel(outputfile,sheet_name = "process_data_3")
    m3.to_excel(outputfile,sheet_name = "result3")
    # outputfile.save() FUTURE WARNING
    outputfile.close()


