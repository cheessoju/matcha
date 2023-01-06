import numpy as np
import pandas as pd
df = pd.read_excel("matrix0106.xlsx", sheet_name="Sheet2")
df = df.iloc[0:5,1:5]
print(df)
A = df.values

'''
A = np.array([[1, 1/2, 1/3,1/4],
   [2, 1, 1/3,1/4],
   [3,3,1, 1/2],
   [4,4,2,1]
   ])
'''
def AHP(A) -> np.array:
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
        print("对应的特征向量为：", abs(y).round(3))
        return abs(y)
    else:
        print("随机一致性指标为{}，判断矩阵不具有满意的一致性。".format(CR))
        return None


if __name__ == '__main__':
    m = AHP(A)
    print("congrats!")