"""
this is a code for fuzzy DANP method coded by Pouria Zaukeri
it should be noted that this is coded for a certain use case and a particular research
many adjustments should be made in order for this template to work for another case
any unauthorised usage is condemned
contact pouria_zaukeri@mail.com for more information
"""

import pandas as pd
import numpy as np

# creating data frame
df1 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 1')
df2 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 2')
df3 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 3')
df4 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 4')
df5 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 5')
df6 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 6')
df7 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 7')
df8 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 8')
df9 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 9')
df10 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 10')
df11 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 11')
df12 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 12')
df13 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 13')
df14 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 14')
df15 = pd.read_excel(r'C:/Users/PotyZed/Desktop/experts.xlsx', sheet_name='expert 15')

# defining a list and counting the rows in data frame to make a list of every row
mylist = []
mylist1 = []
mylist2 = []
mylist3 = []
mylist4 = []
mylist5 = []
mylist6 = []
mylist7 = []
mylist8 = []
mylist9 = []
mylist10 = []
mylist11 = []
mylist12 = []
mylist13 = []
mylist14 = []
mylist15 = []
mylist16 = []

# making a list of each sheet and average of them
# first sheet
count_row = df1.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist1 = df1.iloc[i].tolist()
    del rowlist1[0]
    # fuzzification
    rowlist1 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist1]
    mylist1.append(rowlist1)
# second sheet
count_row = df2.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist2 = df2.iloc[i].tolist()
    del rowlist2[0]
    # fuzzification
    rowlist2 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist2]
    mylist2.append(rowlist2)
# third sheet
count_row = df3.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist3 = df3.iloc[i].tolist()
    del rowlist3[0]
    # fuzzification
    rowlist3 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist3]
    mylist3.append(rowlist3)
# forth sheet
count_row = df4.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist4 = df4.iloc[i].tolist()
    del rowlist4[0]
    # fuzzification
    rowlist4 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist4]
    mylist4.append(rowlist4)
# fifth sheet
count_row = df5.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist5 = df5.iloc[i].tolist()
    del rowlist5[0]
    # fuzzification
    rowlist5 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist5]
    mylist5.append(rowlist5)
# sixth sheet
count_row = df6.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist6 = df6.iloc[i].tolist()
    del rowlist6[0]
    # fuzzification
    rowlist6 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist6]
    mylist6.append(rowlist6)
# seventh sheet
count_row = df7.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist7 = df7.iloc[i].tolist()
    del rowlist7[0]
    # fuzzification
    rowlist7 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist7]
    mylist7.append(rowlist7)
# eighth sheet
count_row = df8.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist8 = df8.iloc[i].tolist()
    del rowlist8[0]
    # fuzzification
    rowlist8 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist8]
    mylist8.append(rowlist8)
# ninth sheet
count_row = df9.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist9 = df9.iloc[i].tolist()
    del rowlist9[0]
    # fuzzification
    rowlist9 = [[0, 0, 0.25] if i == 0 else
                [0, 0.25, 0.5] if i == 1 else
                [0.25, 0.5, 0.75] if i == 2 else
                [0.5, 0.75, 1] if i == 3 else
                [0.75, 1, 1] if i == 4 else
                i for i in rowlist9]
    mylist9.append(rowlist9)
# tenth sheet
count_row = df10.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist10 = df10.iloc[i].tolist()
    del rowlist10[0]
    # fuzzification
    rowlist10 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist10]
    mylist10.append(rowlist10)
# eleventh sheet
count_row = df11.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist11 = df11.iloc[i].tolist()
    del rowlist11[0]
    # fuzzification
    rowlist11 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist11]
    mylist11.append(rowlist11)
# twelfth sheet
count_row = df12.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist12 = df12.iloc[i].tolist()
    del rowlist12[0]
    # fuzzification
    rowlist12 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist12]
    mylist12.append(rowlist12)
# thirteenth sheet
count_row = df13.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist13 = df13.iloc[i].tolist()
    del rowlist13[0]
    # fuzzification
    rowlist13 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist13]
    mylist13.append(rowlist13)
# fourteenth sheet
count_row = df14.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist14 = df14.iloc[i].tolist()
    del rowlist14[0]
    # fuzzification
    rowlist14 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist14]
    mylist14.append(rowlist14)
# fifteenth sheet
count_row = df15.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist15 = df15.iloc[i].tolist()
    del rowlist15[0]
    # fuzzification
    rowlist15 = [[0, 0, 0.25] if i == 0 else
                 [0, 0.25, 0.5] if i == 1 else
                 [0.25, 0.5, 0.75] if i == 2 else
                 [0.5, 0.75, 1] if i == 3 else
                 [0.75, 1, 1] if i == 4 else
                 i for i in rowlist15]
    mylist15.append(rowlist15)

# sum of these sheets and their average, i&j for enough steps(list in lists)
for i in range(0, len(mylist1)):
    # to erase previous row calculation
    mylist16 = []
    mylist.append(mylist16)
    for j in range(0, len(mylist1)):
        summation = [sum(z) / 15 for z in zip(mylist1[i][j], mylist2[i][j], mylist3[i][j],
                                            mylist4[i][j], mylist5[i][j], mylist6[i][j],
                                            mylist7[i][j], mylist8[i][j], mylist9[i][j],
                                            mylist10[i][j], mylist11[i][j], mylist12[i][j],
                                            mylist13[i][j], mylist14[i][j], mylist15[i][j])
                    ]
        mylist16.append(summation)

# fuzzy dematel first step creating arrays for l,m,u
X_a = np.zeros((len(mylist), len(mylist)))
X_b = np.zeros((len(mylist), len(mylist)))
X_c = np.zeros((len(mylist), len(mylist)))
# replacing with actual value
for i in range(0, len(mylist)):
    for j in range(0, len(mylist)):
        a, b, c = mylist[i][j]
        X_a[i, j] = a
        X_b[i, j] = b
        X_c[i, j] = c
# normalizing, axis 0 for column sum, 1 for row sum
X_a = X_a / np.max(np.sum(X_a, axis=1))
X_b = X_b / np.max(np.sum(X_b, axis=1))
X_c = X_c / np.max(np.sum(X_c, axis=1))
# calculating identity matrix minus x and inverting it
Y_a = np.linalg.inv(np.identity(len(mylist)) - X_a)
Y_b = np.linalg.inv(np.identity(len(mylist)) - X_b)
Y_c = np.linalg.inv(np.identity(len(mylist)) - X_c)
# calculating total relation matrix
T_a = np.matmul(X_a, Y_a)
T_b = np.matmul(X_b, Y_b)
T_c = np.matmul(X_c, Y_c)
# calculating R, D & ...
D_a = np.sum(T_a, axis=1)
D_b = np.sum(T_b, axis=1)
D_c = np.sum(T_c, axis=1)
R_a = np.sum(T_a, axis=0)
R_b = np.sum(T_b, axis=0)
R_c = np.sum(T_c, axis=0)
defuzzied_D = (((D_c - D_a) + (D_b - D_a)) / 3) + D_a
defuzzied_R = (((R_c - R_a) + (R_b - R_a)) / 3) + R_a
D_plus_R = defuzzied_D + defuzzied_R
D_minus_R = defuzzied_D - defuzzied_R

# showing the dematel Tc final matrix
supermatrix = []
supermatrix2 = []

for i in range(0, len(mylist)):
    supermatrix2 = []
    supermatrix.append(supermatrix2)
    for j in range(0, len(mylist)):
        supermatrix2.append((((T_c[i, j] - T_a[i, j]) + (T_b[i, j] - T_a[i, j])) / 3) + T_a[i, j])
criteria_names = ['c11', 'c12', 'c13', 'c14', 'c15', 'c21', 'c22', 'c31', 'c32', 'c33',
                  'c41', 'c51', 'c52', 'c53', 'c54', 'c61', 'c62', 'c63', 'c64']
dff = pd.DataFrame(supermatrix, index=criteria_names, columns=criteria_names)
dff.to_excel('Tc_defuzzied.xlsx')

# creating data frame and saving a table to excel
# l
data2 = {'D_l': D_a, 'D_m': D_b, 'D_u': D_c, 'R_l': R_a, 'R_m': R_b, 'R_u': R_c,
         'defuzzied_D': defuzzied_D, 'defuzzied_R': defuzzied_R, 'D+R': D_plus_R, 'D-R': D_minus_R}
dff2 = pd.DataFrame(data2, index=criteria_names)
dff2.to_excel('D_R.xlsx')

# creating output for l,m,u of experts average output
mylist5 = []
for i in range(0, len(mylist1)):
    mylist4 = []
    mylist5.append(mylist4)
    for j in range(0, len(mylist1)):
        each = mylist[i][j][0]
        mylist4.append(each)

dff3 = pd.DataFrame(mylist5, index=criteria_names, columns=criteria_names)
dff3.to_excel('l_experts_average.xlsx')
# m
mylist6 = []
for i in range(0, len(mylist1)):
    mylist7 = []
    mylist6.append(mylist7)
    for j in range(0, len(mylist1)):
        each = mylist[i][j][1]
        mylist7.append(each)

dff4 = pd.DataFrame(mylist6, index=criteria_names, columns=criteria_names)
dff4.to_excel('m_experts_average.xlsx')
# u
mylist8 = []
for i in range(0, len(mylist1)):
    mylist9 = []
    mylist8.append(mylist9)
    for j in range(0, len(mylist1)):
        each = mylist[i][j][2]
        mylist9.append(each)

dff5 = pd.DataFrame(mylist8, index=criteria_names, columns=criteria_names)
dff5.to_excel('u_experts_average.xlsx')

# creating normalized l,m,u output
dff6 = pd.DataFrame(X_a, index=criteria_names, columns=criteria_names)
dff6.to_excel('normalized_l.xlsx')

dff7 = pd.DataFrame(X_b, index=criteria_names, columns=criteria_names)
dff7.to_excel('normalized_m.xlsx')

dff8 = pd.DataFrame(X_c, index=criteria_names, columns=criteria_names)
dff8.to_excel('normalized_u.xlsx')

# creating total matrix l,m,u output
dff9 = pd.DataFrame(T_a, index=criteria_names, columns=criteria_names)
dff9.to_excel('Tc_l.xlsx')

dff10 = pd.DataFrame(T_b, index=criteria_names, columns=criteria_names)
dff10.to_excel('Tc_m.xlsx')

dff11 = pd.DataFrame(T_c, index=criteria_names, columns=criteria_names)
dff11.to_excel('Tc_u.xlsx')

# defining dimensions for total matrix
# l of each dimension
Td_a_c1c1 = T_a[0:5, 0:5]
Td_a_c1c2 = T_a[0:5, 5:7]
Td_a_c1c3 = T_a[0:5, 7:10]
Td_a_c1c4 = T_a[0:5, 10:11]
Td_a_c1c5 = T_a[0:5, 11:15]
Td_a_c1c6 = T_a[0:5, 15:]

Td_a_c2c1 = T_a[5:7, 0:5]
Td_a_c2c2 = T_a[5:7, 5:7]
Td_a_c2c3 = T_a[5:7, 7:10]
Td_a_c2c4 = T_a[5:7, 10:11]
Td_a_c2c5 = T_a[5:7, 11:15]
Td_a_c2c6 = T_a[5:7, 15:]

Td_a_c3c1 = T_a[7:10, 0:5]
Td_a_c3c2 = T_a[7:10, 5:7]
Td_a_c3c3 = T_a[7:10, 7:10]
Td_a_c3c4 = T_a[7:10, 10:11]
Td_a_c3c5 = T_a[7:10, 11:15]
Td_a_c3c6 = T_a[7:10, 15:]

Td_a_c4c1 = T_a[10:11, 0:5]
Td_a_c4c2 = T_a[10:11, 5:7]
Td_a_c4c3 = T_a[10:11, 7:10]
Td_a_c4c4 = T_a[10:11, 10:11]
Td_a_c4c5 = T_a[10:11, 11:15]
Td_a_c4c6 = T_a[10:11, 15:]

Td_a_c5c1 = T_a[11:15, 0:5]
Td_a_c5c2 = T_a[11:15, 5:7]
Td_a_c5c3 = T_a[11:15, 7:10]
Td_a_c5c4 = T_a[11:15, 10:11]
Td_a_c5c5 = T_a[11:15, 11:15]
Td_a_c5c6 = T_a[11:15, 15:]

Td_a_c6c1 = T_a[15:, 0:5]
Td_a_c6c2 = T_a[15:, 5:7]
Td_a_c6c3 = T_a[15:, 7:10]
Td_a_c6c4 = T_a[15:, 10:11]
Td_a_c6c5 = T_a[15:, 11:15]
Td_a_c6c6 = T_a[15:, 15:]

# m
Td_b_c1c1 = T_b[0:5, 0:5]
Td_b_c1c2 = T_b[0:5, 5:7]
Td_b_c1c3 = T_b[0:5, 7:10]
Td_b_c1c4 = T_b[0:5, 10:11]
Td_b_c1c5 = T_b[0:5, 11:15]
Td_b_c1c6 = T_b[0:5, 15:]

Td_b_c2c1 = T_b[5:7, 0:5]
Td_b_c2c2 = T_b[5:7, 5:7]
Td_b_c2c3 = T_b[5:7, 7:10]
Td_b_c2c4 = T_b[5:7, 10:11]
Td_b_c2c5 = T_b[5:7, 11:15]
Td_b_c2c6 = T_b[5:7, 15:]

Td_b_c3c1 = T_b[7:10, 0:5]
Td_b_c3c2 = T_b[7:10, 5:7]
Td_b_c3c3 = T_b[7:10, 7:10]
Td_b_c3c4 = T_b[7:10, 10:11]
Td_b_c3c5 = T_b[7:10, 11:15]
Td_b_c3c6 = T_b[7:10, 15:]
Td_b_c4c1 = T_b[10:11, 0:5]
Td_b_c4c2 = T_b[10:11, 5:7]
Td_b_c4c3 = T_b[10:11, 7:10]
Td_b_c4c4 = T_b[10:11, 10:11]
Td_b_c4c5 = T_b[10:11, 11:15]
Td_b_c4c6 = T_b[10:11, 15:]

Td_b_c5c1 = T_b[11:15, 0:5]
Td_b_c5c2 = T_b[11:15, 5:7]
Td_b_c5c3 = T_b[11:15, 7:10]
Td_b_c5c4 = T_b[11:15, 10:11]
Td_b_c5c5 = T_b[11:15, 11:15]
Td_b_c5c6 = T_b[11:15, 15:]

Td_b_c6c1 = T_b[15:, 0:5]
Td_b_c6c2 = T_b[15:, 5:7]
Td_b_c6c3 = T_b[15:, 7:10]
Td_b_c6c4 = T_b[15:, 10:11]
Td_b_c6c5 = T_b[15:, 11:15]
Td_b_c6c6 = T_b[15:, 15:]

# u
Td_c_c1c1 = T_c[0:5, 0:5]
Td_c_c1c2 = T_c[0:5, 5:7]
Td_c_c1c3 = T_c[0:5, 7:10]
Td_c_c1c4 = T_c[0:5, 10:11]
Td_c_c1c5 = T_c[0:5, 11:15]
Td_c_c1c6 = T_c[0:5, 15:]

Td_c_c2c1 = T_c[5:7, 0:5]
Td_c_c2c2 = T_c[5:7, 5:7]
Td_c_c2c3 = T_c[5:7, 7:10]
Td_c_c2c4 = T_c[5:7, 10:11]
Td_c_c2c5 = T_c[5:7, 11:15]
Td_c_c2c6 = T_c[5:7, 15:]

Td_c_c3c1 = T_c[7:10, 0:5]
Td_c_c3c2 = T_c[7:10, 5:7]
Td_c_c3c3 = T_c[7:10, 7:10]
Td_c_c3c4 = T_c[7:10, 10:11]
Td_c_c3c5 = T_c[7:10, 11:15]
Td_c_c3c6 = T_c[7:10, 15:]

Td_c_c4c1 = T_c[10:11, 0:5]
Td_c_c4c2 = T_c[10:11, 5:7]
Td_c_c4c3 = T_c[10:11, 7:10]
Td_c_c4c4 = T_c[10:11, 10:11]
Td_c_c4c5 = T_c[10:11, 11:15]
Td_c_c4c6 = T_c[10:11, 15:]

Td_c_c5c1 = T_c[11:15, 0:5]
Td_c_c5c2 = T_c[11:15, 5:7]
Td_c_c5c3 = T_c[11:15, 7:10]
Td_c_c5c4 = T_c[11:15, 10:11]
Td_c_c5c5 = T_c[11:15, 11:15]
Td_c_c5c6 = T_c[11:15, 15:]

Td_c_c6c1 = T_c[15:, 0:5]
Td_c_c6c2 = T_c[15:, 5:7]
Td_c_c6c3 = T_c[15:, 7:10]
Td_c_c6c4 = T_c[15:, 10:11]
Td_c_c6c5 = T_c[15:, 11:15]
Td_c_c6c6 = T_c[15:, 15:]

# creating Td
Td_a = np.zeros((6, 6))
Td_b = np.zeros((6, 6))
Td_c = np.zeros((6, 6))

# l
Td_a[0, 0] = np.mean(Td_a_c1c1)
Td_a[0, 1] = np.mean(Td_a_c1c2)
Td_a[0, 2] = np.mean(Td_a_c1c3)
Td_a[0, 3] = np.mean(Td_a_c1c4)
Td_a[0, 4] = np.mean(Td_a_c1c5)
Td_a[0, 5] = np.mean(Td_a_c1c6)

Td_a[1, 0] = np.mean(Td_a_c2c1)
Td_a[1, 1] = np.mean(Td_a_c2c2)
Td_a[1, 2] = np.mean(Td_a_c2c3)
Td_a[1, 3] = np.mean(Td_a_c2c4)
Td_a[1, 4] = np.mean(Td_a_c2c5)
Td_a[1, 5] = np.mean(Td_a_c2c6)

Td_a[2, 0] = np.mean(Td_a_c3c1)
Td_a[2, 1] = np.mean(Td_a_c3c2)
Td_a[2, 2] = np.mean(Td_a_c3c3)
Td_a[2, 3] = np.mean(Td_a_c3c4)
Td_a[2, 4] = np.mean(Td_a_c3c5)
Td_a[2, 5] = np.mean(Td_a_c3c6)

Td_a[3, 0] = np.mean(Td_a_c4c1)
Td_a[3, 1] = np.mean(Td_a_c4c2)
Td_a[3, 2] = np.mean(Td_a_c4c3)
Td_a[3, 3] = np.mean(Td_a_c4c4)
Td_a[3, 4] = np.mean(Td_a_c4c5)
Td_a[3, 5] = np.mean(Td_a_c4c6)

Td_a[4, 0] = np.mean(Td_a_c5c1)
Td_a[4, 1] = np.mean(Td_a_c5c2)
Td_a[4, 2] = np.mean(Td_a_c5c3)
Td_a[4, 3] = np.mean(Td_a_c5c4)
Td_a[4, 4] = np.mean(Td_a_c5c5)
Td_a[4, 5] = np.mean(Td_a_c5c6)

Td_a[5, 0] = np.mean(Td_a_c6c1)
Td_a[5, 1] = np.mean(Td_a_c6c2)
Td_a[5, 2] = np.mean(Td_a_c6c3)
Td_a[5, 3] = np.mean(Td_a_c6c4)
Td_a[5, 4] = np.mean(Td_a_c6c5)
Td_a[5, 5] = np.mean(Td_a_c6c6)

# m
Td_b[0, 0] = np.mean(Td_b_c1c1)
Td_b[0, 1] = np.mean(Td_b_c1c2)
Td_b[0, 2] = np.mean(Td_b_c1c3)
Td_b[0, 3] = np.mean(Td_b_c1c4)
Td_b[0, 4] = np.mean(Td_b_c1c5)
Td_b[0, 5] = np.mean(Td_b_c1c6)

Td_b[1, 0] = np.mean(Td_b_c2c1)
Td_b[1, 1] = np.mean(Td_b_c2c2)
Td_b[1, 2] = np.mean(Td_b_c2c3)
Td_b[1, 3] = np.mean(Td_b_c2c4)
Td_b[1, 4] = np.mean(Td_b_c2c5)
Td_b[1, 5] = np.mean(Td_b_c2c6)

Td_b[2, 0] = np.mean(Td_b_c3c1)
Td_b[2, 1] = np.mean(Td_b_c3c2)
Td_b[2, 2] = np.mean(Td_b_c3c3)
Td_b[2, 3] = np.mean(Td_b_c3c4)
Td_b[2, 4] = np.mean(Td_b_c3c5)
Td_b[2, 5] = np.mean(Td_b_c3c6)

Td_b[3, 0] = np.mean(Td_b_c4c1)
Td_b[3, 1] = np.mean(Td_b_c4c2)
Td_b[3, 2] = np.mean(Td_b_c4c3)
Td_b[3, 3] = np.mean(Td_b_c4c4)
Td_b[3, 4] = np.mean(Td_b_c4c5)
Td_b[3, 5] = np.mean(Td_b_c4c6)

Td_b[4, 0] = np.mean(Td_b_c5c1)
Td_b[4, 1] = np.mean(Td_b_c5c2)
Td_b[4, 2] = np.mean(Td_b_c5c3)
Td_b[4, 3] = np.mean(Td_b_c5c4)
Td_b[4, 4] = np.mean(Td_b_c5c5)
Td_b[4, 5] = np.mean(Td_b_c5c6)

Td_b[5, 0] = np.mean(Td_b_c6c1)
Td_b[5, 1] = np.mean(Td_b_c6c2)
Td_b[5, 2] = np.mean(Td_b_c6c3)
Td_b[5, 3] = np.mean(Td_b_c6c4)
Td_b[5, 4] = np.mean(Td_b_c6c5)
Td_b[5, 5] = np.mean(Td_b_c6c6)

# u
Td_c[0, 0] = np.mean(Td_c_c1c1)
Td_c[0, 1] = np.mean(Td_c_c1c2)
Td_c[0, 2] = np.mean(Td_c_c1c3)
Td_c[0, 3] = np.mean(Td_c_c1c4)
Td_c[0, 4] = np.mean(Td_c_c1c5)
Td_c[0, 5] = np.mean(Td_c_c1c6)

Td_c[1, 0] = np.mean(Td_c_c2c1)
Td_c[1, 1] = np.mean(Td_c_c2c2)
Td_c[1, 2] = np.mean(Td_c_c2c3)
Td_c[1, 3] = np.mean(Td_c_c2c4)
Td_c[1, 4] = np.mean(Td_c_c2c5)
Td_c[1, 5] = np.mean(Td_c_c2c6)

Td_c[2, 0] = np.mean(Td_c_c3c1)
Td_c[2, 1] = np.mean(Td_c_c3c2)
Td_c[2, 2] = np.mean(Td_c_c3c3)
Td_c[2, 3] = np.mean(Td_c_c3c4)
Td_c[2, 4] = np.mean(Td_c_c3c5)
Td_c[2, 5] = np.mean(Td_c_c3c6)

Td_c[3, 0] = np.mean(Td_c_c4c1)
Td_c[3, 1] = np.mean(Td_c_c4c2)
Td_c[3, 2] = np.mean(Td_c_c4c3)
Td_c[3, 3] = np.mean(Td_c_c4c4)
Td_c[3, 4] = np.mean(Td_c_c4c5)
Td_c[3, 5] = np.mean(Td_c_c4c6)

Td_c[4, 0] = np.mean(Td_c_c5c1)
Td_c[4, 1] = np.mean(Td_c_c5c2)
Td_c[4, 2] = np.mean(Td_c_c5c3)
Td_c[4, 3] = np.mean(Td_c_c5c4)
Td_c[4, 4] = np.mean(Td_c_c5c5)
Td_c[4, 5] = np.mean(Td_c_c5c6)

Td_c[5, 0] = np.mean(Td_c_c6c1)
Td_c[5, 1] = np.mean(Td_c_c6c2)
Td_c[5, 2] = np.mean(Td_c_c6c3)
Td_c[5, 3] = np.mean(Td_c_c6c4)
Td_c[5, 4] = np.mean(Td_c_c6c5)
Td_c[5, 5] = np.mean(Td_c_c6c6)

# output of Td
dimension_names = ['c1', 'c2', 'c3', 'c4', 'c5', 'c6']
dff12 = pd.DataFrame(Td_a, index=dimension_names, columns=dimension_names)
dff12.to_excel('Td_l.xlsx')

dff13 = pd.DataFrame(Td_b, index=dimension_names, columns=dimension_names)
dff13.to_excel('Td_m.xlsx')

dff14 = pd.DataFrame(Td_c, index=dimension_names, columns=dimension_names)
dff14.to_excel('Td_u.xlsx')

# calculating R, D for Td
Td_D_a = np.sum(Td_a, axis=1)
Td_D_b = np.sum(Td_b, axis=1)
Td_D_c = np.sum(Td_c, axis=1)
Td_R_a = np.sum(Td_a, axis=0)
Td_R_b = np.sum(Td_b, axis=0)
Td_R_c = np.sum(Td_c, axis=0)

Td_defuzzied_D = (((Td_D_c - Td_D_a) + (Td_D_b - Td_D_a)) / 3) + Td_D_a
Td_defuzzied_R = (((Td_R_c - Td_R_a) + (Td_R_b - Td_R_a)) / 3) + Td_R_a
Td_D_plus_R = Td_defuzzied_D + Td_defuzzied_R
Td_D_minus_R = Td_defuzzied_D - Td_defuzzied_R

data3 = {'D_l': Td_D_a, 'D_m': Td_D_b, 'D_u': Td_D_c, 'R_l': Td_R_a, 'R_m': Td_R_b, 'R_u': Td_R_c,
         'defuzzied_D': Td_defuzzied_D, 'defuzzied_R': Td_defuzzied_R, 'D+R': Td_D_plus_R, 'D-R': Td_D_minus_R}
dff15 = pd.DataFrame(data3, index=dimension_names)
dff15.to_excel('Td_D_R.xlsx')

# defuzzied Td
supermatrix3 = []
supermatrix4 = []

for i in range(0, 6):
    supermatrix4 = []
    supermatrix3.append(supermatrix4)  # for some reason this works i dont know why??
    for j in range(0, 6):
        supermatrix4.append((((Td_c[i, j] - Td_a[i, j]) + (Td_b[i, j] - Td_a[i, j])) / 3) + Td_a[i, j])

dff16 = pd.DataFrame(supermatrix3, index=dimension_names, columns=dimension_names)
dff16.to_excel('Td_defuzzied.xlsx')


# normalizing each matrix in supermatrix
def normalizing(matrix):
    array = matrix
    norm_mat = array
    for i in range(0, len(array)):
        line_sum = 0
        line_sum += np.sum(array[i])
        norm_mat[i] = array[i] / line_sum
    return (norm_mat)


##########################################################################
# sending Tds to functions changes them from here on it should not be used
##########################################################################

# l of normalized
norm_mat_l_c1c1 = normalizing(Td_a_c1c1)
norm_mat_l_c1c2 = normalizing(Td_a_c1c2)
norm_mat_l_c1c3 = normalizing(Td_a_c1c3)
norm_mat_l_c1c4 = normalizing(Td_a_c1c4)
norm_mat_l_c1c5 = normalizing(Td_a_c1c5)
norm_mat_l_c1c6 = normalizing(Td_a_c1c6)

norm_mat_l_c2c1 = normalizing(Td_a_c2c1)
norm_mat_l_c2c2 = normalizing(Td_a_c2c2)
norm_mat_l_c2c3 = normalizing(Td_a_c2c3)
norm_mat_l_c2c4 = normalizing(Td_a_c2c4)
norm_mat_l_c2c5 = normalizing(Td_a_c2c5)
norm_mat_l_c2c6 = normalizing(Td_a_c2c6)

norm_mat_l_c3c1 = normalizing(Td_a_c3c1)
norm_mat_l_c3c2 = normalizing(Td_a_c3c2)
norm_mat_l_c3c3 = normalizing(Td_a_c3c3)
norm_mat_l_c3c4 = normalizing(Td_a_c3c4)
norm_mat_l_c3c5 = normalizing(Td_a_c3c5)
norm_mat_l_c3c6 = normalizing(Td_a_c3c6)

norm_mat_l_c4c1 = normalizing(Td_a_c4c1)
norm_mat_l_c4c2 = normalizing(Td_a_c4c2)
norm_mat_l_c4c3 = normalizing(Td_a_c4c3)
norm_mat_l_c4c4 = normalizing(Td_a_c4c4)
norm_mat_l_c4c5 = normalizing(Td_a_c4c5)
norm_mat_l_c4c6 = normalizing(Td_a_c4c6)

norm_mat_l_c5c1 = normalizing(Td_a_c5c1)
norm_mat_l_c5c2 = normalizing(Td_a_c5c2)
norm_mat_l_c5c3 = normalizing(Td_a_c5c3)
norm_mat_l_c5c4 = normalizing(Td_a_c5c4)
norm_mat_l_c5c5 = normalizing(Td_a_c5c5)
norm_mat_l_c5c6 = normalizing(Td_a_c5c6)

norm_mat_l_c6c1 = normalizing(Td_a_c6c1)
norm_mat_l_c6c2 = normalizing(Td_a_c6c2)
norm_mat_l_c6c3 = normalizing(Td_a_c6c3)
norm_mat_l_c6c4 = normalizing(Td_a_c6c4)
norm_mat_l_c6c5 = normalizing(Td_a_c6c5)
norm_mat_l_c6c6 = normalizing(Td_a_c6c6)

# m
norm_mat_m_c1c1 = normalizing(Td_b_c1c1)
norm_mat_m_c1c2 = normalizing(Td_b_c1c2)
norm_mat_m_c1c3 = normalizing(Td_b_c1c3)
norm_mat_m_c1c4 = normalizing(Td_b_c1c4)
norm_mat_m_c1c5 = normalizing(Td_b_c1c5)
norm_mat_m_c1c6 = normalizing(Td_b_c1c6)

norm_mat_m_c2c1 = normalizing(Td_b_c2c1)
norm_mat_m_c2c2 = normalizing(Td_b_c2c2)
norm_mat_m_c2c3 = normalizing(Td_b_c2c3)
norm_mat_m_c2c4 = normalizing(Td_b_c2c4)
norm_mat_m_c2c5 = normalizing(Td_b_c2c5)
norm_mat_m_c2c6 = normalizing(Td_b_c2c6)

norm_mat_m_c3c1 = normalizing(Td_b_c3c1)
norm_mat_m_c3c2 = normalizing(Td_b_c3c2)
norm_mat_m_c3c3 = normalizing(Td_b_c3c3)
norm_mat_m_c3c4 = normalizing(Td_b_c3c4)
norm_mat_m_c3c5 = normalizing(Td_b_c3c5)
norm_mat_m_c3c6 = normalizing(Td_b_c3c6)

norm_mat_m_c4c1 = normalizing(Td_b_c4c1)
norm_mat_m_c4c2 = normalizing(Td_b_c4c2)
norm_mat_m_c4c3 = normalizing(Td_b_c4c3)
norm_mat_m_c4c4 = normalizing(Td_b_c4c4)
norm_mat_m_c4c5 = normalizing(Td_b_c4c5)
norm_mat_m_c4c6 = normalizing(Td_b_c4c6)

norm_mat_m_c5c1 = normalizing(Td_b_c5c1)
norm_mat_m_c5c2 = normalizing(Td_b_c5c2)
norm_mat_m_c5c3 = normalizing(Td_b_c5c3)
norm_mat_m_c5c4 = normalizing(Td_b_c5c4)
norm_mat_m_c5c5 = normalizing(Td_b_c5c5)
norm_mat_m_c5c6 = normalizing(Td_b_c5c6)

norm_mat_m_c6c1 = normalizing(Td_b_c6c1)
norm_mat_m_c6c2 = normalizing(Td_b_c6c2)
norm_mat_m_c6c3 = normalizing(Td_b_c6c3)
norm_mat_m_c6c4 = normalizing(Td_b_c6c4)
norm_mat_m_c6c5 = normalizing(Td_b_c6c5)
norm_mat_m_c6c6 = normalizing(Td_b_c6c6)

# u
norm_mat_u_c1c1 = normalizing(Td_c_c1c1)
norm_mat_u_c1c2 = normalizing(Td_c_c1c2)
norm_mat_u_c1c3 = normalizing(Td_c_c1c3)
norm_mat_u_c1c4 = normalizing(Td_c_c1c4)
norm_mat_u_c1c5 = normalizing(Td_c_c1c5)
norm_mat_u_c1c6 = normalizing(Td_c_c1c6)

norm_mat_u_c2c1 = normalizing(Td_c_c2c1)
norm_mat_u_c2c2 = normalizing(Td_c_c2c2)
norm_mat_u_c2c3 = normalizing(Td_c_c2c3)
norm_mat_u_c2c4 = normalizing(Td_c_c2c4)
norm_mat_u_c2c5 = normalizing(Td_c_c2c5)
norm_mat_u_c2c6 = normalizing(Td_c_c2c6)

norm_mat_u_c3c1 = normalizing(Td_c_c3c1)
norm_mat_u_c3c2 = normalizing(Td_c_c3c2)
norm_mat_u_c3c3 = normalizing(Td_c_c3c3)
norm_mat_u_c3c4 = normalizing(Td_c_c3c4)
norm_mat_u_c3c5 = normalizing(Td_c_c3c5)
norm_mat_u_c3c6 = normalizing(Td_c_c3c6)

norm_mat_u_c4c1 = normalizing(Td_c_c4c1)
norm_mat_u_c4c2 = normalizing(Td_c_c4c2)
norm_mat_u_c4c3 = normalizing(Td_c_c4c3)
norm_mat_u_c4c4 = normalizing(Td_c_c4c4)
norm_mat_u_c4c5 = normalizing(Td_c_c4c5)
norm_mat_u_c4c6 = normalizing(Td_c_c4c6)

norm_mat_u_c5c1 = normalizing(Td_c_c5c1)
norm_mat_u_c5c2 = normalizing(Td_c_c5c2)
norm_mat_u_c5c3 = normalizing(Td_c_c5c3)
norm_mat_u_c5c4 = normalizing(Td_c_c5c4)
norm_mat_u_c5c5 = normalizing(Td_c_c5c5)
norm_mat_u_c5c6 = normalizing(Td_c_c5c6)

norm_mat_u_c6c1 = normalizing(Td_c_c6c1)
norm_mat_u_c6c2 = normalizing(Td_c_c6c2)
norm_mat_u_c6c3 = normalizing(Td_c_c6c3)
norm_mat_u_c6c4 = normalizing(Td_c_c6c4)
norm_mat_u_c6c5 = normalizing(Td_c_c6c5)
norm_mat_u_c6c6 = normalizing(Td_c_c6c6)

# creating an array of arrays
# l
# first column
array_l_col_11 = norm_mat_l_c1c1
array_l_col_21 = norm_mat_l_c2c1
array_l_col_31 = norm_mat_l_c3c1
array_l_col_41 = norm_mat_l_c4c1
array_l_col_51 = norm_mat_l_c5c1
array_l_col_61 = norm_mat_l_c6c1
tuple_l_col_1 = (array_l_col_11, array_l_col_21, array_l_col_31, array_l_col_41, array_l_col_51, array_l_col_61)
array_l_col_1 = np.vstack(tuple_l_col_1)

# second column
array_l_col_12 = norm_mat_l_c1c2
array_l_col_22 = norm_mat_l_c2c2
array_l_col_32 = norm_mat_l_c3c2
array_l_col_42 = norm_mat_l_c4c2
array_l_col_52 = norm_mat_l_c5c2
array_l_col_62 = norm_mat_l_c6c2
tuple_l_col_2 = (array_l_col_12, array_l_col_22, array_l_col_32, array_l_col_42, array_l_col_52, array_l_col_62)
array_l_col_2 = np.vstack(tuple_l_col_2)

# third column
array_l_col_13 = norm_mat_l_c1c3
array_l_col_23 = norm_mat_l_c2c3
array_l_col_33 = norm_mat_l_c3c3
array_l_col_43 = norm_mat_l_c4c3
array_l_col_53 = norm_mat_l_c5c3
array_l_col_63 = norm_mat_l_c6c3
tuple_l_col_3 = (array_l_col_13, array_l_col_23, array_l_col_33, array_l_col_43, array_l_col_53, array_l_col_63)
array_l_col_3 = np.vstack(tuple_l_col_3)

# forth column
array_l_col_14 = norm_mat_l_c1c4
array_l_col_24 = norm_mat_l_c2c4
array_l_col_34 = norm_mat_l_c3c4
array_l_col_44 = norm_mat_l_c4c4
array_l_col_54 = norm_mat_l_c5c4
array_l_col_64 = norm_mat_l_c6c4
tuple_l_col_4 = (array_l_col_14, array_l_col_24, array_l_col_34, array_l_col_44, array_l_col_54, array_l_col_64)
array_l_col_4 = np.vstack(tuple_l_col_4)

# fifth column
array_l_col_15 = norm_mat_l_c1c5
array_l_col_25 = norm_mat_l_c2c5
array_l_col_35 = norm_mat_l_c3c5
array_l_col_45 = norm_mat_l_c4c5
array_l_col_55 = norm_mat_l_c5c5
array_l_col_65 = norm_mat_l_c6c5
tuple_l_col_5 = (array_l_col_15, array_l_col_25, array_l_col_35, array_l_col_45, array_l_col_55, array_l_col_65)
array_l_col_5 = np.vstack(tuple_l_col_5)

# sixth column
array_l_col_16 = norm_mat_l_c1c6
array_l_col_26 = norm_mat_l_c2c6
array_l_col_36 = norm_mat_l_c3c6
array_l_col_46 = norm_mat_l_c4c6
array_l_col_56 = norm_mat_l_c5c6
array_l_col_66 = norm_mat_l_c6c6
tuple_l_col_6 = (array_l_col_16, array_l_col_26, array_l_col_36, array_l_col_46, array_l_col_56, array_l_col_66)
array_l_col_6 = np.vstack(tuple_l_col_6)

# combining the columns to form array of arrays
tuple_l = (array_l_col_1, array_l_col_2, array_l_col_3, array_l_col_4, array_l_col_5, array_l_col_6)
normal_super_l = np.hstack(tuple_l)

# m
# first column
array_m_col_11 = norm_mat_m_c1c1
array_m_col_21 = norm_mat_m_c2c1
array_m_col_31 = norm_mat_m_c3c1
array_m_col_41 = norm_mat_m_c4c1
array_m_col_51 = norm_mat_m_c5c1
array_m_col_61 = norm_mat_m_c6c1
tuple_m_col_1 = (array_m_col_11, array_m_col_21, array_m_col_31, array_m_col_41, array_m_col_51, array_m_col_61)
array_m_col_1 = np.vstack(tuple_m_col_1)

# second column
array_m_col_12 = norm_mat_m_c1c2
array_m_col_22 = norm_mat_m_c2c2
array_m_col_32 = norm_mat_m_c3c2
array_m_col_42 = norm_mat_m_c4c2
array_m_col_52 = norm_mat_m_c5c2
array_m_col_62 = norm_mat_m_c6c2
tuple_m_col_2 = (array_m_col_12, array_m_col_22, array_m_col_32, array_m_col_42, array_m_col_52, array_m_col_62)
array_m_col_2 = np.vstack(tuple_m_col_2)

# third column
array_m_col_13 = norm_mat_m_c1c3
array_m_col_23 = norm_mat_m_c2c3
array_m_col_33 = norm_mat_m_c3c3
array_m_col_43 = norm_mat_m_c4c3
array_m_col_53 = norm_mat_m_c5c3
array_m_col_63 = norm_mat_m_c6c3
tuple_m_col_3 = (array_m_col_13, array_m_col_23, array_m_col_33, array_m_col_43, array_m_col_53, array_m_col_63)
array_m_col_3 = np.vstack(tuple_m_col_3)

# forth column
array_m_col_14 = norm_mat_m_c1c4
array_m_col_24 = norm_mat_m_c2c4
array_m_col_34 = norm_mat_m_c3c4
array_m_col_44 = norm_mat_m_c4c4
array_m_col_54 = norm_mat_m_c5c4
array_m_col_64 = norm_mat_m_c6c4
tuple_m_col_4 = (array_m_col_14, array_m_col_24, array_m_col_34, array_m_col_44, array_m_col_54, array_m_col_64)
array_m_col_4 = np.vstack(tuple_m_col_4)

# fifth column
array_m_col_15 = norm_mat_m_c1c5
array_m_col_25 = norm_mat_m_c2c5
array_m_col_35 = norm_mat_m_c3c5
array_m_col_45 = norm_mat_m_c4c5
array_m_col_55 = norm_mat_m_c5c5
array_m_col_65 = norm_mat_m_c6c5
tuple_m_col_5 = (array_m_col_15, array_m_col_25, array_m_col_35, array_m_col_45, array_m_col_55, array_m_col_65)
array_m_col_5 = np.vstack(tuple_m_col_5)

# sixth column
array_m_col_16 = norm_mat_m_c1c6
array_m_col_26 = norm_mat_m_c2c6
array_m_col_36 = norm_mat_m_c3c6
array_m_col_46 = norm_mat_m_c4c6
array_m_col_56 = norm_mat_m_c5c6
array_m_col_66 = norm_mat_m_c6c6
tuple_m_col_6 = (array_m_col_16, array_m_col_26, array_m_col_36, array_m_col_46, array_m_col_56, array_m_col_66)
array_m_col_6 = np.vstack(tuple_m_col_6)

# combining the columns to form array of arrays
tuple_m = (array_m_col_1, array_m_col_2, array_m_col_3, array_m_col_4, array_m_col_5, array_m_col_6)
normal_super_m = np.hstack(tuple_m)

# u
# first column
array_u_col_11 = norm_mat_u_c1c1
array_u_col_21 = norm_mat_u_c2c1
array_u_col_31 = norm_mat_u_c3c1
array_u_col_41 = norm_mat_u_c4c1
array_u_col_51 = norm_mat_u_c5c1
array_u_col_61 = norm_mat_u_c6c1
tuple_u_col_1 = (array_u_col_11, array_u_col_21, array_u_col_31, array_u_col_41, array_u_col_51, array_u_col_61)
array_u_col_1 = np.vstack(tuple_u_col_1)

# second column
array_u_col_12 = norm_mat_u_c1c2
array_u_col_22 = norm_mat_u_c2c2
array_u_col_32 = norm_mat_u_c3c2
array_u_col_42 = norm_mat_u_c4c2
array_u_col_52 = norm_mat_u_c5c2
array_u_col_62 = norm_mat_u_c6c2
tuple_u_col_2 = (array_u_col_12, array_u_col_22, array_u_col_32, array_u_col_42, array_u_col_52, array_u_col_62)
array_u_col_2 = np.vstack(tuple_u_col_2)

# third column
array_u_col_13 = norm_mat_u_c1c3
array_u_col_23 = norm_mat_u_c2c3
array_u_col_33 = norm_mat_u_c3c3
array_u_col_43 = norm_mat_u_c4c3
array_u_col_53 = norm_mat_u_c5c3
array_u_col_63 = norm_mat_u_c6c3
tuple_u_col_3 = (array_u_col_13, array_u_col_23, array_u_col_33, array_u_col_43, array_u_col_53, array_u_col_63)
array_u_col_3 = np.vstack(tuple_u_col_3)

# forth column
array_u_col_14 = norm_mat_u_c1c4
array_u_col_24 = norm_mat_u_c2c4
array_u_col_34 = norm_mat_u_c3c4
array_u_col_44 = norm_mat_u_c4c4
array_u_col_54 = norm_mat_u_c5c4
array_u_col_64 = norm_mat_u_c6c4
tuple_u_col_4 = (array_u_col_14, array_u_col_24, array_u_col_34, array_u_col_44, array_u_col_54, array_u_col_64)
array_u_col_4 = np.vstack(tuple_u_col_4)

# fifth column
array_u_col_15 = norm_mat_u_c1c5
array_u_col_25 = norm_mat_u_c2c5
array_u_col_35 = norm_mat_u_c3c5
array_u_col_45 = norm_mat_u_c4c5
array_u_col_55 = norm_mat_u_c5c5
array_u_col_65 = norm_mat_u_c6c5
tuple_u_col_5 = (array_u_col_15, array_u_col_25, array_u_col_35, array_u_col_45, array_u_col_55, array_u_col_65)
array_u_col_5 = np.vstack(tuple_u_col_5)

# sixth column
array_u_col_16 = norm_mat_u_c1c6
array_u_col_26 = norm_mat_u_c2c6
array_u_col_36 = norm_mat_u_c3c6
array_u_col_46 = norm_mat_u_c4c6
array_u_col_56 = norm_mat_u_c5c6
array_u_col_66 = norm_mat_u_c6c6
tuple_u_col_6 = (array_u_col_16, array_u_col_26, array_u_col_36, array_u_col_46, array_u_col_56, array_u_col_66)
array_u_col_6 = np.vstack(tuple_u_col_6)

# combining the columns to form array of arrays
tuple_u = (array_u_col_1, array_u_col_2, array_u_col_3, array_u_col_4, array_u_col_5, array_u_col_6)
normal_super_u = np.hstack(tuple_u)

# Transposing super matrixes to form unweighted super matrix
unweighted_l = normal_super_l.transpose()
unweighted_m = normal_super_m.transpose()
unweighted_u = normal_super_u.transpose()

dff17 = pd.DataFrame(unweighted_l, index=criteria_names, columns=criteria_names)
dff17.to_excel('unweighted_l.xlsx')

dff18 = pd.DataFrame(unweighted_m, index=criteria_names, columns=criteria_names)
dff18.to_excel('unweighted_m.xlsx')

dff19 = pd.DataFrame(unweighted_u, index=criteria_names, columns=criteria_names)
dff19.to_excel('unweighted_u.xlsx')

# normalizing Td
Td_l_norm = Td_a
Td_m_norm = Td_b
Td_u_norm = Td_c
for i in range(0, len(dimension_names)):
    Td_l_norm[i] = (Td_a[i] / Td_D_a[i])
    Td_m_norm[i] = (Td_b[i] / Td_D_b[i])
    Td_u_norm[i] = (Td_c[i] / Td_D_c[i])

# transposing Td
Td_l_norm_trans = Td_l_norm.transpose()
Td_m_norm_trans = Td_m_norm.transpose()
Td_u_norm_trans = Td_u_norm.transpose()

dff20 = pd.DataFrame(Td_l_norm_trans, index=dimension_names, columns=dimension_names)
dff20.to_excel('Td_l_norm_trans.xlsx')
dff21 = pd.DataFrame(Td_m_norm_trans, index=dimension_names, columns=dimension_names)
dff21.to_excel('Td_m_norm_trans.xlsx')
dff22 = pd.DataFrame(Td_u_norm_trans, index=dimension_names, columns=dimension_names)
dff22.to_excel('Td_u_norm_trans.xlsx')

# creating the weighted supermatrix
# l
weighted_l = unweighted_l

weighted_l[0:5, 0:5] = (unweighted_l[0:5, 0:5] * Td_l_norm_trans[0, 0])
weighted_l[0:5, 5:7] = (unweighted_l[0:5, 5:7] * Td_l_norm_trans[0, 1])
weighted_l[0:5, 7:10] = (unweighted_l[0:5, 7:10] * Td_l_norm_trans[0, 2])
weighted_l[0:5, 10:11] = (unweighted_l[0:5, 10:11] * Td_l_norm_trans[0, 3])
weighted_l[0:5, 11:15] = (unweighted_l[0:5, 11:15] * Td_l_norm_trans[0, 4])
weighted_l[0:5, 15:] = (unweighted_l[0:5, 15:] * Td_l_norm_trans[0, 5])

weighted_l[5:7, 0:5] = (unweighted_l[5:7, 0:5] * Td_l_norm_trans[1, 0])
weighted_l[5:7, 5:7] = (unweighted_l[5:7, 5:7] * Td_l_norm_trans[1, 1])
weighted_l[5:7, 7:10] = (unweighted_l[5:7, 7:10] * Td_l_norm_trans[1, 2])
weighted_l[5:7, 10:11] = (unweighted_l[5:7, 10:11] * Td_l_norm_trans[1, 3])
weighted_l[5:7, 11:15] = (unweighted_l[5:7, 11:15] * Td_l_norm_trans[1, 4])
weighted_l[5:7, 15:] = (unweighted_l[5:7, 15:] * Td_l_norm_trans[1, 5])

weighted_l[7:10, 0:5] = (unweighted_l[7:10, 0:5] * Td_l_norm_trans[2, 0])
weighted_l[7:10, 5:7] = (unweighted_l[7:10, 5:7] * Td_l_norm_trans[2, 1])
weighted_l[7:10, 7:10] = (unweighted_l[7:10, 7:10] * Td_l_norm_trans[2, 2])
weighted_l[7:10, 10:11] = (unweighted_l[7:10, 10:11] * Td_l_norm_trans[2, 3])
weighted_l[7:10, 11:15] = (unweighted_l[7:10, 11:15] * Td_l_norm_trans[2, 4])
weighted_l[7:10, 15:] = (unweighted_l[7:10, 15:] * Td_l_norm_trans[2, 5])

weighted_l[10:11, 0:5] = (unweighted_l[10:11, 0:5] * Td_l_norm_trans[3, 0])
weighted_l[10:11, 5:7] = (unweighted_l[10:11, 5:7] * Td_l_norm_trans[3, 1])
weighted_l[10:11, 7:10] = (unweighted_l[10:11, 7:10] * Td_l_norm_trans[3, 2])
weighted_l[10:11, 10:11] = (unweighted_l[10:11, 10:11] * Td_l_norm_trans[3, 3])
weighted_l[10:11, 11:15] = (unweighted_l[10:11, 11:15] * Td_l_norm_trans[3, 4])
weighted_l[10:11, 15:] = (unweighted_l[10:11, 15:] * Td_l_norm_trans[3, 5])

weighted_l[11:15, 0:5] = (unweighted_l[11:15, 0:5] * Td_l_norm_trans[4, 0])
weighted_l[11:15, 5:7] = (unweighted_l[11:15, 5:7] * Td_l_norm_trans[4, 1])
weighted_l[11:15, 7:10] = (unweighted_l[11:15, 7:10] * Td_l_norm_trans[4, 2])
weighted_l[11:15, 10:11] = (unweighted_l[11:15, 10:11] * Td_l_norm_trans[4, 3])
weighted_l[11:15, 11:15] = (unweighted_l[11:15, 11:15] * Td_l_norm_trans[4, 4])
weighted_l[11:15, 15:] = (unweighted_l[11:15, 15:] * Td_l_norm_trans[4, 5])

weighted_l[15:, 0:5] = (unweighted_l[15:, 0:5] * Td_l_norm_trans[5, 0])
weighted_l[15:, 5:7] = (unweighted_l[15:, 5:7] * Td_l_norm_trans[5, 1])
weighted_l[15:, 7:10] = (unweighted_l[15:, 7:10] * Td_l_norm_trans[5, 2])
weighted_l[15:, 10:11] = (unweighted_l[15:, 10:11] * Td_l_norm_trans[5, 3])
weighted_l[15:, 11:15] = (unweighted_l[15:, 11:15] * Td_l_norm_trans[5, 4])
weighted_l[15:, 15:] = (unweighted_l[15:, 15:] * Td_l_norm_trans[5, 5])

# m
weighted_m = unweighted_m

weighted_m[0:5, 0:5] = (unweighted_m[0:5, 0:5] * Td_m_norm_trans[0, 0])
weighted_m[0:5, 5:7] = (unweighted_m[0:5, 5:7] * Td_m_norm_trans[0, 1])
weighted_m[0:5, 7:10] = (unweighted_m[0:5, 7:10] * Td_m_norm_trans[0, 2])
weighted_m[0:5, 10:11] = (unweighted_m[0:5, 10:11] * Td_m_norm_trans[0, 3])
weighted_m[0:5, 11:15] = (unweighted_m[0:5, 11:15] * Td_m_norm_trans[0, 4])
weighted_m[0:5, 15:] = (unweighted_m[0:5, 15:] * Td_m_norm_trans[0, 5])

weighted_m[5:7, 0:5] = (unweighted_m[5:7, 0:5] * Td_m_norm_trans[1, 0])
weighted_m[5:7, 5:7] = (unweighted_m[5:7, 5:7] * Td_m_norm_trans[1, 1])
weighted_m[5:7, 7:10] = (unweighted_m[5:7, 7:10] * Td_m_norm_trans[1, 2])
weighted_m[5:7, 10:11] = (unweighted_m[5:7, 10:11] * Td_m_norm_trans[1, 3])
weighted_m[5:7, 11:15] = (unweighted_m[5:7, 11:15] * Td_m_norm_trans[1, 4])
weighted_m[5:7, 15:] = (unweighted_m[5:7, 15:] * Td_m_norm_trans[1, 5])

weighted_m[7:10, 0:5] = (unweighted_m[7:10, 0:5] * Td_m_norm_trans[2, 0])
weighted_m[7:10, 5:7] = (unweighted_m[7:10, 5:7] * Td_m_norm_trans[2, 1])
weighted_m[7:10, 7:10] = (unweighted_m[7:10, 7:10] * Td_m_norm_trans[2, 2])
weighted_m[7:10, 10:11] = (unweighted_m[7:10, 10:11] * Td_m_norm_trans[2, 3])
weighted_m[7:10, 11:15] = (unweighted_m[7:10, 11:15] * Td_m_norm_trans[2, 4])
weighted_m[7:10, 15:] = (unweighted_m[7:10, 15:] * Td_m_norm_trans[2, 5])

weighted_m[10:11, 0:5] = (unweighted_m[10:11, 0:5] * Td_m_norm_trans[3, 0])
weighted_m[10:11, 5:7] = (unweighted_m[10:11, 5:7] * Td_m_norm_trans[3, 1])
weighted_m[10:11, 7:10] = (unweighted_m[10:11, 7:10] * Td_m_norm_trans[3, 2])
weighted_m[10:11, 10:11] = (unweighted_m[10:11, 10:11] * Td_m_norm_trans[3, 3])
weighted_m[10:11, 11:15] = (unweighted_m[10:11, 11:15] * Td_m_norm_trans[3, 4])
weighted_m[10:11, 15:] = (unweighted_m[10:11, 15:] * Td_m_norm_trans[3, 5])

weighted_m[11:15, 0:5] = (unweighted_m[11:15, 0:5] * Td_m_norm_trans[4, 0])
weighted_m[11:15, 5:7] = (unweighted_m[11:15, 5:7] * Td_m_norm_trans[4, 1])
weighted_m[11:15, 7:10] = (unweighted_m[11:15, 7:10] * Td_m_norm_trans[4, 2])
weighted_m[11:15, 10:11] = (unweighted_m[11:15, 10:11] * Td_m_norm_trans[4, 3])
weighted_m[11:15, 11:15] = (unweighted_m[11:15, 11:15] * Td_m_norm_trans[4, 4])
weighted_m[11:15, 15:] = (unweighted_m[11:15, 15:] * Td_m_norm_trans[4, 5])

weighted_m[15:, 0:5] = (unweighted_m[15:, 0:5] * Td_m_norm_trans[5, 0])
weighted_m[15:, 5:7] = (unweighted_m[15:, 5:7] * Td_m_norm_trans[5, 1])
weighted_m[15:, 7:10] = (unweighted_m[15:, 7:10] * Td_m_norm_trans[5, 2])
weighted_m[15:, 10:11] = (unweighted_m[15:, 10:11] * Td_m_norm_trans[5, 3])
weighted_m[15:, 11:15] = (unweighted_m[15:, 11:15] * Td_m_norm_trans[5, 4])
weighted_m[15:, 15:] = (unweighted_m[15:, 15:] * Td_m_norm_trans[5, 5])

# u
weighted_u = unweighted_u

weighted_u[0:5, 0:5] = (unweighted_u[0:5, 0:5] * Td_u_norm_trans[0, 0])
weighted_u[0:5, 5:7] = (unweighted_u[0:5, 5:7] * Td_u_norm_trans[0, 1])
weighted_u[0:5, 7:10] = (unweighted_u[0:5, 7:10] * Td_u_norm_trans[0, 2])
weighted_u[0:5, 10:11] = (unweighted_u[0:5, 10:11] * Td_u_norm_trans[0, 3])
weighted_u[0:5, 11:15] = (unweighted_u[0:5, 11:15] * Td_u_norm_trans[0, 4])
weighted_u[0:5, 15:] = (unweighted_u[0:5, 15:] * Td_u_norm_trans[0, 5])

weighted_u[5:7, 0:5] = (unweighted_u[5:7, 0:5] * Td_u_norm_trans[1, 0])
weighted_u[5:7, 5:7] = (unweighted_u[5:7, 5:7] * Td_u_norm_trans[1, 1])
weighted_u[5:7, 7:10] = (unweighted_u[5:7, 7:10] * Td_u_norm_trans[1, 2])
weighted_u[5:7, 10:11] = (unweighted_u[5:7, 10:11] * Td_u_norm_trans[1, 3])
weighted_u[5:7, 11:15] = (unweighted_u[5:7, 11:15] * Td_u_norm_trans[1, 4])
weighted_u[5:7, 15:] = (unweighted_u[5:7, 15:] * Td_u_norm_trans[1, 5])

weighted_u[7:10, 0:5] = (unweighted_u[7:10, 0:5] * Td_u_norm_trans[2, 0])
weighted_u[7:10, 5:7] = (unweighted_u[7:10, 5:7] * Td_u_norm_trans[2, 1])
weighted_u[7:10, 7:10] = (unweighted_u[7:10, 7:10] * Td_u_norm_trans[2, 2])
weighted_u[7:10, 10:11] = (unweighted_u[7:10, 10:11] * Td_u_norm_trans[2, 3])
weighted_u[7:10, 11:15] = (unweighted_u[7:10, 11:15] * Td_u_norm_trans[2, 4])
weighted_u[7:10, 15:] = (unweighted_u[7:10, 15:] * Td_u_norm_trans[2, 5])

weighted_u[10:11, 0:5] = (unweighted_u[10:11, 0:5] * Td_u_norm_trans[3, 0])
weighted_u[10:11, 5:7] = (unweighted_u[10:11, 5:7] * Td_u_norm_trans[3, 1])
weighted_u[10:11, 7:10] = (unweighted_u[10:11, 7:10] * Td_u_norm_trans[3, 2])
weighted_u[10:11, 10:11] = (unweighted_u[10:11, 10:11] * Td_u_norm_trans[3, 3])
weighted_u[10:11, 11:15] = (unweighted_u[10:11, 11:15] * Td_u_norm_trans[3, 4])
weighted_u[10:11, 15:] = (unweighted_u[10:11, 15:] * Td_u_norm_trans[3, 5])

weighted_u[11:15, 0:5] = (unweighted_u[11:15, 0:5] * Td_u_norm_trans[4, 0])
weighted_u[11:15, 5:7] = (unweighted_u[11:15, 5:7] * Td_u_norm_trans[4, 1])
weighted_u[11:15, 7:10] = (unweighted_u[11:15, 7:10] * Td_u_norm_trans[4, 2])
weighted_u[11:15, 10:11] = (unweighted_u[11:15, 10:11] * Td_u_norm_trans[4, 3])
weighted_u[11:15, 11:15] = (unweighted_u[11:15, 11:15] * Td_u_norm_trans[4, 4])
weighted_u[11:15, 15:] = (unweighted_u[11:15, 15:] * Td_u_norm_trans[4, 5])

weighted_u[15:, 0:5] = (unweighted_u[15:, 0:5] * Td_u_norm_trans[5, 0])
weighted_u[15:, 5:7] = (unweighted_u[15:, 5:7] * Td_u_norm_trans[5, 1])
weighted_u[15:, 7:10] = (unweighted_u[15:, 7:10] * Td_u_norm_trans[5, 2])
weighted_u[15:, 10:11] = (unweighted_u[15:, 10:11] * Td_u_norm_trans[5, 3])
weighted_u[15:, 11:15] = (unweighted_u[15:, 11:15] * Td_u_norm_trans[5, 4])
weighted_u[15:, 15:] = (unweighted_u[15:, 15:] * Td_u_norm_trans[5, 5])

dff21 = pd.DataFrame(weighted_l, index=criteria_names, columns=criteria_names)
dff21.to_excel('weighted_l.xlsx')
dff22 = pd.DataFrame(weighted_m, index=criteria_names, columns=criteria_names)
dff22.to_excel('weighted_m.xlsx')
dff23 = pd.DataFrame(weighted_u, index=criteria_names, columns=criteria_names)
dff23.to_excel('weighted_u.xlsx')

# calculationg limit of weighted matrix to the power of infinity

limit_weighted_l = np.linalg.matrix_power(weighted_l, 7)
limit_weighted_m = np.linalg.matrix_power(weighted_m, 7)
limit_weighted_u = np.linalg.matrix_power(weighted_u, 7)

dff24 = pd.DataFrame(limit_weighted_l, index=criteria_names, columns=criteria_names)
dff24.to_excel('limit_weighted_l.xlsx')
dff25 = pd.DataFrame(limit_weighted_m, index=criteria_names, columns=criteria_names)
dff25.to_excel('limit_weighted_m.xlsx')
dff26 = pd.DataFrame(limit_weighted_u, index=criteria_names, columns=criteria_names)
dff26.to_excel('limit_weighted_u.xlsx')

defuzzy_lim_weighted = (((limit_weighted_u - limit_weighted_l) + (
            limit_weighted_m - limit_weighted_l)) / 3) + limit_weighted_l
header = ["Weights"]
dff26 = pd.DataFrame(defuzzy_lim_weighted[:, 0], index=criteria_names, columns=header)
dff26.to_excel('defuzzy_lim_weighted.xlsx')
