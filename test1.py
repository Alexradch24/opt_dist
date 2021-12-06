import pandas as pd
import copy
import time as tm
len_baze = 0
len_break = 0
bar = 0
bar_all = 0
def time(ind, start_ind, fuel, fuel_baze, fuel_break, time_i, t_o_f_f, way):
    for i in range(len_break - 1, -1, -1):
        if fuel_break[i][0] > fuel_baze[ind][1]:
            continue
        if ind == len_baze - 1: 
            return fuel_baze
        dist = fuel_baze[ind + 1][0] - fuel_baze[ind][0]
        buff = fuel - dist * fuel_break[i][1] / 100
        if buff < 0: 
            continue
        t = (dist / fuel_break[i][0]) * 60
        way.append([fuel_baze[ind][0], fuel_break[i][0], fuel_baze[ind][9], fuel_baze[ind - 1][1]])
        if fuel_baze[ind + 1][3] == 0 or fuel_baze[ind + 1][3] > time_i + t:
            fuel_baze[ind + 1][3] = time_i + t
            fuel_baze[ind + 1][4] = start_ind
            fuel_baze[ind + 1][5] = fuel_break[i][0]
            fuel_baze[ind + 1][6] = buff
            fuel_baze[ind + 1][7] = copy.copy(way)
            flag = False
        fuel_baze = time(ind + 1, start_ind, buff, fuel_baze, fuel_break, time_i + t, t_o_f_f, way)
        way.pop()
    return fuel_baze
print('Enter File Name:')
file_name = input()
print('Enter Sheet Name (default "Var"):')
m_sheet_name = input()
if m_sheet_name == '':
    m_sheet_name = 'Var'
test_pd = pd.read_excel(file_name, sheet_name = m_sheet_name)
print('ok')
all_dist = 0
all_fuel = test_pd['Tank volume (l)'][0]
way = []
time_of_full_fuel = test_pd['Refueling time (min)'][0]
fuel_baze = [] #[dist, max, baze, time, baze start, speed, fuel, way, name_b, name_u]
fuel_break = [] # [speed, break]
test_pd = test_pd.replace({'speed (km/h) (c)':{0: None}})
for i in range(len(test_pd['speed (km/h) (c)'].dropna())):
    fuel_break.append([test_pd['speed (km/h) (c)'][i], test_pd['consumption (l/100km)'][i]])
fuel_break.sort()
fuel_baze.append([0, 0, 0, 0, -1, 0, 0, 0, 'NaN', 'Start'])
test_pd = test_pd.replace({'plot':{0: None}})
for i in range(len(test_pd['plot'].dropna())):
    fuel_baze.append([0, 0, 0, 0, -1, 0, 0, 0, 0, 0])
    fuel_baze[i + 1][0] = test_pd['km from 0 (p)'][i]
    fuel_baze[i][1] = test_pd['speed (km/h) (p)'][i]
    fuel_baze[i + 1][8] = ' '
    fuel_baze[i][9] = test_pd['plot'][i]
buff_len = len(fuel_baze)
test_pd = test_pd.replace({'refueling':{0: None}})
for i in range(len(test_pd['refueling'].dropna())):
    fuel_baze.append([0, 0, 1, 0, -1, 0, 0, 0, 0, 0])
    fuel_baze[buff_len + i][0] = test_pd['km from 0 (r)'][i]
    fuel_baze[buff_len + i][8] = test_pd['refueling'][i]
fuel_baze.sort()
fuel_baze[-1][9] = 'Finish'
for i in range(len(fuel_baze)):
    if fuel_baze[i][2] == 1:
        fuel_baze[i][1] = fuel_baze[i - 1][1]
        fuel_baze[i][9] = fuel_baze[i - 1][9]
i = 0
while i < len(fuel_baze):
    if i + 1 != len(fuel_baze) and fuel_baze[i][0] == fuel_baze[i + 1][0]:
        fuel_baze[i][1] = fuel_baze[i + 1][1]
        fuel_baze[i][9] = fuel_baze[i + 1][9]
        fuel_baze = fuel_baze[:i + 1] + fuel_baze[i + 2:]
        continue
    i += 1
fuel_baze[0][6] = test_pd['In the tank at the start (l)'][0]
len_baze = len(fuel_baze)
len_break = len(fuel_break)
for i in range(len(fuel_baze)):
    if i == 0:
        fuel_baze = time(i, i, fuel_baze[i][6], fuel_baze, fuel_break, 0, time_of_full_fuel, way)
    elif fuel_baze[i][2] == 1:
        fuel_baze = time(i, i, all_fuel, fuel_baze, fuel_break, fuel_baze[i][3] + time_of_full_fuel, time_of_full_fuel, way)
    elif fuel_baze[i][2] == 0:
        fuel_baze = time(i, i, fuel_baze[i][6], fuel_baze, fuel_break, fuel_baze[i][3], time_of_full_fuel, way)  
k = len(fuel_baze) - 1
answer = [[] for i in range(11)]
for i in range(len(fuel_baze) - 1, -1, -1):
    if i == k and i != 0:
        answer[1].append(fuel_baze[i][0])
        answer[0].append(fuel_baze[i][7][-1][0])
        int_time_all = fuel_baze[i][3]
        answer[2].append(int_time_all)
        answer[3].append(fuel_baze[i][7][-1][2])
        answer[4].append(fuel_baze[i][7][-1][1])
        answer[5].append(0)
        answer[6].append(fuel_baze[i][8])
        answer[7].append(fuel_baze[i][0] - fuel_baze[fuel_baze[i][4]][0])
        answer[8].append(fuel_baze[i][6])
        answer[10].append(fuel_baze[i - 1][1])
        if fuel_baze[i][8] != ' ':
            answer[9].append(all_fuel - fuel_baze[i][6])
        for j in range(len(fuel_baze[i][7]) - 1, 0, -1):
            answer[1].append(fuel_baze[i][7][j][0])
            answer[0].append(fuel_baze[i][7][j - 1][0])
            answer[2].append(0.0)
            answer[3].append(fuel_baze[i][7][j - 1][2])
            answer[4].append(fuel_baze[i][7][j - 1][1])
            answer[5].append(0)
            answer[6].append(' ')
            answer[7].append(fuel_baze[i][7][j][0] - fuel_baze[fuel_baze[i][4]][0])
            answer[8].append('-')
            answer[9].append(' ')
            answer[10].append(fuel_baze[i][7][j][3])
        k = fuel_baze[i][4]
for i in answer:
    i.reverse()
i = 0
while i < len(answer[0]):
    if i + 1 != len(answer[0]):
        if (answer[4][i] == answer[4][i + 1]) and (answer[6][i] == ' '):
            answer[0] = answer[0][0:i + 1] + answer[0][i + 2:]
            for j in range(1, 11 ):
                answer[j] = answer[j][0:i] + answer[j][i + 1:]
            continue
    i += 1
ras = {}
for i in range(len(fuel_break)):
    ras[fuel_break[i][0]] = fuel_break[i][1]      
for i in range(len(answer[0])):
    answer[5][i] = (answer[1][i] - answer[0][i]) / answer[4][i]
    if answer[8][i] == '-':
        if i == 0:
            answer[8][i] =test_pd['In the tank at the start (l)'][0] - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
        elif answer[6][i - 1] != ' ':
            answer[8][i] = all_fuel - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
        else:
            answer[8][i] = answer[8][i - 1] - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
buff_oil = answer[8][-1]
i = len(answer[0]) - 1
while i > -1 :
    if answer[6][i] != ' ':
        buff_oil = answer[8][i]
    if answer[4][i] == answer[10][i]:
        i -= 1
        continue
    else:
        time_b = answer[5][i]
        dist_b = 0
        buff_time = 0
        buff_speed = 0
        flag = False
        for j in range(len(fuel_break)):
            if fuel_break[j][0] <= answer[4][i]: continue
            if fuel_break[j][0] > answer[10][i]: continue
            buff_dist = (100 * (buff_oil + ras[answer[4][i]]/100 * (answer[1][i] - answer[0][i])) - ras[answer[4][i]] * (answer[1][i] - answer[0][i])) / (fuel_break[j][1] - ras[answer[4][i]])
            if buff_dist < 2: continue
            buff_time = buff_dist / fuel_break[j][0] + (answer[1][i] - answer[0][i] - buff_dist) / fuel_break[j][0]
            if buff_time < time_b:
                time_b = buff_time
                dist_b = buff_dist
                buff_speed = fuel_break[j][0]
                flag = True
        if flag:
            answer[0] = answer[0][:i + 1] + [answer[0][i] + dist_b] + answer[0][i + 1:]
            answer[1] = answer[1][:i] + [answer[0][i] + dist_b] + answer[1][i:]
            answer[3] = answer[3][:i + 1] + [answer[3][i]] + answer[3][i + 1:]
            answer[4] = answer[4][:i] + [buff_speed] + answer[4][i:]
            answer[6] = answer[6][:i] +[' '] + answer[6][i:]
            buff_oil = 0
        i -= 1
for i in range(len(answer[0])):
    answer[0][i] = int(answer[0][i])
    answer[1][i] = int(answer[1][i])
answer[2] = [' ' for i in range(len(answer[0]))] 
answer[5] = [' ' for i in range(len(answer[0]))]
answer[7] = [' ' for i in range(len(answer[0]))]
answer[8] = [' ' for i in range(len(answer[0]))]
answer[9] = [' ' for i in range(len(answer[0]))]
for i in range(len(answer[0])):
    answer[5][i] = (answer[1][i] - answer[0][i]) / answer[4][i]
    if i == 0: answer[7][i] = answer[1][i]
    elif answer[6][i - 1] != ' ': answer[7][i] = (answer[1][i] - answer[0][i])
    else: answer[7][i] = answer[7][i - 1] + (answer[1][i] - answer[0][i])
    if i == 0: answer[2][i] = answer[5][i] * 60
    else: 
        answer[2][i] = answer[2][i - 1] + answer[5][i] * 60
        if answer[6][i] != ' ': answer[2][i] += time_of_full_fuel
    if i == 0: answer[8][i] = test_pd['In the tank at the start (l)'][0] - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
    elif answer[6][i - 1] != ' ': answer[8][i] = all_fuel - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
    else: answer[8][i] = answer[8][i - 1] - ras[answer[4][i]] / 100 * (answer[1][i] - answer[0][i])
    if answer[6][i] != ' ': answer[9][i] = (all_fuel - answer[8][i])
for i in range(len(answer[0])):
    answer[2][i] = round(answer[2][i])
    answer[2][i] = str(answer[2][i] // 60 // 10) + str(answer[2][i] // 60 % 10) + ':' + str(answer[2][i] % 60 // 10) + str(answer[2][i] % 60 % 10)
    answer[5][i] = round(answer[5][i], 2)
    answer[8][i] = round(answer[8][i], 1)
answer_pd = pd.DataFrame()
buff_ser = pd.Series(answer[0])
answer_pd['from (km)'] = buff_ser
buff_ser = pd.Series(answer[1])
answer_pd['before (km)'] = buff_ser
buff_ser = pd.Series(answer[2])
answer_pd['time of all'] = buff_ser
buff_ser = pd.Series(answer[3])
answer_pd['plot'] = buff_ser
buff_ser = pd.Series(answer[4])
answer_pd['speed (km/h)'] = buff_ser
buff_ser = pd.Series(answer[5])
answer_pd['segment time'] = buff_ser
buff_ser = pd.Series(answer[6])
answer_pd['refueling'] = buff_ser
buff_ser = pd.Series(answer[7])
answer_pd['from the last gas station (km)'] = buff_ser
buff_ser = pd.Series(answer[8])
answer_pd['in the tank'] = buff_ser
buff_ser = pd.Series(answer[9])
answer_pd['refuel'] = buff_ser
buff_ser = pd.Series([all_fuel])
answer_pd.index = [i for i in range(1, len(answer[0]) + 1)]
if file_name.count('.xlsx') > 0: file_ans = file_name[:file_name.index('.xlsx')] + '_ans.xlsx'
else: file_ans = file_name[:file_name.index('.')] + '_ans.xlsx'
writer = pd.ExcelWriter(file_ans)
answer_pd.to_excel(writer, 'Moto_answ')
writer.save()
print('Press Enter for end...')
a = input()