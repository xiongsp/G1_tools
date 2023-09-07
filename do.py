import xlrd
import decimal

data = xlrd.open_workbook("1.xlsx")
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols

target = """2021302463
2021302568
2021303096
2021303101
2021303106
2021303110
2021303112
2021303113
2021303114
2021303116
2021303123
2021303125
2021303129
2021303133
2021303134
2021303137
2021303140
2021303150
2021303151
2021303155
2021303157
2021303159
2021303162
2021303165
2021303167
2021303168
2021303169
2021303181
2021303182
2021303194
2021303401
2021303589
2021303823
2021303976
2021303979
2021303999"""


target = target.split("\n")

# 二维数组存储数据
all_data = []
for i in range(nrows):
    row_data = []
    if table.cell_value(i, 0) not in target:
        continue
    for j in range(ncols):
        row_data.append(table.cell_value(i, j))
    all_data.append(row_data)

# 初始化存储学分和成绩的字典
all_dict = {}
for i in target:
    all_dict[i] = [0, 0]


try:
    for i in all_data:
        if i[6] == "" and i[5] == "":  # 正常
            # if i[4]=='P' or i[4]=='NP' or i[4]=='优秀': # PNP课程不计算学分
            if not i[4].isdigit() and i[4] != "缺考" and i[4] != "缓考":  # PNP课程不计算学分
                print(i)
            elif i[4] == "缺考":  # 期末缺考0分
                print(i, 0)
                all_dict[i[0]][0] += decimal.Decimal(i[3])
                all_dict[i[0]][1] += 0
            elif i[4] == "缓考":  # 缓考暂未出成绩
                print(i)
            else:  # 正常
                print(i, decimal.Decimal(i[3]) * decimal.Decimal(i[4]))
                all_dict[i[0]][0] += decimal.Decimal(i[3])
                all_dict[i[0]][1] += decimal.Decimal(i[3]) * decimal.Decimal(i[4])
        elif i[4] == "缓考":  # 缓考已经出成绩
            print(i, decimal.Decimal(i[3]) * decimal.Decimal(i[6]))
            all_dict[i[0]][0] += decimal.Decimal(i[3])
            all_dict[i[0]][1] += decimal.Decimal(i[3]) * decimal.Decimal(i[6])
        elif i[5] == "缺考":  # 补考缺考
            print(i, decimal.Decimal(i[3]) * decimal.Decimal(i[4]))
            all_dict[i[0]][0] += decimal.Decimal(i[3])
            all_dict[i[0]][1] += decimal.Decimal(i[3]) * decimal.Decimal(i[4])
        elif i[5].isdigit():  # 补考已经出成绩
            if decimal.Decimal(i[5]) < 60:
                bigger = max(decimal.Decimal(i[4]), decimal.Decimal(i[5]))
                print(i, decimal.Decimal(i[3]) * bigger)
                all_dict[i[0]][0] += decimal.Decimal(i[3])
                all_dict[i[0]][1] += decimal.Decimal(i[3]) * bigger
            else:
                print(i, 60)
                all_dict[i[0]][0] += decimal.Decimal(i[3])
                all_dict[i[0]][1] += decimal.Decimal(i[3]) * 60
        else:  # 异常
            print(i)
            print("异常")
            exit(0)
except Exception as e:
    print(str(e), i)
    exit(0)


print(all_dict)

for i in all_dict:
    if all_dict[i][0] == 0:
        print(i, all_dict[i][0], all_dict[i][1])
    else:
        # print(i, (all_dict[i][1] / all_dict[i][0]))
        result = all_dict[i][1] / all_dict[i][0]
        # 四舍五入保留3位小数
        quantized_result = decimal.Decimal(result).quantize(decimal.Decimal("0.000"))
        print(i, quantized_result)
