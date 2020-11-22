import process_excel as my_excel
import pandas as pd
import random
import numpy as np


def excel_to_list(filename, number):
    my_list = []
    for index in range(number):
        my_list.append(my_excel.read_excel1(filename, index + 1))
    return my_list


# The function to generate the pair dictionary.
def create_2_way_dict(my_list):
    All_pairs = []
    for index1 in range(len(my_list) - 1):
        for index2 in range(index1 + 1, len(my_list)):
            for index3 in range(len(my_list[index1])):
                for index4 in range(len(my_list[index2])):
                    pair = (my_list[index1][index3], my_list[index2][index4])
                    All_pairs.append(pair)
    two_way_dict = {}
    for i in range(len(All_pairs)):
        two_way_dict[All_pairs[i]] = 0
    return two_way_dict


def create_3_way_dict(my_list):
    All_pairs = []
    for index1 in range(len(my_list) - 2):
        for index2 in range(index1 + 1, len(my_list) - 1):
            for index3 in range(index2 + 1, len(my_list)):
                for index4 in range(len(my_list[index1])):
                    for index5 in range(len(my_list[index2])):
                        for index6 in range(len(my_list[index3])):
                            pair = (my_list[index1][index4], my_list[index2][index5], my_list[index3][index6])
                            All_pairs.append(pair)
    three_way_dict = {}
    for i in range(len(All_pairs)):
        three_way_dict[All_pairs[i]] = 0
    print(len(three_way_dict))
    return three_way_dict


def check_dict(dict):
    Keys = list(dict.keys())
    for index in range(len(Keys)):
        if dict[Keys[index]] == 0:
            return 0
    return 1


def find_max1(candi_list, dict, result):
    Max_index = 0
    count_disappear = []
    Max_disappear = 0
    disappearance = {}
    Keys = list(dict.keys())
    for index in range(len(dict)):
        if dict[Keys[index]] == 0:
            if disappearance.get(Keys[index][0]):
                disappearance[Keys[index][0]] += 1
            else:
                disappearance[Keys[index][0]] = 1
            if disappearance.get(Keys[index][1]):
                disappearance[Keys[index][1]] += 1
            else:
                disappearance[Keys[index][1]] = 1
    flag = 0
    for index in range(len(candi_list)):
        if disappearance.get(candi_list[index]):
            flag = 1
            if disappearance[candi_list[index]] > Max_disappear:
                Max_disappear = disappearance[candi_list[index]]
                Max_index = index
    if flag == 0:
        return random.choice(candi_list)
    for index in range(len(candi_list)):
        count_disappear.append(0)
    for index1 in range(len(candi_list)):
        for index2 in range(len(result)):
            tuple1 = (candi_list[index1], result[index2])
            tuple2 = (result[index2], candi_list[index1])
            if dict.get(tuple1) == 0:
                count_disappear[index1] += 1
                continue
            if dict.get(tuple2) == 0:
                count_disappear[index1] += 1
    arr = np.array(count_disappear)
    print(arr)
    if (arr == 0).all():
        return candi_list[Max_index]
    else:
        Max_index = count_disappear.index(max(count_disappear))
    return candi_list[Max_index]


def find_max2(candi_list, dict, result):
    Max_index = 0
    count_disappear = []
    Max_disappear = 0
    disappearance = {}
    Keys = list(dict.keys())
    for index in range(len(dict)):
        if dict[Keys[index]] == 0:
            if disappearance.get((Keys[index][0], Keys[index][1])):
                disappearance[(Keys[index][0], Keys[index][1])] += 1
            else:
                disappearance[(Keys[index][0], Keys[index][1])] = 1
            if disappearance.get((Keys[index][0], Keys[index][2])):
                disappearance[(Keys[index][0], Keys[index][2])] += 1
            else:
                disappearance[(Keys[index][0], Keys[index][2])] = 1
            if disappearance.get((Keys[index][2], Keys[index][1])):
                disappearance[(Keys[index][2], Keys[index][1])] += 1
            else:
                disappearance[(Keys[index][2], Keys[index][1])] = 1
    flag = 0
    for index in range(len(candi_list)):
        for index1 in range(len(result)):
            if disappearance.get((candi_list[index], result[index1])):
                flag = 1
                if disappearance[(candi_list[index], result[index1])] > Max_disappear:
                    Max_disappear = disappearance[(candi_list[index], result[index1])]
                    Max_index = index
            if disappearance.get((result[index1], candi_list[index])):
                flag = 1
                if disappearance[(result[index1], candi_list[index])] > Max_disappear:
                    Max_disappear = disappearance[(result[index1], candi_list[index])]
                    Max_index = index
    if flag == 0:
        return random.choice(candi_list)
    for index in range(len(candi_list)):
        count_disappear.append(0)
    for index1 in range(len(candi_list)):
        for index2 in range(len(result) - 1):
            for index3 in range(index2 + 1, len(result)):
                tuple1 = (candi_list[index1], result[index2], result[index3])
                tuple2 = (candi_list[index1], result[index3], result[index2])
                tuple3 = (result[index3], candi_list[index1], result[index2])
                tuple4 = (result[index3], result[index2], candi_list[index1])
                tuple5 = (result[index2], candi_list[index1], result[index3])
                tuple6 = (result[index2], result[index3], candi_list[index1])
                if dict.get(tuple1) == 0:
                    count_disappear[index1] += 1
                    continue
                if dict.get(tuple2) == 0:
                    count_disappear[index1] += 1
                    continue
                if dict.get(tuple3) == 0:
                    count_disappear[index1] += 1
                    continue
                if dict.get(tuple4) == 0:
                    count_disappear[index1] += 1
                    continue
                if dict.get(tuple5) == 0:
                    count_disappear[index1] += 1
                    continue
                if dict.get(tuple6) == 0:
                    count_disappear[index1] += 1
    arr = np.array(count_disappear)
    if (arr == 0).all():
        return candi_list[Max_index]
    else:
        Max_index = count_disappear.index(max(count_disappear))
    return candi_list[Max_index]


def choose_case1(list, dict, number):
    result = []
    for index in range(number):
        result.append(find_max1(list[index], dict, result))
    return result


def choose_case2(list, dict, number):
    result = []
    for index in range(number):
        result.append(find_max2(list[index], dict, result))
    return result


def two_way_AETG():
    function_name = ("语言切换", "用户名", "密码", "项目名称", "项目简介", "项目类型", "Covering_pairs")
    df = pd.DataFrame(columns=function_name)
    Two_way_list = excel_to_list("test1.xlsx", 6)
    Two_way_dict = create_2_way_dict(Two_way_list)
    count = 1
    Keys = list(Two_way_dict.keys())
    while 1:
        if check_dict(Two_way_dict) == 1:
            break
        count_pairs = 0
        testcase = choose_case1(Two_way_list, Two_way_dict, 6)
        print(testcase)
        for i in range(len(testcase) - 1):
            for j in range(i + 1, len(testcase)):
                new_tuple1 = (testcase[i], testcase[j])
                new_tuple2 = (testcase[j], testcase[i])
                if Two_way_dict.get(new_tuple1) == 0:
                    Two_way_dict[new_tuple1] = 1
                    count_pairs += 1
                if Two_way_dict.get(new_tuple2) == 0:
                    Two_way_dict[new_tuple2] = 1
                    count_pairs += 1
        testcase.append(count_pairs)
        df.loc[count] = testcase
        count = count + 1
    df.to_excel('result1.xlsx')


def three_way_AETG():
    function_name = ("品牌", "屏幕尺寸", "能效等级", "分辨率", "电视类型", "品牌类型", "观看距离", "用户优选", "产品渠道", "Covering_pairs")
    df = pd.DataFrame(columns=function_name)
    Three_way_list = excel_to_list("test2.xlsx", 9)
    Three_way_dict = create_3_way_dict(Three_way_list)
    count = 1
    Keys = list(Three_way_dict.keys())
    while 1:
        if check_dict(Three_way_dict) == 1:
            break
        count_pairs = 0
        testcase = choose_case2(Three_way_list, Three_way_dict, 9)
        # print(testcase)
        for i in range(len(testcase) - 2):
            for j in range(i + 1, len(testcase) - 1):
                for k in range(j + 1, len(testcase)):
                    new_tuple1 = (testcase[i], testcase[j], testcase[k])
                    new_tuple2 = (testcase[i], testcase[k], testcase[j])
                    new_tuple3 = (testcase[j], testcase[k], testcase[i])
                    new_tuple4 = (testcase[j], testcase[i], testcase[k])
                    new_tuple5 = (testcase[k], testcase[j], testcase[i])
                    new_tuple6 = (testcase[k], testcase[i], testcase[j])
                    if Three_way_dict.get(new_tuple1) == 0:
                        Three_way_dict[new_tuple1] = 1
                        count_pairs += 1
                    if Three_way_dict.get(new_tuple2) == 0:
                        Three_way_dict[new_tuple2] = 1
                        count_pairs += 1
                    if Three_way_dict.get(new_tuple3) == 0:
                        Three_way_dict[new_tuple3] = 1
                        count_pairs += 1
                    if Three_way_dict.get(new_tuple4) == 0:
                        Three_way_dict[new_tuple4] = 1
                        count_pairs += 1
                    if Three_way_dict.get(new_tuple5) == 0:
                        Three_way_dict[new_tuple5] = 1
                        count_pairs += 1
                    if Three_way_dict.get(new_tuple6) == 0:
                        Three_way_dict[new_tuple6] = 1
                        count_pairs += 1
        if count_pairs == 0:
            continue
        testcase.append(count_pairs)
        print(testcase)
        df.loc[count] = testcase
        count = count + 1
    df.to_excel('result2.xlsx')


if __name__ == '__main__':
    two_way_AETG()
    three_way_AETG()
