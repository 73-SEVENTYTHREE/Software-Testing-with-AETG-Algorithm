import process_excel as my_excel
import pandas as pd
import random
import numpy as np


def excel_to_list():
    my_list = []
    my_list.append(my_excel.read_excel(1))
    my_list.append(my_excel.read_excel(2))
    my_list.append(my_excel.read_excel(3))
    my_list.append(my_excel.read_excel(4))
    my_list.append(my_excel.read_excel(5))
    my_list.append(my_excel.read_excel(6))
    return my_list


# The function to generate the pair dictionary.
def create_2_way_dict(my_list):
    All_pairs = []
    for index in range(len(my_list) - 1):
        for index3 in range(index + 1, len(my_list)):
            for index1 in range(len(my_list[index])):
                for index2 in range(len(my_list[index3])):
                    pair = (my_list[index][index1], my_list[index3][index2])
                    All_pairs.append(pair)
    two_way_dict = {}
    for i in range(len(All_pairs)):
        two_way_dict[All_pairs[i]] = 0
    return two_way_dict


def check_dict(dict):
    Keys = list(dict.keys())
    for index in range(len(Keys)):
        if dict[Keys[index]] == 0:
            return 0
    return 1


def find_max(list, dict, result):
    Max_index = 0
    count_disappear = []
    if len(result) == 0:
        return random.choice(list)
    for index in range(len(list)):
        count_disappear.append(0)
    for index1 in range(len(list)):
        for index2 in range(len(result)):
            tuple1 = (list[index1], result[index2])
            tuple2 = (result[index2], list[index1])
            if dict.get(tuple1) is None:
                if dict[tuple2] == 0:
                    count_disappear[index1] += 1
            else:
                if dict[tuple1] == 0:
                    count_disappear[index1] += 1
    arr = np.array(count_disappear)
    if (arr == 0).all():
        return  random.choice(list)
    else:
        Max_index = count_disappear.index(max(count_disappear))
    return list[Max_index]


def choose_case(list, dict):
    result = []
    for index in range(6):
        result.append(find_max(list[index], dict, result))
    return result


def two_way_AETG():
    function_name = ("语言切换", "用户名", "密码", "项目名称", "项目简介", "项目类型", "Covering_pairs")
    df = pd.DataFrame(columns=function_name)
    Two_way_list = excel_to_list()
    Two_way_dict = create_2_way_dict(Two_way_list)
    count = 0
    Keys = list(Two_way_dict.keys())
    while 1:
        if check_dict(Two_way_dict) == 1:
            break
        count_pairs = 0
        testcase = choose_case(Two_way_list, Two_way_dict)
        print(testcase)
        for i in range(len(testcase) - 1):
            for j in range(i + 1, len(testcase)):
                new_tuple = (testcase[i], testcase[j])
                if Two_way_dict[new_tuple] == 0:
                    Two_way_dict[new_tuple] = 1
                    count_pairs += 1
        if count_pairs == 0:
            continue
        testcase.append(count_pairs)
        df.loc[count] = testcase
        count = count + 1
    df.to_excel('result1.xlsx')


if __name__ == '__main__':
    two_way_AETG()
