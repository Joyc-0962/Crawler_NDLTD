from time import sleep
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import time
import os
import xlrd
import csv
from collections import defaultdict

def open_csv(filename):
    professor_name_list=[]
    workbook = xlrd.open_workbook(filename)
    # 获取所有工作表的名称
    sheet_names = workbook.sheet_names()
    print("工作表名称列表:", sheet_names)
    # 取得第一個工作表
    for i in range(len(sheet_names)):
        sheet = workbook.sheet_by_name(sheet_names[i])
        # 顯示 row總數 及 column總數
        print(f'第{i}個工作表')
        print('row總數:', sheet.nrows-1)
        print('column總數:', sheet.ncols)
        # 顯示 cell 資料
        for j in range(1, sheet.nrows):
            professor_name_list.append(sheet.cell(j,1).value)
    print("總共有 "+str(len(professor_name_list))+" 個教授")
    print("總共不重複的有 "+str(len(set(professor_name_list)))+" 個教授")
    return list(set(professor_name_list))

def open_csv_dict(filename):
    professor_info_dict = defaultdict(list)
    professor_school_count = defaultdict(int)
    workbook = xlrd.open_workbook(filename)
    # 获取所有工作表的名称
    sheet_names = workbook.sheet_names()
    print("工作表名称列表:", sheet_names)
    # 取得第一個工作表
    for i in range(len(sheet_names)):
        sheet = workbook.sheet_by_name(sheet_names[i])
        # 顯示 row總數 及 column總數
        print(f'第{i}個工作表')
        print('row總數:', sheet.nrows-1)
        print('column總數:', sheet.ncols)
        # 顯示 cell 資料
        for j in range(1, sheet.nrows):
            professor_name = sheet.cell(j, 1).value
            school_name = sheet.cell(j, 3).value
            professor_school_count[(professor_name, school_name)] += 1

    # 将数据构建成列表[professor, school, count]
    professor_school_count_list = [[professor, school, count] for (professor, school), count in professor_school_count.items()]
    # 對學校以及出現次數做成一個list
    for professor, school, count in professor_school_count_list:
        professor_info_dict[professor].append((school, count))

    # for professor, schools_counts in professor_info_dict.items():
    #     if professor == professor_name:
    #         max_school, max_count = max(schools_counts, key=lambda x: x[1])
            # print(f"教授：{professor}，學校：{max_school}，出現次數：{max_count}")
        
    # print(professor_info_dict)
    print("總共有 " + str(len(professor_info_dict)) + " 個教授")
    return professor_info_dict

def print_most_frequent_school(schools_counts):
    max_school, max_count = max(schools_counts, key=lambda x: x[1])
    return max_school

def write_to_csv(df,filename):
    # 取得目前工作目錄路徑
    path = os.getcwd()
    filepath = os.path.join(path,filename)
    # df.to_excel(filepath, sheet_name=professor_name, engine='xlsxwriter')
    # 将DataFrame追加到现有CSV文件中（如果文件不存在，则创建新文件）
    # 如果文件不存在，添加标题
    if not os.path.exists(filepath):
        df.to_csv(filepath, header=True, index=False, encoding='utf-8-sig')
    else:
        # 否则追加到现有CSV文件中
        df.to_csv(filepath, mode='a', header=False, index=False, encoding='utf-8-sig')

def read_name_csv(filename):
    # 如果文件不存在，创建一个新的空CSV文件并写入标题
    if not os.path.exists(filename):
        with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(['Name'])  # 写入标题行

    with open(filename, mode='r', newline='') as file:
        reader = csv.reader(file)
        existing_names = [row[0] for row in reader]            
    return existing_names

def write_name_csv(filename, name):
    with open(filename, mode='a', newline='') as file:
        # 创建CSV写入器
        writer = csv.writer(file)
        # 写入名字
        writer.writerow([name])
