from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import selenium.webdriver.support.ui as ui
from time import sleep
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import time
from tool import *
import logging
#設定 log 的 filename 後只會輸出到檔案不會輸出在 console
logging.basicConfig(level=logging.INFO, filename='./log_nstc.txt', filemode='a',
    format='[%(asctime)s %(levelname)s] %(message)s',
    datefmt='%Y%m%d %H:%M:%S',
    )

def crawler_NSTC(professor_name,professor_school):
    fail_flag=0
    while True:
        try:
            fail_flag+=1
            start = time.time()
            chrome_options = Options()
            # chrome_options.add_argument('--headless=new')
            # 一開始只使用 professor_name 搜尋
            driver = webdriver.Chrome(options=chrome_options)
            driver.get('https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initSearchList&LANG=ch')
            driver.find_element(By.NAME,"nameChi").send_keys(professor_name)
            element = driver.find_element(By.NAME,"send")
            element.click()
            # 找到第一個點下去
            element = driver.find_element(By.XPATH,'//*[@id="zone.content11"]/div/div[2]/table/tbody/tr[2]/td[1]/center/a')
            element.click()
            break
        except:
            try: 
                # 獲取元素中的文本
                data_count_element = driver.find_element(By.CSS_SELECTOR,"div.page span em")
                data_count_text = data_count_element.text
                if int(data_count_text) > 1:
                    # 需要輸入 professor_school
                    # 找到select元素
                    select_element = driver.find_element(By.XPATH,'//*[@id="form1"]/div[1]/table/tbody/tr[3]/td/table/tbody/tr[5]/td/input[1]')
                    select_element.click()
                    # 找到radio元素
                    driver.find_element(By.XPATH,'//*[@id="organDesc"]').send_keys(professor_school)
                element = driver.find_element(By.NAME,"send")
                element.click()
                # 找到第一個點下去
                element = driver.find_element(By.XPATH,'//*[@id="zone.content11"]/div/div[2]/table/tbody/tr[2]/td[1]/center/a')
                element.click()
                break
            except:# 如果搜尋不到該教授
                if fail_flag > 10:
                    filename = 'nstc_blank_professor_title.csv'
                    write_name_csv(filename,professor_name)
                    print("喔喔～沒有這個教授喔")
                    return
            

    element = driver.find_element(By.CLASS_NAME, 'c30Tblist')
    soup = BeautifulSoup(element.get_attribute('innerHTML'), 'html.parser')
    info_rows = soup.find_all('tr')
    professor_info = {}
    for row in info_rows:
        th = row.find('th')
        td = row.find('td')
        if th and td:
            key = th.text.strip()  # 去除首尾空格
            value = td.text.strip()
            # 去除分隔符和换行符号
            value = value.replace('\xa0', '').replace('\n', '').replace('\r', '').strip()
            professor_info[key] = value
    # print(professor_info)
    # 将项目信息列表转换为 DataFrame
    df_professor_info = pd.DataFrame([professor_info])
    filename = "nstc_professor_info.csv"
    write_to_csv(df_professor_info,filename)

    # 获取当前网页的URL
    current_url = driver.current_url
    # 将当前URL更改为第二个URL
    sec_url = current_url.replace("initBasic", "initRsm05")
    driver.get(sec_url)
    work_fail=0
    while True:
        try:
            work_fail+=1
            element = driver.find_element(By.CLASS_NAME, 'c30Tblist2')
            soup = BeautifulSoup(element.get_attribute('innerHTML'), 'html.parser')
            works_list = []
            # 遍历表格中的每一行（除了表头）
            for row in soup.find_all('tr')[1:]:
                cells = row.find_all('td')
                # 提取出版年月、著作类别、著作名称、作者和收录出处，并去除首尾空格
                published_date = cells[0].text.strip()
                work_type = cells[1].text.strip()
                work_title = cells[2].text.strip().replace('\n', '').replace('\t', '')
                authors = cells[3].text.strip().replace('\n', '').replace('\t', '')
                source = cells[4].text.strip()
                work_info = {
                    "教授名稱":professor_name,
                    "出版年月": published_date,
                    "著作類別": work_type,
                    "著作名稱": work_title,
                    "作者": authors,
                    "收錄出處": source
                }
                works_list.append(work_info)
            # print("著作目錄：",len(works_list))
            # print(works_list[:5])
            df_works_list = pd.DataFrame(works_list)
            filename = "nstc_work_list.csv"
            write_to_csv(df_works_list,filename)
            break
        except NoSuchElementException:
            # 如果找不到element = driver.find_element(By.CLASS_NAME, 'c30Tblist2')
            print("找不到指定的元素")
            logging.info(professor_name+" 沒有著作~~")
            break
        except:
            if work_fail > 10:
                print("著作部分怪怪的")
                logging.info(professor_name+" 著作部分怪怪的~~")
                break

    # 将当前URL更改为第三个URL
    third_url = sec_url.replace("initRsm05","initRsm17new")
    driver.get(third_url)
    project_fail=0
    while True:
        try:
            project_fail+=1
            element = driver.find_element(By.CLASS_NAME, 'c30Tblist2')
            soup = BeautifulSoup(element.get_attribute('innerHTML'), 'html.parser')
            projects_list = []
            # 遍历表格中的每一行（除了表头）
            for row in soup.find_all('tr')[1:]:
                cells = row.find_all('td')
                # 提取年度、補助類別、學門代碼、計畫名稱、擔任工作和核定經費，并去除首尾空格
                year = cells[0].text.strip()
                grant_type = cells[1].text.strip().replace('\n', '').replace('\t', '')
                grant_type = ' '.join(grant_type.split()) # 删除额外空格
                field_code = cells[2].text.strip()
                project_name = cells[3].text.strip().replace('\n', '').replace('\t', '')
                role = cells[4].text.strip().replace('\n', '').replace('\t', '')
                approved_fund = cells[5].text.strip()
                # 将提取的信息组成一个字典，并添加到作品列表中
                project_info = {
                    "教授名稱":professor_name,
                    "年度": year,
                    "補助類別": grant_type,
                    "學門代碼": field_code,
                    "計畫名稱": project_name,
                    "擔任工作": role,
                    "核定經費(新台幣)": approved_fund
                }
                projects_list.append(project_info)
            # print("計畫總覽：",len(projects_list))
            # print(projects_list[:5])

            df_projects_list = pd.DataFrame(projects_list)
            filename = "nstc_projects_list.csv"
            write_to_csv(df_projects_list,filename)
            break
        except NoSuchElementException:
            # 如果找不到element = driver.find_element(By.CLASS_NAME, 'c30Tblist2')
            print("找不到指定的元素")
            logging.info(professor_name+" 沒有計畫~~")
            break
        except:
            if project_fail > 10:
                print("計畫部分怪怪的")
                logging.info(professor_name+" 計畫部分怪怪的~~")
                break

    # 寫入成功爬蟲的professor_name
    filename = 'nstc_done_professor_title.csv'
    write_name_csv(filename,professor_name)
    end = time.time()
    logging.info("save to "+professor_name+".xlsx spend "+str(round((end - start)/60,4))+" 分鐘")
    

if __name__=="__main__":
    filename = '自動化學門及控制學門計畫申請案.xls'
    # professor_name = open_csv(filename)
    professor_info_dict = open_csv_dict(filename)
    done_filename = 'nstc_done_professor_title.csv'
    blank_filename = 'nstc_blank_professor_title.csv'
    count = 0
    for i, (name, schools) in enumerate(professor_info_dict.items()):
        print("現在是第"+str(i)+"個教授 "+name)
        done_list = read_name_csv(done_filename)
        blank_list = read_name_csv(blank_filename)
        if name in done_list:
            continue
        else:
            if name in blank_list:
                continue
            else:
                if count>50:
                    sleep(60*5)
                    count=0
                max_school = print_most_frequent_school(schools)
                print(max_school)
                crawler_NSTC(name,max_school)
                count+=1