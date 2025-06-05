# -*- coding: utf-8 -*-

"""
File: WOS_spider.py
Author: Chream
Email: yaom7917@gmail.com
Date: 2025-5-19
Version: 1.0

本脚本诞生于自动控制原理课程作业
描述: 该脚本使用Selenium和BeautifulSoup从Web of Science (WOS)网站抓取详细的论文信息。
它浏览每篇论文的详情页面，提取关键信息如标题、引用次数、国家、期刊等，
并将收集的数据保存到CSV文件中。

请注意，此脚本仅用于教育目的，使用本脚本或任何衍生作品时，
应遵守Web of Science的服务条款和使用政策。

Readme:
    1. 确保已安装所需的Python库，如Selenium、BeautifulSoup和pandas。
    2. 替换url_root为要爬取的WOS论文列表页面的URL。
    3. 调整papers_need和file_path以满足您的需求。
    4. 确保已安装Chrome浏览器和相应的WebDriver。
    5. 运行脚本后，它将自动打开Chrome浏览器并开始爬取数据。
    6. 爬取完成后，数据将保存到指定的CSV文件中。（当继续使用时，程序会提示是否读取csv，选择“是”，即可继续爬取）
    7. 脚本会在爬取过程中检测键盘输入，按下0/方向下键可以手动结束程序。
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import time, keyboard
import winsound

# 在import部分确保已导入
import sys
import keyboard
import winsound


# 解析html
def parse_html(html, driver=None):
    soup = BeautifulSoup(html, 'html.parser')
    
    # 创建一个空的字典
    data_dict = {}
    
    try:
        containers = soup.find_all('div', class_='cdx-two-column-grid-container')
        for container in containers:
            # 在这个容器内找到所有的标签和数据
            labels = container.find_all(class_='cdx-grid-label')
            datas = container.find_all(class_='cdx-grid-data')
            label = labels[0].text.strip()
            data_texts = [data.text.strip() for data in datas] # 提取数据列表中的文本
            text = '\n'.join(data_texts) # 将文本连接成一个字符串，使用换行符分隔
            
            # 存储到字典中
            data_dict[label] = text
    except:
        print("解析容器失败")
    
    try:
        # 首先尝试使用原有的class方式获取标题
        class_title = soup.find(class_="title text--large cdx-title")
        if class_title:
            data_dict['title'] = class_title.text.strip()
            print('\t'+class_title.text.strip())
        else:
            # 如果失败，尝试使用用户提供的XPath获取标题
            if driver:
                try:
                    title_xpath = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[1]/div/div/div/h2'
                    title_element = driver.find_element(By.XPATH, title_xpath)
                    if title_element:
                        data_dict['title'] = title_element.text.strip()
                        print('\t通过XPath获取标题：'+title_element.text.strip())
                    else:
                        print("通过XPath获取标题失败")
                        data_dict['title'] = "未获取到标题"
                except Exception as e:
                    print(f"通过XPath获取标题失败：{str(e)}")
                    data_dict['title'] = "未获取到标题"
            else:
                print("获取标题失败：driver未提供")
                data_dict['title'] = "未获取到标题"
    except Exception as e:
        print(f"获取标题失败：{str(e)}")
        data_dict['title'] = "未获取到标题"

    # 添加发表年份抓取逻辑
    try:
        if driver:
            try:
                year_xpath = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[7]/div[2]/span'
                year_element = driver.find_element(By.XPATH, year_xpath)
                if year_element:
                    data_dict['publication_year'] = year_element.text.strip()
                    print('\t发表年份：'+year_element.text.strip())
                else:
                    print("获取发表年份失败")
                    data_dict['publication_year'] = "未获取到发表年份"
            except Exception as e:
                print(f"通过XPath获取发表年份失败：{str(e)}")
                data_dict['publication_year'] = "未获取到发表年份"
        else:
            print("获取发表年份失败：driver未提供")
            data_dict['publication_year'] = "未获取到发表年份"
    except Exception as e:
        print(f"获取发表年份失败：{str(e)}")
        data_dict['publication_year'] = "未获取到发表年份"

    try:
        class_citation = soup.find(class_="mat-tooltip-trigger medium-link-number link ng-star-inserted")
        data_dict['citation'] = class_citation.text.strip()
    except:
        data_dict['citation'] = '0'

    try:
        class_addresses = soup.find('span', class_='ng-star-inserted', id='FRAOrgTa-RepAddressFull-0')
        print('\t'+class_addresses.text.strip())
        data_dict['country'] = class_addresses.text.split(',')[-1].strip()
    except:
        try:
            class_addresses = soup.find('span', class_='value padding-right-5--reversible')
            print('\t查询规则2：'+class_addresses.text.strip())
            data_dict['country'] = class_addresses.text.split(',')[-1].strip()
        except:
            print("获取国家失败")
    
    try:
        class_journal = soup.find(class_="mat-focus-indicator mat-tooltip-trigger font-size-14 summary-source-title-link remove-space no-left-padding mat-button mat-button-base mat-primary font-size-16 ng-star-inserted")
        data_dict['journal'] = class_journal.text.strip()
    except:
        print("获取期刊失败")
    
    try:
        input_box = soup.find(class_='wos-input-underline page-box')  # 获取包含输入框的标签
        index = int(input_box['aria-label'].split()[-1].replace(",", ""))
    except:
        print("获取页码失败")
        
    return index, data_dict    
        

if __name__ == "__main__": 
    # 0000391627
    # adolescent depression 1: https://webofscience-clarivate-cn-s.era.lib.swjtu.edu.cn/wos/alldb/full-record/WOS:000653016400005 
    url_root = 'https://www.webofscience.com/wos/woscc/summary/1441e534-ad57-4f71-a279-b047e2bc331d-0160969cfb/relevance/1'
    papers_need = 100000
    file_path = 'result.csv'    
    wait_time = 5
    pause_time = 0.1
    
    # 变量
    judge_xpath = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/app-full-record-mini-crl/div/div/app-records-list/app-record[1]/div/div/div[1]/span'
    xpath_nextpaper = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[1]/app-page-controls/div/form/div/button[2]'
    df = pd.DataFrame()
    index = 0
    duration = 5000  # 提示音时间 millisecond
    freq = 440  # 提示音Hz
    flag = 0
    
    # 读取df
    ifread = input("是否读取已有的CSV文件？(y/n)")
    if ifread == 'y':
        df = pd.read_csv(file_path, index_col=0)
        index = int(df.index[-1].split('_')[-1])
        print(f"读取已有的CSV文件，当前行索引为{index},即第{index+1}篇论文")
    
    # 创建ChromeOptions对象
    chrome_options = webdriver.ChromeOptions()

    # 禁止加载图片等资源
    chrome_options.add_argument("--disable-images")
    chrome_options.add_argument("--disable-plugins")
    chrome_options.add_argument("--disable-extensions")

    # 创建WebDriver对象时传入ChromeOptions
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url_root) # 打开的页面
    
    # 手动操作，比如切换标签页等
    input("请手动操作至论文详情页面,完成后按Enter键继续...")
    
    # 获取获取当前所有窗口的句柄
    window_handles = driver.window_handles
    
    # 假设新窗口是最后被打开的
    new_window_handle = window_handles[-1]

    # 切换到新窗口
    driver.switch_to.window(new_window_handle)

    # 在新窗口上进行操作，例如获取新窗口的标题
    print("新窗口的标题(请确保页面正确):", driver.title)

    # 在主循环中添加手动结束功能
    while index <= papers_need:
        print("正在处理第", index+1, "篇论文")
        
        # 检查是否按下了退出键(Esc)
        if keyboard.is_pressed('esc'):
            choice = input("确认退出程序？(y/n): ")
            if choice.lower() == 'y':
                print("程序已手动结束，数据已保存")
                driver.quit()
                sys.exit(0)
        
        # 等待页面加载 
        try: 
            # 或者等待直到某个元素可见 
            element = WebDriverWait(driver, wait_time).until( 
                EC.visibility_of_element_located((By.XPATH, judge_xpath)) 
            ) 
        except Exception as e: 
            print("等待超时，页面不存在该元素，也可能是页面加载失败") 
        time.sleep(pause_time)
        
        # 解析HTML
        # 在主循环的try块中
        try:
            html = driver.page_source
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_nextpaper))).click() # 切换到下一页
            index, data = parse_html(html, driver)  # 传入driver参数
            row_index = f'Row_{index}'
            if row_index in df.index:
                df.loc[row_index] = pd.Series(data, name=row_index) # 如果行索引存在，则覆盖对应行的数据
            else:
                df = df._append(pd.Series(data, name=row_index)) # 如果行索引不存在，则追加新的行
            df.to_csv(file_path, index=True)  # 将DataFrame保存为CSV文件,保留行索引作为第一列
    
            # 增强的键盘中断检测
            if keyboard.is_pressed('down') or keyboard.is_pressed('0'):  # 按下键盘上的向下键或者数字0 进入中断检查
                t = input("程序中断，输入Enter键继续，输入'exit'退出程序，其他用于调试...")
                if t.lower() == 'exit':
                    print("程序已手动结束，数据已保存")
                    driver.quit()
                    sys.exit(0)
                while t != '':
                    html = driver.page_source
                    index,data = parse_html(html)
                    t = input("程序中断，输入Enter键继续，输入'exit'退出程序...")
                    if t.lower() == 'exit':
                        print("程序已手动结束，数据已保存")
                        driver.quit()
                        sys.exit(0)
            flag = flag - 1
        except Exception as e:
            print("An error occurred:", e)
            if flag <= 0:
                print("尝试重新加载页面...")
                flag = 2
                driver.back()
                time.sleep(pause_time)
            else:
                winsound.Beep(freq, duration)  # 提示音
                input("网页出现问题等待手动解决...")
            
        

    # 关闭浏览器
    driver.quit()