'''
:@Author: KIKO_KEKE
:@Date: 2024/4/9 13:09:59
:@LastEditors: KIKO_KEKE
:@LastEditTime: 2024/4/9 13:09:59
:Description: 
:Copyright: Copyright (©)}) 2024 KIKO_KEKE. All rights reserved.
'''
# -*- coding:utf-8 -*-

from this import d
from bs4 import BeautifulSoup
#from matplotlib.pyplot import title
import pandas as pd
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep

df = pd.read_excel('111.xlsx')
names = df["keyword_2"].tolist()
# r = open("result.csv", "a+", encoding="utf-8", newline='')
# f = open("error.csv", "a+", encoding="utf-8", newline='')
# 加载驱动，括号里为下载的引擎存放地址
option = webdriver.ChromeOptions()
option.add_experimental_option("detach", True)
# 页面隐藏
option.add_argument('--no-sandbox')
option.add_argument('--disable-dev-shm-usage')
option.add_argument('window-size=1920x3000')  # 指定浏览器分辨率
option.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
option.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
# option.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度
# option.add_argument('--headless')  # 浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
driver = webdriver.Chrome(chrome_options=option)
#first_url = "http://www.baikelib.com"
first_url = "http://www.ume188.com/"
# 相当于在浏览器输入网络地址
url_1 = driver.get(first_url)
usr = '2877705123@qq.com'
passwd = 'KYL20760213'

def signin():
    sleep(2)
    #driver.find_element(
       # By.XPATH, value='/html/body/header/div/nav/ul[2]/li[2]/a').click()
    driver.find_element(
        By.XPATH, value='//*[@id="bs-navbar"]/ul[2]/li[1]/a').click()
    print('进入登录')
    sleep(2)
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/form/div[1]/div/input').clear()
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/form/div[1]/div/input').send_keys(usr)
    print('输入用户名')
    # sleep(3)
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/form/div[2]/div/input').clear()
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/form/div[2]/div/input').send_keys(passwd)
    print('输入密码')
    sleep(5)
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/form/div[4]/input').click()
    print('输入验证码')
    sleep(3)


def entry():
    driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
    #driver.find_element(
        #By.XPATH, value='/html/body/div[2]/div[3]/form/input[2]').click()
    driver.find_element(
        By.XPATH, value='/html/body/div/div[13]/div[2]/div/a').click()
    driver.find_element(
        By.XPATH, value='/html/body/div[2]/div/a[2]').click()
    print('进入数据库集')
    sleep(15)
    #driver.find_element(By.XPATH, value='//*[@id="lib129"]').click()
    #print('Factiva道琼斯')
    #sleep(3)
    

    print('Factiva道琼斯--')

    handle = driver.current_window_handle
    print(handle)
    handles = driver.window_handles
    print(handles)
    for j in handles:
        if j != handle:
            driver.switch_to.window(j)
            print(j)
    sleep(30)#手动人机验证
    driver.find_element(
        By.XPATH, value=' //*[@id="navmbm0"]/a').click()
    print('进入成功')
   

def puttime():
    frm = '01'
    frd = '01'
    fry = '2022'
    tom = '12'
    tod = '31'
    toy = '2022'
    driver.find_element(
        By.XPATH, value='//*[@id="dr"]').click()
    sleep(1)
    driver.find_element(
        By.XPATH, value='//*[@id="dr"]/option[10]').click()
    print('点击日期')
    driver.find_element(
        By.XPATH, value='//*[@id="frm"]').send_keys(frm)
    driver.find_element(
        By.XPATH, value='//*[@id="frd"]').send_keys(frd)
    sleep(1)
    driver.find_element(
        By.XPATH, value='//*[@id="fry"]').send_keys(fry)
    driver.find_element(
        By.XPATH, value='//*[@id="tom"]').send_keys(tom)
    driver.find_element(
        By.XPATH, value='//*[@id="tod"]').send_keys(tod)
    driver.find_element(
        By.XPATH, value='//*[@id="toy"]').send_keys(toy)
    print('输入日期成功')
    sleep(2)

def search(i):
    cleartext = driver.find_element(
        By.CLASS_NAME, value='ace_text-input')
    cleartext.clear()
    cleartext.send_keys(Keys.CONTROL + 'a')
    cleartext.send_keys(Keys.BACKSPACE)
    print('清除空格')
    sleep(1)
    driver.find_element(By.CLASS_NAME, value='ace_text-input').send_keys(i)
    # text=i
    # js = "var sum=driver.find_element(By.CSS_SELECTOR, value='#editor > textarea');sum.value='" + text +"';"
    # driver.execute_script(js);
    print('输入关键词')
    sleep(1)

    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    driver.find_element(
        By.XPATH,
        value='//*[@id="btnSBSearch"]/div/span').click()
    print('点击搜索')
    sleep(10)

    # num1 = driver.find_element(
    #     By.XPATH,
    #     value='//*[@id="oddEvenPreview"]/tbody/tr[11]/td[2]').text
    try:
        num1 = driver.find_element(
            By.XPATH, value='//*[@id="headlineTabs"]/table[1]/tbody/tr/td/span[2]/a/span').text
    except:
        driver.find_element(
        By.XPATH, value='//*[@id="headlineTabs"]/table[1]/tbody/tr/td/span[1]/a').click()
        sleep(5)
        num1 = driver.find_element(
            By.XPATH, value='//*[@id="headlineTabs"]/table[1]/tbody/tr/td/span[2]/a/span').text
    sleep(5)

    print(num1)
    with open("112.xlsx","a",newline = "",encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(
        (n+1, df.iloc[n, 0],df.iloc[n, 5],i, num1))
    print(n+1, df.iloc[n, 0],  df.iloc[n, 5], i, num1)
    print("成功")
    sleep(5)
    driver.find_element(
        By.XPATH,
        value='//*[@id="btnModifySearch"]/div/span').click()
    sleep(2)
    #点击修改搜索，避免重复输入日期


signin()
entry()
puttime()

e = 0
#for n in range(e,len(names)):
for n in range(e, len(names)):
    #for i in df['name']:
        # print(i)
    i=names[n]
    try:
        search(i)
    except:
        with open("道琼斯失败_2.xlsx","a",newline = "",encoding='utf-8') as fi:
            writer = csv.writer(fi)
            writer.writerow
            ((n+1, df.iloc[n, 0], df.iloc[n, 1], df.iloc[n, 2],df.iloc[n, 3],df.iloc[n, 4],'nan'))
        print(n+1,df.iloc[n, 5]+"失败")
    n += 1

