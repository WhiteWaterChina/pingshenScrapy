#!/usr/bin/env python
# -*- coding:cp936 -*-
import time
import os
import sys
import re
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import tkMessageBox
import xlsxwriter
import datetime
import Tkinter


def sumtimesplit(strtimeList):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimeList:
        if re.search(u'��', item):
            timeList = item.split("��".decode('gbk'))
            timeOne = int(timeList[0]) * 86400
            timeTwo = int(timeList[1].split("Сʱ".decode('gbk'))[0]) * 3600
            totalTimeTemp = timeOne + timeTwo
            tempTimeFunc.append(totalTimeTemp)
        else:
            timeList = item.split("Сʱ".decode('gbk'))
            totalTimeTemp = int(timeList[0]) * 3600
            tempTimeFunc.append(totalTimeTemp)
    for item in tempTimeFunc:
        totalTime += item
    dayTime, hourtimeTemp = divmod(totalTime, 86400)
    hourTime = divmod(hourtimeTemp, 3600)[0]
    dataReturn = "%d��%dСʱ".decode('gbk') % (dayTime, hourTime)
    return dataReturn


def get_data():
    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    username = var_char_entry_username.get()
    password = var_char_entry_password.get()
    chromedriverPath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
    auditList = []
    nameList = []
    productNameList = []
    submitTimeList = []
    closeTimeList = []
    handleTimeList = []
    statusList = []
    reportList = []
    reportNameList = []
    totalTestTimeList = []
    start_data = int(var_char_entry_date_start.get())
    end_date = int(var_char_entry_data_end.get())
    browser = webdriver.Chrome(chromedriverPath)
    #browser = webdriver.PhantomJS()
    url = "http://218.57.146.175/techAudit/welcome.htm"
    browser.get(url)
    browser.find_element_by_id("loginName").send_keys(username)
    browser.find_element_by_id("password").send_keys(password)
    browser.find_element_by_xpath("//a[@class='alogin fl']").click()
    time.sleep(3)
    try:
        testEle = browser.find_element_by_css_selector("#msg")
        tkMessageBox.showerror("�û���/�����������".decode('gbk'), "�û������������������������ִ�г���".decode('gbk'))
        browser.close()
        sys.exit()
    except selenium.common.exceptions.NoSuchElementException:
        pass
    browser.switch_to.frame('ta')
    abc = browser.find_element_by_xpath("//div[@id='_easyui_tree_2']")
    abc.click()
    browser.implicitly_wait(10)
    browser.switch_to.default_content()
    browser.switch_to.frame(1)
    WebDriverWait(browser, 100).until(
        ec.element_to_be_clickable((By.XPATH, "//div[@class='datagrid-pager pagination']/table/tbody/tr/td[10]/a")))
    totalPagesEle = browser.find_element_by_css_selector(
        "body > div.panel.layout-panel.layout-panel-center > div > div > div > div.datagrid-pager.pagination > table > tbody > tr > td:nth-child(8) > span")
    time.sleep(5)
    totalPagesUnicode = totalPagesEle.text
    patternPages = re.compile(r'\d+', re.U)
    totalPages = int(re.findall(patternPages, totalPagesUnicode)[0])
    for pageCount in range(2, totalPages+2):
        lineTotal = len(browser.find_elements_by_css_selector(
            "body > div.panel.layout-panel.layout-panel-center > div > div > div > div.datagrid-view > div.datagrid-view2 > div.datagrid-body > table > tbody > tr"))
        for count in range(0, lineTotal):
            auditNo = browser.find_element_by_css_selector(
                "#datagrid-row-r1-2-%d > td:nth-child(2) > div > span" % count).get_attribute("title")
            submitTimeTemp = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(7) > div" % count).text.split(" ")[0]
            submitTimeTempTemp = int("".join(submitTimeTemp.split('-')))
            if auditNo not in auditList and end_date >= submitTimeTempTemp >= start_data:
                name = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(3) > div > span" % count).get_attribute("title")
                productName = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(5) > div" % count).text
#                submitTimeTemp = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(7) > div" % count).text.split(" ")[0]
                submitTimeTempTemp = "".join(submitTimeTemp.split('-'))
                submitTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(submitTimeTemp, '%Y-%m-%d')))
                closeTimeTemp = \
                browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(8) > div" % count).text.split(
                    " ")[0]
                if closeTimeTemp != u'':
                    closeTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(closeTimeTemp, '%Y-%m-%d')))
                else:
                    closeTime = '���������'.decode('gbk')
                handleTime = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(9) > div" % count).text
                status = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(12) > div" % count).text
                lineData = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d" % count)
                ActionChains(browser).double_click(lineData).perform()
                time.sleep(2)
                browser.switch_to.default_content()
                browser.switch_to.frame(2)
                try:
                    reportTemp = browser.find_element_by_css_selector(
                        "form#testUploadForm > table > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(1)")
                    reportList.append("�б���".decode('gbk'))
                    nameTemp = reportTemp.text[:-6]
                    reportNameList.append(nameTemp)
                except selenium.common.exceptions.NoSuchElementException:
                    reportList.append("�ޱ���".decode('gbk'))
                    reportNameList.append("��".decode('gbk'))
                try:
                    timeTestTemp = []
                    totalTd = browser.find_elements_by_tag_name("td")
                    for item in totalTd:
                        textTemp = item.text
                        if textTemp == "����".decode('gbk'):
                            timeTest = item.find_element_by_xpath("parent::tr/td[3]").text
                            timeTestTemp.append(timeTest)
                    timeTestData = sumtimesplit(timeTestTemp)
                    totalTestTimeList.append(timeTestData)
                except selenium.common.exceptions.NoSuchElementException:
                    totalTestTimeList.append("�޲��Բ���".decode('gbk'))
                browser.switch_to.default_content()
                closeButton = browser.find_element_by_css_selector(
                    "#tabs > div.tabs-header.tabs-header-noborder > div.tabs-wrap > ul > li.tabs-selected > a.tabs-close")
                closeButton.click()
                browser.switch_to.frame(1)
                auditList.append(auditNo)
                nameList.append(name)
                productNameList.append(productName)
                submitTimeList.append(submitTime)
                closeTimeList.append(closeTime)
                handleTimeList.append(handleTime)
                statusList.append(status)
        print("��ǰ����ץȡ��%dҳ���ܹ�%dҳ".decode('gbk') % (pageCount - 1, totalPages))
        inNum = browser.find_element_by_css_selector("input.pagination-num")
        inNum.clear()
        inNum.send_keys(pageCount)
        inNum.send_keys(Keys.ENTER)
        time.sleep(5)
    browser.quit()

    #���������ݴ�������������ٷ�����ϵ
    TitleItem = ['������'.decode('gbk'), '��������'.decode('gbk'), '��Ŀ����'.decode('gbk'), '�ύʱ��'.decode('gbk'),
                 '�ر�ʱ��'.decode('gbk'), '����ʱ��'.decode('gbk'), '���Ի���ʱ��'.decode('gbk'), '״̬'.decode('gbk'), '�Ƿ��б��渽��'.decode('gbk'),
                 '��������'.decode('gbk'),]
    timestamp = time.strftime('%Y%m%d', time.localtime())
    WorkBook = xlsxwriter.Workbook("����ϵͳץȡ��Ϣ-%s.xlsx".decode('gbk') % timestamp)
    SheetOne = WorkBook.add_worksheet('����ϵͳץȡ��Ϣ'.decode('gbk'))
    formatOne = WorkBook.add_format()
    formatOne.set_border(1)
    formatTwo = WorkBook.add_format()
    formatTwo.set_border(1)
    formatTwo.set_num_format('yy/mm/dd')
    SheetOne.set_column('A:J', 14)
    for i in range(0, len(TitleItem)):
        SheetOne.write(0, i, TitleItem[i], formatOne)
    lineCount = 1
    for index, item in enumerate(auditList):
        SheetOne.write(lineCount, 0, auditList[index], formatOne)
        SheetOne.write(lineCount, 1, nameList[index], formatOne)
        SheetOne.write(lineCount, 2, productNameList[index], formatOne)
        SheetOne.write(lineCount, 3, submitTimeList[index], formatTwo)
        SheetOne.write(lineCount, 4, closeTimeList[index], formatTwo)
        SheetOne.write(lineCount, 5, handleTimeList[index], formatOne)
        SheetOne.write(lineCount, 6, totalTestTimeList[index], formatOne)
        SheetOne.write(lineCount, 7, statusList[index], formatOne)
        SheetOne.write(lineCount, 8, reportList[index], formatOne)
        SheetOne.write(lineCount, 9, reportNameList[index], formatOne)
        lineCount += 1
    WorkBook.close()
    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    tkMessageBox.showinfo('�����'.decode('gbk'), 'ץ��%s��������Ѿ������д�롶����ϵͳץȡ��Ϣ.xlsx���������в��ģ�����EXIT�˳�����'.decode('gbk') % len(auditList))



root = Tkinter.Tk()
root.title("����ϵͳץȡ����".decode('gbk'))
#root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_username = Tkinter.StringVar()
var_char_entry_password = Tkinter.StringVar()
var_char_entry_date_start = Tkinter.StringVar()
var_char_entry_data_end = Tkinter.StringVar()
var_char_entry_day_start = Tkinter.StringVar()

Tkinter.Label(root, text='����������������ϵͳ���û���������'.decode('gbk'), bg='Yellow', justify='left').grid(row=0, column=0, columnspan=4)
Tkinter.Label(root, text='�û���'.decode('gbk'), bg='Red').grid(row=1, column=0, sticky='W', padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_username).grid(row=1, column=0, columnspan=2, padx=10, pady=5)
Tkinter.Label(root, text='����'.decode('gbk'), bg='Red').grid(row=2, column=0, sticky='W', padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_password).grid(row=2, column=0, columnspan=2, padx=10, pady=5)

Tkinter.Label(root, text='��������������Ҫץȡ����Ϣ����ֹ����(�������/��/����Ϣ����ʽΪ20170101.��λ�����º���һ��Ҫ��0����'.decode('gbk'), bg='Yellow', justify='left').grid(row=3, column=0, columnspan=4)
Tkinter.Label(root, text='��ʼ����'.decode('gbk'), bg='Red').grid(row=4, column=0, sticky='W', padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_date_start).grid(row=4, column=0, columnspan=2, padx=10, pady=5)
Tkinter.Label(root, text='��������'.decode('gbk'), bg='Red').grid(row=5, column=0, sticky='W', padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_data_end).grid(row=5, column=0, columnspan=2, padx=10, pady=5)

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=6, column=0, columnspan=2, sticky='WESN', padx=20, pady=10)
Tkinter.Button(root, text='Exit'.decode('gbk'), command=root.destroy).grid(row=6, column=2, columnspan=2, sticky='WESN', padx=20, pady=10)


root.mainloop()
