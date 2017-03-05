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


def sumtimesplit(strtimeList):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimeList:
        if re.search(u'天', item):
            timeList = item.split("天".decode('gbk'))
            timeOne = int(timeList[0]) * 86400
            timeTwo = int(timeList[1].split("小时".decode('gbk'))[0]) * 3600
            totalTimeTemp = timeOne + timeTwo
            tempTimeFunc.append(totalTimeTemp)
        else:
            timeList = item.split("小时".decode('gbk'))
            totalTimeTemp = int(timeList[0]) * 3600
            tempTimeFunc.append(totalTimeTemp)
    for item in tempTimeFunc:
        totalTime += item
    dayTime, hourtimeTemp = divmod(totalTime, 86400)
    hourTime = divmod(hourtimeTemp, 3600)[0]
    dataReturn = "%d天%d小时".decode('gbk') %(dayTime, hourTime)
    return dataReturn

print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
username = raw_input("Please input Username:")
password = raw_input("Please input Password:")
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
    tkMessageBox.showerror("用户名/密码输入错误".decode('gbk'), "用户名或者密码输入错误，请重新执行程序".decode('gbk'))
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
        if auditNo not in auditList:
            name = browser.find_element_by_css_selector(
                "#datagrid-row-r1-2-%d > td:nth-child(3) > div > span" % count).get_attribute("title")
            productName = browser.find_element_by_css_selector(
                "#datagrid-row-r1-2-%d > td:nth-child(5) > div" % count).text
            submitTimeTemp = \
            browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(7) > div" % count).text.split(
                " ")[0]
            submitTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(submitTimeTemp, '%Y-%m-%d')))
            closeTimeTemp = \
            browser.find_element_by_css_selector("#datagrid-row-r1-2-%d > td:nth-child(8) > div" % count).text.split(
                " ")[0]
            if closeTimeTemp != u'':
                closeTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(closeTimeTemp, '%Y-%m-%d')))
            else:
                closeTime = '评审进行中'.decode('gbk')
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
                reportList.append("有报告".decode('gbk'))
                nameTemp = reportTemp.text[:-6]
                reportNameList.append(nameTemp)
            except selenium.common.exceptions.NoSuchElementException:
                reportList.append("无报告".decode('gbk'))
                reportNameList.append("无".decode('gbk'))
            try:
                timeTestTemp = []
                totalTd = browser.find_elements_by_tag_name("td")
                for item in totalTd:
                    textTemp = item.text
                    if textTemp == "测试".decode('gbk'):
                        timeTest = item.find_element_by_xpath("parent::tr/td[3]").text
                        timeTestTemp.append(timeTest)
                timeTestData = sumtimesplit(timeTestTemp)
                totalTestTimeList.append(timeTestData)
            except selenium.common.exceptions.NoSuchElementException:
                totalTestTimeList.append("无测试参与".decode('gbk'))
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
    print("当前正在抓取第%d页，总共%d页".decode('gbk') % (pageCount - 1, totalPages))
    inNum = browser.find_element_by_css_selector("input.pagination-num")
    inNum.clear()
    inNum.send_keys(pageCount)
    inNum.send_keys(Keys.ENTER)
    time.sleep(5)
browser.quit()

#如下是数据处理，有浏览器不再发生关系
TitleItem = ['评审编号'.decode('gbk'), '评审名称'.decode('gbk'), '项目名称'.decode('gbk'), '提交时间'.decode('gbk'),
             '关闭时间'.decode('gbk'), '处理时长'.decode('gbk'), '测试花费时间'.decode('gbk'), '状态'.decode('gbk'), '是否有报告附件'.decode('gbk'),
             '报告名称'.decode('gbk'),]
WorkBook = xlsxwriter.Workbook("评审系统抓取信息.xlsx".decode('gbk'))
SheetOne = WorkBook.add_worksheet('sheet1')
formatOne = WorkBook.add_format()
formatOne.set_border(1)
formatTwo = WorkBook.add_format()
formatTwo.set_border(1)
formatTwo.set_num_format('yy/mm/dd')
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
tkMessageBox.showinfo('完成啦'.decode('gbk'), '抓到%s个结果！已经将结果写入《评审系统抓取信息.xlsx》，请自行查看！'.decode('gbk') % len(auditList))
