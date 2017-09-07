#!/usr/bin/env python
# -*- coding:cp936 -*-
import time
import os
import re
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import datetime
import wx
from threading import Thread


def sumtimesplit(strtimelist):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimelist:
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
    dataReturn = "%d天%d小时".decode('gbk') % (dayTime, hourTime)
    return dataReturn


class PingShenFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"评审系统信息抓取工具", pos=wx.DefaultPosition,
                          size=wx.Size(504, 460), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))
        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下输入用户名和密码", wx.DefaultPosition, wx.DefaultSize,
                                         wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)
        self.text_title1.SetFont(wx.Font(12, 70, 90, 90, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer3, 0, wx.EXPAND, 5)

        gSizer2 = wx.GridSizer(2, 2, 0, 0)

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"用户名", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_username.Wrap(-1)
        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        gSizer2.Add(self.input_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"密码", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_password.Wrap(-1)
        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        gSizer2.Add(self.input_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(gSizer2, 0, 0, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_title2 = wx.StaticText(self.m_panel1, wx.ID_ANY,
                                         u"请在如下输入想要抓取的信息的起止日期(需包含年/月/日信息！格式为20170101.\n个位数的月和日一定要带0！）",
                                         wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTRE)
        self.text_title2.Wrap(-1)
        self.text_title2.SetFont(wx.Font(9, 70, 90, 90, False, wx.EmptyString))
        self.text_title2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_title2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer4, 0, wx.EXPAND, 5)

        gSizer3 = wx.GridSizer(0, 2, 0, 0)

        self.text_startdate = wx.StaticText(self.m_panel1, wx.ID_ANY, u"开始日期", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_startdate.Wrap(-1)
        self.text_startdate.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_startdate.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer3.Add(self.text_startdate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_startdate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                           0)
        gSizer3.Add(self.input_startdate, 0, wx.ALL, 5)

        self.text_enddate = wx.StaticText(self.m_panel1, wx.ID_ANY, u"结束日期", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_enddate.Wrap(-1)
        self.text_enddate.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_enddate.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer3.Add(self.text_enddate, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_enddate = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                         0)
        gSizer3.Add(self.input_enddate, 0, wx.ALL, 5)

        bSizer10.Add(gSizer3, 0, 0, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.text_3 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下选择想要在最后文件中显示的项目", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        self.text_3.Wrap(-1)
        self.text_3.SetFont(wx.Font(12, 70, 90, 90, False, wx.EmptyString))
        self.text_3.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_3.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer9.Add(self.text_3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.VERTICAL)

        bSizer61 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkBox_audit = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"评审编号", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_audit.SetValue(True)
        bSizer61.Add(self.checkBox_audit, 0, wx.ALL, 5)

        self.checkBox_name = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"评审名称", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_name.SetValue(True)
        bSizer61.Add(self.checkBox_name, 0, wx.ALL, 5)

        self.checkBox_productName = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"项目名称", wx.DefaultPosition, wx.DefaultSize,
                                                0)
        self.checkBox_productName.SetValue(True)
        bSizer61.Add(self.checkBox_productName, 0, wx.ALL, 5)

        self.checkBox_submitTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"提交时间", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_submitTime.SetValue(True)
        bSizer61.Add(self.checkBox_submitTime, 0, wx.ALL, 5)

        self.checkBox_closeTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"关闭时间", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_closeTime.SetValue(True)
        bSizer61.Add(self.checkBox_closeTime, 0, wx.ALL, 5)

        bSizer6.Add(bSizer61, 0, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.checkBox_handleTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"处理时长", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_handleTime.SetValue(True)
        bSizer8.Add(self.checkBox_handleTime, 0, wx.ALL, 5)

        self.checkBox_totalTestTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"测试花费时间", wx.DefaultPosition,
                                                  wx.DefaultSize, 0)
        self.checkBox_totalTestTime.SetValue(True)
        bSizer8.Add(self.checkBox_totalTestTime, 0, wx.ALL, 5)

        self.checkBox_status = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"当前状态", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_status.SetValue(True)
        bSizer8.Add(self.checkBox_status, 0, wx.ALL, 5)

        self.checkBox_report = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"是否有报告附件", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_report.SetValue(True)
        bSizer8.Add(self.checkBox_report, 0, wx.ALL, 5)

        self.checkBox_summary = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"评审要点", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkBox_summary.SetValue(True)
        bSizer8.Add(self.checkBox_summary, 0, wx.ALL, 5)

        bSizer6.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer10.Add(bSizer6, 0, wx.EXPAND, 5)

        bSizer21 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self.m_panel1, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer10.Add(bSizer21, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.textctrl_display = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.textctrl_display, 1, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer91, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

        self._thread = Thread(target=self.run, args=())
        self._thread.daemon = True


    def __del__(self):
        pass

    def close(self, event):
        self.Close()

    def run(self):
        self.updatedisplay("开始抓取".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
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
        summaryList = []
        start_data = int(self.input_startdate.GetValue())
        end_date = int(self.input_enddate.GetValue())
        browser = webdriver.Chrome(chromedriverPath)
        # browser = webdriver.PhantomJS()
        url = "http://218.57.146.175/techAudit/welcome.htm"
        browser.get(url)
        browser.find_element_by_id("loginName").send_keys(username)
        browser.find_element_by_id("password").send_keys(password)
        browser.find_element_by_css_selector(
            "#loginForm > div > div.communalForm.clear > dl:nth-child(5) > dd > a").click()
        time.sleep(3)
        try:
            testEle = browser.find_element_by_css_selector("#msg")
            dlg_error = wx.MessageDialog(None, "用户名或者密码输入错误，请重新执行程序".decode('gbk'), "用户名/密码输入错误".decode('gbk'),
                                         wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            browser.close()
            if dlg_error.ShowModal() == wx.ID_OK:
                self.Close()
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
        for pageCount in range(2, totalPages + 2):
            lineTotal = len(browser.find_elements_by_css_selector(
                "body > div.panel.layout-panel.layout-panel-center > div > div > div > div.datagrid-view > div.datagrid-view2 > div.datagrid-body > table > tbody > tr"))
            for count in range(0, lineTotal):
                auditNo = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(2) > div > span" % count).get_attribute("title")
                submitTimeTemp = browser.find_element_by_css_selector(
                    "#datagrid-row-r1-2-%d > td:nth-child(7) > div" % count).text.split(" ")[0]
                submitTimeTempTemp = int("".join(submitTimeTemp.split('-')))
                if auditNo not in auditList and end_date >= submitTimeTempTemp >= start_data:
                    name = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(3) > div > span" % count).get_attribute("title")
                    productName = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(5) > div" % count).text

#                    submitTimeTempTemp = "".join(submitTimeTemp.split('-'))

                    submitTime = datetime.datetime.fromtimestamp(time.mktime(time.strptime(submitTimeTemp, '%Y-%m-%d')))
                    closeTimeTemp = \
                        browser.find_element_by_css_selector(
                            "#datagrid-row-r1-2-%d > td:nth-child(8) > div" % count).text.split(
                            " ")[0]
                    if closeTimeTemp != u'':
                        closeTime = datetime.datetime.fromtimestamp(
                            time.mktime(time.strptime(closeTimeTemp, '%Y-%m-%d')))
                    else:
                        closeTime = '评审进行中'.decode('gbk')
                    handleTime = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(9) > div" % count).text
                    status = browser.find_element_by_css_selector(
                        "#datagrid-row-r1-2-%d > td:nth-child(12) > div" % count).text
                    lineData = browser.find_element_by_css_selector("#datagrid-row-r1-2-%d" % count)
                    ActionChains(browser).double_click(lineData).perform()
                    time.sleep(2)
                    browser.switch_to.default_content()
                    browser.switch_to.frame(2)
                    summary = browser.find_element_by_css_selector(
                        "body > div.panel.layout-panel.layout-panel-center > div > table.TableCssList > tbody > tr:nth-child(9) > td > p:nth-child(1)").text
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
                    summaryList.append(summary)
           # print("当前正在抓取第%d页，总共%d页".decode('gbk') % (pageCount - 1, totalPages))
            self.updatedisplay("当前正在抓取第%d页，总共%d页".decode('gbk') % (pageCount - 1, totalPages))
            inNum = browser.find_element_by_css_selector("input.pagination-num")
            inNum.clear()
            inNum.send_keys(pageCount)
            inNum.send_keys(Keys.ENTER)
            time.sleep(5)
        browser.quit()
        # 如下是数据处理，与浏览器不再发生关系
        TitleItem = ['评审编号'.decode('gbk'), '评审名称'.decode('gbk'), '项目名称'.decode('gbk'), '提交时间'.decode('gbk'),
                     '关闭时间'.decode('gbk'), '处理时长'.decode('gbk'), '测试花费时间'.decode('gbk'), '状态'.decode('gbk'),
                     '是否有报告附件'.decode('gbk'), '报告名称'.decode('gbk'), '评审要点'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("评审系统抓取信息-%s.xlsx".decode('gbk') % timestamp)
        SheetOne = WorkBook.add_worksheet('评审系统抓取信息'.decode('gbk'))
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
            if self.checkBox_audit.GetValue():
                SheetOne.write(lineCount, 0, auditList[index], formatOne)
            if self.checkBox_name.GetValue():
                SheetOne.write(lineCount, 1, nameList[index], formatOne)
            if self.checkBox_productName.GetValue():
                SheetOne.write(lineCount, 2, productNameList[index], formatOne)
            if self.checkBox_submitTime.GetValue():
                SheetOne.write(lineCount, 3, submitTimeList[index], formatTwo)
            if self.checkBox_closeTime.GetValue():
                SheetOne.write(lineCount, 4, closeTimeList[index], formatTwo)
            if self.checkBox_handleTime.GetValue():
                SheetOne.write(lineCount, 5, handleTimeList[index], formatOne)
            if self.checkBox_totalTestTime.GetValue():
                SheetOne.write(lineCount, 6, totalTestTimeList[index], formatOne)
            if self.checkBox_status.GetValue():
                SheetOne.write(lineCount, 7, statusList[index], formatOne)
            if self.checkBox_report.GetValue():
                SheetOne.write(lineCount, 8, reportList[index], formatOne)
                SheetOne.write(lineCount, 9, reportNameList[index], formatOne)
            if self.checkBox_summary.GetValue():
                SheetOne.write(lineCount, 10, summaryList[index], formatOne)
            lineCount += 1
        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("抓到%s个结果！已经将结果写入《评审系统抓取信息.xlsx》，请自行查阅！请点击EXIT退出程序！".decode('gbk') % len(auditList))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def onbutton(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("完成第%s页".decode('gbk') % t)
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText("%s".decode('gbk') % t)
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    app = wx.App()
    frame = PingShenFrame(None)
    frame.Show()
    app.MainLoop()
