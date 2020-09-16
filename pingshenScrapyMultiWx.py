#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import multiprocessing
import xlsxwriter
import os
import time
import datetime
from threading import Thread
import wx
from multiprocessing import Pool
import base64
import sys
from bs4 import element

sys.setrecursionlimit(3000)

ver = "Ward Yan-20200915"
web_address = "218.57.146.175:8114"


def sumtimesplit(strtimelist):
    tempTimeFunc = []
    totalTime = int(0)
    for item in strtimelist:
        if re.search(u'天', item):
            timeList = item.split("天")
            timeOne = int(timeList[0]) * 86400
            timeTwo = int(timeList[1].split("小时")[0]) * 3600
            totalTimeTemp = timeOne + timeTwo
            tempTimeFunc.append(totalTimeTemp)
        else:
            timeList = item.split("小时")
            totalTimeTemp = int(timeList[0]) * 3600
            tempTimeFunc.append(totalTimeTemp)
    for item in tempTimeFunc:
        totalTime += item
    dayTime, hourtimeTemp = divmod(totalTime, 86400)
    hourTime = divmod(hourtimeTemp, 3600)[0]
    dataReturn = "{}天{}小时".format(dayTime, hourTime)
    return dataReturn


def get_status(status):
    switcher = {
        "0": "保存",
        "1": "提交",
        "audit-submit": "修改评审信息",
        "shenhepeizhi-sq": "售前审核配置",
        "shenhe-chanpinjingli": "产品经理审核",
        "querenxuanpei-ddy": "订单员确认选配",
        "shenhepingshen-yf": "研发接口人审核评审",
        "shenhepingshen-csjk": "测试接口人审核评审",
        "ceshi-cs": "测试人员测试",
        "shenheceshibaogao-yf": "研发接口审核测试报告",
        "shenheceshibaogao-xmjl": "项目经理审核测试报告",
        "shenheceshibaogao-csfzr": "测试负责人审核测试报告",
        "shenheceshibaogao-csjk": "测试接口人审核测试报告",
        "xfzl-gc": "工程人员确认是否下发指令",
        "100": "关闭",
        "101": "异常关闭",
        "102": "暂停",
        "103": "终止",
        "vm-audit": "VM审核",
        "npi-audit": "NPI处理",
        "leader-test-judge": "测试teamleader测试决策",
        "exec-test": "执行测试",
        "leader-audit": "测试teamleader审核测试结果",
        "vm-test-audit": "VM审核测试结果",
        "test_report_concordance": "测试报告整合",
        "os-comp-test": "OS兼容性测试",
        "product-verification": "生产验证",
        "oqc-valication": "OQC验证",
        "material-add": "物料追加",
        "material-assess": "物料评估",
        "om-maintaince": "BOM维护",
        "custom-dev": "定制化开发",
        "hadware-dev": "固件研发",
    }
    return switcher.get(status, status)


def get_detail(link, login_session):
    headers_data_all = {
        'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.9,en;q=0.8",
        'Connection': "keep-alive",
        'Host': "{}".format(web_address),
        'Referer': "http://{}/techAudit/welcome.htm".format(web_address),
        'Upgrade-Insecure-Requests': "1",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36"
    }
    get_page = login_session.get(link, headers=headers_data_all)
    data_page = get_page.text
    print("Get link:{} with return code {}".format(link, get_page.status_code))
    data_filter = BeautifulSoup(data_page, "html5lib")
    # 获取附件信息
    attachment_temp = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr")
    if len(attachment_temp) <= 1:
        has_report = "无报告"
        report_filename = "None"
    else:
        has_report = "有报告"
        filename_temp = []
        for index_report, item_report in enumerate(attachment_temp):
            if index_report != 0:
                filename = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr:nth-of-type(2) > td:nth-of-type(1)")[0].contents[0].strip()
                filename_temp.append(filename)
        filename_write = ";".join(filename_temp)
        report_filename = filename_write
    # 评审要点信息
    try:
        keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(9) > td:nth-of-type(1)")[0].get_text()
    except IndexError:
        try:
            keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(7) > td:nth-of-type(1)")[0].get_text()
        except IndexError:
            keywords_temp = "None"
    keywords = keywords_temp
    # 测试花费时间
    timeTestTemp = []
    try:
        test_item = data_filter.select("body > div:nth-of-type(1) > table:nth-of-type(3) > tbody > tr")
        for item_tr in test_item:
            item_td = item_tr.select("td:nth-of-type(5)")[0].get_text().strip()
            if item_td == "测试":
                test_time = item_tr.select("td:nth-of-type(3)")[0].get_text().strip()
                timeTestTemp.append(test_time)
        if len(timeTestTemp) == 0:
            timeTestData = "None"
        else:
            timeTestData = sumtimesplit(timeTestTemp)
        test_time = timeTestData
    except IndexError:
        test_time = "None"
    # NPI最后一次的评论信息
    npi_comment = ""
    npi_comment_list = data_filter.find_all('td',text="NPI处理")
    # npi_comment_list = data_filter.select('td[text="NPI处理"]')
    if npi_comment_list is not None:
        print("Got one!")
        if len(npi_comment_list) != 0:
            npi_comment_element = npi_comment_list[-1]
            npi_comment_1 = npi_comment_element.next_sibling.next_sibling.next_sibling.next_sibling
            print(type(npi_comment_1))
            if isinstance(npi_comment_1, element.Tag):
                npi_comment = npi_comment_1.text
                print(npi_comment)
        print(1)

    return link, has_report, report_filename, keywords, test_time, npi_comment


class PingShenFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"评审系统信息抓取工具-{}".format(ver), pos=wx.DefaultPosition,
                          size=wx.Size(504, 785), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

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

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
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
                                         u"请在如下输入想要抓取的信息的起止日期(需包含年/月/日信息！\n格式为20170101.个位数的月和日一定要带0！）",
                                         wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.text_title2.Wrap(-1)

        self.text_title2.SetFont(
            wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer4.Add(self.text_title2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

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

        bSizer13 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText9 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请按如下按钮来获取项目信息", wx.DefaultPosition,
                                           wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.m_staticText9.Wrap(-1)

        self.m_staticText9.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText9.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer13.Add(self.m_staticText9, 0, wx.ALL | wx.EXPAND, 5)

        self.button_get_productname = wx.Button(self.m_panel1, wx.ID_ANY, u"获取项目名称", wx.DefaultPosition, wx.DefaultSize,
                                                0)
        bSizer13.Add(self.button_get_productname, 0, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer13, 0, wx.EXPAND, 5)

        bSizer11 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText8 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下选择需要获取数据的产品型号!可以多选！", wx.DefaultPosition,
                                           wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.m_staticText8.Wrap(-1)

        self.m_staticText8.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText8.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer11.Add(self.m_staticText8, 1, wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        listbox_productnameChoices = []
        self.listbox_productname = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                              listbox_productnameChoices, wx.LB_HSCROLL | wx.LB_MULTIPLE | wx.LB_SORT)
        bSizer11.Add(self.listbox_productname, 0, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer11, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.text_3 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下选择想要在最后文件中显示的项目", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        self.text_3.Wrap(-1)

        self.text_3.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
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
        self.button_get_productname.Bind(wx.EVT_BUTTON, self.get_productname)
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def get_productname(self, event):
        self.updatedisplay("开始获取项目名称信息,请耐心等待")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue().encode()).decode()
        # 登录
        url_login = "http://{}/techAudit/login.htm".format(web_address)
        login_session = requests.session()
        headers_login = {
            'accept': "tapplication/json, text/javascript, */*; q=0.01",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9,en;q=0.8",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/logout.htm?service=http://{}/techAudit/welcome.htm&reload=true".format(web_address, web_address),
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest",
        }
        payload_login = "loginName={}&password={}".format(username, password)
        #post登录
        login_session.post(url_login, headers=headers_login, data=payload_login)
        # 开始获取数据
        headers_data = {
            'Accept': "application/json, text/javascript, */*; q=0.01",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/auditList/queryAuditList.htm".format(web_address),
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest"
        }
        # 先用1获取最大数据条数
        payload_data_test = "row=1"
        url_data = "http://{}/techAudit/AuditListController/getQueryAuditList.htm?".format(web_address)
        response_data_test = login_session.post(url_data, headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        # 抓取按照每页5000条来进行，根据最大数量/5000来计算抓取的页面次数pages_number
        if int(total_item) % 1000 == 0:
            pages_number = int(total_item) // 1000
        else:
            pages_number = int(total_item) // 1000 + 1

        productname_list_all = []

        # 然后使用每页1000条数逐页抓取
        for item_pages_temp in range(pages_number):
            item_pages = item_pages_temp + 1
            payload_data = "page={}&rows=1000".format(item_pages)
            print(payload_data)
            response_data_get = login_session.post(url_data, headers=headers_data, data=payload_data)
            response_data = response_data_get.text
            # 获取productname的信息
            productname_list_temp = re.findall(r'"modeName":"(.*?)",', response_data)
            for item_productname_temp in productname_list_temp:
                if item_productname_temp not in productname_list_all:
                    productname_list_all.append(item_productname_temp)
        for item in productname_list_all:
            self.listbox_productname.Append(item)
        self.updatedisplay("抓取项目信息结束,请在如下选择需要抓取信息的项目名称，然后点击GO开始抓取！")
        diag_finish_project = wx.MessageDialog(None, "获取项目信息完成！请选择需要抓取信息的项目名称，然后点击GO开始抓取！", '提示',
                                               wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish_project.ShowModal()

    def run_all(self):
        self.button_go.Disable()
        self.updatedisplay("开始抓取，请耐心等待...")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        # 获取用户名和密码
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue().encode()).decode()
        # 获取开始和结束时间
        start_time = self.input_startdate.GetValue()
        end_time = self.input_enddate.GetValue()
        # 获取选择的项目
        productname_selected_list = []
        productname_selected_index_list = self.listbox_productname.GetSelections()
        for item in productname_selected_index_list:
            productname_selected_list.append(self.listbox_productname.GetString(item))
        # 登录
        login_session = requests.session()

        url_login = "http://{}/techAudit/login.htm".format(web_address)
        headers_login = {
            'accept': "tapplication/json, text/javascript, */*; q=0.01",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9,en;q=0.8",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/logout.htm?service=http://{}/techAudit/welcome.htm&reload=true".format(web_address, web_address),
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest",
        }
        payload_login = "loginName={}&password={}".format(username, password)
        # 使用以上获取的信息post登录
        login_session.post(url_login, headers=headers_login, data=payload_login)
        # 开始获取数据
        headers_data = {
            'Accept': "application/json, text/javascript, */*; q=0.01",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "{}".format(web_address),
            'Origin': "http://{}".format(web_address),
            'Referer': "http://{}/techAudit/auditList/queryAuditList.htm".format(web_address),
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest"
        }
        # 先用1获取最大数据条数
        payload_data_test = "rows=1"
        url_data = "http://218.57.146.175/techAudit/AuditListController/getQueryAuditList.htm?"
        response_data_test = login_session.post(url_data, headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        # 抓取按照每页1000条来进行，根据最大数量/1000来计算抓取的页面次数pages_number
        if int(total_item) % 1000 == 0:
            pages_number = int(total_item) // 1000
        else:
            pages_number = int(total_item) // 1000 + 1
        auditno_list = []
        id_list = []
        create_date_list = []
        close_date_list = []
        projectname_list = []
        productname_list = []
        status_list = []
        url_list = []
        total_time_list = []
        # 然后使用每页1000条数逐页抓取
        for item_pages_temp in range(pages_number):
            item_pages = item_pages_temp + 1
            payload_data = "page={}&rows=1000".format(item_pages)
            print(payload_data)
            response_data_get = login_session.post(url_data, headers=headers_data, data=payload_data)
            response_data = response_data_get.text
            # 先获取autidno/id/creattime/updatetime/projectname/productname的原始数据
            auditno_list_temp = re.findall(r'"auditNo":"(\w+)",', response_data)
            id_list_temp = re.findall(r'"id":(\d+),', response_data)
            create_date_list_temp = re.findall(r'"submitTime":"(\d+-\d+-\d+)', response_data)  # 2016-09-23
            close_date_list_temp_1 = re.findall(r'"closeTime":(.*?),', response_data)  # "closeTime":"2015-04-10 08:42:21", or "closeTime":null
            projectname_list_temp = re.findall(r'"projectName":"(.*?)",', response_data)
            productname_list_temp = re.findall(r'"modeName":"(.*?)",', response_data)
            status_list_temp = re.findall(r'"status":"(.*?)",', response_data)
            total_time_list_temp = re.findall(r'"handleTime":"(.*?)"', response_data)
            # 转换要求的开始和结束时间相对1970年1月1日的秒数
            start_date = time.mktime(time.strptime(start_time, '%Y%m%d'))
            end_date = time.mktime(time.strptime(end_time, '%Y%m%d'))
            # 处理关闭时间的多种情况
            close_date_list_temp = []
            for item_close_date in close_date_list_temp_1:
                if item_close_date == "null":
                    close_date_list_temp.append("None")
                else:
                    temp_close_date = item_close_date.strip().split(" ")[0][1:]
                    close_date_list_temp.append(temp_close_date)

            # 过滤一遍，去除时间不符合要求的、不是选择的项目名称的
            auditno_list_every_page = []
            id_list_every_page = []
            create_date_list_every_page = []
            close_date_list_every_page = []
            projectname_list_every_page = []
            productname_list_every_page = []
            status_list_every_page = []
            url_list_every_page = []
            total_time_list_every_page = []

            for index_creattime, item_creattime in enumerate(create_date_list_temp):
                create_date_format = time.mktime(time.strptime(item_creattime, '%Y-%m-%d'))
                pruductname_every = productname_list_temp[index_creattime]
                # 时间判断 + 产品名称判断
                if float(start_date) <= float(create_date_format) <= float(end_date) and pruductname_every in productname_selected_list:
                    # 创建时间，第四列
                    create_date_list_every_page.append(item_creattime)
                    # 评审编号，第一列
                    auditno_list_every_page.append(auditno_list_temp[index_creattime])
                    url = "http://{}/techAudit/v1Details/viewBill.htm?id=".format(web_address) + id_list_temp[index_creattime]
                    # 每个评审的url中的唯一编号，作为连接前面获取的总数据和后面每项评审的数据的纽带
                    id_list_every_page.append(id_list_temp[index_creattime])
                    # 关闭时间，第五列
                    close_date_list_every_page.append(close_date_list_temp[index_creattime])
                    # 项目名称，第二列
                    projectname_list_every_page.append(projectname_list_temp[index_creattime])
                    # 产品名称，第三列
                    productname_list_every_page.append(productname_list_temp[index_creattime])
                    # 当前状态，第八列
                    status_list_every_page.append(get_status(status_list_temp[index_creattime]))
                    url_list_every_page.append(url)
                    # 总花费时间，第六列
                    total_time_list_every_page.append(total_time_list_temp[index_creattime])


            auditno_list.extend(auditno_list_every_page)
            id_list.extend(id_list_every_page)
            create_date_list.extend(create_date_list_every_page)
            close_date_list.extend(close_date_list_every_page)
            projectname_list.extend(projectname_list_every_page)
            productname_list.extend(productname_list_every_page)
            status_list.extend(status_list_every_page)
            url_list.extend(url_list_every_page)
            total_time_list.extend(total_time_list_every_page)

        print("auditno_list:{}".format(len(auditno_list)))
        print("id_list:{}".format(len(id_list)))
        print("create_date_list:{}".format(len(create_date_list)))
        print("close_date_list:{}".format(len(close_date_list)))
        print("projectname_list:{}".format(len(projectname_list)))
        print("productname_list:{}".format(len(productname_list)))
        print("status_list:{}".format(len(status_list)))
        print("url_list:{}".format(len(url_list)))
        print("total_time_list:{}".format(len(total_time_list)))
        self.updatedisplay("共搜索到{}个符合时间和项目名称要求的评审，请等候抓取数据~~~~".format(len(auditno_list)))

        # 获取每个评审的详细页面信息
        dict_data_detail = {}
        for item_1 in id_list:
            dict_data_detail["{}".format(item_1)] = []
        temp_detail = []
        pool_detail = Pool()
        for index, item_2 in enumerate(url_list):
            temp_detail.append(pool_detail.apply_async(get_detail, args=(item_2, login_session)))
        pool_detail.close()
        pool_detail.join()
        # return link, has_report, report_filename, keywords, test_time
        for item_detail in temp_detail:
            data_detail_temp = item_detail.get()
            if data_detail_temp is not None:
                # if data_detail_temp[0] != "None":
                num = data_detail_temp[0].split("=")[-1]
                index_to_log = id_list.index(num)
                # 评审编号 audit_number
                dict_data_detail["{}".format(num)].append(auditno_list[index_to_log])
                # 评审名称 project name
                dict_data_detail["{}".format(num)].append(projectname_list[index_to_log])
                # 产品名称 product name
                dict_data_detail["{}".format(num)].append(productname_list[index_to_log])
                # 提交时间 create time
                dict_data_detail["{}".format(num)].append(create_date_list[index_to_log])
                # 关闭时间 close time
                dict_data_detail["{}".format(num)].append(close_date_list[index_to_log])
                # 总处理时长 total time
                dict_data_detail["{}".format(num)].append(total_time_list[index_to_log])
                # 测试花费时间 test time
                dict_data_detail["{}".format(num)].append(data_detail_temp[4])
                # 状态 status
                dict_data_detail["{}".format(num)].append(status_list[index_to_log])
                # 是否有报告 has_report
                dict_data_detail["{}".format(num)].append(data_detail_temp[1])
                # 报告名称， report filename
                dict_data_detail["{}".format(num)].append(data_detail_temp[2])
                # 评审要点 keywords
                dict_data_detail["{}".format(num)].append(data_detail_temp[3])
                # NPI备注
                dict_data_detail["{}".format(num)].append(data_detail_temp[5])

        autidno_list_write = []
        projectname_list_write = []
        productname_list_write = []
        create_date_list_write = []
        close_date_list_write = []
        total_time_list_write = []
        test_time_list_write = []
        status_list_write = []
        has_report_list_write = []
        report_filename_list_write = []
        keywords_list_write = []
        npi_comment_list_write = []

        for item_write in dict_data_detail:
            autidno_list_write.append(dict_data_detail[item_write][0])
            projectname_list_write.append(dict_data_detail[item_write][1])
            productname_list_write.append(dict_data_detail[item_write][2])
            create_date_list_write.append(dict_data_detail[item_write][3])
            close_date_list_write.append(dict_data_detail[item_write][4])
            total_time_list_write.append(dict_data_detail[item_write][5])
            test_time_list_write.append(dict_data_detail[item_write][6])
            status_list_write.append(dict_data_detail[item_write][7])
            has_report_list_write.append(dict_data_detail[item_write][8])
            report_filename_list_write.append(dict_data_detail[item_write][9])
            keywords_list_write.append(dict_data_detail[item_write][10])
            npi_comment_list_write.append(dict_data_detail[item_write][11])

        # 写入到本地xlsx文档
        TitleItem = ['评审编号', '评审名称', '项目名称', '提交时间', '最后更新时间', '处理时长', '测试花费时间', '状态', '是否有报告附件', '报告名称', '评审要点', 'NPI备注']
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("评审系统抓取信息-{}.xlsx".format(timestamp))
        SheetOne = WorkBook.add_worksheet('评审系统抓取信息')
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)

        SheetOne.set_column('A:J', 15)
        already_write_list = []
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], formatOne)
        for index_write, item_write in enumerate(autidno_list_write):
            if len(item_write) != 0:
                if item_write not in already_write_list:
                    already_write_list.append(item_write)
                    if self.checkBox_audit.GetValue():
                        SheetOne.write(1 + index_write, 0, item_write, formatOne)
                    if self.checkBox_name.GetValue():
                        SheetOne.write(1 + index_write, 1, projectname_list_write[index_write], formatOne)
                    if self.checkBox_productName.GetValue():
                        SheetOne.write(1 + index_write, 2, productname_list_write[index_write], formatOne)
                    if self.checkBox_submitTime.GetValue():
                        SheetOne.write_datetime(1 + index_write, 3,
                                                datetime.datetime.strptime(create_date_list_write[index_write], '%Y-%m-%d'),
                                                WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                    if self.checkBox_closeTime.GetValue():
                        if re.search(r'\d+', close_date_list_write[index_write]) is None:
                            SheetOne.write(1 + index_write, 4, "评审进行中", formatOne)
                        else:
                            SheetOne.write(1 + index_write, 4, datetime.datetime.strptime(close_date_list_write[index_write], '%Y-%m-%d'), WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                    if self.checkBox_handleTime.GetValue():
                        SheetOne.write(1 + index_write, 5, total_time_list_write[index_write], formatOne)
                    if self.checkBox_totalTestTime.GetValue():
                        SheetOne.write(1 + index_write, 6, test_time_list_write[index_write], formatOne)
                    if self.checkBox_status.GetValue():
                        SheetOne.write(1 + index_write, 7, status_list_write[index_write], formatOne)
                    if self.checkBox_report.GetValue():
                        SheetOne.write(1 + index_write, 8, has_report_list_write[index_write], formatOne)
                        SheetOne.write(1 + index_write, 9, report_filename_list_write[index_write], formatOne)
                    if self.checkBox_summary.GetValue():
                        SheetOne.write(1 + index_write, 10, keywords_list_write[index_write], formatOne)
                    SheetOne.write(1 + index_write, 11, npi_comment_list_write[index_write], formatOne)

        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay(
            "抓到{}个结果！已经将结果写入《评审系统抓取信息-{}.xlsx》，请自行查阅！请点击EXIT退出程序！".format(len(already_write_list), timestamp))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def close(self, event):
        self.Close()

    def newthread(self):
        Thread(target=self.run_all).start()

    def onbutton(self, event):
        self.button_go.Disable()
        self.newthread()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("完成第{}页".format(t))
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText(t)
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = PingShenFrame(None)
    frame.Show()
    app.MainLoop()
