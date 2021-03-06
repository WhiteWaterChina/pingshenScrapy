#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import os
import time
import datetime
from threading import Thread
import wx
import base64
import urllib2
import json


ver="20180407"
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

def get_status(status):
    switcher = {
        "0": "保存".decode('gbk'),
        "1": "提交".decode('gbk'),
        "shenhepeizhi-sq": "售前审核配置".decode('gbk'),
        "shenhe-chanpinjingli":"产品经理审核".decode('gbk'),
        "querenxuanpei-ddy": "订单员确认选配".decode('gbk'),
        "shenhepingshen-yf": "研发接口人审核评审".decode('gbk'),
        "shenhepingshen-csjk": "测试接口人审核评审".decode('gbk'),
        "ceshi-cs": "测试人员测试".decode('gbk'),
        "shenheceshibaogao-yf": "研发接口审核测试报告".decode('gbk'),
        "shenheceshibaogao-xmjl": "项目经理审核测试报告".decode('gbk'),
        "shenheceshibaogao-csfzr": "测试负责人审核测试报告".decode('gbk'),
        "shenheceshibaogao-csjk":"测试接口人审核测试报告".decode('gbk'),
        "xfzl-gc": "工程人员确认是否下发指令".decode('gbk'),
        "100": "关闭".decode('gbk'),
        "101": "异常关闭".decode('gbk')
    }
    return switcher.get(status, status)


class PingShenFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"评审系统信息抓取工具-Version:%s" % ver, pos=wx.DefaultPosition,
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

        self.checkBox_closeTime = wx.CheckBox(self.m_panel1, wx.ID_ANY, u"最后更新时间", wx.DefaultPosition, wx.DefaultSize, 0)
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

    def run(self):
        self.updatedisplay("开始抓取".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        username = self.input_username.GetValue()
        password = base64.b64encode(self.input_password.GetValue())
        start_time = self.input_startdate.GetValue()
        end_time = self.input_enddate.GetValue()
        get_data = requests.session()
        #获取登录页面的lt/execution/eventid信息
        url_login = "http://218.57.146.175/inspurSSO/login"
        headers_login = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.8",
            'cache-control': "max-age=0",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded",
            'host': "218.57.146.175",
            'upgrade-insecure-requests': "1",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        }
        response_login_test = get_data.get(url_login, headers=headers_login).text
        data_soup_tobe_filter = BeautifulSoup(response_login_test, "html.parser")
        lt = data_soup_tobe_filter.find('input',{'name':'lt'})['value']
        execution = data_soup_tobe_filter.find('input',{'name':'execution'})['value']
        eventid = data_soup_tobe_filter.find('input',{'name':'_eventId'})['value']
        #使用以上获取的信息post登录
        headers_base = {
        'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.8",
        'Cache-Control': "max-age=0",
        'Connection': "keep-alive",
        'Content-Length': "125",
        'Content-Type': "application/x-www-form-urlencoded",
        'Host': "218.57.146.175",
        'Origin': "http://218.57.146.175",
        'Referer': "http://218.57.146.175/inspurSSO/login",
        'Upgrade-Insecure-Requests': "1",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        }
        payload_login = {
            'username': "%s" % username,
            'password': "%s" % password,
            'lt': "%s" % lt,
            'execution': "%s" % execution,
            '_eventId': "%s" % eventid
        }
        log_in = get_data.post(url_login,headers=headers_base, data=payload_login)
        #开始获取数据
        headers_data = {
        'Accept': "application/json, text/javascript, */*; q=0.01",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.8",
        'Connection': "keep-alive",
        'Content-Length': "34",
        'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
        'Host': "218.57.146.175",
        'Origin': "http://218.57.146.175",
        'Referer': "http://218.57.146.175/techAudit/auditList/queryAuditList.htm",
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        'X-Requested-With': "XMLHttpRequest"
        }
        #先用1获取最大数据条数
        payload_data_test = "rows=1"
        url_data = "http://218.57.146.175/techAudit/AuditListController/getQueryAuditList.htm?"
        response_data_test = get_data.post(url_data,headers=headers_data, data=payload_data_test)
        total_item = re.search(r'"total":(\d+?),', response_data_test.text).groups()[0]
        #然后使用最大条数直接1页显示
        payload_data = "rows=%s" % total_item
        response_data = get_data.post(url_data,headers=headers_data, data=payload_data).text

        #先获取autidno/id/creattime/updatetime/projectname/productname的原始数据
        auditno_list_temp = re.findall(r'"auditNo":"(\w+)",', response_data)
        id_list_temp = re.findall(r'"id":(\d+),', response_data)
        create_date_list_temp = re.findall(r'"createTime":"(\d+-\d+-\d+)', response_data)#2016-09-23
        lastupdate_date_list_temp = re.findall(r'"updateTime":"(\d+-\d+-\d+)', response_data)
        projectname_list_temp = re.findall(r'"projectName":"(.*?)",', response_data)
        productname_list_temp = re.findall(r'"modeName":"(.*?)",', response_data)
        status_list_temp = re.findall(r'"status":"(.*?)",', response_data)
        total_time_list_temp = re.findall(r'"handleTime":"(.*?)"', response_data)
        #转换要求的开始和结束时间相对1970年1月1日的秒数
        start_date = time.mktime(time.strptime(start_time,'%Y%m%d'))
        end_date = time.mktime(time.strptime(end_time,'%Y%m%d'))

        #过滤一遍，去除时间不符合要求的
        autidno_list = []
        id_list = []
        create_date_list = []
        lastupdate_date_list = []
        projectname_list = []
        productname_list = []
        status_list = []
        url_list = []
        total_time_list = []
        report_list = []
        report_filename_list = []
        keywords_list = []
        test_time_list = []
        for index_creattime, item_creattime in enumerate(create_date_list_temp):
            now_date =  time.mktime(time.strptime(item_creattime,'%Y-%m-%d'))
            #print now_date
            if float(start_date) <= float(now_date) <= float(end_date):
                create_date_list.append(item_creattime)
                autidno_list.append(auditno_list_temp[index_creattime])
                url = "http://218.57.146.175/techAudit/details/viewBill.htm?id=" + id_list_temp[index_creattime]
                id_list.append(id_list_temp[index_creattime])
                if len(lastupdate_date_list_temp[index_creattime]) == 0:
                    lastupdate_date_list.append("评审进行中".decode('gbk'))
                else:
                    lastupdate_date_list.append(lastupdate_date_list_temp[index_creattime])
                projectname_list.append(projectname_list_temp[index_creattime])
                productname_list.append(productname_list_temp[index_creattime])
                status_list.append(get_status(status_list_temp[index_creattime]))
                url_list.append(url)
                total_time_list.append(total_time_list_temp[index_creattime])
        self.updatedisplay("共搜索到%s个符合时间要求的评审，请等候抓取数据~~~~".decode('gbk') % len(autidno_list))
        #分页获取
        headers_data_all = {
        'Accept':"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'Accept-Encoding':"gzip, deflate",
        'Accept-Language':"zh-CN,zh;q=0.8",
        'Connection':"keep-alive",
        'Host':"218.57.146.175",
        'Referer':"http://218.57.146.175/techAudit/welcome.htm",
        'Upgrade-Insecure-Requests':"1",
        'User-Agent':"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"
        }

        for index_url, item_url in enumerate(url_list):
            data_page = get_data.get(item_url, headers=headers_data_all).text
            data_filter = BeautifulSoup(data_page, "html.parser")
            # print data_filter
            self.updatedisplay("正在获取第%s/%s个评审".decode('gbk') % (index_url+1, len(autidno_list)))
            # 获取附件信息
            attachment_temp = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr")
            if len(attachment_temp) <= 1:
                report_list.append("无报告".decode('gbk'))
                report_filename_list.append("None")
            else:
                report_list.append("有报告".decode('gbk'))
               # print data_filter
                filename_temp = []
                for index_report, item_report in enumerate(attachment_temp):
                    if index_report != 0:
                        filename = data_filter.select(".testAttachmenttable > tr:nth-of-type(1) > td:nth-of-type(2) > table:nth-of-type(1) > tr:nth-of-type(2) > td:nth-of-type(1)")[0].contents[0].strip()
                        filename_temp.append(filename)
                filename_write = ";".join(filename_temp)
                report_filename_list.append(filename_write)
            # 评审要点信息
            try:
                keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(9) > td:nth-of-type(1)")[0].get_text()
            except IndexError:
                try:
                    keywords_temp = data_filter.select(".TableCssList > tr:nth-of-type(7) > td:nth-of-type(1)")[0].get_text()
                except IndexError:
                    keywords_temp = "None"
           # print  keywords_temp
            keywords_list.append(keywords_temp)
            # 测试花费时间
            timeTestTemp = []
            try:
                test_item = data_filter.select("body > div:nth-of-type(1) > table:nth-of-type(3) > tbody > tr")
                for item_tr in test_item:
                    item_td = item_tr.select("td:nth-of-type(5)")[0].get_text().strip()
                    if item_td == "测试".decode('gbk'):
                        test_time = item_tr.select("td:nth-of-type(3)")[0].get_text().strip()
                        timeTestTemp.append(test_time)
                if len(timeTestTemp) == 0:
                    timeTestData = "None"
                else:
                    timeTestData = sumtimesplit(timeTestTemp)
                test_time_list.append(timeTestData)
            except IndexError:
                test_time_list.append("None")


        # 如下是数据处理，与浏览器不再发生关系
        TitleItem = ['评审编号'.decode('gbk'), '评审名称'.decode('gbk'), '项目名称'.decode('gbk'), '提交时间'.decode('gbk'),
                     '最后更新时间'.decode('gbk'), '处理时长'.decode('gbk'), '测试花费时间'.decode('gbk'), '状态'.decode('gbk'),
                     '是否有报告附件'.decode('gbk'), '报告名称'.decode('gbk'), '评审要点'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("评审系统抓取信息-%s.xlsx".decode('gbk') % timestamp)
        SheetOne = WorkBook.add_worksheet('评审系统抓取信息'.decode('gbk'))
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)

        SheetOne.set_column('A:J', 14)
        already_write_list = []
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], formatOne)
        for index_write, item_write in enumerate(autidno_list):
            if item_write not in already_write_list:
                already_write_list.append(item_write)
                if self.checkBox_audit.GetValue():
                    SheetOne.write(1 + index_write, 0, item_write, formatOne)
                if self.checkBox_name.GetValue():
                    SheetOne.write(1 + index_write, 1, projectname_list[index_write], formatOne)
                if self.checkBox_productName.GetValue():
                    SheetOne.write(1 + index_write, 2, productname_list[index_write], formatOne)
                if self.checkBox_submitTime.GetValue():
                    SheetOne.write_datetime(1 + index_write, 3, datetime.datetime.strptime(create_date_list[index_write], '%Y-%m-%d'), WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                if self.checkBox_closeTime.GetValue():
                    SheetOne.write(1 + index_write, 4, datetime.datetime.strptime(lastupdate_date_list[index_write], '%Y-%m-%d'), WorkBook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
                if self.checkBox_handleTime.GetValue():
                    SheetOne.write(1 + index_write, 5, total_time_list[index_write], formatOne)
                if self.checkBox_totalTestTime.GetValue():
                    SheetOne.write(1 + index_write, 6, test_time_list[index_write], formatOne)
                if self.checkBox_status.GetValue():
                    SheetOne.write(1 + index_write, 7, status_list[index_write], formatOne)
                if self.checkBox_report.GetValue():
                    SheetOne.write(1 + index_write, 8, report_list[index_write], formatOne)
                    SheetOne.write(1 + index_write, 9, report_filename_list[index_write], formatOne)
                if self.checkBox_summary.GetValue():
                    SheetOne.write(1 + index_write, 10, keywords_list[index_write], formatOne)
        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay("抓到%s个结果！已经将结果写入《评审系统抓取信息-%s.xlsx》，请自行查阅！请点击EXIT退出程序！".decode('gbk') % (len(autidno_list), timestamp))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def close(self, event):
        self.Close()

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