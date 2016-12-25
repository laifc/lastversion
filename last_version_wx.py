# coding=utf-8
'''
Created on Feb 17, 2016

@author: 95
'''
''''
功能
获取     softlist=["QQ","360","qq pinyin","qq guanjia","qq game","kingsoft","xunlei","baofeng","pps","ali wangwang","pptv","baidu music","wps"]
13个软件最新版本的版本号和更新时间  保存到excel中 

特点 
默认保存文件位置   os.getcwd()+"\\last_version.xls"
可以根据爬虫时间段增加excel记录
13个 软件中   kingsoft  ali wangwang  暂时不支持爬取信息  qq中获取的版本信息不够完善，有多余信息

后期可以考虑将配置爬虫规则配置到配置文件中，必要时可以修改

'''
import urllib2

import xlrd  #读excel
import xlwt

from xlutils.copy import copy  #用于复制表格
import time
import wx, os, sys

import re
import chardet
import gzip
from StringIO import StringIO


class MyFrame(wx.Frame):
	#初始化数据
	def __init__(self):
		wx.Frame.__init__(self, None, -1, u'最新版本信息获取 V1.0', size=(450, 180),
						  style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)  #style=wx.DEFAULT_FRAME_STYLE
		self.panel = wx.Panel(self, -1)
		#设置背景颜色
		self.panel.SetBackgroundColour("")

		g_name = u"获取13个软件最新版本信息，保存到excel中"
		g_label = wx.StaticText(self.panel, -1, g_name, pos=(10, 15), size=(450, 25), style=wx.ALIGN_LEFT)

		self.pathBtn = wx.Button(self.panel, -1, u"选择保存目录", size=(80, 25), pos=(10, 60))
		self.path_basicText = wx.TextCtrl(self.panel, -1, value="", size=(300, 25), pos=(100, 60))
		#添加确认按钮
		button_name1 = u"开始"
		self.startBtn = wx.Button(self.panel, -1, button_name1, size=(60, 25),
								  pos=(100, 110))  #这个startBtn  前面加没加self.有什么区别
		#添加确认按钮
		button_name1 = u"关闭"
		self.closeBtn = wx.Button(self.panel, -1, button_name1, size=(60, 25),
								  pos=(250, 110))  #这个startBtn  前面加没加self.有什么区别

		'''
        #-----------------------分割线
        line = wx.StaticLine(self.panel,-1,(8,185),(450,3))   #  分割线  pos 和size                 

        #搜索结果
        wx.StaticText(self.panel, -1,u'下载状态：',pos=(10,200),size=(80,50),style=wx.ALIGN_LEFT)
        self.statue_basicText = wx.TextCtrl(self.panel, -1, size=(340, 180),pos=(100,200),style=wx.TE_MULTILINE) #|wx.HSCROLL                    
        #-----------------------      
        '''
		self.path_basicText.SetValue(os.getcwd() + "\\last_version.xls")  #默认值c:\\Last_Version\\last_version.xls


		#按钮事件
		#self.button.Bind(wx.EVT_BUTTON ,self.post_url)  #这个方法和 下面的方法有什么区别  应该是一样的  self后面的button对应下一句的startBtn
		self.Bind(wx.EVT_BUTTON, self.OnPathButton, self.pathBtn)
		self.Bind(wx.EVT_BUTTON, self.OnStartButton, self.startBtn)
		self.Bind(wx.EVT_BUTTON, self.OnCloseButton, self.closeBtn)
		#初始化数据
		self.softlist = ["QQ", "360", "qq pinyin", "qq guanjia", "qq game", "kingsoft", "xunlei", "baofeng", "pps",
						 "ali wangwang", "pptv", "baidu music", "wps", \
						 "APP_taobao", u"APP_百度外卖", u"APP_美图秀秀", u"APP_搜狗输入法", u"APP_滴滴出行", u"APP_酷狗音乐", u"APP_芒果tv",
						 u"APP_腾讯视频", u"APP_YY", u"APP_高德地图"]
		self.data = {
			"QQ":
				{
					"support": "yes",
					"Geturl": "http://im.qq.com/pcqq/",
					"re_soft": r'QQ(.+)\.exe',
					"re_lasttime": r'更新日期： (.+)</span>'
				},
			"360":
				{
					"support": "yes",
					"Geturl": "http://www.360.cn/weishi/updatelog.html",
					"re_soft": r'<h2><span class="tit">(.+)</span><span',
					"re_lasttime": r'</span><span class="date">(.+)</span></h2>'
				},
			"qq pinyin":
				{
					"support": "no",  #宁可放过也不要抓错了
					"Geturl": "http://qq.pinyin.cn/",
					"re_soft": r'<p>版本：(.+)</p>',  #<br /><b   注释的内容也有这种格式的，会导致抓错了
					"re_lasttime": r'<p>更新时间：(.+)</p>'
				},
			"qq guanjia":
				{
					"support": "yes",
					"Geturl": "http://guanjia.qq.com/product/update.html",
					"re_soft": r'class="" title=" ">(.+) </a> </h2>',
					"re_lasttime": r'</a> </h2><span class="">(.+)</span></div>'
				},
			"qq game":
				{
					"support": "yes",
					"Geturl": "http://qqgame.qq.com/download.shtml",
					"re_soft": r'<h2 class="intro-title">(.+)</h2>',
					"re_lasttime": r'MB</span><span>更新日期：(.+)</span></p>'
				},
			"kingsoft":
				{
					"support": "no",
					"Geturl": "http://qqgame.qq.com/download.shtml",
					"re_soft": r'class="" title=" ">(.+) </a> </h2>',
					"re_lasttime": r'</a> </h2><span class="">(.+)</span></div>'
				},
			"xunlei":
				{
					"support": "yes",
					"Geturl": "http://dl.xunlei.com/xl7.9/intro.html",
					"re_soft": r'<p>版本：(.+)</p><p>支持系统：',
					"re_lasttime": r'<p>版本：(.+)</p><p>支持系统：'
				},
			"baofeng":
				{
					"support": "yes",
					"Geturl": "http://home.baofeng.com/history.html",
					"re_soft": r'exe">(.+)</a>',
					"re_lasttime": r'</a>【更新时间：(.+)】</dt>'
				},
			"pps":
				{
					"support": "yes",
					"Geturl": "http://app.iqiyi.com/pc/player/index.html#pcplayer <http://app.iqiyi.com/pc/player/index.html",
					"re_soft": r'<span>最新版本： (.+)</span><span>',
					"re_lasttime": r'</span><span>发布时间：(.+)</span>'
				},
			"ali wangwang":  #不行
				{
					"support": "no",
					"Geturl": "http://download.ww.taobao.com/AliIm_taobao.php?spm=0.0.0.0.YwoAEz",
					"re_soft": r'\((.+)\)\.exe',
					"re_lasttime": r''  #没有时间
				},
			"pptv":
				{
					"support": "yes",
					"Geturl": "http://app.pptv.com/pg_get_clt",
					"re_soft": r'更新<span>(.+)</span><span>',
					"re_lasttime": r'<p>(.+)更新'
				},
			"baidu music":
				{
					"support": "yes",
					"Geturl": "http://music.baidu.com/pc/index.html",
					"re_soft": r'版本：(.+)</span>',
					"re_lasttime": r'更新日期：(.+)</span>'
				},
			"wps":
				{
					"support": "yes",
					"Geturl": "http://www.wps.cn/product/preview/",
					"re_soft": r'/download/W\.P\.S.(.+)\.exe"  title="免费',
					"re_lasttime": r'<span class="txt_date">(.+)</span>'
				},

			"APP_taobao":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.taobao.taobao",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_百度外卖":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.baidu.lbs.waimai",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_美图秀秀":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.mt.mtxx.mtxx",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_搜狗输入法":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.sohu.inputmethod.sogou",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_滴滴出行":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.sdu.didi.psnger",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_酷狗音乐":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.kugou.android",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_芒果tv":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.hunantv.imgo.activity",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},

			u"APP_腾讯视频":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.tencent.qqlive",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_YY":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.duowan.mobile",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				},
			u"APP_高德地图":
				{
					"support": "yes",
					"Geturl": "http://www.wandoujia.com/apps/com.autonavi.minimap",
					"re_soft": r'<dt>版本</dt>\n\s*<dd>(.+)</dd>',
					"re_lasttime": r'">(.+)</time></dd>'
				}
		}
		#print self.data

	def OnStartButton(self, evt):
		self.savepath = self.path_basicText.GetValue()
		if (self.savepath != ""):
			mySpider = LastvSpider(self.data, self.softlist, self.savepath)
			mySpider.main_search()
			wx.MessageBox(u'获取信息完成', caption=u"完成", style=wx.OK)
		else:
			wx.MessageBox(u'请确认保存路径不为空', caption=u"错误信息", style=wx.OK)

	def OnCloseButton(self, evt):
		sys.exit(0)

	def OnPathButton(self, evt):
		'''
                        弹出文件保存对话框
        '''
		file_wildcard = "Excel2003(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx"
		dlg = wx.FileDialog(self,
							"Save  as ...",
							os.getcwd(),
							style=wx.SAVE | wx.CHANGE_DIR,  #|
							wildcard=file_wildcard)
		if dlg.ShowModal() == wx.ID_OK:
			filename = dlg.GetPath()
			if not os.path.splitext(filename)[1]:  #如果没有文件名后缀
				filename = filename + '.xls'
			self.path_basicText.SetValue(filename)
		dlg.Destroy()


class LastvSpider:
	# 申明相关的属性
	def __init__(self, data_dict, softlist, savepath):
		self.data_dict = data_dict
		self.softlist = softlist
		self.savepath = savepath

		print u'已经启动爬虫，咔嚓咔嚓'

	#主函数
	def main_search(self):
		softlist = []
		lasttimelist = []
		for softname in self.softlist:
			(soft, lasttime) = self.get_search(softname)
			softlist.append(soft)
			lasttimelist.append(lasttime)

		self.save_excel(self.savepath, softlist, lasttimelist)

	#保存的excel  filename不一定存在，   存在就正常打开，然后返回excel中的数据，不存在就完善表头等信息，再返回数据
	def open_excel(self, filename):

		try:
			data = xlrd.open_workbook(filename, formatting_info=True)
			return data
		except Exception, e:
			#print str(e)
			workbook = xlwt.Workbook()
			sheet = workbook.add_sheet("Sheet Name")
			i = 2
			sheet.write(0, 0, u"软件")
			for softname in self.softlist:
				sheet.write(i, 0, softname)
				i = i + 1
			workbook.save(filename)
			data = xlrd.open_workbook(filename, formatting_info=True)
			return data

	#数据保存到excel
	def save_excel(self, filename, softlist, lasttimelist):

		dataexcel = self.open_excel(filename)  #返回保存位置excel中的数据，
		table = dataexcel.sheets()[0]
		rowdata0 = table.row_values(0)  #
		#rowdata1=table.row_values(1)
		print rowdata0, len(rowdata0)

		newWb = copy(dataexcel)
		print newWb;  #<xlwt.Workbook.Workbook object at 0x000000000315F470>
		newWs = newWb.get_sheet(0)
		t1 = time.strftime("%Y-%m-%d", time.localtime(time.time()))  # %H-%M-%S"
		newWs.write(0, len(rowdata0), t1)
		newWs.write(1, len(rowdata0), u"版本")
		newWs.write(1, len(rowdata0) + 1, u"时间")
		print u"匹配到的结果列表：", lasttimelist
		i = 1
		n = 1
		for soft in softlist:
			i = i + 1
			newWs.write(i, len(rowdata0), soft)
		for lasttime in lasttimelist:
			n = n + 1
			newWs.write(n, len(rowdata0) + 1, lasttime)
		print "write new values ok"
		print filename
		newWb.save(filename)
		print "save with same name ok"


	def get_search(self, solfname):
		self.support = self.data_dict[solfname]["support"]
		self.myUrl = self.data_dict[solfname]["Geturl"]
		self.re_soft = self.data_dict[solfname]["re_soft"]
		self.re_lasttime = self.data_dict[solfname]["re_lasttime"]

		print self.myUrl  #,self.re_soft,self.re_lasttime

		if self.support == "yes":
			req = urllib2.Request(self.myUrl)  #,data=self.postdata,headers = self.Gethearderdict1
			try:
				resp = urllib2.urlopen(req)
			except:
				#logging.info(u"查询 帐号"+str(custAcctid)+u"出错。错误代码;")
				print "open url error"
				return "requsert error", "requsert error"
			data = self.deal_gzip(resp)  #2312编码
			data = self.replaceTab(data)

			soft = self.re_data(data, self.re_soft)
			lasttime = self.re_data(data, self.re_lasttime)

			print  soft, lasttime
			return soft.decode("utf-8"), lasttime.decode("utf-8")
		else:
			return "", ""

	def replaceTab(self, html):
		replacetab = [("&lt;", "<"), ("&gt;", ">"), ("&amp;", "&"), ("&amp;", "\""), ("&nbsp;", " ")]
		for t in replacetab:
			html = html.replace(t[0], t[1])
			#print html
		return html

	def re_data(self, html, re_rule):
		htmlre = re.compile(re_rule)
		job = htmlre.findall(html)
		#g=r"<p>版本：(.+)</p>"
		#h = r'<p>更新时间：(.+)</p>'
		if (job != None) and (job != []):
			#if (re_rule == g or re_rule == h):
			#	return job[1]
			return job[0]
		else:
			print u"无匹配值"
			return ""


	def deal_gzip(self, resp):
		if resp.info().get('Content-Encoding') == 'gzip':
			#print "gzip"
			buf = StringIO(resp.read())
			f = gzip.GzipFile(
				fileobj=buf)  # gzip.GzipFile(mode="rb", fileobj=open('d:\\test\\sitemap.log.gz', 'rb'))
			data = f.read()
		else:
			#print u"正常"
			data = resp.read()

		'''
			if resp.info().getparam('charset')  == 'GB2312':    #resp.info().getparam('charset')  可能是none
				data = data.decode("gbk").encode("utf-8")
			'''

		result = chardet.detect(data)  #detect（）检测方法 返回的是字典类型 'confidence':0.99,'encoding':'utf-8'  表示有百分之99的概率编码类型是utf-8
		if result['encoding'] == 'GB2312':
			data = data.decode("gbk").encode("utf-8")

			#print resp.info().getparam('charset')    #这个地方无法输出
		return data


if __name__ == "__main__":
	app = wx.App()  #实例化APP，
	frame = MyFrame()  #frame的实例
	frame.Show()
	app.MainLoop()  #wxpython的启动函数



    