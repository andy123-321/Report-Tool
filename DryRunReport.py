import csv
import os
import json
import xlwt
import datetime

#color index
COL_GREEN = 11
COL_ORANGE = 52
COL_YELLOW = 13
COL_BLUE = 71
COL_GRAY = 22
COL_RED = 10
# COL_GRAY = '#C0C0C0'
COL_LIGHT_GRAY = 41


LINE_Title       = 0
LINE_Project     = 1
LINE_DisplayName = 2
LINE_PDCAName    = 3
LINE_UpperLimit  = 4
LINE_LowerLimit  = 5
LINE_Unit        = 6

class SaveExcel(object):
	def __init__(self):
		super(SaveExcel, self).__init__()
		self.workbook = xlwt.Workbook()
		self.font_w_max_list = []
		self.sheet = ''
		# self.ProductDict = {}

	def setStyle(self,name=None, height=None,font_color=None, bg_color = None,bold=False,horz= 0x02,vert = 0x01,wrap=1):
		style = xlwt.XFStyle()  # 初始化样式

		font = xlwt.Font()  # 为样式创建字体
		# 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
		if name:
			font.name = name
		# 设置字体颜色
		if font_color:
			font.colour_index = font_color
		# 字体大小
		if height:
			font.height = height
		# 定义格式
		style.font = font
		# borders.left = xlwt.Borders.THIN
		# NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
		# THIN： 官方代码中THIN所表示的值为1，边框为实线
		borders = xlwt.Borders()
		if bold:
			borders.left = xlwt.Borders.THIN
			borders.right = xlwt.Borders.THIN
			borders.top = xlwt.Borders.THIN
			borders.bottom = xlwt.Borders.THIN

		# 定义格式
		style.borders = borders
		if bg_color:
			# 设置背景颜色
			pattern = xlwt.Pattern()
			# 设置背景颜色的模式
			pattern.pattern = xlwt.Pattern.SOLID_PATTERN
			# 背景颜色
			pattern.pattern_fore_colour = bg_color
			style.pattern = pattern

		# 设置单元格对齐方式
		alignment = xlwt.Alignment()
		# 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
		alignment.horz = horz
		# 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
		alignment.vert = vert
		# 设置自动换行
		alignment.wrap = wrap
		style.alignment = alignment
		return style

	def writeRow(self,sheet,startpoint,content_list,cell_style_list):
		for i in range(startpoint,len(content_list)+startpoint):
			str(content_list[i]).strip()
			sheet.write(self.row_count,i, content_list[i-startpoint],cell_style_list[i-startpoint])
		self.row_count += 1

	def writeLineRow(self,sheet,row_count,content_list,cell_style_list):
		for i in range(0,len(content_list)):
			str(content_list[i]).strip()
			sheet.write(row_count,i, content_list[i],cell_style_list[i])

	def writeLineColumnRow(self,sheet,row_count,column,content_list,cell_style_list):
		for i in range(0,len(content_list)):
			str(content_list[i]).strip()
			sheet.write(row_count,column, content_list[i],cell_style_list[i])
			column += 1

	def saveAction(self,ReportPath):
		print("\nsave report to excel file...")
		# current_path = self.getPath(0)
		#print(current_path)
		filename = 'DryRunReport.xls'
		# save_file_name = os.path.join(self.filePath,filename)
		save_file_name = ReportPath + "/"+filename
		# print(save_file_name)
		self.workbook.save(save_file_name)
		print(save_file_name,'\n')
		print("save done.\n\n")

	def Outputsummary(self,summuryDict):
		print(summuryDict,'\n')
		self.row_count = 0
		sheet = self.workbook.add_sheet("Dry run summary")

		titlelist = ['Dry run date','Bundle','Station','Purpose','Overlay Version','Path','Diags','BBFW','BBlib','NFC','PertOS','Phleet','Rose FW','BT LPEM FW','WIFI FW','BT FW','BT PHY FW','WIFI PHY FW','Project','Config','Vendor','Result','Fail Symptom','Test time(s)','Radar','DRI']

		#title
		style1=self.setStyle(bold = True,bg_color=COL_GRAY,height=250)
		cell_style_list = [style1]*len(titlelist)

		# 设置行宽
		coumn_count=0
		coumnWidthList = [13,24,15,10,20,10,16,20,27,28,20,20,25,15,15,8,13,22,10,10,10,10,10,10,10,10]
		for item in coumnWidthList:
			sheet.col(coumn_count).width = item*256
			coumn_count += 1

		self.writeRow(sheet,0,titlelist,cell_style_list)

		styletall = self.setStyle(height=850)
		cell_style_listtall = [styletall]
		self.writeLineColumnRow(sheet,0,30+30,' ',cell_style_listtall)

		style2=self.setStyle(bold = True,height=11*20)
		style21=self.setStyle(bold = True,horz=0x01,height=11*20)
		cell_style_list = [style2]*30
		

		styletall1 = self.setStyle(height=400)
		cell_style_listtall1 = [styletall1]

		count = 0
		datalist = []
		for title in titlelist:
			if title in summuryDict.keys():
				if len(summuryDict[title])>23:
					cell_style_list[count]=style21
				datalist.append(summuryDict[title])
			else:
				datalist.append(' ')
			count += 1
		self.writeRow(sheet,0,datalist,cell_style_list)
		self.writeLineColumnRow(sheet,1,30+30,' ',cell_style_listtall1)

	def OutputDryrunDatail(self,dataDict):
		self.row_count = 0

		# 按照project整理product
		ProductDict = {}
		for Product in dataDict.keys():
			if Product[0:5] not in ProductDict.keys():
				ProductDict[Product[0:5]] = [Product]
				stationName = dataDict[Product]['stationName']
			else:
				ProductDict[Product[0:5]].append(Product)

		sheetName = stationName + ' DryRunReport'
		sheet = self.workbook.add_sheet(sheetName)

		summuryDict = {}
		ProjiectList = list(set(ProductDict.keys()))
		for Project in ProjiectList:
			# ProductdataDict 传递同一个project当中每个项目的行数
			ProductdataDict = {}
			# 获取每个product当中的report数据
			for Product in ProductDict[Project]:
				if Project not in ProductdataDict.keys():
					ProductdataDict[Project] = {}
					ProductdataDict[Project]['sncount'] = 0
					ProductdataDict[Project]['failsn'] = 0
					ProductdataDict[Project]['retestsn'] = 0
				ProductdataDict = self.getProductDict(Product,dataDict,ProductdataDict[Project])
				summuryDict[Product] = ProductdataDict
			# 更新同一个project(X2010)当中的sncount、failsn、retestsn 对准同一个project当中的行
			for Product in ProductDict[Project]:
				summuryDict[Product][Project] = ProductdataDict[Project]
		# 写入dryrun report
		for Product in summuryDict.keys():
			Project = Product[0:5]
			Productlist = ProductDict[Project]
			self.addProduct(sheet,summuryDict[Product],Project,ProjiectList.index(Project),Productlist.index(Product))

	
	def getProductDict(self,Product,dataDict,ProjectDict):
		listDict = {}
		snDict = {}
		truefailDict = {}
		RetestDict = {}

		totalTest = 0
		totalFail = 0
		totalRetest = 0

		passcount = 0
		sumtesttime = 0
		index = 0

		for sn in dataDict[Product]['SerialNumber']:
			if dataDict[Product]['Pass/Fail'][index] == 'PASS':
				passcount += 1
				sumtesttime += int(dataDict[Product]['testTime'][index])

			FixtureID = dataDict[Product]['fixtureid'][index]
			stationID = dataDict[Product]['stationID'][index]
			CFG = dataDict[Product]['cfg'][index]

			if 'Product' not in listDict.keys():
				listDict['Product'] = []
			if 'stationID' not in listDict.keys():
				listDict['stationID'] = []
			if 'FixtureID' not in listDict.keys():
				listDict['FixtureID'] = []
			if 'cfg' not in listDict.keys():
				listDict['cfg'] = []

			if Product not in listDict['Product']:
				listDict['Product'].append(Product)
			if stationID not in listDict['stationID']:
				listDict['stationID'].append(stationID)
			if FixtureID not in listDict['FixtureID']:
				listDict['FixtureID'].append(FixtureID)
			if CFG not in listDict['cfg']:
				listDict['cfg'].append(CFG)

			testCount = len(dataDict[Product]['SerialNumber'])
					
			if sn not in snDict.keys():
				snDict[sn] = {}
				snDict[sn]['Test count'] = 1
				snDict[sn]['Fail count'] = 0
				snDict[sn]['Retest count'] = 0
				snDict[sn]['index'] = [index]
				snDict[sn]['result'] = [dataDict[Product]['Pass/Fail'][index]]
				snDict[sn]['Channel ID'] = [dataDict[Product]['Channel ID'][index]]
				snDict[sn]['cfg'] = [CFG]
				snDict[sn]['FailMsg'] = [dataDict[Product]['FailMsg'][index]]
			else:
				snDict[sn]['index'].append(index)
				snDict[sn]['result'].append(dataDict[Product]['Pass/Fail'][index])
				snDict[sn]['FailMsg'].append(dataDict[Product]['FailMsg'][index])
				snDict[sn]['Test count'] += 1
				if dataDict[Product]['Channel ID'][index] not in snDict[sn]['Channel ID']:
					snDict[sn]['Channel ID'].append(dataDict[Product]['Channel ID'][index])
				if CFG not in snDict[sn]['cfg']:
					snDict[sn]['cfg'].append(CFG)
			index += 1
				
		for sn in snDict.keys():
			if 'FAIL' == snDict[sn]['result'][-1]:
				snDict[sn]['Fail count'] += 1
				if sn not in truefailDict.keys():
					truefailDict[sn] = [Product,snDict[sn]['cfg'],sn,snDict[sn]['FailMsg'][-1]]
			elif 'FAIL' in snDict[sn]['result']:
				snDict[sn]['Retest count'] += 1
				failindex = snDict[sn]['result'].index('FAIL')
				if sn not in RetestDict.keys():
					RetestDict[sn] = [Product,snDict[sn]['cfg'],sn,snDict[sn]['FailMsg'][failindex]]

			FailRate = snDict[sn]['Fail count']/snDict[sn]['Test count']
			RetestRate = snDict[sn]['Retest count']/snDict[sn]['Test count']
			ChannelID = ''
			for channelID in snDict[sn]['Channel ID']:
				ChannelID = ChannelID + channelID.strip() + '\n'

			snDict[sn]['datalist'] = [sn,ChannelID.strip(),snDict[sn]['Test count'],snDict[sn]['Fail count'],snDict[sn]['Retest count'],FailRate,RetestRate,'{:.2f}%'.format(FailRate*100),'{:.2f}%'.format(RetestRate*100)]

			totalTest += snDict[sn]['Test count']
			totalFail += snDict[sn]['Fail count']
			totalRetest += snDict[sn]['Retest count']

		totalFailRate = totalFail/totalTest
		totalRetestRate = totalRetest/totalTest

		totallist = ['','Total',totalTest,totalFail,totalRetest,totalFailRate,totalRetestRate,'{:.2f}%'.format(totalFailRate*100),'{:.2f}%'.format(totalRetestRate*100)]

		stationIDlist = ''
		for stationID in listDict['stationID']:
			stationIDlist = stationID + '\n' + stationIDlist

		fixtureIDlist = ''
		for fixtureID in listDict['FixtureID']:
			fixtureIDlist =  fixtureID + '\n' + fixtureIDlist
		cfglist = ''
		for cfg in listDict['cfg']:
			cfglist = cfglist + cfg + '\n'

		Productlist = [listDict['Product'],stationIDlist.strip(),fixtureIDlist.strip(),cfglist.strip()]
		Productlist.append(str(totalFail)+'F/'+str(testCount)+'T')
		Productlist.append(str(totalRetest)+'R/'+str(testCount)+'T')

		station = dataDict[Product]['stationName']
		overlayVersion = dataDict[Product]['overlayVersion']
		titleSummary = '1.Station:\n        '+station+'\n2.OVL:\n        '+overlayVersion+'\n3.Path#2:'
		Config = '\n4.Config:\n        '
		for cfg in list(set(dataDict[Product]['cfg'])):
			Config = Config + cfg +'   '
		
		result = '\n5.Result:\n        '+Product+'  '+str(totalFail)+'F/'+str(testCount)+'T'+';  '+str(totalRetest)+'R/'+str(testCount)+'T'
		avgtesttime = sumtesttime/passcount
		testtime = '\n6.Test Time:\n        '+Product+':  '+str(avgtesttime) +' s'
		titleSummary = titleSummary+Config + result + testtime
		# print(titleSummary) 
		ProductdataDict = {}
		ProductdataDict['titleSummary'] = titleSummary
		ProductdataDict['Productlist'] = Productlist
		ProductdataDict['snDict'] = snDict
		ProductdataDict['truefailDict'] = truefailDict
		ProductdataDict['RetestDict'] = RetestDict
		ProductdataDict['totallist'] = totallist
		ProductdataDict[Product[0:5]] = {}

		if len(snDict.keys()) < ProjectDict['sncount']:
			ProductdataDict[Product[0:5]]['sncount'] = ProjectDict['sncount']
		else:
			ProductdataDict[Product[0:5]]['sncount'] = len(snDict.keys())

		if len(truefailDict.keys()) < ProjectDict['failsn']:
			ProductdataDict[Product[0:5]]['failsn'] = ProjectDict['failsn']
		else:
			ProductdataDict[Product[0:5]]['failsn'] = len(truefailDict.keys())

		if len(RetestDict.keys()) < ProjectDict['retestsn']:
			ProductdataDict[Product[0:5]]['retestsn'] = ProjectDict['retestsn']
		else:
			ProductdataDict[Product[0:5]]['retestsn'] = len(RetestDict.keys())

		return ProductdataDict

	def addProduct(self,sheet,ProductDict,Project,ProjectIndex,ProductIndex):
		# row: 7为固定行：summary；product title；product list；sip sn；total；failsn title；retest title
		row_count = ProjectIndex * (7 + ProductDict[Project]['sncount'] + ProductDict[Project]['failsn'] + ProductDict[Project]['retestsn'])
		# column：7为title数量
		column = ProductIndex * 7

		# 设置各种表格样式
		style_summary = self.setStyle(bold = True,bg_color=COL_BLUE,height=230,horz=0x01,vert=0x00,wrap=1)
		style_summary_tall = self.setStyle(height=2800)

		style_title=self.setStyle(bold = True,bg_color=COL_GRAY,height=240)
		style_title_tall = self.setStyle(height=420)
		title_style_list = [style_title]*7

		style_text=self.setStyle(height=220,bold = True)
		style_text_tall = self.setStyle(height=320)
		text_style_list = [style_text]*7

		style_Green=self.setStyle(bold = True,bg_color=COL_GREEN,height=240)
		Green_style_list = [style_Green]*2

		style_Red=self.setStyle(bold = True,bg_color=COL_RED,height=240)
		Red_style_list = [style_Red]*2

		style_Yellow=self.setStyle(bold = True,bg_color=COL_YELLOW,height=240)
		Yellow_style_list = [style_Yellow]*2

		# write summary
		sheet.write_merge(row_count,row_count,column,column+6,ProductDict['titleSummary'],style_summary)
		# write row tall
		if ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_summary_tall])
		row_count += 1

		# write Product
		title_Product = ['Product','Fixture ID','CFG','Result','Retest']
		columnwidth = [22,22,22,20,20,13,13]
		coumn_count = 0
		for width in columnwidth:
			sheet.col(coumn_count+column).width = width*256
			coumn_count += 1

		sheet.write(row_count,column, title_Product[0],style_title)
		self.writeLineColumnRow(sheet,row_count,column+3,title_Product[1:],title_style_list)
		sheet.write_merge(row_count,row_count,column+1,column+2,'Station ID',style_title)
		if ProjectIndex == 0 and ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_title_tall])
		row_count += 1

		sheet.write(row_count,column, ProductDict['Productlist'][0],style_text)
		sheet.write_merge(row_count,row_count,column+1,column+2,ProductDict['Productlist'][1].strip(),style_text)
		self.writeLineColumnRow(sheet,row_count,column+3,ProductDict['Productlist'][2:],text_style_list)
		if ProjectIndex == 0 and ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+60,' ',text_style_list)
		row_count += 1

		title_SIPSN = ['SIP SN','Channel ID','Test count','Fail count','Retest count','Fail Rate','Retest Rate']
		self.writeLineColumnRow(sheet,row_count,column,title_SIPSN,title_style_list)
		if ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_title_tall])
		row_count += 1

		snDict = ProductDict['snDict']
		for sn in snDict.keys():
			self.writeLineColumnRow(sheet,row_count,column,snDict[sn]['datalist'][0:5],text_style_list)
			if snDict[sn]['datalist'][5] > snDict[sn]['datalist'][6]:
				if snDict[sn]['datalist'][5] < 0.03:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],Green_style_list)
				elif snDict[sn]['datalist'][5] < 0.99 and snDict[sn]['datalist'][5] > 0.03:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],Yellow_style_list)
				else:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],Red_style_list)
				self.writeLineColumnRow(sheet,row_count,column+6,' ',text_style_list)
			else:
				if snDict[sn]['datalist'][5] < 0.03:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],Green_style_list)
				elif snDict[sn]['datalist'][5] < 0.99 and snDict[sn]['datalist'][5] > 0.03:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],Yellow_style_list)
				else:
					self.writeLineColumnRow(sheet,row_count,column+5,snDict[sn]['datalist'][7:8],cell_style_listR)

				if snDict[sn]['datalist'][6] < 0.03:
					self.writeLineColumnRow(sheet,row_count,column+6,snDict[sn]['datalist'][8:],Green_style_list)
				elif snDict[sn]['datalist'][6] < 0.99 and snDict[sn]['datalist'][6] > 0.03:
					self.writeLineColumnRow(sheet,row_count,column+6,snDict[sn]['datalist'][8:],Yellow_style_list)
				else:
					self.writeLineColumnRow(sheet,row_count,column+6,snDict[sn]['datalist'][8:],Red_style_list)
			if ProductIndex == 0:
				self.writeLineColumnRow(sheet,row_count,column+30,' ',text_style_list)
				
			row_count += 1

		listRow = ['','','','','','','']
		# print(Project,ProductDict[Project]['sncount'],ProductDict[Project]['failsn'],ProductDict[Project]['retestsn'],len(snDict.keys()),'**********')
		for i in range(ProductDict[Project]['sncount']-len(snDict.keys())):
			self.writeLineColumnRow(sheet,row_count,column,listRow,text_style_list)
			row_count += 1

		self.writeLineColumnRow(sheet,row_count,column,ProductDict['totallist'][0:5],text_style_list)

		if ProductDict['totallist'][5] < 0.03:
			self.writeLineColumnRow(sheet,row_count,column+5,ProductDict['totallist'][7:8],Green_style_list)
		elif ProductDict['totallist'][5] < 0.99 and ProductDict['totallist'][5] > 0.03:
			self.writeLineColumnRow(sheet,row_count,column+5,ProductDict['totallist'][7:8],Yellow_style_list)
		else:
			self.writeLineColumnRow(sheet,row_count,column+5,ProductDict['totallist'][7:8],Red_style_list)
		
		if ProductDict['totallist'][6] < 0.03:
			self.writeLineColumnRow(sheet,row_count,column+6,ProductDict['totallist'][8:],Green_style_list)
		elif ProductDict['totallist'][6] < 0.99 and ProductDict['totallist'][6] > 0.03:
			self.writeLineColumnRow(sheet,row_count,column+6,ProductDict['totallist'][8:],Yellow_style_list)
		else:
			self.writeLineColumnRow(sheet,row_count,column+6,ProductDict['totallist'][8:],Red_style_list)

		row_count += 1

		title_FailSN = ['Product','Config','Fail SN','Radar']
		self.writeLineColumnRow(sheet,row_count,column,title_FailSN[0:3],title_style_list)
		sheet.write_merge(row_count,row_count,column+3,column+5,'Fail item',style_title)
		sheet.write(row_count,column+6, title_FailSN[-1].strip(),style_title)
		if ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_title_tall])
		row_count += 1

		for truefailsn in ProductDict['truefailDict'].keys():
			self.writeLineColumnRow(sheet,row_count,column,ProductDict['truefailDict'][truefailsn][0:3],text_style_list)
			sheet.write_merge(row_count,row_count,column+3,column+5,ProductDict['truefailDict'][truefailsn][3].strip(),style_text)
			sheet.write(row_count,column+6, '',style_text)
			row_count += 1

		for i in range(ProductDict[Project]['failsn']-len(ProductDict['truefailDict'].keys())):
			self.writeLineColumnRow(sheet,row_count,column,listRow[0:3],text_style_list)
			sheet.write_merge(row_count,row_count,column+3,column+5,listRow[3].strip(),style_text)
			sheet.write(row_count,column+6, '',style_text)
			row_count += 1

		title_retest = ['Product','Config','Retest SN','Radar']
		self.writeLineColumnRow(sheet,row_count,column,title_retest[0:3],title_style_list)
		sheet.write_merge(row_count,row_count,column+3,column+5,'Retest item',style_title)
		sheet.write(row_count,column+6, title_retest[-1].strip(),style_title)
		if ProductIndex == 0:
			self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_title_tall])
		row_count += 1
		
		for Retestsn in ProductDict['RetestDict'].keys():
			self.writeLineColumnRow(sheet,row_count,column,ProductDict['RetestDict'][Retestsn][0:3],text_style_list)
			sheet.write_merge(row_count,row_count,column+3,column+5,ProductDict['RetestDict'][Retestsn][3].strip(),style_text)
			sheet.write(row_count,column+6, '',style_text)
			if ProductIndex == 0:
				self.writeLineColumnRow(sheet,row_count,column+30,' ',[style_text_tall])
			row_count += 1

		for i in range(ProductDict[Project]['retestsn']-len(ProductDict['RetestDict'].keys())):
			self.writeLineColumnRow(sheet,row_count,column,listRow[0:3],text_style_list)
			sheet.write_merge(row_count,row_count,column+3,column+5,listRow[3].strip(),style_text)
			sheet.write(row_count,column+6, '',style_text)
			row_count += 1

class ReadCSV(object):
	def __init__(self):
		super(ReadCSV, self).__init__()

		self.Title_fixtureID = 'Fixture Id'
		self.Title_channelID = 'Head Id'
		self.Title_SN = 'SerialNumber'
		self.Title_Config = 'Special Build Description'
		
		self.Title_Ver = 'Version'
		self.Title_FailMsg = 'List of Failing Tests'
		self.Title_result = 'Test Pass/Fail Status'

		self.filePath = ''

	def GetReportDict(self,dictAll):
		dictHeader = dictAll['dictHeader']

		mixDict = {}
		index = 0

		for Product in dictHeader['Product']:
			if Product not in mixDict.keys():
				mixDict[Product] = {}
				mixDict[Product]['overlayVersion'] = dictAll['overlayVersion']
				mixDict[Product]['stationName'] = dictAll['StationName']

			if 'SerialNumber' not in mixDict[Product]:
				mixDict[Product]['SerialNumber'] = []
			if 'stationID' not in mixDict[Product]:
				mixDict[Product]['stationID'] = []
			if 'fixtureid' not in mixDict[Product]:
				mixDict[Product]['fixtureid'] = []
			if 'cfg' not in mixDict[Product]:
				mixDict[Product]['cfg'] = []
			if 'Channel ID' not in mixDict[Product]:
				mixDict[Product]['Channel ID'] = []
			if 'StartTime' not in mixDict[Product]:
				mixDict[Product]['StartTime'] = []
			if 'testTime' not in mixDict[Product]:
				mixDict[Product]['testTime'] = []
			if 'EndTime' not in mixDict[Product]:
				mixDict[Product]['EndTime'] = []
			if 'Pass/Fail' not in mixDict[Product]:
				mixDict[Product]['Pass/Fail'] = []
			if 'FailMsg' not in mixDict[Product]:
				mixDict[Product]['FailMsg'] = []

			mixDict[Product]['SerialNumber'].append(dictHeader['SerialNumber'][index])
			mixDict[Product]['stationID'].append(dictHeader['Station ID'][index])
			mixDict[Product]['fixtureid'].append(dictHeader['Fixture Id'][index])
			mixDict[Product]['cfg'].append(dictHeader['Special Build Description'][index])
			mixDict[Product]['Channel ID'].append(dictHeader['Fixture Id'][index] + '-' + dictHeader['Head Id'][index])
			mixDict[Product]['Pass/Fail'].append(dictHeader['Test Pass/Fail Status'][index])
			mixDict[Product]['FailMsg'].append(dictHeader['List of Failing Tests'][index])
			mixDict[Product]['StartTime'].append(dictHeader['StartTime'][index])
			mixDict[Product]['EndTime'].append(dictHeader['EndTime'][index])

			t1 = datetime.datetime.strptime(dictHeader['StartTime'][index], "%Y/%m/%d %H:%M")
			t2 = datetime.datetime.strptime(dictHeader['EndTime'][index], "%Y/%m/%d %H:%M")
			interval_time = (t2 - t1).seconds  
			mixDict[Product]['testTime'].append(interval_time)  

			index = index + 1

		return mixDict


	def ReadCSVData(self,filePath):
		Title_SN = 'SerialNumber'
		Title_Config = 'Special Build Description'

		Title_FailMsg = 'List of Failing Tests'
		dictAll = {}
		dictHeader = {}
		with open(filePath,encoding="utf-8") as f:
			reader = csv.reader(f)
			line = 0
			for row in reader:
				if(row[0] == 'Site'):
					titleRow = row[0:14]
					dictAll['TitleRow'] = row[0:14]
					for item in titleRow:
						dictHeader[item] = []
				elif(line == LINE_Project or line == LINE_Title):
					dictAll['StationName'] = row[0]
					dictAll['overlayVersion'] = row[1]
					

				elif(line > LINE_Unit):
					for item in titleRow:
						dictHeader[item].append(row[titleRow.index(item)])
				line = line + 1
		dictAll['dictHeader'] = dictHeader
		print('read CSVdata successful...............\n')
		return dictAll

class GetInfo(object):
	"""docstring for GetInfo"""
	def __init__(self):
		super(GetInfo, self).__init__()

	def read_txt_high(self,filename):
	    print('reading log file, please wait a moment.........')
	    with open(filename,'r',errors='ignore') as file_to_read:
	        list1 = [] 
	        while True:
	            lines = file_to_read.readline()  # 整行读取数据
	            if not lines:
	                break
	            item = lines
	            if len(item)>1:
	                item = item.split('\n')
	                itemStr = ''
	                for Istr in item:
	                    itemStr = itemStr+str(Istr).strip()
	                itemStr.strip()
	                list1.append(itemStr)
	    print("read log Info successful............\n")
	    return list1

	def Get_summuryItem(self,fileItem):
		titlelist = ['Dry run date','Bundle','Station','Purpose','Overlay Version','Path','Diags','BBFW','BBlib','NFC','PertOS','Phleet','Rose FW','BT LPEM FW','WIFI FW','BT FW','BT PHY FW','WIFI PHY FW','Project','Config','Vendor','Result','Fail Symptom','Test time(s)','Radar','DRI']
		summuryItem = {}
		for title in titlelist:
			summuryItem[title] = ''

		for item in fileItem:
			if summuryItem["Dry run date"] == "" and '-' in item and '202' == item[0:3]:
				summuryItem["Dry run date"] = item.split("-")[1]+"/"+item.split("-")[2][0:2]+"D"
			if summuryItem["Dry run date"] == "" and '/' in item and '202' == item[0:3]:
				summuryItem["Dry run date"] = item.split("/")[1]+"/"+item.split("/")[2][0:2]+"D"
			if "softwareversion =" in item and summuryItem['Bundle'] == "":
				summuryItem['Bundle'] = item.split("=")[1].split("_")[0]
			elif "0x00 Passed 0 0 2 310323220" in item and summuryItem['Bundle'] == "":
				summuryItem['Bundle'] = item.split("220 ")[1][0:len(item.split("220 ")[1])-4]
			elif 'Bundle name from gh_station_info is ' in item and summuryItem['Bundle'] == "":
				summuryItem['Bundle'] = item.split("is")[1]
			elif "Version  -" in item and summuryItem['Diags'] == "":
				summuryItem['Diags'] = item.split("- ")[1]
			elif '"diags_version" = "' in item and summuryItem['Diags'] == "":
				summuryItem['Diags'] = item.split('= "')[1][0:len(item.split('= "')[1])-3]
			elif '"diags_version":' in item and summuryItem['Diags'] == "":
				summuryItem['Diags'] = item.split('"diags_version":')[1].split(',')[0]
			elif "[REPORT]: BB_FIRMWARE_VERSION = " in item and summuryItem['BBFW'] == "":
				summuryItem['BBFW'] = item.split("=")[1]
			elif 'firmware-version: "' in item and summuryItem['BBFW'] == "":
				summuryItem['BBFW'] = item.split(': "')[1][0:len(item.split(': "')[1])-1]
			elif "[REPORT]: BBLIB_VER = " in item and summuryItem['BBlib'] == "":
				summuryItem['BBlib'] = item.split("=")[1]
			elif "<Info> firmware-version:" in item or "<Info> firmware-revision:" in item:
				if summuryItem['NFC'] == '':
					summuryItem['NFC'] = item.split("<Info>")[1]
					nfc1 = item.split("<Info>")[1]
				if len(summuryItem['NFC'])<30 and nfc1 not in item:
					summuryItem['NFC'] = summuryItem['NFC'] + "   " + item.split("<Info>")[1]
			elif '> firmware-version:' in item and len(summuryItem['NFC']) < 50:
				if summuryItem['NFC'] == "" or summuryItem['NFC'] not in item:
					summuryItem['NFC'] = summuryItem['NFC'] + '  f' + item.split('> f')[1]
			elif '> firmware-revision:' in item and len(summuryItem['NFC']) < 50:
				if summuryItem['NFC'] == "" or summuryItem['NFC'] not in item:
					summuryItem['NFC'] = summuryItem['NFC'] + '  f' + item.split('> f')[1]

			elif " 		firmware-version: 0x" in item or " 		firmware-revision: 0x" in item:
				if summuryItem['NFC'] == '':
					summuryItem['NFC'] = item.split(" : \t\t")[1][0:len(item.split(" : \t\t")[1])-1]
					nfc1 = item.split(" : \t\t")[1][0:len(item.split(" : \t\t")[1])-1]
				if len(summuryItem['NFC'])<30 and nfc1 not in item:
					summuryItem['NFC'] = summuryItem['NFC'] + "   " + item.split(" : \t\t")[1][0:len(item.split(" : \t\t")[1])-1]
			elif '"softwarename" = "ATLAS' in item and summuryItem['Station'] == "":
				summuryItem['Station'] = item.split('ATLAS-')[1][0:len(item.split('ATLAS-')[1])-3]
			elif '"softwarename" = "' in item and summuryItem['Station'] == "":
				summuryItem['Station'] = item.split('"softwarename" = "')[1][0:len(item.split('"softwarename" = "')[1])-3]
			elif " OFW revision " in item and summuryItem['PertOS'] == "":
				summuryItem['PertOS'] = item.split("revision ")[1]
			elif "firmware [Rev " in item and summuryItem['WIFI FW'] == "":
				summuryItem['WIFI FW'] = item.split("firmware")[1]
			elif '"bt_mac_fw" = "' in item:
				summuryItem['BT FW'] = item.split('= "')[1][0:len(item.split('= "')[1])-2]
			elif "BT MAC FW" in item and summuryItem['BT FW'] == "":
				summuryItem['BT FW'] = item.split("FW")[1]
			elif "TAG:       " in item and summuryItem['Phleet'] == "":
				summuryItem['Phleet'] = item.split("TAG:       ")[1]
			elif '"vendor_id":' in item and summuryItem['Vendor'] == "":
				summuryItem['Vendor'] = item.split('"vendor_id":')[1].split(':')[1][0:2]
			elif '"VENDOR_ID" = "' in item:
				summuryItem['Vendor'] = item.split('= "')[1].split(':')[1][0:2]
			elif "Loaded FW Version " in item and summuryItem['Rose FW'] == "":
				summuryItem['Rose FW'] = item.split("Version ")[1]
			elif "BT PHY FW" in item and summuryItem['BT PHY FW'] == "":
				summuryItem['BT PHY FW'] = item.split("FW")[1]
			elif "phy [" in item and summuryItem['WIFI PHY FW'] == "":
				summuryItem['WIFI PHY FW'] = item.split("phy [")[1][0:len(item.split("phy [")[1])-2]
			elif 'cfg = "' in item:
				if summuryItem['Project'] == "":
					summuryItem['Project'] = item.split('= "')[1].split("/")[0]
				if summuryItem['Config'] == "":
					summuryItem['Config'] = item.split('= "')[1].split("/")[2][0:len(item.split('= "')[1].split("/")[2])-2]
			elif 'CFG#: ' in item:
				if summuryItem['Project'] == "":
					summuryItem['Project'] = item.split('CFG#:')[1].split("/")[0]
				if summuryItem['Config'] == "":
					summuryItem['Config'] = item.split('CFG#:')[1].split("/")[2]
			elif '"STATION_OVERLAY" : "' in item:
				summuryItem['Overlay Version'] = item.split(': "')[1][0:len(item.split(': "')[1])-2]
			elif 'STATION_OVERLAY_VERSION=' in item:
				summuryItem['Overlay Version'] = item.split('=')[1].split('[')[0]
			elif '"STATION_TYPE" :' in item:
				summuryItem['Station'] = item.split(': "')[1][0:len(item.split(': "')[1])-2]
			elif 'STATION=' in item:
				summuryItem['Station'] = item.split('=')[1]

		for item in summuryItem.keys():
			summuryItem[item] = summuryItem[item].strip()
		if len(summuryItem['Station'])>15:
			summuryItem['Station'] = ''

		return summuryItem

def executeAction(infoPathlist='',csvPathlist='',ReportPath=''):
	summuryDict = {}
	if len(infoPathlist)>0: 
		item = []
		for infoPath in infoPathlist:
			getInfo = GetInfo()
			item += getInfo.read_txt_high(infoPath)
		summuryDict = getInfo.Get_summuryItem(item)

	csvDict = {}
	readcsv = ReadCSV()
	if len(csvPathlist)>0:
		csvdata = {}
		for csvPath in csvPathlist:
			if len(csvdata.keys()) == 0:
				csvdata = readcsv.ReadCSVData(csvPath)
			else:
				csv = readcsv.ReadCSVData(csvPath)
				for item in csvdata['dictHeader'].keys():
					csvdata['dictHeader'][item] += csv['dictHeader'][item]
		csvDict = readcsv.GetReportDict(csvdata)

	se = SaveExcel()
	if len(summuryDict.keys())>0:
		se.Outputsummary(summuryDict)
	if len(csvDict.keys())>0:
		se.OutputDryrunDatail(csvDict)
	se.saveAction(ReportPath)

def judgeFile(filePath):
	if os.path.exists(filepath):
		pass
	else:
		print('Wrong path:',filepath)
		filePath = input('Please input the correct path:')
		filePath = filePath.strip()
		return filePath

def getFilePath(filepath):
	pathDict = {}
	pathDict['csv'] = []
	pathDict['log'] = []
	for item in os.listdir(filePath):
		if os.path.isdir(filePath + '/' + item):
			logpath = filePath + '/' + item
			for item1 in os.listdir(logpath):
				logpath1 = logpath + '/' + item1
				if 'efi0-uart.log' in item1 or 'unit.log' in item1 or 'gh_station_info.json' in item1 or 'Restore Info.txt' in item1:
					pathDict['log'].append(logpath1)
					# break
				elif os.path.isdir(logpath1):
					for item2 in os.listdir(logpath1):
						# print(item2)
						logpath2 = logpath1 + '/' + item2
						if 'efi0-uart.log' in item2 or 'unit.log' in item2 or 'gh_station_info.json' in item2 or 'Restore Info.txt' in item2:
							pathDict['log'].append(logpath2)
							# break
						elif os.path.isdir(logpath2):
							for item3 in os.listdir(logpath2):
								logpath3 = logpath2 + '/' + item3
								# print(item3)
								if 'efi0-uart.log' in item3 or 'unit.log' in item3 or 'station_info.json' in item3 or 'Restore Info.txt' in item3:
									pathDict['log'].append(logpath3)
									# break
								elif os.path.isdir(logpath3):
									for item4 in os.listdir(logpath3):
										logpath4 = logpath3 + '/' + item4
										# print(item4)
										if 'efi0-uart.log' in item4 or 'unit.log' in item4 or 'station_info.json' in item4 or 'Restore Info.txt' in item4:
											# print('*********',item4)
											pathDict['log'].append(logpath4)
											# break
		elif '.csv' in item:
			pathDict['csv'].append(filePath+'/'+item)
		elif 'efi0-uart.log' in item or 'unit.log' in item or 'Restore Info.txt' in item:
			pathDict['log'] = filePath+'/'+item
	# print(pathDict)
	return pathDict


if __name__ == '__main__':

	print('           *******************************************************')
	print('           |                                                     |')
	print('           |                                                     |')
	print('           |                    SIP DryRunReport                 |')
	print('           |                                                     |')
	print('           |                文件绝对路径中不允许有空格           |')
	print('           |                                                     |')
	print('           |                                                     |')
	print('           |                                         Ver:1.0.3   |')
	print('           |                                          Andy Li    |')
	print('           |                                         2021.06.23  |')
	print('           |                                                     |')
	print('           *******************************************************')
	
	filePath = input('please drop the file:')
	filePath = filePath.strip()

	# filePath = '/Users/andy/Documents/python/SIP-Tool/DryRunReport/Dry_run_report/PanelFCT'

	while not os.path.exists(filePath):
		filePath = judgeFile(filePath)
		if judgeFile(filePath) == True:
			break
	ReportPath = filePath
	pathDict = getFilePath(filePath)

	if len(pathDict['log'])>0 and len(pathDict['csv'])>0:
		executeAction(infoPathlist=pathDict['log'],csvPathlist=pathDict['csv'],ReportPath=ReportPath)
	elif len(pathDict['log'])>0:
		executeAction(infoPathlist=pathDict['log'],ReportPath=ReportPath)
	elif len(pathDict['csv'])>0:
		executeAction(csvPathlist=pathDict['csv'],ReportPath=ReportPath)

