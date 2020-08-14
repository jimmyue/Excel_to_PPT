import win32com
import xlwings as xw
from win32com.client import Dispatch
import os
import re
#https://docs.microsoft.com/zh-cn/office/vba/api/powerpoint.chart 查找相关ppt方法
#win32com.gen_py.91493440-5A91-11CF-8700-00AA0060263Bx0x2x11' has no attribute 'CLSIDToClassMap' 解决办法，删除C:\Users\yuejing\AppData\Local\Temp\gen_py\3.7下对应文件夹

def findstr(text):
	strall=[]
	#获取字符串出现的次数
	n=text.count('-')
	#获取字符串text中subStr第n次出现的位置
	def findn(text,subStr,n):
		listStr = text.split(subStr,n)
		if len(listStr) <= n:
			return -1
		return len(text)-len(listStr[-1])-len(subStr)
	#循环获取字符开始结束位置
	for i in range(1,n+1):
		a=findn(text,'-',i)
		b=text.find('%',a)+1
		strsub=[]
		strsub.append(a)
		strsub.append(b)
		strall.append(strsub)
	return strall

#读取PPT
ppt = Dispatch('PowerPoint.Application')
ppt.Visible = 1
ppt.DisplayAlerts = 0 
filepath=os.path.join(os.getcwd(), 'template/市场数据汇报.pptx')
pptSel = ppt.Presentations.Open(filepath)
slide_count=pptSel.Slides.Count

#读取EXCEL,默认设置：程序可见，只打开不新建工作薄，禁用警告框,屏幕更新关闭。
app = xw.App(visible=True,add_book=False)
app.display_alerts = False
app.screen_updating = False
wb = app.books.open(os.path.join(os.getcwd(),'data/市场数据.xlsx'))
ws = wb.sheets[0]

#替换PPT对应表格数据，查看PPT的shape（开始->选择->选择窗格）
for i in range(1,slide_count+1):
	gen = (shape for shape in pptSel.Slides(i).Shapes)
	for item in gen:
		if '_图表_' in item.Name:
			#PPT图表类型数据替换，PPT窗格名称与excel标签名一致且为图标
			row = ws.cells.columns(1).api.Find(item.Name)
			if row is not None:
				#获取excel对应表格值
				data = ws.range(row.Row,3).options(index=False, expand='table').value
				#替换PPT对应表格值
				item.Chart.ChartData.Activate()
				pwb = item.Chart.ChartData.Workbook
				pws = pwb.Sheets(1)
				pws.Cells.Rows(2).ClearContents() #清除第二行数据
				pws.Cells.Rows(3).ClearContents() #清除第三行数据
				pws.Range(pws.Cells(2,2),pws.Cells(2,2).GetOffset(len(data)-2,len(data[0])-1)).Value=[data[1],data[2]]
				pws.Range(pws.Cells(1,2),pws.Cells(1,2).GetOffset(0,len(data[0])-1)).Value=[data[0]]
				item.Chart.Refresh()
				pwb.Close()
			
		elif '_文本框_' in item.Name:
			#PPT图表类型数据替换，PPT窗格名称与excel标签名一致且为文本框
			row = ws.cells.columns(1).api.Find(item.Name)
			if row is not None:
				#获取excel对应表格值
				data = ws.range(row.Row,3).options(index=False, expand='table').value
				#获取PPT文本框
				txt = item.TextFrame.TextRange.Text
				pattern = re.compile(r'([-+]\d*\.\d\%)')
				result=pattern.split(txt)
				#替换数据
				txt = re.sub(re.escape(result[1]), data[1], txt, count=1)
				txt = re.sub(re.escape(result[3]), data[3], txt, count=1)
				txt = re.sub(re.escape(result[5]), data[5], txt, count=1)
				txt = re.sub(re.escape(result[7]), data[7], txt, count=1)
				#获取负的百分比并标红
				per=findstr(txt)
				item.TextFrame.TextRange.Text = txt
				for i in range(len(per)):
					item.TextFrame.TextRange.Characters(per[i][0],per[i][1]-per[i][0]+1).Font.Color.RGB=0xFF0000FF #ARPG

wb.close()         #关闭打开的Excel文档
app.quit()         #关闭office
app.kill()
pptSel.SaveAs(os.path.join(os.getcwd(),'result/市场数据汇报.pptx'))
pptSel.Close()     #关闭打开的PowerPoint文档
ppt.Quit()         #关闭office

