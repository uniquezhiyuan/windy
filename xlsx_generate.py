import json
import codecs
import xlrd
import xlsxwriter
import os
from openpyxl import load_workbook
from datetime import datetime, timedelta


#获取天气集以备后用
work_book = xlrd.open_workbook('./points.xls')
sheet_2 = work_book.sheet_by_name('对照表')
code_info = sheet_2.col_values(0)[1:]
code_id = [str(int(i)) for i in sheet_2.col_values(1)[1:]]
code_dic = {}
for i in range(len(code_id)):
    code_dic[code_id[i]] = code_info[i]

work_book.release_resources()
del work_book

weather_set = [code_dic[i] for i in reversed(code_dic)]  #天气集

#name = "喀什地区"
# 读取最新json
dir_name = './temp/json'
file_list = os.listdir(dir_name)
file_list.sort(key=lambda fn:os.path.getmtime(dir_name+'/'+fn))

f = codecs.open(dir_name + '/' + file_list[-1], 'r', encoding='utf-8')
resault_dic = json.load(f)
f.close()

def transform_json(name, future):
	#开始处理数据，字典转列表，以备写入xlsx
	time_now = datetime.now().strftime("%Y年%m月%d日 %H时%M分")
	target_date = (datetime.now() + timedelta(days=future)).strftime("%Y-%m-%d")


	#获取目标预报时间索引
	info_index = [resault_dic[name]['hour_steps'].index(i) for i in resault_dic[name]['hour_steps'] if target_date in i]


	transform_list = []
	#1.地名
	transform_list.append(name)

	#2.纬度
	transform_list.append(str(resault_dic[name]['lat']))

	#3.经度
	transform_list.append(str(resault_dic[name]['lon']))

	#4.海拔
	transform_list.append(str(resault_dic[name]['alt']))

	#5.日期
	transform_list.append(target_date)

	#6.主要天气,算法：遍历天气集列表，找到优先级最高的为当日天气
	weather_list = [resault_dic[name]['weather'][i] for i in info_index]
	for i in weather_set:
	    if i in weather_list:
	        transform_list.append(i)
	        break

	#7.气温
	tempature_list = [int(resault_dic[name]['tempature'][i]) for i in info_index]
	tempature_str = str(min(tempature_list)) + '～' + str(max(tempature_list))
	transform_list.append(tempature_str)

	#8.雨量
	rain_list = [float(resault_dic[name]['rain'][i]) for i in info_index]
	transform_list.append(str(round(sum(rain_list))))

	#9.平均风速
	wind_list = [int(resault_dic[name]['wind'][i]) for i in info_index]
	transform_list.append(str(round(sum(wind_list)/len(wind_list))))

	#10.最大风速
	gust_list = [int(resault_dic[name]['gust'][i]) for i in info_index]
	transform_list.append(str(round(sum(gust_list)/len(gust_list))))

	#11.风向,算法：正午风向
	direction_list = [resault_dic[name]['direction'][i] for i in info_index]
	transform_list.append(direction_list[4])

	#12.预警
	transform_list.append('无')
	return transform_list


def generate_xlsx(future):
	data_set = []
	for i in list(resault_dic.keys()):
	    data_set.append(transform_json(i, future))

	# 开始生成报表
	options = {
	    'default_format_properties':{
	    'font_name':'仿宋_GB2312',
	    'font_size':10,
	    'align':'center',
	    'valign':'center',
	    'text_wrap':1,
	    }
	}

	title_date = (datetime.now() + timedelta(days=future)).strftime("%Y年%m月%d日")
	work_book = xlsxwriter.Workbook('./temp/xlsx/'+ title_date +'.xlsx', options)
	work_sheet = work_book.add_worksheet('预报结果')

	work_sheet.set_column('A:A', 22) #列宽
	work_sheet.set_column('B:L', 10)
	work_sheet.set_row(0, 15)
	work_sheet.set_row(1, 13) #行高默认13
	work_sheet.set_row(2, 23)
	for i in range(len(data_set)):
	    work_sheet.set_row(i+3, 9)

	work_sheet.set_paper(8) #A3,A4为9
	work_sheet.set_portrait()  #纸张竖向，横向为set_landscape()
	work_sheet.set_margins(left=0.6, right=0.6, top=0.2, bottom=0) #边距
	work_sheet.set_footer(footer=' ')
	work_sheet.set_header(header=' ')
	work_sheet.center_horizontally()  #页面居中
	#work_sheet.protect('123')  #设置保护视图防止更改

	#首行
	fmt = work_book.add_format({'align': 'center', 'valign': 'vcenter', 'font_size':13, 'bold':True,})
	work_sheet.merge_range(0,0,0,11, '南疆地区'+ title_date +'气象预报', cell_format=fmt)

	time_generate = resault_dic['喀什地区']['generate_time']
	#注记
	fmt = work_book.add_format({'align': 'left', 'valign': 'vcenter', 'font_size':10, 'bold':False,})
	work_sheet.merge_range(1,0,1,3, '制表：气象室', cell_format=fmt)
	fmt = work_book.add_format({'align': 'center', 'valign': 'vcenter', 'font_size':10, 'bold':False,})
	work_sheet.merge_range(1,4,1,7, '数据源：ECMWF', cell_format=fmt)
	fmt = work_book.add_format({'align': 'right', 'valign': 'vcenter', 'font_size':10, 'bold':False,})
	work_sheet.merge_range(1,8,1,11, '生成时间：'+time_generate, cell_format=fmt)

	#表头
	title_list = ['地名','纬度\n（°）','经度\n（°）','海拔\n（m）','日期','主要天气','气温\n（℃）','雨量/雪深\n（mm）','平均风速\n（m/s）','最大风速\n（m/s）','风向\n（°）','灾害预警',]
	fmt = work_book.add_format({'align': 'center', 'valign': 'vcenter', 'font_size':10, 'bold':True, 'top':1, 'bottom':1, 'left':1, 'right':1,})
	work_sheet.write_row('A3', title_list, cell_format=fmt)

	fmt = work_book.add_format({'align': 'center', 'valign': 'vcenter', 'font_size':8, 'bold':False, 'top':1, 'bottom':1, 'left':1, 'right':1,})
	for i in range(len(data_set)):
	    start_cell = 'A' + str(4 + i)
	    work_sheet.write_row(start_cell, data_set[i], cell_format=fmt)

	work_book.close()
	print(title_date + "生成成功。")


if __name__ == "__main__":
	generate_xlsx(1)
	generate_xlsx(2)
	generate_xlsx(3)
	generate_xlsx(4)
	generate_xlsx(5)

