import xlrd
import xlsxwriter
from forecast_linux import forecast
from datetime import datetime
from selenium import webdriver
import json
import codecs


work_book = xlrd.open_workbook('C:/Users/zhiyuan/Code/windy/points.xls')
sheet = work_book.sheet_by_name('点位表')
names = sheet.col_values(0)[3:]
coordinates = sheet.col_values(1)[3:]
altitudes = sheet.col_values(2)[3:]
work_book.release_resources()
del work_book

lats = [round(int(i[1:3]) + int(i[4:6])/60 + int(i[7:9])/3600, 4) for i in coordinates]
lons = [round(int(i[12:14]) + int(i[15:17])/60 + int(i[18:20])/3600, 4) for i in coordinates]
name_dic = {}
for i in range(len(names)):
    name_dic[names[i]] = [round(lats[i], 3), round(lons[i], 3), int(altitudes[i])]


driver = webdriver.Chrome("C:/Program Files/Google/Chrome/Application/chromedriver.exe")
driver.implicitly_wait(20)


resault_dic = {}
for i in list(name_dic.keys()):
    while True:
        if i in list(resault_dic.keys()):
            break
        else:
            try:
                resault = forecast(i, name_dic[i], driver)
                resault_dic[i] = resault
                print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                print(i + '气象预报已生成完毕。')
            except:
                continue


driver.close()
driver.quit()


time_now = datetime.now().strftime("%Y年%m月%d日 %H时%M分%S秒")
file_name = './temp/json/' + time_now + '.json'
json_string = json.dumps(resault_dic, ensure_ascii=False, sort_keys=False, indent=4, separators=(',', ':'))
f = codecs.open(file_name, 'w', encoding='utf-8')

f.write(json_string)
f.close()




#a = forecast('喀什市', name_dic['喀什市'], driver)
