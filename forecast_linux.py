import time
from datetime import datetime, timedelta
import xlrd
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import json
import codecs


work_book = xlrd.open_workbook('./points.xls')
sheet_2 = work_book.sheet_by_name('对照表')
code_info = sheet_2.col_values(0)[1:]
code_id = [str(int(i)) for i in sheet_2.col_values(1)[1:]]
code_dic = {}
for i in range(len(code_id)):
    code_dic[code_id[i]] = code_info[i]

work_book.release_resources()
del work_book


def forecast(name, position, driver):
        
    #name = "疏勒县"
    #position = name_dic['疏勒县']

    lat = position[0]
    lon = position[1]
    alt = position[2]
    time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = 'https://www.windy.com/' + str(lat) + '/' + str(lon)
    #https://www.windy.com/39.468/75.994
    print('发送请求...')
    driver.get(url) 
    time.sleep(2)
    print('响应已获取...')
    driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/div[2]').click()  # 未来十天
    time.sleep(2)
    table = driver.find_element_by_xpath('''/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table''')  # 数据表格

    #获得标签元素
    day = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[1]')
    hour = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[2]')
    weather = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[3]')
    tempature = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[4]')
    rain = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[5]')
    wind = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[6]')
    gust = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[7]')
    direction = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table/tbody/tr[8]')
    print('标签数据获取完毕...')
    # 处理日期,取得起始日期和小时，timedelta逐次累加3小时，得到48元素小时列表

    hours = hour.text.split(' ')[0:48]
    hours_int = [int(i[:-2]) for i in hours] #linux改动，获取小时后缀带AM或PM，去除

    # DAY_FIRST = int(day.text.split(' ')[1]) #linux改动，Wednesday 23\nThursday 24 Friday 25 Saturday 26 不便处理
    HOUR_FIRST = hours_int[0]

    NUMBER_FIRST_DAY = hours_int.index(max(hours_int)) + 1

    YEAR = int(time.strftime("%Y", time.localtime()))
    MONTH = int(time.strftime("%m", time.localtime()))
    DAY = int(time.strftime("%d", time.localtime()))

    start_date_hour = datetime(YEAR, MONTH, DAY, HOUR_FIRST)
    day_hour_list = []
    for i in range(48):
        day_hour_list.append(start_date_hour.strftime("%Y-%m-%d %H"))
        start_date_hour += timedelta(hours=3)

    print('时间标签处理完毕...')


    weathers = weather.find_elements_by_tag_name('img')
    weather_img_link = [i.get_attribute('src') for i in weathers]
    weather_list = []
    for i in range(48):
        code = weather_img_link[i][42:44]
        if '.' in code or '_' in code:
            code = code[0]
        weather_list.append(code_dic[code])

    print('天气标签处理完毕...')

    tempatures = tempature.text.split(' ')
    tempature_list_F = []
    for i in range(48):
        tempature_list_F.append(tempatures[i][:-1])

    tempature_list = []  #linux F转C
    for i in tempature_list_F:
        tempature_list.append(str(int(round((int(i)-32)*5/9))))


    print('气温标签处理完毕...')

    rains = [i.text for i in rain.find_elements_by_tag_name('td')][0:48]
    rain_list = []
    for i in rains:
        if 'in' in i:
            rain_list.append(str(round(float(i[:-2]) * 25.4)))  #Linux特别处理，inch转mm
        else:
            if i == '':
                rain_list.append('0')
            else:
                rain_list.append(str(round(float(i) * 25.4)))

    print('降水量标签处理完毕...')

    #1 kt = 0.514444 m/s
    winds = [str(round(int(i) * 0.514444)) for i in wind.text.split(' ')][0:48]
    wind_list = winds 

    print('平均风速标签处理完毕...')

    gusts = [str(round(int(i) * 0.514444)) for i in gust.text.split(' ')][0:48]
    gust_dic = gusts

    print('阵风风速标签处理完毕...')

    directions = [i.get_attribute('style') for i in direction.find_elements_by_tag_name('div')]
    direction_list_int = [int(i[i.find('(')+1:-5])-180 for i in directions][:48]
    for i in range(48):
        if direction_list_int[i] < 0:
            direction_list_int[i] += 360

    direction_list = [str(i) for i in direction_list_int]

    print('风向标签处理完毕...')

    weather_dic = {}
    weather_dic['name'] = name
    weather_dic['lat'] = lat
    weather_dic['lon'] = lon
    weather_dic['alt'] = alt
    weather_dic['generate_time'] = time_now
    weather_dic['hour_steps'] = day_hour_list
    weather_dic['weather'] = weather_list
    weather_dic['tempature'] = tempature_list
    weather_dic['rain'] = rain_list
    weather_dic['wind'] = wind_list
    weather_dic['gust'] = gust_dic
    weather_dic['direction'] = direction_list
    print('数据封装完成。')
    return weather_dic

