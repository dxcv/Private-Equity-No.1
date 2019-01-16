from WindPy import w
import xlrd
import xlwt
import os
import pandas as pd
import getData
import convertDate
import matplotlib.pyplot as plt
#import Plotting
#from datetime import *
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False

bondCode = '1103'
fundCode = '1105'
stockCode = '1102'
SH = '.SH'
SZ = '.SZ'
IB = '.IB'
OF = '.OF'

w.start()

d = {}
d[bondCode] = {}
d[fundCode] = {}
d[stockCode] = {}
dateIndex = []
arr = []

book2 = xlwt.Workbook(encoding = 'utf=8', style_compression = 0)
sheet2 = book2.add_sheet('Sheet1', cell_overwrite_ok = True)
sheet2.write(0, 0, '日期')
sheet2.write(0, 1, '久期')
sheet2.write(0, 2, '十年国开债收益率(%)')
sheet2.write(0, 3, '累计净值')
sheet2.write(0, 4, '可转债占比')
sheet2.write(0, 5, '分级A占比')
sheet2.write(0, 6, 'AA-级债券')
sheet2.write(0, 7, 'AA级债券')
sheet2.write(0, 8, 'AA+级债券')
sheet2.write(0, 9, 'AAA级债券')

# create a database for future reference
dictionary = book2.add_sheet('数据库', cell_overwrite_ok = True)
dictionary.write(0, 0, 'Wind代码')
dictionary.write(0, 1, '证券名称')
dictionary.write(0, 2, '证券类型')
dictionary.write(0, 3, '债券发行日期债项评级')
dictionary.write(0, 4, '债券发行日期主体评级')
dictionary.write(0, 5, '债券最新债项评级')
dictionary.write(0, 6, '债券最新主体评级')
dictionary.write(0, 7, '评级机构')

no_dur = book2.add_sheet("久期缺失债券", cell_overwrite_ok = True)
r3 = 0

files = os.listdir('瑞泽平衡')
sorted_files = convertDate.sort_files(files)
total_files = len(sorted_files[0])
beg_date = sorted_files[1][0]
end_date = sorted_files[1][-1]
#print(sorted_files)
# 获取十年国债收益率
cdb10 = w.edb("M1004271", beg_date, end_date, "Fill=Previous")

for i in range(0, total_files):
    # get date
    dt = sorted_files[1][i]
    dateIndex.append(dt)
    print("Current date is: " + dt)

    # open spreadsheet to read
    file= '瑞泽平衡\\' + files[sorted_files[0][i]]
    book1 = xlrd.open_workbook(file)
    sheet1 = book1.sheets()[0]
    row = sheet1.nrows

    # initialization
    total = 0
    dur = 0
    net = 0
    convBond = 0
    rateA = 0
    rateB = 0
    etf = 0
    stock = 0
    creditRate1 =0
    issuerRate1 = 0
    creditRate2 = 0
    issuerRate2 = 0
    creditRate3 = 0
    issuerRate3 = 0
    creditRate4 = 0
    issuerRate4 = 0
    exMarket = ''
    isETF = 0

    # iterating through rows
    for r in range(row - 1, 0, -1):
        index = str(sheet1.cell(r, 0).value)
        if index == "资产类合计：" or index == "资产合计":
            total = float(sheet1.cell(r, 11).value)
            break
        elif index == "累计单位净值：" or index == "累计单位净值":
            net = float(sheet1.cell(r, 1).value)

    #获取今日的十年国债收益率
    currDate = 0
    while (1):
        if (str(cdb10.Times[currDate]) == dt):
            break
        currDate += 1
    ytm = cdb10.Data[0][currDate]

    for r in range(10, row):
        index = str(sheet1.cell(r, 0).value)
        name = str(sheet1.cell(r, 1).value)
        index_len = len(index)
        if index_len == 0 or (index[0] != '1' and index[0] != '2'):
            break
        elif index_len < 8 and index_len > 4 and \
                (index[0:4] == bondCode or index[0:4] == fundCode or index[0:4] == stockCode):
            if name.find('深交所') != -1:
                exMarket = SZ
            elif name.find('上交所') != -1:
                exMarket = SH
            elif name.find('银行间') != -1:
                exMarket = IB
            elif name.find('场外') != -1:
                exMarket = OF
            if name.find('ETF'):
                isETF = 1
            else:
                isETF = 0
            #print("testing exMarket: " + exMarket + " and etf now is: " + str(etf) + ' ' + name)
        # position
        if len(index) >= 14 and \
                (index[0:4] == bondCode or index[0:4] == fundCode or index[0:4] == stockCode):
            windCode = getData.convert(index, name, d, beg_date, end_date, exMarket, etf)
            if index[0:4] == stockCode:
                stock += float(sheet1.cell(r, 11).value)/total
            elif index[0:4] == fundCode:
                type = d[fundCode][windCode][1]
                if type == '分级A':
                    rateA += float(sheet1.cell(r, 11).value)/total
                elif type == '分级B':
                    rateB += float(sheet1.cell(r, 11).value)/total
                elif type == 'ETF':
                    etf += float(sheet1.cell(r, 11).value)/total
            elif index[0:4] == bondCode:
                tempcr = d[bondCode][windCode][2]
                tempir = d[bondCode][windCode][3]
                if d[bondCode][windCode][1] == '可转债':
                    convBond += float(sheet1.cell(r, 11).value)/total
                if  tempcr == 'AA-':
                    creditRate1 += float(sheet1.cell(r, 11).value)/total
                elif tempcr == 'AA':
                    creditRate2 += float(sheet1.cell(r, 11).value)/total
                elif tempcr == 'AA+':
                    creditRate3 += float(sheet1.cell(r, 11).value)/total
                elif tempcr == 'AAA':
                    creditRate4 += float(sheet1.cell(r, 11).value)/total
                elif tempir == 'AA-':
                    creditRate1 += float(sheet1.cell(r, 11).value)/total
                elif tempir == 'AA':
                    creditRate2 += float(sheet1.cell(r, 11).value)/total
                elif tempir == 'AA+':
                    creditRate3 += float(sheet1.cell(r, 11).value)/total
                elif tempir == 'AAA':
                    creditRate4 += float(sheet1.cell(r, 11).value)/total

                try:
                    currDate = 0
                    while str(d[bondCode][windCode][7][currDate]) != dt:
                        #print("date: " + str(d[bondCode][windCode][7][currDate]) + ' ' + dt)#test
                        currDate += 1
                except:
                    print("Duration is not provided; date out of range")
                    date_format = xlwt.XFStyle()
                    date_format.num_format_str = 'YYYY-MM-DD'
                    no_dur.write(r3, 0, dt, date_format)
                    no_dur.write(r3, 1, windCode)
                    no_dur.write(r3, 2, name)
                    r3 += 1
                    tempDur = 0
                tempDur = float(d[bondCode][windCode][8][currDate])

                if str(tempDur) == 'nan':
                    tempDur = 0
                    date_format = xlwt.XFStyle()
                    date_format.num_format_str = 'YYYY-MM-DD'
                    no_dur.write(r3, 0, dt, date_format)
                    no_dur.write(r3, 1, windCode)
                    no_dur.write(r3, 2, name)
                    r3 += 1
                    print ("Warning: Duration is not provided")
                    tempDur = 0
                #print(windCode + " tempDur is: " + str(tempDur))
                dur = dur + float(sheet1.cell(r, 12).value) * float(tempDur)
        elif len(index) >= 14 and index[0] == '2' and index[3] == '2':
            if index[13] == '7':
                dur = dur + 7 / 365 * sheet1.cell(r, 12).value
            elif index[13] == '1':
                dur = dur + 1 / 365 * sheet1.cell(r, 12).value
            elif index[13] == '8':
                dur = dur + 28 / 365 * sheet1.cell(r, 12).value

    arr.append([dur, ytm, net, convBond * 100, rateA * 100,\
              creditRate1 * 100, creditRate2 * 100, creditRate3 * 100, creditRate4 * 100])
    i = i + 1


    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'YYYY-MM-DD'
    sheet2.write(i, 0, dt, date_format)
    sheet2.write(i, 1, dur)
    sheet2.write(i, 2, ytm)
    sheet2.write(i, 3, net)
    sheet2.write(i, 4, convBond * 100)
    sheet2.write(i, 5, rateA * 100)
    sheet2.write(i, 6, creditRate1 * 100)
    sheet2.write(i, 7, creditRate2 * 100)
    sheet2.write(i, 8, creditRate3 * 100)
    sheet2.write(i, 9, creditRate4 * 100)


#construct dataframe
df = pd.DataFrame(arr)
df.columns = [['久期','十年国开债收益率(%)','累计净值', '可转债占比','分级A占比', \
                 'AA-级债券','AA级债券','AA+级债券', 'AAA级债券']]
df.index = pd.to_datetime(dateIndex)

# plotting
fig, ax1 = plt.subplots()
ax2 = ax1.twinx()
plt.xlabel(u'日期')
plt.title(u'瑞泽平衡久期和十年国开债收益率对比图')
df[u'久期'].plot(ax = ax1, color='b', lw=1.5, legend=True)
plt.ylabel(u'久期')
plt.legend(loc=2)
df[u'十年国开债收益率(%)'].plot(ax = ax2, color='g', lw=1.5, legend=True)
plt.legend(loc=1)
plt.ylabel(u'收益率(%)')
plt.grid(True)
plt.show() #test

# 将dictionary里的数据写入数据库
r = 1
for k in d:
    if k == bondCode:
        for c in d[k]:
            dictionary.write(r, 0, c)
            for item in range(0 , 7):
                dictionary.write(r, item + 1, d[k][c][item])
            r += 1
    if k == fundCode:
        for c in d[k]:
            dictionary.write(r, 0, c)
            for item in range(0 , 2):
                dictionary.write(r, item + 1, d[k][c][item])
            r += 1

try:
    book2.save('瑞泽平衡归因.xls')
except:
    print("Please Close the Book before Running")
    exit(1)

