# -*- coding:utf-8 -*-
from WindPy import w

fund = '1105'
bond = '1103'
stock = '1102'

def convert(index, name, d, beg_date, end_date, exMarket, etf):
    # Get Rid of Unwanted "."
    if index[4] == '.' and index[7] == '.' and index[10] == '.':
        index = index[0:4] + index[5:7] + index[8: 10] + index[11: -1]
    elif index[4] == '.' and index[7] == '.':
        index = index[0:4] + index[5: 7] + index[8: -1]
    elif index[4] == '.':
        index = index[0:4] + index[5:-1]
    windCode = index[8:14] + exMarket

    # 基金:
    if index[0:4] == fund:
        type = ''
        if etf:
            type = 'ETF'
        elif name.find('A'):
            type = '分级A'
        elif name.find('B'):
            type = '分级B'
        elif index[4:6] == '38': #to be fixed
            type = '货币基金'
        elif index[4:6] == '24':
            type = '场外货币'
        else:
            print('Warning-unidentified fund: ' + windCode + ' ' + name)
            exit(401)
        # putting the fund information into the dictionary
        if windCode not in d[fund]:
            d[fund][windCode] = [name, type]
    # 股票
    elif index[0:4] == stock:
        if windCode not in d[stock]:
            d[stock][windCode] = [name]
    # 债券
    elif index[0:4] == bond:
        type = ''
        if index[4: 6] == '04' or index[4: 6] == '12' or index[4:6]== '23':
            type = '可转债'
        elif index[4: 6] == '01' or index[4:6] == '11':
            type = '国债'
        elif index[4: 6] == '10':
            type =  '政策金融债'
        elif index[4: 6] == '13' or index[4: 6] == '02':
            type = '企业债'
        elif index[4: 6] == '03':
            type = '公司债'
        elif index[4: 6] == 'I1' or index[4: 6] == '15':
            type = '可交换债'
        elif index[4:6] == '05':
            type = '金融债'
        elif index[4: 6] == "15":
            type = '可交换债'

        # 深交所
        elif index[4: 6] == '61'or index[4: 6] == '99' or index[4: 6] == '80' or index[4: 6] == '32':
            type = '可转债'
        elif index[4: 6] == '33':
            type = '企业债'
        elif index[4: 6] == '35':
            type = '可交换'
        #银行间
        elif index[4: 6] == '69' or index[4: 6] == 'C4':
            type = '政策金融债'
        elif index[4: 6] == '51' or index[4: 6] == 'B5':
            type = '国债'
        elif index[4: 6] == 'H2':
            windCode = windCode[:-4] + index[14:17] + windCode[-4:]
            type = '同业存单'
        elif index[4: 6] == 'B9':
            type = '银行间金融债'
        elif name[-6: -3] == 'SCP':
            windCode = '01' + name[0: 2] + index[10: 12] + '0' + index[12: 14] + '.IB'
            type = '超短期融资券'
        elif name[-5: -3] == 'CP':
            windCode = '04' + name[0: 2] + index[10: 12] + '0' + index[12: 14] + '.IB'
            type = '短期融资券'
        else:
            print("Warning-unidentified bond: " + index + ' ' + name)
            exit(404)

        # putting the bond information into the dictionary
        if windCode not in d[bond]:
            w.start()
            getRate = w.wss(windCode, "creditrating, issuer_rating, amount, \
            latestissurercreditrating, issurercreditratingcompany").Data
            getDur = w.wsd(windCode, "duration", beg_date, end_date, "")
            d[bond][windCode] = [name, type, getRate[0][0], getRate[1][0], getRate[2][0], getRate[3][0], getRate[4][0], \
                                 getDur.Times, getDur.Data[0]]
    else:
        print("Warning-不明证券信息 " + index + " " + name)
        exit(404)
    print('testing windCode: ' + windCode)
    return windCode