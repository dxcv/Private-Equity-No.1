# -*- coding:utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt

plt.rcParams['axes.edgecolor'] = 'white'
plt.rcParams['xtick.color'] = 'grey'
plt.rcParams['ytick.color'] = 'grey'
plt.rcParams['axes.labelcolor'] = "black"
plt.rcParams['font.sans-serif']=['KaiTi']
plt.rcParams['axes.unicode_minus']=False
#zhfont1 = matplotlib.font_manager.FontProperties(fname='/usr/share/fonts/cjkunifonts-ukai/ukai.ttc')

def plotting_data():
    df = pd.read_excel('对比.xlsx').transpose()
    print(df)
    #dates = pd.date_range(start = df['日期'][0], periods = numDate,freq= 'W')
    df.columns = ['PE', 'Profit']
    #df.index = pd.to_datetime(df['Date'])
    #print(df.index)
    fig, ax1 = plt.subplots()
    plt.xlabel(u'日期',fontsize = 13)
    plt.xticks(np.arange(2008, 2017, 1))
    plt.title(u'指数PE与盈利变化对比图', fontsize=15)
    df[u'PE'].plot(c=(0.5,0.5,0.5), lw=1.5, legend=True, figsize=(10, 5),rot=0)
    df[u'Profit'].plot(c=(0.8,0.2,0), lw=1.5, legend=True, figsize=(10, 5),rot=0)
    #df[u'中证500'].plot(c=(0,0.3,0.9), lw=1.5, legend=True, figsize=(10, 5), rot=0)
    plt.legend(loc=2, edgecolor='w')
    ax1.yaxis.grid()
    #ax1.legend(bbox_to_anchor=(0.35, -0.15))
    #ax1.legend(bbox_to_anchor=(0.85, -0.15))

    plt.show()

plotting_data()