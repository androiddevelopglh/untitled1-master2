# -*- coding: utf-8 -*-
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rc
import textwrap
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
rc('mathtext', default='regular')
df1=pd.DataFrame({"A":[14,4,5,4,1],
                  "B":[5,2,54,3,2],
                  "C":[20,20,7,3,8],
                  "D":[14,3,6,2,6]})

df2=pd.DataFrame({"A":[12,4,5,2,1],
                  "B":[7,2,54,3,3],
                  "C":[20,16,11,3,8],
                  "D":[14,3,4,2,6]})

df3=df1*df2

df1.mul(df2)

plt.rcParams['font.sans-serif'] = ['SimHei']#让lable可以是中文
plt.rcParams['axes.unicode_minus'] = False #让lable可以是中文
time = np.arange(10)
temp = np.random.random(10)*30
Swdown = np.random.random(10)*100-10
Rn = np.random.random(10)*100-10


list1=[Rn,temp,temp]
a='''操纵'''
dedented_text = textwrap.dedent(a).strip()
nn=textwrap.fill(a)
print(nn)
print(dedented_text)
fig=plt.figure(figsize=[10,8])
ax= fig.add_subplot(111)

with PdfPages(r'ddd.pdf') as pdf:
    for li in list1:
        if (li==temp).all():
            ax2 = ax.twinx()
            ax2.plot(time, li, '-r', label='temp')
            ax2.legend(loc='upper left')
            ax2.set_ylabel(r"Temperature ($^\circ$C)")
            ax2.set_ylim(0, 35)
        else:

            ax.plot(time, li, '-', label = 'Swdown')
            ax.legend(loc='upper right')
            ax.grid()
            ax.set_xlabel("Time (h)")
            ax.set_ylabel(r"Radiation ($MJ\,m^{-2}\,d^{-1}$)")
            ax.set_ylim(-20,100)


    #plt.title(dedented_text, )    #设置字体旋转角度
    plt.title(nn,loc='center', fontsize='large', fontweight='bold', horizontalalignment='left',color='blue', wrap=True,
                  bbox=dict(facecolor='g', edgecolor='blue', alpha=0.65))  # 设置字体大小与格式
    plt.show()
    pdf.savefig()  # saves the current figure into a pdf page
    plt.close()
