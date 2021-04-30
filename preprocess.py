import os
import pandas as pd
import numpy as np
import random
import gc
import time
import re
from tqdm import tqdm
import matplotlib.pyplot as plt
import numpy as np
import tensorflow as tf
from sklearn import metrics
import pickle
rawdatadir='../data/'


def process(dataset):
    list_stuid = []
    list_skillid = []
    list_stu = []  # 学生ID
    list_skill = []  # 技能
    list_correct = []  # 是否正确
    first_skill = []  # 每个学生第一次做的题
    first_correct = []  # 每个学生第一次做的题是否正确
    index_all = []  # 学生的每次做的题量
    index_singl = []
    index_minus = []
    list_num=[]
    index_for=[]
    stu_id=0
    for subdirs in os.listdir(rawdatadir):
        if dataset not in subdirs:
            continue
        istest1=True
        dir=os.path.join(rawdatadir, subdirs)
        with open(dir, 'r') as f1:
            lines1 = f1.readlines()
            len_num=int(len(lines1)/3)
            for num in range(0,len_num):
                #line = line.strip().split('\t')
                try:
                    lenlist = int(re.findall(r"\d+", lines1[num * 3])[0])
                except:
                    lenlist = eval(lines1[num * 3])
                if lenlist>1 :
                    eval_skill=list(map(int, re.findall(r"\d+", lines1[num*3+1])))
                    eval_correct = list(map(int, re.findall(r"\d+", lines1[num * 3 + 2])))
                #else :
                    #if lenlist>100:
                        #eval_skill=list(eval(lines1[num*3+1]))
                        #eval_correct=list(eval(lines1[num*3+2]))
                    #list_skill = list_skill + eval_skill
                    #list_correct = list_correct + eval_correct
                    list_stu = list_stu +lenlist* [stu_id]
                    stu_id=stu_id+1
                    list_num = list_num+lenlist*[lenlist]#单个学生做题的总数量
                    index_all=index_all+list(range(0,lenlist))
                    for new_skill in list(set(eval_skill)):
                        start=eval_skill.index(new_skill)
                        #找到含有同一skill的下标，并将其取出，并且对应下标减去开始下标取出
                        index_singl=index_singl+[m-start for m, n in enumerate(eval_skill) if n == new_skill]
                        minus_for=[m-start for m, n in enumerate(eval_skill) if n == new_skill]
                        minus2=minus_for.copy()
                        minus2.pop()
                        minus2.insert(0, 0)
                        index_minus=index_minus+(np.array(minus_for)-np.array(minus2)).tolist()
                        list_skill=list_skill+[eval_skill[m] for m, n in enumerate(eval_skill) if n == new_skill]

                        list_correct = list_correct + [eval_correct[m] for m, n in enumerate(eval_skill) if n == new_skill]
                        for_skillcor=[eval_correct[m] for m, n in enumerate(eval_skill) if n == new_skill]
                        for_skillcor.pop()
                        for_skillcor.insert(0, 0)
                        index_for=index_for+for_skillcor
                        first_skill.append(new_skill)
                        first_correct.append(eval_correct[eval_skill.index(new_skill)])
    print(len(index_minus),len(index_all),len(first_skill),len(first_correct))
    constru = {"list_stu": list_stu, "list_skill": list_skill, "list_correct": list_correct,'index_singl':index_singl,
               "index_all": index_all,"list_num":list_num,'index_minus':index_minus,'index_for':index_for}
    dataframe = pd.DataFrame(constru, columns=['list_stu', 'list_skill', 'list_correct','index_singl','index_all','list_num','index_minus','index_for'])
    first_cons = {'first_skill': first_skill, 'first_correct': first_correct}
    df_first_cons = pd.DataFrame(first_cons, columns=['first_skill', 'first_correct'])
    return dataframe,df_first_cons

# 定义画散点图的函数
def draw_scatter(x1, y1):
    """
    :return: None
   """
    # 加载数据
    # 创建画图窗口
    fig = plt.figure()
    # 将画图窗口分成1行1列，选择第一块区域作子图
    ax1 = fig.add_subplot(1, 1, 1)
    # 设置标题
    ax1.set_title('做题数量统计')
    # 设置横坐标名称
    ax1.set_xlabel('学生做题的数量')
    # 设置纵坐标名称
    ax1.set_ylabel('做该数量的题学生个数')
    # 画散点图
    ax1.scatter(x1, y1, s=20, c='k', marker='.')
    # 调整横坐标的上下界
    plt.show()


def getlistnum(li):#这个函数就是要对列表的每个元素进行计数
    li = list(li)
    set1 = set(li)
    dict1 = {}
    for item in set1:
        dict1.update({item:li.count(item)})
    return dict1
#处理训练集/测试集数据
def gen(flag_train_test,isplt=False):
    #isPlt 是否画图
    dataframe,df_first_cons=process(flag_train_test)
    dataframe=dataframe[dataframe["list_num"]!=1]
    if isplt:
        numhist=dataframe['list_stu']
        dict1=getlistnum(numhist)
        draw_scatter(list(dict1.keys()),list(dict1.values()))
        print(np.mean(list(dict1.values())),np.min(list(dict1.values())),np.max(list(dict1.values())))
    #dataframe=dataframe[dataframe["list_num"]<400]
    df1 = dataframe
    dataframe=dataframe[['list_stu', 'list_skill', 'list_correct']]
    #dataframe.to_csv('../data/'+flag_train_test+'.txt', sep='\t', header=False, index=False)
    set_skill_corr={}
    fir_skill_corr={}
    for skill in set(dataframe['list_skill']):
        #计算某个skill正确率
        dfskill_correct=dataframe[dataframe['list_skill']==skill]['list_correct']
        set_skill_corr[str(skill)]=sum(dfskill_correct)/len(dfskill_correct)
        #计算skill第一次做正确的概率
        fir_skill_correct=df_first_cons[df_first_cons['first_skill']==skill]['first_correct']
        fir_skill_corr[str(skill)]=sum(fir_skill_correct)/len(fir_skill_correct)
    return df1,dataframe,set_skill_corr,fir_skill_corr


df1,_,test_corr,test_skill_corr=gen('test',isplt=False)
#dft,_,set_skill_corr,fir_skill_corr=gen('train.csv',isplt=False)
#
#pickle.dump(set_skill_corr,open('set_skill_corr.txt', 'wb') )
#pickle.dump(fir_skill_corr,open('fir_skill_corr.txt', 'wb') )

set_skill_corr=pickle.load(open('set_skill_corr.txt', 'rb'))
fir_skill_corr=pickle.load(open('fir_skill_corr.txt', 'rb'))


df1['forget']=None
df1['hard']=None
df1['hardfirst']=None
df1['cor_can']=None
df1['pred']=None

Pforget,P1_can=0.1,0.8
auclist=[]
for skill in set(df1['list_skill']):

    ds=df1[df1['list_skill'] == skill]
    ds.loc[ds['index_singl'] == 0, 'hardfirst']=fir_skill_corr[str(skill)]
    ds.loc[ds['index_singl'] != 0, 'hardfirst'] = 0.5
    pred_x=set_skill_corr[str(skill)] +ds['hardfirst']-ds['index_minus']*Pforget+ds['index_for']*P1_can
    df1.loc[df1['list_skill'] == skill, 'pred']=pred_x.apply(lambda x: (np.exp(x)-(np.exp(-x)))/(np.exp(x)+(np.exp(-x))))
    pred1=df1.loc[df1['list_skill'] == skill, 'pred']
    print(pred1)
    y =df1.loc[df1['list_skill'] == skill, 'list_correct']
    print(y)
    #fpr, tpr, thresholds = metrics.roc_curve(y, pred)
    #metrics.auc(fpr, tpr)
    try:
        AUC= metrics.roc_auc_score(y, pred1)
    except:
        pass
    auclist.append(AUC)
    m = 1

pred2=df1['pred']
y2 =df1['list_correct']
fpr, tpr, thresholds = metrics.roc_curve(y2, pred2)
draw_scatter(fpr, tpr)
#metrics.auc(fpr, tpr)
AUCtotal= metrics.roc_auc_score(y2, pred2)
print(AUCtotal)
m=1










