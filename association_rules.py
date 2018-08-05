import openpyxl
import csv

def loaddataset(filename):
    dataset=[]                  #保存数据集
    data=[]                     #保存归纳数据集
    peoset=set()                #购物人
    wb=openpyxl.load_workbook(filename)
    sheet = wb['Sheet1']
    for line in sheet.values:
        dataset.append(line)
    dataset = [[str(i) for i in row] for row in dataset]#全部字符串化
    cleanout(dataset)
    for row in dataset:
        peoset.add(row[0])
    for peo in peoset:#对于每个peopel，查找在数据集中对应的每一条购买记录，并将其购买的物品去重
        lit=[]
        for line in dataset:
            if line[0]==peo:
                if line[1] not in lit:
                    lit.append(line[1])
        data.append(lit)
    return data

def cleanout(dataset):#数据清洗
    for i in range(7):
        for item in dataset:
            if not item[0].isdigit() or item[0]=='':
                dataset.remove(item)
            else:
                if not item[1].isdigit() or item[0]=='':
                    dataset.remove(item)

def apriori(dataset,minsupport):
    C1=createC1(dataset)                         #创建1-项集
    D= list(map(set,dataset))                   #数据集的list
    L1,supportdata = scanD(D,C1,minsupport)     #1-频繁项集
    L=[]                                        #保存各个频繁项集
    L.append(L1)                                #将1频繁项集加入到L中
    k=2
    while(len(L[k-2])>0):
        ck=apriorigen(L[k-2],k)                 #构造k候选项集
        lk,supK=scanD(D,ck,minsupport)          #得到k频繁项集
        supportdata.update(supK)                #将新产生的键值对存放到supportData中，并将k-频繁项集放到L中
        L.append(lk)
        k+=1
    return L,supportdata

def createC1(dataset):#构造1频繁项集
    C1=[]
    for i in dataset:
        for j in i:
            if not [j] in C1:
                C1.append([j])
    C1.sort()
    return list(map(frozenset,C1))

def scanD(D,ck,minsupport):#从1候选集产生1频繁集
    ssCnt={}
    for tid in D:
        #对于每条交易记录，如果k-项集中的商品存在于这条购物记录中，则增加这些商品的取值
        for can in ck:
            if can.issubset(tid):#如果候选集中存在商品
                if not can in ssCnt:
                    ssCnt[can]=1
                else: ssCnt[can]+=1
    numItems=float(len(D))
    retList=[]
    supportdata={}
    for key in ssCnt:
        support=ssCnt[key]/numItems#求得每个候选集的支持度
        if support>=minsupport:
            retList.insert(0,key)
        supportdata[key]=support#将符合的候选集加入到频繁集中
    return retList,supportdata

def apriorigen(lk,k):#合并形成新的项集
    retList=[]
    lenlk=len(lk)
    for i in range(lenlk):
        for j in range(i+1,lenlk):
            l1=list(lk[i])[:k-2]
            l2=list(lk[j])[:k-2]
            l1.sort()
            l2.sort()
            if l1==l2:
                retList.append(lk[i]|lk[j])
    return retList

def generateRules(l,supportdata,minconf=0.8):#生成关联规则
    bigrulelist=[]
    for i in range(1,len(l)):
        for freqset in l[i]:#对于每一个频繁集
            H1 = [frozenset([item])for item in freqset]
            if(i>1):
                rulesfromconseq(freqset,H1,supportdata,bigrulelist,minconf)
            else:
                calcconf(freqset,H1,supportdata,bigrulelist,minconf)#找到所有支持最小可信度的规则
    return bigrulelist

def calcconf(freqset,h,supportdata,bigrulelist,minconf):
    prunedh=[]
    for conseq in h:#遍历H中的所有项集并计算他们的可信度
        conf = supportdata[freqset]/supportdata[freqset-conseq]
        if conf>=minconf:
            print(freqset-conseq,'-->',conseq,'confidence:',conf)
            bigrulelist.append((freqset-conseq,conseq,conf))#将这一条规则加入到规则列表中
            prunedh.append(conseq)#同时加入到当前项数的规则列表中
    return  prunedh

def rulesfromconseq(freqset,h,supportdata,bigrulelist,minconf):
    m=len(h[0])
    if(len(freqset)>(m+1)):
        hmp=apriorigen(h,m+1)#生成H中元素的无重复组合
        hmp=calcconf(freqset,hmp,supportdata,bigrulelist,minconf)#计算新生成的组合的可信度
        if(len(hmp)>1):
            rulesfromconseq(freqset,hmp,supportdata,bigrulelist,minconf)#继续生成

def check(dataset):
    lit=set()
    for item in dataset:
        for j in item:
            lit.add(j)
    print(len(lit))
    sum=0
    for item in dataset:
        sum+=len(item)
    print(sum)

def createdataset():
    return [[1,3,4],[2,3,5],[1,2,3,5],[2,5]]

def load(filename):
    csv_reader=csv.reader(open(filename))
    return csv_reader

# dataset=load(r'C:\Users\yang\Desktop\groceries.csv')
# data=[]
# for row in dataset:
#     data.append(row)
dataset=loaddataset(r'C:\Users\yang\Desktop\OnlineRetail5.xlsx')   #加载数据集
# with open(r'C:\Users\yang\Desktop\association_rule.txt','w') as f:
#     f.write(str(dataset))
# dataset=createdataset()
# check(dataset)

L,supportdata=apriori(dataset,0.01)                                         #apriori算法
rule=generateRules(L,supportdata,0.6)                                      #关联规则
with open(r'C:\Users\yang\Desktop\association_rule.txt','w') as f:
    f.write(str(rule))