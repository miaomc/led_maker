import excel
import random

def cutNumber(l):
    # 去头部 序号
    spec = '0 123456789.()、'
    new_l = []
    for line in l:
        for n,i in enumerate(line):
            if i not in spec:
                new_l.append(line[n:])
                break

    #print(new_l)
    return new_l


def quBaoGao(s):
    i_jcbg = s.find("检测报告")
    if i_jcbg == -1:
        return s
    i_end = s.find('）',i_jcbg)

    i = -1
    i_start = i_end
    while True:
        i = s.find('（',i+1,i_jcbg)
        if i!= -1:
            i_start = i
        else:
            break

    return s[:i_start]+s[i_end+1:]        


def combineCanShu(l1,l2,m=10,n=2,nn=1):
    """全部的普通CanShu
       m条带 检测报告 参数，删除检测报告
       n条 带星参数，删除星
       1条 2星 ★ 
    """
    XING = "★"
    
    pt_l=[]
    dx_l=[]
    d2x_l=[]
    for i in l2:
        if i[0] == XING:
            if i[1] == XING:
                d2x_l.append(i[2:]) # 去星不去报告
            else:
                dx_l.append(i[1:])  # 去星不去报告
        else:
            pt_l.append(quBaoGao(i))

    # 随机组合
    new_l = l1
    for tmp_l in [pt_l,dx_l,d2x_l]:
        if tmp_l == pt_l:
            num = m
        elif tmp_l == dx_l:
            num = n
        elif tmp_l == d2x_l:
            num = nn
            
        if num<=len(tmp_l):
            i_lm = random.sample(range(0,len(tmp_l)),num)
            i_lm.sort()
            #print(i_lm,tmp_l,num)
            #print(range(0,len(tmp_l)),num)
            for i in i_lm:
                new_l.append(tmp_l[i])
        else:
            new_l = new_l + tmp_l

    return new_l

    
def makeCanShu():
    """
    读取data.xlsx中 data-CANSHU 和 招标指引-控标参数
    生成到data-SHANGWUCANSHU中去
    """
    itemDict,detailDict = excel.getFromExcel('data.xlsx','data')
    
    # 读取data.xlsx CANSHU 到 baseCanShu {('HDMI 线缆-5m','CAB-HI-05M') : ["line1","line2",..],...}
    baseCanShu = {}
    for item in itemDict: # 'XIANCAI'
        for key in itemDict[item]: # 'HDMI 线缆-5m'
            if detailDict[key][item+"_CANSHU"]:
                baseCanShu[(key,detailDict[key][item+'_XINGHAO'])] = cutNumber(detailDict[key][item+"_CANSHU"].split('\n'))
    # print(baseCanShu[('1.8mm间距 LED显示模组','MW-M18-RB')])
    
    # 读取data.xlsx 招标指引 到 kongbiaoCanShu = kbCanShu {('MW7212','MW7215','MW7218'): ['line1','line2',...,...]}
    itemDict,detailDict = excel.getFromExcel('data.xlsx',u'招标指引')

    kbCanShu ={}
    for item in itemDict: # 'D06589'
        for key in itemDict[item]: # 'MW72\nMW72\n'
            kbCanShu[tuple(key.split('\n'))] = cutNumber(detailDict[key][item+"_KONGBIAOCANSHU"].split('\n'))
    #print(kbCanShu[('MW7915-P@SC','MW7918-P@SC')])
            
    # 生成 shangwuCanShu  shangwuCanShu = swCanShu {('HDMI 线缆-5m','CAB-HI-05M'): 'line1 line2...'}
    swCanShu = {}
    for key,xinghao in baseCanShu:
        l1 = baseCanShu[(key,xinghao)]
        l2 = []
        for kb_xinghaoList in kbCanShu.keys():
            if xinghao in kb_xinghaoList:
                l2=kbCanShu[kb_xinghaoList]
                break
            
##        if l2:
##            for i in (combineCanShu(l1,l2,10,2,1)):
##                print(i)
##            print
        canshuList = combineCanShu(l1,l2,10,2,1)
        tmp=''
        for n,i in enumerate(canshuList):
            tmp += str(n+1)+'. '+i+'\n'
        swCanShu[(key,xinghao)] = tmp
        
    # 写入 data.xlsx data SHANGWUCANSHU
    excel.writeInCanShu('data.xlsx','data',swCanShu,'SHANGWUCANSHU2')

if __name__ == "__main__":
    makeCanShu()
