# -*- coding: cp936 -*-
import json


def roundup(i):
    if int(i) < i:
        return int(i) + 1
    else:
        return i
    

def calcBanZi(chang=1, gao=1, chicun=[320,160]):
    banzi_chang = round(chang*1000 / chicun[0])
    banzi_gao = round(gao*1000 / chicun[1])
    return banzi_chang,banzi_gao


def calcChangGao(banzi_chang=1, banzi_gao=1, chicun=[320,160]):
    shiji_chang = banzi_chang * chicun[0] /1000
    shiji_gao = banzi_gao * chicun[1] /1000
    return shiji_chang,shiji_gao


def calcFengBianLv(banzi_chang, banzi_gao, fenbianlv):
    return banzi_chang*fenbianlv[0], banzi_gao*fenbianlv[1]


def calcJieShouKa(banzi_chang,banzi_gao,JieShouKaList,ITEM,ledDict):
    """
    INPUT:
    12 12
    [['LN-DH7508-S', 76], ['LN-DH7512-S', 87], ['LN-DH7516-S', 99]]
    SHINEILED
    {'SHINEILED_XINGHAO': 'MW-M12-RB', 'SHINEILED_MINGCHENG': '1.25mm间距LED显示模组',
    'SHINEILED_CANSHU': '点间距:P1.2\n（Kg/pcs）:0.45±0.1\n', 'SHINEILED_DANJIA': 350,
    'SHINEILED_FENBIANLV': '256*128', 'SHINEILED_CHICUN': '320*160', 'SHINEILED_GONGLV': 600, 'SHINEILED_LN-DH7508-S': None,
    'SHINEILED_LN-DH7512-S': '[1,8],[2,4]', 'SHINEILED_LN-DH7516-S': None}
    """
    #print(banzi_chang,banzi_gao,JieShouKaList,ITEM,ledDict)
    zuizhong_xinghao=''
    zuizhong_chang=0
    zuizhong_gao=0
    zuizhong_jiage=0
    for i0 in JieShouKaList:
        #print(i0)
        i = i0[0]
        if ledDict[ITEM+'_'+i]:
            l0 = json.loads('['+ledDict[ITEM+'_'+i]+']') # l =[[2,4],[1,8]]
            for l in l0:
                jieshou_chang = roundup(banzi_chang / l[0])
                jieshou_gao = roundup(banzi_gao / l[1])
                jiage = jieshou_chang*jieshou_gao*i0[1]
                if jiage<zuizhong_jiage or zuizhong_jiage == 0:
                    zuizhong_chang = jieshou_chang
                    zuizhong_gao = jieshou_gao
                    zuizhong_jiage = jiage
                    zuizhong_xinghao = i0[0]

    return zuizhong_chang,zuizhong_gao,zuizhong_jiage,zuizhong_xinghao


def calcChuLiQi(chuliqi,keyList,detailDict,fenbianlv_chang,fenbianlv_gao):
    zuizhong_key = ""
    zuizhong_jiage = 0
    for i in keyList:
        #print(detailDict[i])
        if detailDict[i][chuliqi+"_DAIZAI"]>fenbianlv_chang*fenbianlv_gao:
            if detailDict[i][chuliqi+"_DANJIA"]<zuizhong_jiage or zuizhong_jiage==0:
                zuizhong_key = i
                zuizhong_jiage = detailDict[i][chuliqi+"_DANJIA"]
        
    return zuizhong_key


def calcFaSongKa(fasongka,keyList,detailDict,danyuanFBL=[416,312],danyuanCG=[26,7]):
    DWKdaizai=650000
    
    PKzuigaoGC=[2048,2048]
    danyuanFBL_zong = danyuanFBL[0]*danyuanFBL[1]

    zuizhong_key=None
    for fsk in keyList:
        ZJdaizai=int(detailDict[fsk][fasongka+"_DAIZAI"])
        ZJwangkoushu=int(detailDict[fsk][fasongka+"_WANGKOU"])
        # print(ZJdaizai,ZJwangkoushu)
        # 计算单网口最多可以带几个单元
        dangwangkou_zuiduo = int(DWKdaizai/danyuanFBL_zong)
        #print(dangwangkou_zuiduo)

        # 从单网口带满开始计算
        rList = []
        for i in range(dangwangkou_zuiduo,0,-1):
            # 用满整机单网口开始计算
            for j in range(ZJwangkoushu,0,-1):
                #print(i,j)
                # 不能超过单机总带载,不能超过拼控的最长和高
                if i*j*danyuanFBL_zong > ZJdaizai:
                    continue
                # print(i,j)
                # 假设单网口是1xi,或者ix1的方式
                if i*danyuanFBL[0] <= PKzuigaoGC[0] and j*danyuanFBL[1] <= PKzuigaoGC[1]:
                    rList.append([i,j])
                    zuizhong_key=fsk
                if j*danyuanFBL[0] <= PKzuigaoGC[0] and i*danyuanFBL[1] <= PKzuigaoGC[1]:
                    rList.append([j,i])
                    zuizhong_key=fsk
    
    # 计算需要多少张发送卡
    zuizhong_shuliang = 0
    zuizhongList = []
    for i in rList:
        # print(i,roundup(danyuanCG[0]/i[0]),roundup(danyuanCG[1]/i[1]),roundup(danyuanCG[0]/i[0])*roundup(danyuanCG[1]/i[1]))
        if roundup(danyuanCG[0]/i[0])*roundup(danyuanCG[1]/i[1]) < zuizhong_shuliang or zuizhong_shuliang==0:
            zuizhong_shuliang = roundup(danyuanCG[0]/i[0])*roundup(danyuanCG[1]/i[1])
            zuizhongList.append(i)

    return zuizhong_key,zuizhong_shuliang,zuizhongList
            
    


def calcFaSongKa1(fasongka,keyList,detailDict,fenbianlv_chang,fenbianlv_gao):
    zuizhong_key = None
    zuizhong_shuliang = 0
    zuizhong_jiage = 0
    for i in keyList:
        shuliang = fenbianlv_chang*fenbianlv_gao / detailDict[i][fasongka+"_DAIZAI"] # ！！！这个计算过程太粗略，有问题，实际上是按长和高带多少数量计算的
        shuliang = roundup(shuliang)
        if shuliang*detailDict[i][fasongka+"_DANJIA"]<zuizhong_jiage or zuizhong_jiage==0:
            zuizhong_shuliang = shuliang
            zuizhong_jiage = shuliang*detailDict[i][fasongka+"_DANJIA"]
            zuizhong_key = i

    return zuizhong_key, zuizhong_shuliang


def calcPingKong(fasongka_shuliang):
    kou8 = roundup(fasongka_shuliang/8)
    if kou8 <= 3: # 小于等于3张8口输出板卡就用2U拼控
        keyU = "2U"
    else:
        keyU = "3U" # !!!没有考虑超过10张的情况

    return keyU, kou8

            
def calcPeiDianXiang(keyList,detailDict,gonglv):
    gonglv=gonglv*1.3  # 向上抛30%
    zuizhong_key = ""
    zuizhong_jiage = 0
    for i in keyList:
        #print(i,detailDict[i]["PEIDIANXIANG_DANJIA"])
        if detailDict[i]["PEIDIANXIANG_DAIZAI"]>gonglv and (detailDict[i]["PEIDIANXIANG_DANJIA"]<zuizhong_jiage or zuizhong_key==""):
            zuizhong_key = i
            zuizhong_jiage = detailDict[i]["PEIDIANXIANG_DANJIA"]
            

    return zuizhong_key,zuizhong_jiage


def calcDianYuan(banzi_chang, banzi_gao):
    return roundup(banzi_chang * banzi_gao / 6)
