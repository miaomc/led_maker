import json


def roundup(i):
    if int(i) < i:
        return int(i) + 1
    else:
        return i
    

def calcBanZi(chang=1, gao=1, chicun=[320,160]):
    banzi_chang = int(chang*1000 / chicun[0])
    banzi_gao = int(gao*1000 / chicun[1])
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


def calcFaSongKa(fasongka,keyList,detailDict,fenbianlv_chang,fenbianlv_gao):
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

            
def calcPeiDianXiang(keyList,detailDict,gonglv):
    gonglv=gonglv*1.3  # 向上抛30%
    zuizhong_key = ""
    zuizhong_jiage = 0
    for i in keyList:
        if detailDict[i]["PEIDIANXIANG_DAIZAI"]>gonglv and detailDict[i]["PEIDIANXIANG_DANJIA"]<zuizhong_jiage or zuizhong_jiage==0:
            zuizhong_key = i
            zuizhong_jiage = detailDict[i]["PEIDIANXIANG_DANJIA"]

    return zuizhong_key,zuizhong_jiage


def calcDianYuan(banzi_chang, banzi_gao):
    return roundup(banzi_chang * banzi_gao / 6)
