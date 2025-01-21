# -*- coding: cp936 -*-
import tkinter as tk
from tkinter import ttk,messagebox
from datetime import datetime
import json

import excel
import calc

originFileName = 'LED_template.xlsx'
dataFileName = 'data.xlsx'
TITLE=u'一键LED配单生成工具V3.4-20250121'
makingTimes = 0
newFileName = ""
doReplaceDict = None
doSheetName = ""

def main():

    # 处理函数
    def doReplace(makeExcel):
        # 获取界面数据,计算LED数据
        replaceDict = {}
        try:
            replaceDict["CHANG"] = float(itemDict["CHANG"].get())
            replaceDict["GAO"] = float(itemDict["GAO"].get())
        except:
            messagebox.showwarning(u'警告',u'输入的长或者高不是数字')
            return
        replaceDict["KEY"] = itemDict["LED"].get() # "2.5mm间距 LED显示模组"

        # 开始
        # label = tk.Label(mainframe,text='    making..').grid(row=len(showDict.keys())*2+2) 
        LEDMingCheng = replaceDict["KEY"]
        LEDLEIBIE = list(detailDict[LEDMingCheng].keys())[0].split("_")[0] # 挑第一个key的名字取"_"前的部分
        
        if LEDLEIBIE == "SHINEILED": # 作sheet的分类
            sheetName = "LED"
        elif LEDLEIBIE == "XIANGTILED":
            sheetName = u"LED箱体"
        elif LEDLEIBIE == "HUWAILED":
            sheetName = u"户外LED"
        #print(LEDLEIBIE)
        # replaceDict["XIANGMUMINGCHENG"] = itemDict["XIANGMUMINGCHENG"].get()

        # 开始计算
        r = replaceDict

        # 单元的长和高 mm
        r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]=[float(i) for i in detailDict[LEDMingCheng][LEDLEIBIE+'_CHICUN'].split('*')]
        # 单元的几长 和 几高 整数
        r["BANZI_CHANG"],r["BANZI_GAO"] = calc.calcBanZi(r["CHANG"],r["GAO"], chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        # 实际 长和高 面积 m
        r["SHIJI_CHANG"],r["SHIJI_GAO"] = calc.calcChangGao(r["BANZI_CHANG"],r["BANZI_GAO"],chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        r["MIANJI"]=r["SHIJI_CHANG"]*r["SHIJI_GAO"]

        # 点间距 mm
        r["JIANJU"] = float(detailDict[LEDMingCheng][LEDLEIBIE+'_JIANJU'])

        # LED单元分辨率,可以是模组或者箱体
        r["MOZU_FENBIANLV_CHANG"] = round(r["MOZU_CHICUN_CHANG"]/r["JIANJU"])
        r["MOZU_FENBIANLV_GAO"] = round(r["MOZU_CHICUN_GAO"]/r["JIANJU"])

        # 整体大屏分辨率
        r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"] = calc.calcFengBianLv(r["BANZI_CHANG"],r["BANZI_GAO"],[r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"]])

        # 整屏功率
        r["PINGMUZONGGONGLV"] = r["MIANJI"] * detailDict[LEDMingCheng][LEDLEIBIE+"_GONGLV"]

        # 计算接收卡
        r['JIESHOUKA_MINGCHENG'] = None
        if itemDict['SHOUFALEIBIE'].get() == 0:
            jieshouka = "JIESHOUKA-LC"
        else:
            jieshouka = "JIESHOUKA-LN"
        JieShouKaList = [[detailDict[k][jieshouka+"_XINGHAO"],detailDict[k][jieshouka+"_DANJIA"]] for k in keyDict[jieshouka]]  #['LN-DH7508-S', 'LN-DH7512-S', 'LN-DH7516-S']
        # print(JieShouKaList)
        # print(replaceDict)
        
        r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_JIAGE"],r["JIESHOUKA_XINGHAO"]=calc.calcJieShouKa(r["BANZI_CHANG"],r["BANZI_GAO"],JieShouKaList,LEDLEIBIE,detailDict[LEDMingCheng])
        r["JIESHOUKA_SHULIANG"]=r["JIESHOUKA_CHANG"]*r["JIESHOUKA_GAO"]
        # print(r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_JIAGE"],r["JIESHOUKA_XINGHAO"])
        for i in keyDict[jieshouka]:
            if r["JIESHOUKA_XINGHAO"]  == detailDict[i][jieshouka+"_XINGHAO"]:
                for j in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                    r["JIESHOUKA_"+j] = detailDict[i][jieshouka+"_"+j]

        # 计算处理器 或者 发送卡
        if itemDict['SHOUFALEIBIE'].get() == 0:
            chuliqi = "LEDCHULIQI-LC"
        else:
            chuliqi = "LEDCHULIQI-LN"
        key = calc.calcChuLiQi(chuliqi,keyDict[chuliqi],detailDict,r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"])
        if key:  # 如果处理器能带的动
            r["LEDCHULIQI_SHULIANG"] = 1
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["LEDCHULIQI_"+i] = detailDict[key][chuliqi+"_"+i]
        else:
            # key为空，就是处理器超载了，需要计算发送卡
            if itemDict['SHOUFALEIBIE'].get() == 0:
                fasongka = "FASONGKA-LC"
            elif itemDict['SHOUFALEIBIE'].get() == 1:
                fasongka = "FASONGKA-LN"
            #key, shuliang = calc.calcFaSongKa1(fasongka,keyDict[fasongka],detailDict,r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"])
            key,shuliang,r["FASONGKAZUWANG_LIST"] = calc.calcFaSongKa(fasongka,keyDict[fasongka],detailDict,[r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"]],[r["BANZI_CHANG"],r["BANZI_GAO"]])    
            r["LEDCHULIQI_SHULIANG"] = shuliang
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["LEDCHULIQI_"+i] = detailDict[key][fasongka+"_"+i]
            # 切换为 LED及拼控 的sheet & 计算拼控
            if LEDLEIBIE == "SHINEILED":
                sheetName = u"LED及拼控"
            elif LEDLEIBIE == "XIANGTILED":
                sheetName = u"LED箱体及拼控"
            keyU, kou8_shuliang = calc.calcPingKong(shuliang)

            r["PINGKONG-SHURU_SHULIANG"] = 1
            r["PINGKONG-SHUCHU_SHULIANG"] = kou8_shuliang
            
            if keyU == "2U":
                key = u"视频拼接处理服务器(2U)"
                keyShuRu = u"4路HDMI输入板卡(2U)"
                keyShuChu = u"8路HDMI输出板卡(2U)"
            elif keyU == "3U":
                key = u"视频拼接处理服务器(3U)"
                keyShuRu = u"8路HDMI输入板卡(3U)"
                keyShuChu = u"8路HDMI输出板卡(3U)"
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["PINGKONG_"+i] = detailDict[key]["PINGKONG_"+i]
                r["PINGKONG-SHURU_"+i] = detailDict[keyShuRu]["PINGKONG-SHURU_"+i]
                r["PINGKONG-SHUCHU_"+i] = detailDict[keyShuChu]["PINGKONG-SHUCHU_"+i]
        #print(r)
        
        # 计算配电箱
        key, r["PEIDIANXIANG_JIAGE"] = calc.calcPeiDianXiang(keyDict["PEIDIANXIANG"],detailDict,r["PINGMUZONGGONGLV"])
        # print(key,r["PEIDIANXIANG_JIAGE"],r["PINGMUZONGGONGLV"])
        if not key:
            key="PEIDIANXIANG_ELSE"
        for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
            r["PEIDIANXIANG_"+i] = detailDict[key]["PEIDIANXIANG_"+i]

        # 计算LED电源数量
        r["DIANYUAN_SHULIANG"] = calc.calcDianYuan(r["BANZI_CHANG"],r["BANZI_GAO"])

        # 其他参数传递
        for i in ["CITIE","ANZHUANGFUWU","DIANYUAN",LEDLEIBIE]:
            if i == LEDLEIBIE:
                key = replaceDict["KEY"]
                #print(key)
            else:
                key = keyDict[i][0]
            #print(detailDict.keys())
            for j in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                if i=="SHINEILED" and j=="DANJIA":
                    r[i+"_"+j] = detailDict[key][i+"_"+j]*20
                else:
                    r[i+"_"+j] = detailDict[key][i+"_"+j]
        
        # 显示计算内容
        tmpList = []
        tmpList.append(u"实际宽高: %s * %s = %s 平方米"%(str(r["SHIJI_CHANG"]),str(r["SHIJI_GAO"]),str(r["MIANJI"])))
        tmpList.append(u"单元宽高: %s * %s "%(str(r["BANZI_CHANG"]),str(r["BANZI_GAO"])))
        tmpList.append(u"%s: %d * %d = %d 张"%(r["JIESHOUKA_MINGCHENG"],r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_SHULIANG"]))
        tmpList.append(u"发送数量：%s"%r["LEDCHULIQI_SHULIANG"])
        tmpList.append(u"总分辨率: %d * %d = %d 像素"%(r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"],r["FENBIANLV_CHANG"]*r["FENBIANLV_GAO"]))
        tmpList.append(u"电源数量: %d 个"%r["DIANYUAN_SHULIANG"])
        tmpList.append(u"整体功率: %d W"%r["PINGMUZONGGONGLV"])
        
        textItem.delete('0.0',tk.END)
        for n,i in enumerate(tmpList):
            textItem.insert('%d.0'%(n+1),i+'\n')

        # 写入config.json
        with open('config.json','w') as f0:
            tmpDict = {}
            for key in itemDict:
                tmpDict[key]=itemDict[key].get()
                
            f0.write(json.dumps(tmpDict))
            
        global makingTimes
        # 生成Excel文件
        if makeExcel:
            newFileName = u'宇视'+itemDict['LED'].get().split(' ')[0]+u'LED大屏'+'('+str(r["SHIJI_CHANG"])+'x'+str(r["SHIJI_GAO"])+u')-'+datetime.now().strftime('%Y%m%d%S')+str(makingTimes)+'.xlsx'
            excel.copyExcel(originFileName, newFileName, r, sheetName=sheetName)
            # 计数显示
            
            makingTimes += 1
            # 更新状态显示
            label = tk.Label(mainframe,text=u'生成第'+str(makingTimes)+u'份ok: '+newFileName).grid(row=len(showDict.keys())*2+2,column=0,columnspan=4,sticky=tk.W)#靠右
        
    
    # --------------------------- 分隔符 ----------------------------

    
    # 标题
    mainframe = tk.Tk()
    mainframe.title(TITLE)

    # 读取excel数据
    keyDict,detailDict = excel.getFromExcel(dataFileName, worksheet='data')
    keyDict["LED"] = keyDict["SHINEILED"]+keyDict["HUWAILED"]+keyDict["XIANGTILED"] # 创建并汇总 三类LED到 "LED"这个key
    
    # 显示Dict
    showDict = {
                'CHANG':u'长(m)',
                'GAO':u'高(m)',
                'LED':u'LED点间距',
                'SHOUFALEIBIE':u'收发系统类别'
                }

    # 读取config.json
    configDict={}
    try:
        with open('config.json','r') as f1:
            configDict=json.loads(f1.read())
    except:
        pass
    
    itemDict = {} # 组件字典
    for n,i in enumerate(showDict.keys()):
        label = tk.Label(mainframe,text=showDict[i]+' :').grid(row=n,sticky=tk.E) #靠东        
        if i == "CHANG" or "GAO":
            content = ttk.Entry(mainframe, width=48)
            if i in configDict:
                content.insert(0,configDict[i])
            else:
                content.insert(0,3)
        if i == "LED":
            content = ttk.Combobox(mainframe,width=45)
            content['value'] = keyDict[i]
            if i in configDict:
                content.insert(0,configDict[i])
            else:
                content.current(0)  # 选择默认第0个
        if i == "SHOUFALEIBIE":
            v = tk.IntVar()
            if i in configDict:
                v.set(configDict[i])
            else:
                v.set(0)
            tk.Radiobutton(mainframe,text=u"宇视(卡莱特)",variable=v,value=0).grid(row=n, column=1)
            tk.Radiobutton(mainframe,text=u"宇视(诺瓦)",variable=v,value=1).grid(row=n, column=2)
            tk.Radiobutton(mainframe,text=u"凯仕达",variable=v,value=2).grid(row=n, column=3)
            itemDict[i]=v

        # 显示内容占横向3个格子
        if i != "SHOUFALEIBIE":
            itemDict[i] = content
            content.grid(row=n,column=1,columnspan=3)
    
    # 显示次数和点击按钮
    # label = tk.Label(mainframe,text='  0').grid(sticky=tk.W)
    tk.Button(mainframe,text=u'计算',width=15,height=2,command=lambda:doReplace(False)).grid(row=n+1,column=2)
    tk.Button(mainframe,text=u'生成大屏清单',width=15,height=2,command=lambda:doReplace(True)).grid(row=n+1,column=3)

    # 显示计算结果
    textItem = tk.Text(mainframe,height=9,width=60)
    textItem.grid(columnspan=4,pady=5,padx=5)

    
    mainframe.mainloop()


if __name__=='__main__':
    main()
