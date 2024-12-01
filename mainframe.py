# -*- coding: cp936 -*-
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import json

import excel
import calc

originFileName = 'LED_template.xlsx'
dataFileName = 'data.xlsx'
TITLE=u'一键LED配单生成工具V20241201'

makingTimes = 0

def main():
    
    # 工具标题
    mainframe = tk.Tk()
    mainframe.title(TITLE)

    # 读取excel数据
    keyDict,detailDict = excel.getFromExcel(dataFileName)

    # 显示Dict
    showDict = {'XIANGMUMINGCHENG':u'项目名称',
                'CHANG':u'长',
                'GAO':u'高',
                'SHINEILED':u'LED点间距',
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
        label = tk.Label(mainframe,text=showDict[i]+' :')
        label.grid(row=n,sticky=tk.E) #靠东
        if i == "XIANGMUMINGCHENG":
            content = tk.Entry(mainframe, width=48)           
        if i == "CHANG":
            content = ttk.Entry(mainframe,width=48)
            if "CHANG" in configDict:
                content.insert(0,configDict[i])
            else:
                content.insert(0,4)
        if i == "GAO":
            content = ttk.Entry(mainframe,width=48)
            if i in configDict:
                content.insert(0,configDict[i])
            else:
                content.insert(0,2)
        if i == "SHINEILED":
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
            tk.Radiobutton(mainframe,text="LC-*",variable=v,value=0).grid(row=n,column=1)
            tk.Radiobutton(mainframe,text="LN-*",variable=v,value=1).grid(row=n,column=2)
            itemDict[i]=v

        # 开始显示
        if i != "SHOUFALEIBIE":
            itemDict[i] = content
            content.grid(row=n,column=1,columnspan=2)  
            

    #itemDict['XIANGMUMINGCHENG'].insert(0,'LED大屏清单-宇视'+datetime.now().strftime('%Y%m%d'))

    # 处理函数
    def doReplace():

        # 计数显示
        global makingTimes
        makingTimes += 1
        label = tk.Label(mainframe,text='    making..').grid(row=len(showDict.keys())*2+2,sticky=tk.W) 

        # 获取界面数据,计算LED数据
        replaceDict = {}

        replaceDict["CHANG"] = float(itemDict["CHANG"].get())
        replaceDict["GAO"] = float(itemDict["GAO"].get())
        replaceDict["KEY"] = itemDict["SHINEILED"].get()
        LEDMingCheng = replaceDict["KEY"]
        replaceDict["XIANGMUMINGCHENG"] = itemDict["XIANGMUMINGCHENG"].get()

        r = replaceDict
        r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]=[float(i) for i in detailDict[LEDMingCheng]['SHINEILED_CHICUN'].split('*')]
        r["BANZI_CHANG"],r["BANZI_GAO"] = calc.calcBanZi(r["CHANG"],r["GAO"], chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        r["SHIJI_CHANG"],r["SHIJI_GAO"] = calc.calcChangGao(r["BANZI_CHANG"],r["BANZI_GAO"],chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        
        r["MIANJI"]=r["SHIJI_CHANG"]*r["SHIJI_GAO"]
        # print(detailDict[LEDMingCheng])
        r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"] = [float(i) for i in detailDict[LEDMingCheng]['SHINEILED_FENBIANLV'].split('*')]
        r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"] = calc.calcFengBianLv(r["BANZI_CHANG"],r["BANZI_GAO"],[r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"]])

        r["PINGMUZONGGONGLV"] = r["MIANJI"] * detailDict[LEDMingCheng]["SHINEILED_GONGLV"]

        # 计算接收卡
        if itemDict['SHOUFALEIBIE'].get() == 0:
            jieshouka = "JIESHOUKA-LC"
        else:
            jieshouka = "JIESHOUKA-LN"
        JieShouKaList = [[detailDict[k][jieshouka+"_XINGHAO"],detailDict[k][jieshouka+"_DANJIA"]] for k in keyDict[jieshouka]]  #['LN-DH7508-S', 'LN-DH7512-S', 'LN-DH7516-S']
        # print(JieShouKaList)
        # print(replaceDict)
        
        r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_JIAGE"],r["JIESHOUKA_XINGHAO"]=calc.calcJieShouKa(r["BANZI_CHANG"],r["BANZI_GAO"],JieShouKaList,"SHINEILED",detailDict[LEDMingCheng])
        r["JIESHOUKA_SHULIANG"]=r["JIESHOUKA_CHANG"]*r["JIESHOUKA_GAO"]
        # print(r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_JIAGE"],r["JIESHOUKA_XINGHAO"])
        for i in keyDict[jieshouka]:
            if r["JIESHOUKA_XINGHAO"]  == detailDict[i][jieshouka+"_XINGHAO"]:
                for j in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                    r["JIESHOUKA_"+j] = detailDict[i][jieshouka+"_"+j]

        # 计算处理器
        if itemDict['SHOUFALEIBIE'].get() == 0:
            chuliqi = "LEDCHULIQI-LC"
        else:
            chuliqi = "LEDCHULIQI-LN"
        key, r["LEDCHULIQI_JIAGE"] = calc.calcChuLiQi(chuliqi,keyDict[chuliqi],detailDict,r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"])
        if key:
            r["LEDCHULIQI_SHULIANG"] = 1
        for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
            r["LEDCHULIQI_"+i] = detailDict[key][chuliqi+"_"+i]

        # 计算配电箱
        key, r["PEIDIANXIANG_JIAGE"] = calc.calcPeiDianXiang(keyDict["PEIDIANXIANG"],detailDict,r["PINGMUZONGGONGLV"])
        for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
            r["PEIDIANXIANG_"+i] = detailDict[key]["PEIDIANXIANG_"+i]

        # 计算LED电源数量
        r["DIANYUAN_SHULIANG"] = calc.calcDianYuan(r["BANZI_CHANG"],r["BANZI_GAO"])

        # 其他参数传递
        for i in ["CITIE","ANZHUANGFUWU","DIANYUAN","SHINEILED"]:
            if i == "SHINEILED":
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
        #print(r)
                    
        # 生成Excel文件
        if itemDict['XIANGMUMINGCHENG'].get():
            newFileName = itemDict['XIANGMUMINGCHENG'].get()+datetime.now().strftime('%Y%m%d%H%M')+'.xlsx'
        else:
            newFileName = 'LED清单-宇视'+datetime.now().strftime('%Y%m%d%H%M')+str(makingTimes)+'_'+str(r["SHIJI_CHANG"])+'x'+str(r["SHIJI_GAO"])+'.xlsx'
        #print(newFileName)
        excel.copyExcel(originFileName, newFileName, replaceDict)

        # 写入config.json
        with open('config.json','w') as f0:
            tmpDict = {}
            for key in itemDict:
                tmpDict[key]=itemDict[key].get()
                
            f0.write(json.dumps(tmpDict))

        # 更新状态显示
        label = tk.Label(mainframe,text='      '+str(makingTimes)+'  ok. '+newFileName).grid(row=len(showDict.keys())*2+2,column=0,stick="W",columnspan=2)#靠右
        
    # 显示次数和点击按钮
    label = tk.Label(mainframe,text='      0').grid(row=len(showDict.keys())*2+2,sticky=tk.W)#靠右
    tk.Button(mainframe,text=u'生成大屏清单',width=15,height=2,command=doReplace).grid(row=len(showDict.keys())*2+2, column=2)

    mainframe.mainloop()


if __name__=='__main__':
    main()
