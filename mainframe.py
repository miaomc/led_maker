# -*- coding: cp936 -*-
import tkinter as tk
from tkinter import ttk,messagebox
from datetime import datetime
import json

import excel
import calc

originFileName = 'LED_template.xlsx'
dataFileName = 'data.xlsx'
TITLE=u'һ��LED�䵥���ɹ���V3.4-20250121'
makingTimes = 0
newFileName = ""
doReplaceDict = None
doSheetName = ""

def main():

    # ������
    def doReplace(makeExcel):
        # ��ȡ��������,����LED����
        replaceDict = {}
        try:
            replaceDict["CHANG"] = float(itemDict["CHANG"].get())
            replaceDict["GAO"] = float(itemDict["GAO"].get())
        except:
            messagebox.showwarning(u'����',u'����ĳ����߸߲�������')
            return
        replaceDict["KEY"] = itemDict["LED"].get() # "2.5mm��� LED��ʾģ��"

        # ��ʼ
        # label = tk.Label(mainframe,text='    making..').grid(row=len(showDict.keys())*2+2) 
        LEDMingCheng = replaceDict["KEY"]
        LEDLEIBIE = list(detailDict[LEDMingCheng].keys())[0].split("_")[0] # ����һ��key������ȡ"_"ǰ�Ĳ���
        
        if LEDLEIBIE == "SHINEILED": # ��sheet�ķ���
            sheetName = "LED"
        elif LEDLEIBIE == "XIANGTILED":
            sheetName = u"LED����"
        elif LEDLEIBIE == "HUWAILED":
            sheetName = u"����LED"
        #print(LEDLEIBIE)
        # replaceDict["XIANGMUMINGCHENG"] = itemDict["XIANGMUMINGCHENG"].get()

        # ��ʼ����
        r = replaceDict

        # ��Ԫ�ĳ��͸� mm
        r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]=[float(i) for i in detailDict[LEDMingCheng][LEDLEIBIE+'_CHICUN'].split('*')]
        # ��Ԫ�ļ��� �� ���� ����
        r["BANZI_CHANG"],r["BANZI_GAO"] = calc.calcBanZi(r["CHANG"],r["GAO"], chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        # ʵ�� ���͸� ��� m
        r["SHIJI_CHANG"],r["SHIJI_GAO"] = calc.calcChangGao(r["BANZI_CHANG"],r["BANZI_GAO"],chicun=[r["MOZU_CHICUN_CHANG"],r["MOZU_CHICUN_GAO"]])
        r["MIANJI"]=r["SHIJI_CHANG"]*r["SHIJI_GAO"]

        # ���� mm
        r["JIANJU"] = float(detailDict[LEDMingCheng][LEDLEIBIE+'_JIANJU'])

        # LED��Ԫ�ֱ���,������ģ���������
        r["MOZU_FENBIANLV_CHANG"] = round(r["MOZU_CHICUN_CHANG"]/r["JIANJU"])
        r["MOZU_FENBIANLV_GAO"] = round(r["MOZU_CHICUN_GAO"]/r["JIANJU"])

        # ��������ֱ���
        r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"] = calc.calcFengBianLv(r["BANZI_CHANG"],r["BANZI_GAO"],[r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"]])

        # ��������
        r["PINGMUZONGGONGLV"] = r["MIANJI"] * detailDict[LEDMingCheng][LEDLEIBIE+"_GONGLV"]

        # ������տ�
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

        # ���㴦���� ���� ���Ϳ�
        if itemDict['SHOUFALEIBIE'].get() == 0:
            chuliqi = "LEDCHULIQI-LC"
        else:
            chuliqi = "LEDCHULIQI-LN"
        key = calc.calcChuLiQi(chuliqi,keyDict[chuliqi],detailDict,r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"])
        if key:  # ����������ܴ��Ķ�
            r["LEDCHULIQI_SHULIANG"] = 1
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["LEDCHULIQI_"+i] = detailDict[key][chuliqi+"_"+i]
        else:
            # keyΪ�գ����Ǵ����������ˣ���Ҫ���㷢�Ϳ�
            if itemDict['SHOUFALEIBIE'].get() == 0:
                fasongka = "FASONGKA-LC"
            elif itemDict['SHOUFALEIBIE'].get() == 1:
                fasongka = "FASONGKA-LN"
            #key, shuliang = calc.calcFaSongKa1(fasongka,keyDict[fasongka],detailDict,r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"])
            key,shuliang,r["FASONGKAZUWANG_LIST"] = calc.calcFaSongKa(fasongka,keyDict[fasongka],detailDict,[r["MOZU_FENBIANLV_CHANG"],r["MOZU_FENBIANLV_GAO"]],[r["BANZI_CHANG"],r["BANZI_GAO"]])    
            r["LEDCHULIQI_SHULIANG"] = shuliang
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["LEDCHULIQI_"+i] = detailDict[key][fasongka+"_"+i]
            # �л�Ϊ LED��ƴ�� ��sheet & ����ƴ��
            if LEDLEIBIE == "SHINEILED":
                sheetName = u"LED��ƴ��"
            elif LEDLEIBIE == "XIANGTILED":
                sheetName = u"LED���弰ƴ��"
            keyU, kou8_shuliang = calc.calcPingKong(shuliang)

            r["PINGKONG-SHURU_SHULIANG"] = 1
            r["PINGKONG-SHUCHU_SHULIANG"] = kou8_shuliang
            
            if keyU == "2U":
                key = u"��Ƶƴ�Ӵ��������(2U)"
                keyShuRu = u"4·HDMI����忨(2U)"
                keyShuChu = u"8·HDMI����忨(2U)"
            elif keyU == "3U":
                key = u"��Ƶƴ�Ӵ��������(3U)"
                keyShuRu = u"8·HDMI����忨(3U)"
                keyShuChu = u"8·HDMI����忨(3U)"
            for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
                r["PINGKONG_"+i] = detailDict[key]["PINGKONG_"+i]
                r["PINGKONG-SHURU_"+i] = detailDict[keyShuRu]["PINGKONG-SHURU_"+i]
                r["PINGKONG-SHUCHU_"+i] = detailDict[keyShuChu]["PINGKONG-SHUCHU_"+i]
        #print(r)
        
        # ���������
        key, r["PEIDIANXIANG_JIAGE"] = calc.calcPeiDianXiang(keyDict["PEIDIANXIANG"],detailDict,r["PINGMUZONGGONGLV"])
        # print(key,r["PEIDIANXIANG_JIAGE"],r["PINGMUZONGGONGLV"])
        if not key:
            key="PEIDIANXIANG_ELSE"
        for i in ["MINGCHENG","XINGHAO","CANSHU","DANJIA"]:
            r["PEIDIANXIANG_"+i] = detailDict[key]["PEIDIANXIANG_"+i]

        # ����LED��Դ����
        r["DIANYUAN_SHULIANG"] = calc.calcDianYuan(r["BANZI_CHANG"],r["BANZI_GAO"])

        # ������������
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
        
        # ��ʾ��������
        tmpList = []
        tmpList.append(u"ʵ�ʿ��: %s * %s = %s ƽ����"%(str(r["SHIJI_CHANG"]),str(r["SHIJI_GAO"]),str(r["MIANJI"])))
        tmpList.append(u"��Ԫ���: %s * %s "%(str(r["BANZI_CHANG"]),str(r["BANZI_GAO"])))
        tmpList.append(u"%s: %d * %d = %d ��"%(r["JIESHOUKA_MINGCHENG"],r["JIESHOUKA_CHANG"],r["JIESHOUKA_GAO"],r["JIESHOUKA_SHULIANG"]))
        tmpList.append(u"����������%s"%r["LEDCHULIQI_SHULIANG"])
        tmpList.append(u"�ֱܷ���: %d * %d = %d ����"%(r["FENBIANLV_CHANG"],r["FENBIANLV_GAO"],r["FENBIANLV_CHANG"]*r["FENBIANLV_GAO"]))
        tmpList.append(u"��Դ����: %d ��"%r["DIANYUAN_SHULIANG"])
        tmpList.append(u"���幦��: %d W"%r["PINGMUZONGGONGLV"])
        
        textItem.delete('0.0',tk.END)
        for n,i in enumerate(tmpList):
            textItem.insert('%d.0'%(n+1),i+'\n')

        # д��config.json
        with open('config.json','w') as f0:
            tmpDict = {}
            for key in itemDict:
                tmpDict[key]=itemDict[key].get()
                
            f0.write(json.dumps(tmpDict))
            
        global makingTimes
        # ����Excel�ļ�
        if makeExcel:
            newFileName = u'����'+itemDict['LED'].get().split(' ')[0]+u'LED����'+'('+str(r["SHIJI_CHANG"])+'x'+str(r["SHIJI_GAO"])+u')-'+datetime.now().strftime('%Y%m%d%S')+str(makingTimes)+'.xlsx'
            excel.copyExcel(originFileName, newFileName, r, sheetName=sheetName)
            # ������ʾ
            
            makingTimes += 1
            # ����״̬��ʾ
            label = tk.Label(mainframe,text=u'���ɵ�'+str(makingTimes)+u'��ok: '+newFileName).grid(row=len(showDict.keys())*2+2,column=0,columnspan=4,sticky=tk.W)#����
        
    
    # --------------------------- �ָ��� ----------------------------

    
    # ����
    mainframe = tk.Tk()
    mainframe.title(TITLE)

    # ��ȡexcel����
    keyDict,detailDict = excel.getFromExcel(dataFileName, worksheet='data')
    keyDict["LED"] = keyDict["SHINEILED"]+keyDict["HUWAILED"]+keyDict["XIANGTILED"] # ���������� ����LED�� "LED"���key
    
    # ��ʾDict
    showDict = {
                'CHANG':u'��(m)',
                'GAO':u'��(m)',
                'LED':u'LED����',
                'SHOUFALEIBIE':u'�շ�ϵͳ���'
                }

    # ��ȡconfig.json
    configDict={}
    try:
        with open('config.json','r') as f1:
            configDict=json.loads(f1.read())
    except:
        pass
    
    itemDict = {} # ����ֵ�
    for n,i in enumerate(showDict.keys()):
        label = tk.Label(mainframe,text=showDict[i]+' :').grid(row=n,sticky=tk.E) #����        
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
                content.current(0)  # ѡ��Ĭ�ϵ�0��
        if i == "SHOUFALEIBIE":
            v = tk.IntVar()
            if i in configDict:
                v.set(configDict[i])
            else:
                v.set(0)
            tk.Radiobutton(mainframe,text=u"����(������)",variable=v,value=0).grid(row=n, column=1)
            tk.Radiobutton(mainframe,text=u"����(ŵ��)",variable=v,value=1).grid(row=n, column=2)
            tk.Radiobutton(mainframe,text=u"���˴�",variable=v,value=2).grid(row=n, column=3)
            itemDict[i]=v

        # ��ʾ����ռ����3������
        if i != "SHOUFALEIBIE":
            itemDict[i] = content
            content.grid(row=n,column=1,columnspan=3)
    
    # ��ʾ�����͵����ť
    # label = tk.Label(mainframe,text='  0').grid(sticky=tk.W)
    tk.Button(mainframe,text=u'����',width=15,height=2,command=lambda:doReplace(False)).grid(row=n+1,column=2)
    tk.Button(mainframe,text=u'���ɴ����嵥',width=15,height=2,command=lambda:doReplace(True)).grid(row=n+1,column=3)

    # ��ʾ������
    textItem = tk.Text(mainframe,height=9,width=60)
    textItem.grid(columnspan=4,pady=5,padx=5)

    
    mainframe.mainloop()


if __name__=='__main__':
    main()
