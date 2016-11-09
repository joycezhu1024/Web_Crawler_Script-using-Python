#coding=utf-8
__author__ = 'Lin'
import os
import json
import xlwt

db = xlwt.Workbook()
sheet = db.add_sheet('sheet1')
title = ['id','shipIMOnumber','shipName','flag','age','shipType','grossTonnnage','keelDate',
         'companyIMOnumber','companyName','companyCountry',
         'issuingAuthority','issueDate','expiryDate',
        'inspectionType','inspectionPlace','firstVisitDate','finalVisitDate','numberDeficiencies','numberDeficienciesDetainable',
        'inspectionAreas','inspectionOperationalControls',
         'detentionReason','detentionDate','detentionEndDate','detentionDuration','nextPort',
         'validated','refusalReason','isActive'
         ]
for i,t in enumerate(title):
    sheet.write(0,i,t)


tot=1
def deleteComma(text):
    okay = False
    res = ''
    for c in text:
        if(c=='"'):
            okay= not okay
        elif((c in {',',':','\''} )and okay):
            c='.'
        res+=c
    return res
ss={}
ss["None"]='None'
#process()
def toJson(iden):
    global ss
    with open("ParisDetailInfo/paris_detail_%d.txt"%(iden),"r") as f:
        whole = f.read()
    #delete comments
    head = whole[:161]
    #print(head)
    #load ss
    for i in range(0,40):
        cur = whole.find('s%d.name'%(i))
        if(cur != -1):
            cur = whole.find('"',cur)
            start = cur+1
            end = whole.find('"',cur+1)
            ss['s%d'%(i)] = whole[start:end]
        cur = whole.find('s%d.description'%(i))
        if(cur != -1):
            cur = whole.find('"',cur)
            start = cur+1
            end = whole.find('"',cur+1)
            ss['s%d'%(i)] = whole[start:end]
        cur = whole.find('s%d.areaDescription'%(i))
        if(cur != -1):
            cur = whole.find('"',cur)
            start = cur+1
            end = whole.find('"',cur+1)
            ss['s%d'%(i)] = whole[start:end]
        cur = whole.find('s%d.subAreaDescription'%(i))
        if(cur != -1):
            cur = whole.find('"',cur)
            start = cur+1
            end = whole.find('"',cur+1)
            ss['s%d'%(i)] +='--'+ whole[start:end]
    #print(ss)
    text = whole[161:]
    text = deleteComma(text)
    text = text.replace('\\r','')
    text = text.replace('\\','')
    text = text.replace('\\n','')
    text = text.replace("'",'"')#单引号变双引号？
    text = text.replace("\"s",'\'s')
    #text = text.replace("ers\"",'ers\'')#tankers不行

    for i in range(0,40):
        text = text.replace('s%d.php'%(i),'')
        text = text.replace('egs%d'%(i),'')
        text = text.replace('s%d,'%(i),'\"s%d\",'%(i))
        text = text.replace('s%d}'%(i),'\"s%d\"}'%(i))
        text = text.replace('s%d]'%(i),'\"s%d\"]'%(i))

    sep = text.find("deficiencies:")
    text = text[sep-1:]

    js = ''
    for i,c in enumerate(text):
        if(c in {'{','[',','}):
            if(text[i+1] not in {'0','{','[',','}):
                js += c+ '"'
            else:
                js += c
        elif(c in {':'}):
            if('http' in text[i-5:i]):
                js += c
            else:
                js += '"'+c
        elif(c ==']' and text[i-1]!='"' and text[i-1]!='}') :
            js += '"' + c
        else:
            js+=c
    return js[:-4]




def saveOne(iden):
    global ss
    try:
        t = toJson(iden)

        #print(t)
        j = {}
        j = json.loads(t)
    except Exception as e:
        print(e)
        return False
    global tot
    #############
    isNullforSHIPissuingAuthority = False
    SHIPissuingAuthority = j["inspectionClassCertificates"][0]
    if(type(SHIPissuingAuthority)!=dict):
        isNullforSHIPissuingAuthority=True
    for i,t in enumerate(title):
        content = ''
        try:
            #print(t)
            if(t=='id'):
                content=iden
            elif(t=='shipIMOnumber'):
                content = j["inspectionShip"]["imoNumber"]
            elif(t=='shipName'):
                content = j["inspectionShip"]["name"]
            elif(t=='age'):
                content = j["inspectionShip"]["age"]
            elif(t=='keelDate'):
                content = j["inspectionShip"]["keelDate"]
            elif(t=='flag'):
                content = j["inspectionShip"]["flag"]["countryDescription"]
            elif(t=='shipType'):
                content = j["inspectionShip"]["shipType"]["description"]
            elif(t=='grossTonnnage'):
                content = j["inspectionShip"]["grossTonnage"]

            elif(t=='companyIMOnumber'):
                content = j["ismCompany"]["imoNumber"]
            elif(t=='companyName'):
                content = j["ismCompany"]["name"]
            elif(t=='companyCountry'):
                content = j["inspectionShip"]["flag"]["countryDescription"]

            elif(t=='issuingAuthority' and not isNullforSHIPissuingAuthority):
                SHIPissuingAuthority =j["inspectionClassCertificates"][0]["issuingAuthority"]
                if(type(SHIPissuingAuthority)==dict):
                     SHIPissuingAuthority=j["inspectionClassCertificates"][0]["issuingAuthority"]['name']
                else:
                    SHIPissuingAuthority=ss[j["inspectionClassCertificates"][0]["issuingAuthority"]]
                content=SHIPissuingAuthority
            elif(t=='issueDate' and not isNullforSHIPissuingAuthority):
                content=j["inspectionClassCertificates"][0]["issueDate"]
            elif(t=='expiryDate' and not isNullforSHIPissuingAuthority):
                content=j["inspectionClassCertificates"][0]["expiryDate"]
            elif(t=='inspectionType'):
                content=j['inspectionType']['description']
            elif(t=='inspectionPlace'):
                content=j['port']['name']
            elif(t=='firstVisitDate'):
                content=j['firstVisitDate']
            elif(t=='finalVisitDate'):
                content=j['finalVisitDate']
            elif(t=='numberDeficiencies'):
                content=j['numberDeficiencies']
            elif(t=='numberDeficienciesDetainable'):
                content=j['numberDeficienciesDetainable']
            elif(t=='inspectionAreas'):
                iatext = '{'
                inspectionAreas = j['inspectionShipAreas']
                if(type(inspectionAreas)!=list):
                    continue
                for ia in inspectionAreas:
                    if(type(ia)==str):
                        continue
                    iatext += ia['description']+','
                iatext = iatext[:-1]
                iatext +='}'
                content = iatext
            elif(t=='inspectionOperationalControls'):
                octext = '{'
                inspectionOperationalControls = j['inspectionOperationalControls']
                if(type(inspectionOperationalControls)!=list):
                    continue
                for oc in inspectionOperationalControls:
                    if(type(oc)==dict):
                        octext += oc['description']+','
                octext = octext[:-1]
                octext +='}'
                content = octext
                if(content=='}'):
                    content='{}'
            else:
                detention = j['detention']
                if(detention!=None):

                    if(t=='detentionReason'):
                        content=detention['additionalInformation']==None and 'None' or detention['additionalInformation']
                    elif(t=='detentionDate'):
                        content=detention['detentionDate']
                    elif(t=='detentionEndDate'):
                        content=detention['detentionEndDate']
                    elif(t=='detentionDuration'):
                        content=detention['duration']
                    elif(t=='nextPort'):
                        content = j['inspectionRepairOrNextPort']['repairOrNextCountry']['description']
                    elif(t=='validated'):
                            content=detention['ban']['validated']
                    elif(t=='refusalReason'):
                        content=detention['ban']['banReason']['description']
                    elif(t=='isActive'):
                        content=detention['ban']['banOrderStatus']['active']
        except Exception as e:
            pass
        sheet.write(tot,i,content)
    tot+=1
    return True 
# for i in range(60000,110995):
#     if(not saveOne(i)):
#         print((i,'failed'))
# db.save("dbParis_BasicInfo2.xls")