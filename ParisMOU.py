#coding=utf8
import urllib2
import urllib
import urlparse
import re 
import time
import os
from multiprocessing import Pool
from multiprocessing.dummy import Pool as TPool
from datetime import *
import xlwt #xlwt3 xlrd
import os
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
bad_ips=[]
ips = []
ip_url = 'ip2.txt'
is_ip_ok = False
opener = None
list_filename = ''
year = 2015
month = 10
count = 0
total = 0
def LoadIP():
    global ips,ip_url
    fo = open(ip_url,'r')
    lines = fo.read().strip().split('\n')
    for l in lines:
        end = l.find('@')
        ip = l[:end]
        ips.append(ip)
    print 'LoadIP.... ok'

def GetOpener():
    import random
    global ips,ip_url,is_ip_ok,opener,total
    N = len(ips)
    if(N<20):
        print 'ip.txt need to update'
    try:
        while(1):
            is_ip_ok = False
            i = random.randint(0,len(ips)-1)
            ip = ips[i]
            #ip = '49.1.245.241:3128'
            print 'Connecting :',ip
            try:
                proxy_handler = urllib2.ProxyHandler({'http':ip})
                url = "https://www.parismou.org"
                opener = urllib2.build_opener(urllib2.HTTPHandler,proxy_handler)
                response = opener.open(url,timeout=5)
                text = str(response.read())
                okay = text.find("hoppinger")
                if(okay!=-1):
                    print('okay~~~~~~~~~~~Current IP: ' + ip)
                    is_ip_ok = True
                    return opener                
            except Exception as e:
                print 'Error',e     
                continue    
    except Exception as e:
        print 'Error',e
        is_ip_ok=False
        return None



def get(year,month):

    #year = 14 month =11 01 02..
    
    #proxy_auth_handler = urllib.request.
    url = "https://portal.emsa.europa.eu/portlet-public-site-inspection/dwr/call/plaincall/PublicInspectionRemoteServices.searchInspections.dwr"
    values={}

    values={'callCount':'1','windowName':'DWR-FDCE2F19201BB65EBBCAE725C78CA6A8','c0-scriptName':'PublicInspectionRemoteServices','c0-methodName':'searchInspections','c0-id':'0','c0-e1':'string:','c0-e2':'string:','c0-e3':'Array:[]','c0-e4':'Array:[]','c0-e5':'string:','c0-e6':'string:','c0-e7':'string:','c0-e8':'string:','c0-e9':'string:','c0-e10':'string:','c0-e11':'Array:[]','c0-e12':'Array:[]','c0-e13':'string:','c0-e14':'string:','c0-e15':'boolean:true','c0-e16':'Array:[]','c0-e17':'Array:[]','c0-e18':'Array:[]','c0-e19':'Array:[]','c0-e20':'string:','c0-e21':'string:','c0-e22':'Array:[]','c0-e23':'string:','c0-e24':'string:','c0-param0':'Object_Object:{shipImoNumber:reference:c0-e1, shipName:reference:c0-e2, flags:reference:c0-e3, shipTypes:reference:c0-e4, grossTonnageMin:reference:c0-e5, grossTonnageMax:reference:c0-e6, ageMin:reference:c0-e7, ageMax:reference:c0-e8, ismCompanyImoNumber:reference:c0-e9, ismCompanyName:reference:c0-e10, classificationSocieties:reference:c0-e11, recognizedOrganizations:reference:c0-e12, periodMin:reference:c0-e13, periodMax:reference:c0-e14, pscInspectionRegime:reference:c0-e15, inspectionTypes:reference:c0-e16, portStates:reference:c0-e17, ports:reference:c0-e18, inspectionResults:reference:c0-e19, deficienciesNumberMin:reference:c0-e20, deficienciesNumberMax:reference:c0-e21, deficiencyAreas:reference:c0-e22, detentionDurationMin:reference:c0-e23, detentionDurationMax:reference:c0-e24}','c0-e25':'number:2','c0-e26':'number:20','c0-e27':'null:null','c0-e28':'string:ASCENDING','c0-param1':'Object_Object:{currentPage:reference:c0-e25, numberOfRows:reference:c0-e26, sortingParameter:reference:c0-e27, sortingDirection:reference:c0-e28}','c0-param2':'boolean:false','batchId':'16','page':'/web/thetis/inspections','httpSessionId':'','scriptSessionId':'B2A156858D487302ADBD61E535005975'}
    values['c0-e25']='number:%d'%(0)
    #values['c0-e26']='number:%d'%(1000)
    values['c0-e13']='string:01%2F'+month+'%2F'+year
    values['c0-e14']='string:28%2F'+month+'%2F'+year

    postdata = urllib.urlencode(values).encode(encoding='UTF8')


    opener = None
    while opener==None:
        opener = GetOpener()
    text=''
    while(text==''):
        try:
            response = opener.open(url,postdata)
            text = str(response.read())
        except Exception,e:
            print e
            opener = GetOpener()
    print'ok!',datetime.now()
    return text 
#get()
def GetAllInspId(y,m):
    global list_filename,year,month
    year = y
    month = m
    res = ''
    t = get(str(y),str(m))
    nex = 0
    f = open('Paris_%s_%s.txt'%(y,m),'w')
    f.write(t)
    while(1):
        nex = t.find('false},id:')
        if(nex==-1):
            break
        nex = nex+len('false},id:')
        #end = text[nex+len('false},id:'):].find(',')
        #idlist.append(t[nex:nex+9])
        res+=t[nex:nex+9]+'\n'
        t = t[nex:]
    f = open('Paris_%s_%s_list.txt'%(y,m),'w')
    f.write(res)
    list_filename = 'Paris_%s_%s_list.txt'%(y,m)
    print("GetAllInspId Done!")


def GetDetailInfoByInspId(InspId):
    InspId = str(InspId)
    global count,opener,is_ip_ok
    if not os.path.exists('ParisDetailInfo'):  #
        os.mkdir('ParisDetailInfo')
    name = "ParisDetailInfo/paris_detail_%s.txt"%(str(InspId).strip())

    if(os.path.exists(name) and os.path.getsize(name)>100):
        print(name,'exists.')
        return True

    global isIpOKay,opener
    #InspId = "EED9EC6E-A397-4E29-9F71-0268BABEF797"
    url = "https://portal.emsa.europa.eu/portlet-public-site-inspection/dwr/call/plaincall/PublicInspectionRemoteServices.getInspectionDetail.dwr"
    values={}
    values['callCount']='1'
    values['windowName']='DWR-616AC023BBA89903DEEDF8965319C573'
    values['c0-scriptName']='PublicInspectionRemoteServices'
    values['c0-methodName']='getInspectionDetail'
    values['c0-id']='0'
    values['c0-param0']='number:'+InspId
    values['batchId']='10'
    values['page']='/widget/web/thetis/inspections/-/publicSiteInspection_WAR_portletpublicsiteinspection'
    values['httpSessionId']=''
    values['scriptSessionId']='B2A156858D487302ADBD61E535005975'
    postdata = urllib.urlencode(values).encode(encoding='UTF8')

    text = ''
    while text=='':
        try:
            if(not is_ip_ok):
                opener = GetOpener()
            while(opener==None):
                opener = GetOpener()
            response = opener.open(url,postdata,timeout=5)
            text = str(response.read())
            if(len(text)<200):
                is_ip_ok = False
                print text
                text = '' 
                continue
            else:
                break
        except Exception as e:
            is_ip_ok=False
            print e
            continue
    f =  open(name,"w")
    f.write(text)
    f.close()
    count+=1
    print count,'/',total,InspId,"is ok!", datetime.now()

def GetDetailInfo():
    global count,total
    count = 0
    InspID = set()
    f = open(list_filename,'r')
    for l in f.xreadlines():
        InspID.add(l)
    f.close()
    total = len(InspID)
    print 'Total:',total
    p = TPool(10)
    p.map(GetDetailInfoByInspId,InspID)
    p.close()
    p.join() 

    print "All detail pages are done ..."

#LoadIP()
#GetAllInspId(2015,10)
#GetDetailInfo()
