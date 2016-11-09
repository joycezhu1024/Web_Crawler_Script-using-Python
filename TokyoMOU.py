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

bad_ips = []
total = 0
startDate = ''
endDate = ''
InspID = set()
ips = []
ip_url = 'ip.txt'
is_ip_ok = False
count = 0
opener = None
db = xlwt.Workbook()
sheet = db.add_sheet('sheet1')
insp_list_filename = ''
def LoadIP():
    global ips,ip_url
    fo = open(ip_url,'r')
    lines = fo.read().strip().split('\n')
    for l in lines:
        end = l.find('@')
        ip = l[:end]
        ips.append(ip)
    print 'LoadIP.... ok Connecting'
def GetOpener():
    import random
    global ips,ip_url,is_ip_ok,opener,total
    N = len(ips)
    if(N<20):
        print 'ip.txt need to update'
    try:
        while(1):
            is_ip_ok = False
            ip = ips[random.randint(0,len(ips)-1)]
            while ip in bad_ips:
                ip = ips[random.randint(0,len(ips)-1)]
            #print 'Connecting :',ip
            try:
                proxy_handler = urllib2.ProxyHandler({'http':ip})
                url = "http://212.45.16.136/isss/public_apcis.php?Action=getSearchForm"
                opener = urllib2.build_opener(urllib2.HTTPHandler,proxy_handler)
                response = opener.open(url,timeout=10)
                text = str(response.read())
                okay = text.find("hours")
                if(okay==-1):
                    print('okay~~~~~~~~~~~Current IP: ' + ip)
                    is_ip_ok = True
                    return opener
            except Exception as e:
                #print 'Error',e     
                bad_ips.append(ip)
                continue    
    except Exception as e:
        #print 'Error',e
        is_ip_ok=False
        return None


def Prepare(from_,to_):
    global total,startDate,endDate,is_ip_ok,opener,total,insp_list_filename
    startDate = from_
    endDate = to_
    insp_list_filename = 'insp_from%s_to%s.txt'%(startDate,endDate)
    url = 'http://212.45.16.136/isss/public_apcis.php?Action=searchInsp&from='+startDate+'&to='+endDate+'&SortOrder1=Ascending&SortOrder2=Ascending&MOU=TMOU&Skip=0'
    
    html = ''
    while html == '':
        if(not is_ip_ok): 
            opener = GetOpener()
        while opener==None:
            opener = GetOpener()
        try:
            html = opener.open(url).read()
        except Exception as e:
            #print e
            is_ip_ok = False
    totalStart = html.find('Records found: ') + len('Records found: ')
    totalEnd = html.find('<br>',totalStart)
    total = int(html[totalStart:totalEnd].strip())


def GetInspBySkip(i):
    global opener,InspID,is_ip_ok,total
    skip = i*25
    
    url = 'http://212.45.16.136/isss/public_apcis.php?Action=searchInsp&from='+startDate+'&to='+endDate+'&SortOrder1=Ascending&SortOrder2=Ascending&MOU=TMOU&Skip='+str(skip)
    while(1):
        if(not is_ip_ok):   
            opener = GetOpener()
        while opener==None:
            opener = GetOpener()
        text = ''
        try:
            response = opener.open(url,timeout=10)
            text = response.read()
        except Exception as e:
            #print e 
            is_ip_ok = False
            continue
        ids = re.findall(r'InspID=[^\s]*&',text,re.I)
        for inspid in ids:
            InspID.add(inspid[7:-1])
        if(len(ids)>0):
            print 'Downloading List: '+str(skip)+'/'+str(total),'ok'
            break
        else:
            print 'Downloading List: '+str(skip)+'/'+str(total),'failed'

def GetAllInspId():
    global opener,InspID,is_ip_ok,total
    p = TPool(10)
    p.map(GetInspBySkip,range(0,total/25+1))
    p.close()
    p.join() 

    print 'Okay..Begin to save InspID...'
    f = open(insp_list_filename,'w')
    for i in InspID:
        f.write(i+'\r\n')
    f.close()
    print len(InspID),'inspids have been saved ~'


def GetDetailInfoByInspId(InspId):
    global opener,count,is_ip_ok,total
    if not os.path.exists('detailInfo'):  #如果文件夹不存在则新建
        os.mkdir('detailInfo')
    name = "detailInfo/%s.html"%(InspId.strip())
    if(os.path.exists(name) and os.path.getsize(name)>100):
        print count,'/',total,InspId,"is ok!", datetime.now()
        return True
    global is_ip_ok,opener
    #InspId = "EED9EC6E-A397-4E29-9F71-0268BABEF797"
    url = "http://212.45.16.136/isss/public_apcis.php?Action=searchInspByID&InspID=%s&Skip=0"%(InspId)
    
    text = ''
    while text=='':
        try:
            if(not is_ip_ok):
                opener = GetOpener()
            while(opener==None):
                opener = GetOpener()
            response = opener.open(url,timeout=5)
            text = str(response.read())
            if(len(text)<3000):
                is_ip_ok = False
                text = ''
                continue
            else:
                break
        except Exception as e:
            is_ip_ok=False
            #print e
            continue
    f =  open(name,"w")
    f.write(text)
    f.close()
    count+=1
    print count,'/',total,InspId,"is ok!", datetime.now()

def GetDetailInfo():
    global InspID,count,total
    count = 0
    if(len(InspID)!=total):
        InspID = set()
        f = open(insp_list_filename,'r')
        for l in f.xreadlines():
            InspID.add(l)
        f.close()
    
    p = TPool(10)
    p.map(GetDetailInfoByInspId,InspID)
    p.close()
    p.join() 

    print "All detail pages are done ..."

def getText_1(html,name):
    seg ='\r\n\t\t'
    tag = '<td align="right"><font >'+name+'</font></td>'+seg+'<td align="left"><font ><b><u>&nbsp;'
    s = html.find(tag)+len(tag)
    e = html.find('&nbsp;</u></b></font></td>',s)
    return html[s:e].strip()

def getText_2(html,name):
    seg ='\r\n\t\t'
    tag = '<td align="right"><font >'+name+'</font></td>'+seg+'<td align="left"colspan="5"><font ><b><u>&nbsp;'    
    s = html.find(tag)+len(tag)
    e = html.find('&nbsp;</u></b></font></td>',s)
    return html[s:e].strip()
def getText_3(html,name):
    seg ='\r\n\t\t'
    tag = '<td align="right" valign="top"><font >'+name+'</font></td>'+seg+'<td align="left"colspan="5"><font ><b><u>&nbsp;'    
    s = html.find(tag)+len(tag)
    e = html.find('&nbsp;</u></b></font></td>',s)
    return html[s:e].strip()

def getDateKeelLaid(html):
    tag1 = '<td align="right"><font >date keel laid'
    tag2 = '</font></td>'
    s = html.find(tag2,html.find(tag1))+67
    e = html.find('&nbsp;</u></b></font></td></tr>',s) 
    return html[s:e].strip()
def getDeficenciesNumber(html):
    if('(total / new)' in html):
        return '0'
    tag='number of deficiencies'
    tag2='<b><u>&nbsp;'
    s = html.find(tag2,html.find(tag))+len(tag2)
    e = html.find('</u></b>',s)
    text = html[s:e].strip()
    dig = ''
    for c in text:
        if(c in '0123456789'):
            dig+=c
    return dig
def getDeficencies(number,html):
    s = html.find('including')
    tag = '<td colspan="3" align="right">&nbsp;'
    text = ''
    for i in range(number):
        s = html.find(tag,s)+len(tag)
        if(s<len(tag)):
            break
        e = html.find('</td>',s)

        text += html[s:e].strip()
        s = html.find('<td colspan="2">&nbsp;<b>',e)+len('<td colspan="2">&nbsp;<b>')
        e = html.find('</b></td>',s)
        text += html[s:e]
        text += '; '
    return text
def ProcessDetailByInspID(insp,row=1):
    fo = open('detailInfo/'+insp+'.html','r')
    html = fo.read()
    fo.close() 
    val = range(18)
    val[0]  = getText_1(html,'name of ship:')
    val[1]  = getText_1(html,'call sign:')
    val[2]  = getText_1(html,'IMO number:')
    val[3]  = getDateKeelLaid(html)
    val[4]  = getText_2(html,' gross tonnage:')
    val[5]  = getText_2(html,'deadweight:')
    val[6]  = getText_2(html,'type of ship:')
    val[7]  = getText_2(html,'flag of ship:')
    val[8]  = getText_2(html,'classification society:')
    val[9]  = getText_3(html,'company IMO No:')
    val[10] = getText_3(html,'particulars of company:')
    val[11] = getText_2(html,'name of reporting authority:')
    val[12] = getText_2(html,'place of inspection:')
    val[13] = getText_2(html,'date of inspection:')
    val[14] = getText_2(html,'deficiencies:')
    val[15] = getText_2(html,'ship detained:')
    val[16] = getDeficenciesNumber(html)
    all_deficiencies = ''
    if(val[16]!='0'):
        all_deficiencies = getDeficencies(int(val[16]),html)
    else:
        val[14]='no'
    val[17]= all_deficiencies
    sheet.write(row,0,insp)
    for i in range(18):
        sheet.write(row,i+1,val[i])
    print 'Process ',row,'/',total,'is okay'
def ProcessDetail():
    global InspID,startDate,endDate
    InspID = set()
    f = open(insp_list_filename,'r')
    for l in f.xreadlines():
        InspID.add(l)
    f.close()
    vals = '''
     
    name_of_ship               
    call_sign                  
    imo_number                 
    date_keel_laid             
    gross_tonnage              
    deadweight                 
    type_of_ship               
    flag_of_ship               
    classification_society     
    company_imo_no             
    particulars_of_company     
    name_of_reporting_authority
    place_of_inspection        
    date_of_inspection         
    deficiencies               
    ship_detained              
    number_of_deficiencies
    all_deficiencies
    '''
    sheet.write(0,0,'InspID')
    for i in range(0,18):
        sheet.write(0,i+1,vals.split()[i])
        #print vals.split()[i]
    row = 1
    for insp in InspID:
        try:
            ProcessDetailByInspID(insp.strip(),row)
            row += 1
        except Exception,e:
            #print insp.strip(),e
            pass
    filename = 'database/Tokyo_FROM'+startDate+'_TO'+endDate+'.xls'
    db.save(filename)
    print 'Excel has been created as '+filename
# LoadIP()
# Prepare('23.10.2015','24.10.2015')
# print total
# GetAllInspId()
# GetDetailInfo()
#ProcessDetail()