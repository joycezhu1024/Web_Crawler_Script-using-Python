import string,urllib2
import os
import xlwt
from string import*

def Getinfo(start_date, end_date,page):



    ProjectName=["Ship Name","IMO Number","Call Sigh","MMSI Number","Gross Tonnage","Deadweight","Flag","Date Keel Laid",
                 "Type of Ship","Classification society","Place of Inspection","Date of Inspection"
                 ,"Type of Inspection","IMO Company Number","Particulars of a Company ","Detained","Inspecting Authority","Deficiencies"]
    
    IMO=[" "]
    ShipName=[" "]
    Call=[" "]
    MMSI=[" "]
    Gross=[" "]
    Deadw=[" "]
    Flag=[" "]
    DKL=[" "]
    Tship=[" "]
    Class=[" "]
    Pinspection=[" "]
    Dinspection=[" "]
    Tinspection=[" "]
    IMOcom=[" "]
    Part=[" "]
    IA=[" "]
    Detained=[" "]
    Dressn=[" "]
    Defi=[[" "]]
    Defic=[[" "]]
    Defians=[[" "]]

    Defi=Defi+[[" "]]
    Defic=Defic+[[" "]]
    Defians=Defians+[[" "]]


    tot=0

    workbook = xlwt.Workbook() 
    sheet = workbook.add_sheet("Information")
    for i in range(0,18):
        sheet.write(0,i,ProjectName[i])

    for i in range(1,page+1):

     fs=start_date+end_date+string.zfill(i,6)+".txt"
     f=open("India_htmlbase/"+fs,"r")
     s=f.read()
     f.close()

     END=len(s)
     START=1
     while START!=-1:
        
        tot=tot+1

        col=0
        
        sub="Ship Name"
       # print sub
        start=s.find(sub,START,END)+52
        sub="</"
        end=s.find(sub,start,END)
        ShipName=ShipName+[s[start:end]]
        
        sheet.write(tot,col,ShipName[tot])
        col=col+1
      #  Output=open('Ship Name.txt','a')
      #  print >>Output,ShipName[tot]
      #  Output.close()
        
        START=end
        sub="IMO Number"
       # print sub
        start=s.find(sub,START,END)+31
        sub="</"
        end=s.find(sub,start,END)
        IMO=IMO+[s[start:end]]

        sheet.write(tot,col,IMO[tot])
        col=col+1
      #  Output=open('IMO Number.txt','a')
      #  print >>Output,IMO[tot]
      #  Output.close()
        
        START=end
        sub="Call Sign"
       # print sub
        start=s.find(sub,START,END)+30
        sub=sub="</"
        end=s.find(sub,start,END)
        Call=Call+[s[start:end]]

        sheet.write(tot,col,Call[tot])
        col=col+1
       # Output=open('Call Sign.txt','a')
       # print >>Output,Call[tot]
       # Output.close()
        
        START=end
        sub="MMSI Number"
       # print sub
        start=s.find(sub,START,END)+31
        sub=sub="</"
        end=s.find(sub,start,END)
        MMSI=MMSI+[s[start:end]]

        sheet.write(tot,col,MMSI[tot])
        col=col+1
     #   Output=open('MMSI Number.txt','a')
     #   print >>Output,MMSI[tot]
     #   Output.close()
        
        START=end
        sub="Gross Tonnage"
      #  print sub
        start=s.find(sub,START,END)+34
        sub=sub="</"
        end=s.find(sub,start,END)
        Gross=Gross+[s[start:end]]

        sheet.write(tot,col,Gross[tot])
        col=col+1
      #  Output=open('Gross Tonnage.txt','a')
       # print >>Output,Gross[tot]
      #  Output.close()
        
        START=end
        sub="Deadweight"
       # print sub
        start=s.find(sub,START,END)+31
        sub=sub="</"
        end=s.find(sub,start,END)
        Deadw=Deadw+[s[start:end]]

        sheet.write(tot,col,Deadw[tot])
        col=col+1
     #   Output=open('Deadweight.txt','a')
     #   print >>Output,Deadw[tot]
      #  Output.close()
        
        START=end
        sub="Flag"
     #   print sub
        start=s.find(sub,START,END)+25
        sub=sub="</"
        end=s.find(sub,start,END)
        Flag=Flag+[s[start:end]]

        sheet.write(tot,col,Flag[tot])
        col=col+1
     #   Output=open('Flag.txt','a')
     #   print >>Output,Flag[tot]
     #   Output.close()
        
        START=end
        sub="Date Keel Laid"
      #  print sub
        start=s.find(sub,START,END)+35
        sub=sub="</"
        end=s.find(sub,start,END)
        DKL=DKL+[s[start:end]]

        sheet.write(tot,col,DKL[tot])
        col=col+1
     #   Output=open('Date Keel Laid.txt','a')
     #   print >>Output,DKL[tot]
     #   Output.close()
        
        START=end
        sub="Type of Ship"
      #  print sub
        start=s.find(sub,START,END)+33
        sub=sub="</"
        end=s.find(sub,start,END)
        Tship=Tship+[s[start:end]]

        sheet.write(tot,col,Tship[tot])
        col=col+1
        
     #   Output=open('Type of Ship.txt','a')
     #   print >>Output,Tship[tot]
     #   Output.close()
        
        START=end
        sub="Classification Society"
     #   print sub
        start=s.find(sub,START,END)+43
        sub=sub="</"
        end=s.find(sub,start,END)
        Class=Class+[s[start:end]]

        sheet.write(tot,col,Class[tot])
        col=col+1
     #   Output=open('Classification Society.txt','a')
     #   print >>Output,Class[tot]
     #   Output.close()
        
        START=end
        sub="Place of Inspection"
     #   print sub
        start=s.find(sub,START,END)+40
        sub=sub="</"
        end=s.find(sub,start,END)
        Pinspection=Pinspection+[s[start:end]]

        sheet.write(tot,col,Pinspection[tot])
        col=col+1
      #  Output=open('Place of Inspection.txt','a')
      #  print >>Output,Pinspection[tot]
      #  Output.close()
        
        START=end
        sub="Date of Inspection"
     #   print sub
        start=s.find(sub,START,END)+39
        sub=sub="</"
        end=s.find(sub,start,END)
        Dinspection=Dinspection+[s[start:end]]

        sheet.write(tot,col,Dinspection[tot])
        col=col+1
      #  Output=open('Date of Inspection.txt','a')
     #   print >>Output,Dinspection[tot]
      #  Output.close()
        
      
        START=end
        sub="Type of Inspection"
      #  print sub
        start=s.find(sub,START,END)+39
        sub=sub="</"
        end=s.find(sub,start,END)
        Tinspection=Tinspection+[s[start:end]]

        sheet.write(tot,col,Tinspection[tot])
        col=col+1
      #  Output=open('Type of Inspection.txt','a')
       # print >>Output,Tinspection[tot]
      #  Output.close()
        
        START=end
        sub="IMO Company Number"
     #   print sub
        start=s.find(sub,START,END)+39
        sub=sub="</"
        end=s.find(sub,start,END)
        IMOcom=IMOcom+[s[start:end]]

        sheet.write(tot,col,IMOcom[tot])
        col=col+1
       # Output=open('IMO Company Number.txt','a')
        #print >>Output,IMOcom[tot]
       # Output.close()
        
        START=end
        sub="Particulars of a Company"
      #  print sub
        start=s.find(sub,START,END)+45
        sub=sub="</"
        end=s.find(sub,start,END)
        Part=Part+[s[start:end]]

        sheet.write(tot,col,Part[tot])
        col=col+1
       # Output=open('Particulars of a Company.txt','a')
       # print >>Output,Part[tot]
      #  Output.close()
        
        START=end
        sub="Detained"
      #  print sub
        start=s.find(sub,START,END)+29
        sub=sub="</"
        end=s.find(sub,start,END)
        Detained=Detained+[s[start:end]]

        sheet.write(tot,col,Detained[tot])
        col=col+1
      #  Output=open('Detained.txt','a')
      #  print >>Output,Detained[tot]
      #  Output.close()
        
        START=end
        sub="Inspecting Authority"
      #  print sub
        start=s.find(sub,START,END)+90
        sub=sub="</"
        end=s.find(sub,start,END)
        IA=IA+[s[start:end]]

        sheet.write(tot,col,IA[tot])
        col=col+1
       # Output=open('Inspecting Authority.txt','a')
       # print >>Output,IA[tot]
       # Output.close()
        
        START=end
        sub="ShowDeficiencies"
        start=s.find(sub,START,START+500)
      #  print tot
        if start!=-1:
            start=start+28
            sub="\""
            end=s.find(sub,start,start+20)
         #   print start
          #  print end
         #   print s[start:end]
         #   print Dressn
            Dressn=Dressn+[s[start:end]]
            temps=s[start:end]
        #    print Dressn[tot]
            url="http://www.iomou.org/php/ShowDeficiencies.php?InspNo="+temps
            sName=temps+"from"+fs
            f=open("India_htmlbase/"+sName,"w+")
            m=urllib2.urlopen(url).read()
            f.write(m)
            f.close()
            print "downloading deficiencies "+"save as "+sName+"\n"
            if os.path.exists("India_htmlbase/"+sName)==1:
                f=open("India_htmlbase/"+sName,"r")
                m=f.read()
                f.close()
                FLAG=True
            else:
                m=' '   
                FLAG=False
            sub="Deficiency Rectified"
            start=m.find(sub)+76
            send=len(m)
          #  print "Show Defic detail-----------"
            while start!=-1 and FLAG :
              if(start>=send): break
              
              
              sub="</TD"
              end=m.find(sub,start)
              if m[start]=='>':
                  start=start+1
            #  print "++++++"
            #  print m[start:end]
            #  print tot
            #  print Defic[tot]
            #  print "------"
         #     print m[start:end]
              Defic[tot]=Defic[tot]+[m[start:end]]
              
            #  print Defic
            
              if(end==-1): break
              start=end+9
              end=m.find(sub,start)
              if(start>=send): break
              if(end==-1): break
              Defi[tot]=Defi[tot]+[m[start:end]]
              sheet.write(tot,col,m[start:end])
              col=col+1
          #    print m[start:end]
              
              start=end+24
              end=m.find(sub,start)
              if(start>=send): break
              if(end==-1): break
              Defians[tot]=Defians[tot]+[m[start:end]]
          #    print m[start:end]
              
              sub="\"center\""
              start=m.find(sub,end)
              
              if(start>=send): break
              
              if start!=-1:
                  start=start+19
            
       # for q in range(1,len(Defi[tot])):
       #     text += str(tot)+' '+Defi[tot][q]+' $'
       # text += '\n'

       # Output=open('Defi.txt','a')
       # Output.write(text)
       # Output.close()
        
        Dressn=Dressn+[" "]
        Defi=Defi+[[" "]]
        Defic=Defic+[[" "]]
        Defians=Defians+[[" "]]
        

        sub="Ship Name"
        START=s.find(sub,START,END)

    
    workbook.save("database/"+"India_MOU_"+start_date+"_to_"+end_date+".xls")   
   # print tot


def Getlist(start_date,end_date):
    
    if not os.path.exists("India_htmlbase"):
        os.mkdir("India_htmlbase")
        
    s1="http://www.iomou.org/php/InspData.php?lstLimit=30&StartOffset="
    s2="&FindInspAction=Find&txtStartDate="
    s3="&txtEndDate="
    s4="&opt_txtISC=I&txtISC=&opt_lstFCS=F&lstFCS=&lstAuth=000&chkDet=All&InspType=All&SortOrder=NoSort&AscDsc=Desc"


   # start_date=start str(raw_input("start date is:\n"))
   # end_date=end str(raw_input("end date is:\n"))

    i=1
    endflag=False
    while(endflag==False):
        
        url=s1+str(i)+s2+start_date+s3+end_date+s4
        nowpage=i
        sName=start_date+end_date+string.zfill(nowpage,6)+".txt"

        
        
        m=urllib2.urlopen(url).read()
       # print url
        
        if m.find('No Matching Records')==-1 and m.find('Records Not Found')==-1 :
            print "downloading "+str(nowpage)+" page,save as "+sName+"\n"
            f=open("India_htmlbase/"+sName,"w+")
            f.write(m)
            f.close()
        else:
           endflag=True
           
        i=i+1
    i=i-2
    return i

#def main(start_date, end_date):
  #  
   # trans=getlist(start_date, end_date)

  #  getinfo(start_date, end_date,trans)
    
    
        
