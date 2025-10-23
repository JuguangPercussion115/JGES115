"""
Spyder Editor
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-

This is a temporary script file.
"""

import configparser
import urllib
import pandas as pd
import hashlib
import datetime
from datetime import datetime
from datetime import timedelta

import os

from sqlite3 import Error
import sqlite3
import shutil
import random


#Upload To Google Colud
##conda install -c conda-forge google-cloud-storage
from google.cloud import storage

#pip install python-telegram-bot
#import telegram
#from telegram.ext import Dispatcher, MessageHandler, Filters


def procbar(a,b):
    return ""

def check_md5(filename):
    # 檢查Excel內容是否有改變
    
    #----檔案不存在則回傳空白
    if os.path.exists(filename)==False:
        return ""
    
    pd.read_excel(filename).to_csv(filename+".md5")
    # 檢查MD5
    m = hashlib.md5()
    with open(filename+".md5",'rb') as f:
        line = f.read()
        m.update(line)
    md5code = m.hexdigest()
    return md5code

def DownLoadFile(url,file_name=""):   
    
    # 下載檔案用 
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.96 Safari/537.36'} 
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64{})'} 
    request = urllib.request.Request(url,headers=headers)
    
    response = urllib.request.urlopen(request)
    
    if file_name=="":
       file_name=response.info().get_filename()
    if file_name=="":
       file_name="no_filename"
    
    
    file_size=int(response.headers.get('content-length',0))
    #print("---Download---")
    #print("Downloading: %s Bytes: %s" % (filename, file_size))
    file_size_dl = 0
    block_sz = 8192
    cnt=0
    
    #print("FileName:",filename)
    #input("A")
    status="   "
    with open(file_name, 'wb') as f:
        while True:
            buffer = response.read(block_sz)
            if not buffer:
                break
            cnt=cnt+1
            file_size_dl += len(buffer)
            f.write(buffer)
            
            if file_size>0:
               status = r"%10d  [%3.2f%%]" % (file_size_dl, file_size_dl * 100. / file_size)
               status = status + chr(8)*(len(status)+1) + procbar(file_size_dl,file_size)
            else:
               status = r"%10d" % (file_size_dl)
            
            
            if cnt % 10 ==1:
               #print(status)
               print("\rDownloading:"+status,end='')
               
    print("\r                                  \r File Size=({})".format(file_size))





def GenReport(df):
    break_line='''
<tr bgcolor="#000000">
<tr bgcolor="#000000">
  <td bgcolor="#000000" rowspan=1 colspan="{}">
  </td>
</tr>
</tr>
'''    


    #html_file_name="html/莒光兒樂比賽曲排名.html"
    gcp_file_name=PGM_INFO['產出GCP檔名']
    html_file_name=PGM_INFO['產出HTML檔名']
    
    
    temp_html_name="ini/"+"Sample_Race.html"
    
    pgm_info=dict()
    today_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    PGM_INFO["出勤資訊"]=PGM_INFO.get( "ADJ_UPDATE" )
    
    
    html_header=LoadPatten(temp_html_name,"HEADER")
    
    
    html_header=html_header.replace("#TITLE#",PGM_INFO['產出HTML標題'])
    #html_header=html_header.replace("#UPDATE_DATE#",today_str)     
    html_header=html_header.replace("#UPDATE_DATE#",PGM_INFO["_UPDATE"]     )    
    #                  
    html_header=html_header.replace("#UPDATE_DATE1#",str(PGM_INFO["出勤資訊"]))
    
    html_header1=LoadPatten(temp_html_name,"HEADER_R1")
    html_header1_1=LoadPatten(temp_html_name,"HEADER_R1_1")
    html_header1_2=LoadPatten(temp_html_name,"HEADER_R1_2")
        
    html_student=LoadPatten(temp_html_name,"STUDENT_R")
    html_student_rs=LoadPatten(temp_html_name,"STUDENT_RS")
    
    max_score=df['總積分'].max()
        
    
    
    with open(html_file_name, mode='w',encoding='utf-8') as f:
            f.write(html_header)
            f.write(html_header1)
            
            for song in song_list:
                tdstr=html_header1_1.replace("#SONG_NAME#",song)
                f.write(tdstr)
            f.write("</tr>")   
            
            f.write("<tr>") 
            for song in song_list:
                f.write(html_header1_2)
            f.write("</tr>")     
            
            student_cnt=0
            
            
            
            last_class=""
            
            for student_idx in range(0,len(df.index)):
                student_cnt=student_cnt+1
                
                
                stu_class=  str( df.iloc[student_idx,1]).strip() 
                stu_name=   str( df.iloc[student_idx,2]).strip()
                
                
                if (last_class[0:1]!=stu_class[0:1]) and (last_class!=""):
                    f.write(break_line.format(7+len(song_list)*4) )
                
                last_class=stu_class
                
                f.write("<tr>")
                
                html_student_str=html_student.replace("#STUDENT_NO#",str(student_cnt) )
                
                                                    
                
                 
                stu_rank=   str(df.iloc[student_idx]['總排名'])
                if stu_rank=="1":
                    stu_rank="🏆1"
                
                stu_score_tol=   str(df.iloc[student_idx]['總積分'])
                stu_score_abs=   str(df.iloc[student_idx]['出勤扣分'])
                stu_score_teach=   str(df.iloc[student_idx]['老師加分'])
                
                
                #print(stu_class,stu_name,stu_rank)
                                    
                
                html_student_str=html_student_str.replace("#STUDENT_CLASS#",stu_class)
                html_student_str=html_student_str.replace("#STUDENT_NAME#",stu_name)
                html_student_str=html_student_str.replace("#STUDENT_RANK#",stu_rank)
                
                
                #total_score=score_point.get(student[0]+"_"+student[3],None)
                
                    
                #input("A") 
                if stu_score_abs!="0":
                    html_student_str=html_student_str.replace("#ATTENT_SCORE#",stu_score_abs)
                else:    
                    html_student_str=html_student_str.replace("#ATTENT_SCORE#","")
                 
                if stu_score_teach!="0":   
                    html_student_str=html_student_str.replace("#TEACHER_SCORE#",stu_score_teach)
                else:
                    html_student_str=html_student_str.replace("#TEACHER_SCORE#","")
                    
                
                if stu_score_tol==None:
                    html_student_str=html_student_str.replace("#STUDENT_SCORE#","NA")
                else:    
                   # str(total_score) +
                   #html_student_st=str(stu_score_tol)+"<h1 class=left><progress class=a max={} value={}></progress></h1>".format(max_score,stu_score_tol)
                   html_student_st=str(stu_score_tol)+"<progress class=a max={} value={}></progress>".format(max_score,stu_score_tol)
                
                    
                html_student_str=html_student_str.replace("#STUDENT_SCORE#",html_student_st)
                
                
                if student_cnt%2==0:
                    html_student_str=html_student_str.replace("#BGCOLOR#","#DCECEA")
                else:
                    html_student_str=html_student_str.replace("#BGCOLOR#","#FFFF2CC")
                f.write(html_student_str)
                
                
                for song_id,song_name in enumerate(song_list):
                    html_student_rs_str=html_student_rs
                    
                    #info_key=student[0]+"_"+student[3]+"_"+song
                    #info=score_info.get(info_key,None)
                    #if info==None:
                    info=dict()
                    info['上傳時間']='<font size=4>-</font>'
                    info['優先通過分數']='-'
                    info['很好通過分數']=''
                    info['示範版分數']=''
                    
                    info['優先通過分數']=str(df.iloc[student_idx]['{}_通過分數'.format(song_id+1)])
                    if info['優先通過分數']=="0":
                        info['優先通過分數']="-"
                    
                    info['上傳時間']=    str(df.iloc[student_idx]['{}_通過時間'.format(song_id+1)])
                    
                    info['很好通過分數']= str(df.iloc[student_idx]['{}_很好加分'.format(song_id+1)])
                    if info['很好通過分數']=="0":
                        info['很好通過分數']=""
                    
                    info['示範版分數']=   str(df.iloc[student_idx]['{}_示範加分'.format(song_id+1)])
                    if info['示範版分數']=="0":
                        info['示範版分數']=""
                    

                        
                    html_student_rs_str=html_student_rs_str.replace("#上傳時間#",str(info['上傳時間']))
                    html_student_rs_str=html_student_rs_str.replace("#優先通過分數#",str(info['優先通過分數']))
                    html_student_rs_str=html_student_rs_str.replace("#很好通過分數#",str(info['很好通過分數']))
                    html_student_rs_str=html_student_rs_str.replace("#示範版分數#",str(info['示範版分數']))
                    f.write(html_student_rs_str)
                f.write("</tr>")    
    print("-已產生檔案:{}".format(html_file_name))           
    #upload_to_bucket(gcp_file_name,html_file_name,PGM_INFO['GCP_BUCKET_NAME'])
        

def get_modification_date(filename):
    t = os.path.getmtime(filename)
    return datetime.fromtimestamp(t)

    

def load_adj(bak=False):
    # 讀入調整檔
    
    #新鮮期為一小時,60秒*60分鐘*1小時
    refresh_time=60*60*1
    
    try:
        
        if bak==False:
           adj_file_url=PGM_INFO.get("ADJ_URL")
           download_file_name="加減分.xlsx"
           print("-下載出勤檔_{}".format(download_file_name))
           
           fresh_file=False
           
           if os.path.exists(download_file_name):
              file_date_time=get_modification_date(download_file_name) 
              print( "Time Dif:" ,(datetime.now() -  file_date_time).seconds)
              if (datetime.now() -  file_date_time).seconds < (refresh_time):
                 fresh_file=True
           
         #----如果現在-檔案時間小於　1小時,則不下載
           if fresh_file:
               print(" 上次下載檔案[{}]尚在新鮮期中,暫不下載。".format(download_file_name) )
           else:
               DownLoadFile(adj_file_url,download_file_name)
        else:
           download_file_name="加減分_BAK.xlsx" 
           print(" 使用備用,{}".format(download_file_name))
                             
    
    
    
        ret_list=list()
    
    
        #df1=pd.read_excel(file_name , names=["編號","班級","姓名","出勤加分","老師加分"] )
        df1=pd.read_excel(download_file_name )
        df2=df1.copy(deep=True)
        for chk_col in ['編號','班級','姓名','出勤扣分','老師加分']:
            if chk_col not in list(df2.columns):
               print("FORMAT ERROR")
               return pd.DataFrame()
    

        
        df2['班級']=df2['班級'].astype(int).astype(str).str.strip()
        #df['姓名']=df['姓名'].astype(object)
        df2['姓名']=df2['姓名'].astype(str).str.strip()
      
        df2['出勤扣分'].fillna(value=0, inplace=True)
        df2['老師加分'].fillna(value=0, inplace=True)  
      
        df2['出勤扣分']=pd.to_numeric(df2['出勤扣分'],errors='coerce')
        df2['老師加分']=pd.to_numeric(df2['老師加分'],errors='coerce')
        #df2['老師加分']=df2['老師加分'].astype(int)
        

        
        #print(dataTypeSeries)
        
        
    except Exception  as e:
        print("ERRPOR" , e)
        return pd.DataFrame()

    return df2

 
def LoadConfig(CONFIGFILENAME,pgm_info):
    # 載入設定檔 (LoadConfig)
    configfile=CONFIGFILENAME
    
    full_configfile=os.path.join(os.getcwd(),configfile)
        
    if os.path.exists(full_configfile)==False:
           SaveConfig(CONFIGFILENAME,pgm_info)
        
    config = configparser.ConfigParser()
        
    with open(full_configfile, mode='rb') as f:
             content = f.read()
    if content.startswith(b'\xef\xbb\xbf'):
        content=content[3:]
        #config.read(full_configfile,encoding="utf-8")
    config.read_string(content.decode('utf8'))
        #cfg.read_string(content.decode('utf8'))
        

    # 如果未定義則以字串方式讀入 ----
    for key in config['DEFAULT']:
        if key.upper() not in pgm_info:
            pgm_info[key.upper()]=config['DEFAULT'].get(key.upper())
    
    #print(config['DEFAULT'])
    for i in pgm_info:
        #print("para:",i,type(pgm_info[i]))
        if i[0:1]!="_":
        
            if isinstance(pgm_info[i], (str,)):
               try: 
                if config['DEFAULT'].get(i,"")!="":
                   pgm_info[i]=config['DEFAULT'].get(i).replace("#百分比#","%")
                           #print("Get Str:",pgm_info[i])
                   #input("Get Str")
                else:
                   pgm_info[i]=""
               except:  
                err=1    
                #print("Config2:",config['DEFAULT'][i])
                
                #pgm_info[i]=config['DEFAULT'].get[i]
                #print(i,"=",pgm_info[i])
            elif isinstance(pgm_info[i], (bool,)):    
                #print(i,"=",config['DEFAULT'][i],type(pgm_info[i]))
                #print("Config:",config['DEFAULT'].get(i),i)
                try:
                    c=config['DEFAULT'].getboolean(i)
                    if c==True:
                        pgm_info[i]=True
                        #print("True")
                    #elif c==False:
                    #    print("False")
                    else:
                        #print("Other",c,type(c))
                        pgm_info[i]=False
                except:
                    err=1        
                    #ccc
                #boolstr=config['DEFAULT'][i].upper()
                #if boolstr=="TRUE":
                #    pgm_info[i]=True
                #else:
                #    pgm_info[i]=False
                #print("Bool:","["+boolstr+"]")
            elif isinstance(pgm_info[i], (int,)):  
                print("INT",i)
                #try:
                if True:   
                    intstr=config['DEFAULT'].get(i,"")
                    #print("Get:",intstr)
                    #input("Get---")
                    if intstr.isdigit():
                       pgm_info[i]=int(intstr)
                       #print(pgm_info[i],type(pgm_info[i]))
                #except:
                #    print("ERR!!")
                #    input("")
                #    err=1
            elif isinstance(pgm_info[i], (list,)):   
                 #try:
                 if True:    
                     liststr=decode(config['DEFAULT'].get(i,""))
                     if liststr=="":
                        pgm_info[i] =list()
                     else:   
                        
                        liststr=liststr[1:-1]
                        #print(liststr)
                        #input("A") 
                        listarr=liststr.split(",")
                        newlist=list()
                        for list_i in listarr:
                            newlist.append(list_i.replace("'","").replace('"',"").strip())    
                        pgm_info[i]=newlist
                        #print(pgm_info[i])
                 #except:
                 #   err=1    
                 
                
            else:
                print("!"*10,i,type(pgm_info[i]))
    
    #print(pgm_info)        
    del config    

    # Check Config ------    
    #teachers=pgm_info["_teachers"]   
    
    if pgm_info.get("TELEGRAM_TOKEN","")=="":
       pgm_info["TELEGRAM_TOKEN"]=input("請輸入Telgegram Token,ex:1275534965:AAF8o3vM0Q0lxFdfIUQAcnoI6JH_nKYhg3c\n如不使用Telegram通知功能則輸入NONE\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)
    
    '''    
    if pgm_info.get("TELEGRAM_TOKEN","")!="" and pgm_info.get("TELEGRAM_CHAT_ID","")=="":
       pgm_info["TELEGRAM_CHAT_ID"]=input("請輸入Telgegram  Chat ID,ex: 792730143\n如不使用Telegram通知功能則輸入NONE\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)
    '''    
       
       
       
    
    if pgm_info.get("GCP金鑰檔名","")=="":
       pgm_info["GCP金鑰檔名"]=input("請輸入GCP金鑰檔名,ex:ini/Score111-bk.json\n如不使用GCP上傳功能則輸入NONE\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)
       
    if pgm_info.get("GCP_BUCKET_NAME","")=="":
       pgm_info["GCP_BUCKET_NAME"]=input("請輸入GCP bucket_name,ex:score111_2020\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)   
       
    if pgm_info.get("產出GCP檔名","")=="":
       pgm_info["產出GCP檔名"]=input("請輸入產出GCP檔名,ex: 2021比賽表.html\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)      
       
    if pgm_info.get("產出HTML檔名","")=="":
       pgm_info["產出HTML檔名"]=input("請輸入產出HTML檔名,ex: html/2021比賽表.html\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)   
       
    if pgm_info.get("產出HTML標題","")=="":
       pgm_info["產出HTML標題"]=input("請輸入產出HTML標題,ex: 2022全國賽_選拔排名\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)     
       
    if pgm_info.get("ADJ_URL","").strip()=='':   
       arrstr=input("請輸入加減分檔案URL,(請注意應該是https://drive.google.com/uc?export=download&id=1ecdMu3AD)\n")  
       pgm_info["ADJ_URL"]=arrstr
       SaveConfig(CONFIGFILENAME,pgm_info)
    
    if pgm_info.get("GRADE","").strip()=='':   
       arrstr=input("請輸入群組名稱陣列,ex:108兒樂,109兒樂,110兒樂,111兒樂\n")  
       pgm_info["GRADE"]=arrstr
       
       SaveConfig(CONFIGFILENAME,pgm_info)
       
    if pgm_info["GRADE"].strip()=="":
        input("無群組資訊須處理，程式中斷!")
        quit()
    else:
        pgm_info["_GRADE"]=pgm_info["GRADE"].split(",")    
        
    for grade in pgm_info["_GRADE"]:
        if pgm_info.get("{}_URL".format(grade),'').strip()=='':
           pgm_info["{}_URL".format(grade)]=input("請輸入{} 團的成績檔來源URL,(請注意應該https://drive.google.com/uc?export=download&id=1ecdMu3AD)\n".format(grade) ) 
           SaveConfig(CONFIGFILENAME,pgm_info)
    if pgm_info.get("曲目資訊","")=="":
       pgm_info["曲目資訊"]=input("請輸入產出曲目資訊,ex: 木棉道-1,木棉道-2\n") 
       SaveConfig(CONFIGFILENAME,pgm_info)            

    if pgm_info.get("REMOVE_STUDENT","").strip()=='':   
       pgm_info["REMOVE_STUDENT"]="601_林祈薰,602_陳育芸"
       
       SaveConfig(CONFIGFILENAME,pgm_info)

    
def load_one_group(group,song_list,bak=False):
    # 依照各團載入比賽曲成績
    
    # 讀入一個成績檔
    #print("Loading {}".format(group))
    ret_list=list()
    
    cnt=0
    
    
    

    if bak==False:
        file_name="{}.xlsx".format(group)
        print("-載入檔 {}".format(file_name))
    else:
        file_name="{}_BAK.xlsx".format(group)
        print(" 載入備份 {}".format(file_name))
    
    if os.path.exists(file_name)==False:
        print("  {} 不存在".format(file_name))
        return None
    else:
        try:
           df=pd.read_excel(file_name)
        except Exception as e:
            print("錯誤:",str(e))   
            print(" \n載入Excel錯誤，請檢查 {} 是否為正常Excel檔案".format(file_name),end="")
            input(" 請按 Enter 繼續!!!")
            
            return None
            
        
    #print("Size:",len(df.index))
    
    #讀入學生清單
    student_list=list()
    for student_idx in range(0,len(df.index)):
        # 學生編號
        #Stu_idx=group+str(df.iloc[student_idx,0]).strip().zfill(2)
        Stu_idx=group+str(df.iloc[cnt,0]).strip().zfill(2)
        cnt=cnt+1
        
        # 學生班級 ---
        Class=str(df.iloc[student_idx,1]).strip()
        # 學生姓名 ---
        Name=df.iloc[student_idx,2].strip()
        
        #print(df.iloc[student_idx,1])
        #print(df.iloc[student_idx,2])
        student_list.append((Class,Name))
        
        #print("學生班級:",Class ,end=" ")
        #print("學生姓名:",Name,end=" ")
        
        # 取出所有曲目成績 ------
        sco=list()
        for song_id,song_name in enumerate(song_list):


            pass_mode=""
            pass_time=""
            demo_score=0
            good_score=0
            pass_score=0
            pass_mode=""
            
            try:
                #print("比賽曲目:",song_name,end=" ")
                if song_name in list(df):
                    # 計算本曲目所在位置 (0,1,2 是編號,班級,	姓名) 
                    
                    song_id=(list(df).index(song_name)-3)//2
                    # 欄位算法  (0編號	1班級	2姓名	3木棉道-1	4時間) 
                    # 3,4 5,6 
                    # 3,4 開始,每首歌偏移 2  , 3+song_id*2
                    #通過模式
                    pass_mode=str(df.iloc[student_idx,3+song_id*2]).strip()
                    #通過時間
                    pass_time=str(df.iloc[student_idx,4+song_id*2]).strip().replace("-","/")
                    
                    
                    if "★" in pass_mode:
                        demo_score=5
                    if "●" in pass_mode:
                        good_score=5

                    if pass_mode in ("nan","","X","□"):
                       pass_mode=""
                       pass_time=""
                       
                    if pass_mode not in ("★","●","◎","○","X","nan","","★◎","◎★","★●","●★","★○","○★","□"):
                        print("["+pass_mode+"]")
                        input("A")
            except:
                return None
  
                #pass_score=0
                
            sco=sco+[ pass_mode,demo_score,good_score,pass_time,pass_score ]   
            #print("通過方式",pass_mode,end=" ")
            #print("示範版加分:",demo_score ,end=" ")    
            #print("很好加分:",good_score,end=" ")    
            #print("通過時間:",pass_time,end=" ")
            #print("通過分數:",pass_score,end)

                
            #print("Data:",df.iloc[student_idx,2+2])
        ret_list.append( [Stu_idx,Class,Name]+ sco    ) 
            
                
        
        #print(Class,Name)
    
    #print()
    #print(ret_list)
    #input("Aaaa")
    return ret_list 

def load_one_group_from_db(group):
    pass_code={"示範版":"★","很好通過":"●","通過":"◎","通過有評語":"○"}

    
    # 組裝欄位清單 ----
    column_name=['編號','班級','姓名']
    for song_id,song_name in enumerate(song_list):
        column_name.extend( [song_name,'通過時間'] )

    students=list()
    
    
    conn = sqlite3.connect("Score.db")
    sqlcmd='''
         select 編號,班級,姓名 from Student_info
        where 團=? order by 編號
        '''
    cur = conn.cursor()
    cur1 = conn.cursor()
    
    cur.execute(sqlcmd,(group,))
    rows = cur.fetchall()
    
    
    for row in rows:
        student=list(row)
        #print(student)
        for song_id,song_name in enumerate(song_list):
            sqlcmd='''
select 通過方式,strftime('%Y/%m/%d %H:%M:%S',datetime(貼文時間U,'unixepoch', 'localtime')) as 通過時間,示範版 from score_info 
               where 團=? and 曲目=? and 學生姓名=? and 通過方式 in ('示範版','很好通過','通過','通過有評語') and Processed='Y' order by 貼文時間u               
'''
         
            #print(song_name,student[2])
            #print(sqlcmd)
            cur1.execute(sqlcmd,(group,song_name,student[2]))
            data=cur1.fetchone()
            
             
            
            if data==None:
                student.extend(['','']) 
            else:
                pass_mode=pass_code[data[0]]
                #data[0]=pass_code[data[0]]
                if data[2]=='示範版':
                  pass_mode="★"+pass_mode 
                
                student.extend( [pass_mode, data[1]] ) 
        #print(student)
        students.append(student)
        
        #input("A")
    #print(column_name)
    #print(students)
    df=pd.DataFrame(students,columns=column_name)     
    #print(df)
    print("-由DB匯出本地檔案 {}.xlsx".format(group))
    df.to_excel("{}.xlsx".format(group),index=False)

    

           
           


        



def LoadPatten(filename,patten_tag):
    #print("FileName:",filename)
    
    ret=""
    with open(filename, mode='r',encoding='utf-8') as f:
             content = f.read()
             #print("Cont:",content)

             #print(content)
             #print("#"+patten_tag.upper())
             f1=content.find("#"+patten_tag.upper()+"#")
             if f1>=0:
                 
                f2=content.find("#"+patten_tag+"#",f1+1)
                if f2>=0:                
                   ret=content[f1+len(patten_tag)+3:f2]
                else:
                   ret="Not Found"
    return ret
                   
def message(msg):
    print(msg)   
                                                       
                    
def SaveConfig(CONFIGFILENAME,pgm_info):
    # 儲存設定檔 (SavedConfig)
    configfile=CONFIGFILENAME
    full_configfile=os.path.join(os.getcwd(),configfile)
    
    config = configparser.ConfigParser()
    
    #if pgm_info["GRADE"]!=""
    #pgm_info["GRADE"]=",".join(pgm_info["GRADE"])
    #print(pgm_info["GRADE"],type(pgm_info["GRADE"]))
    #input("A")
    
    try:
     config.add_section("DEFAULT")
    #except configparser.DuplicateSectionError:
    except Exception as e:
     #print("ERR1",str(e))   
     #input("ERR1")
     Fail=1
    
    #print(pgm_info)
    #input("Before save")
    for i in pgm_info:
        if i[0:1]!="_":
            #print(pgm_info[i],type(pgm_info[i]))
            if isinstance(pgm_info[i], str):
            #config.set("DEFAULT",'pgm_title',pgm_info['pgm_title'])
                #print("str:=",pgm_info[i])
                config.set("DEFAULT",i,pgm_info[i].replace("%","#百分比#"))
                
                #print(i,"===#",pgm_info[i],type(pgm_info[i]))
            if isinstance(pgm_info[i], (int,bool)):
                config.set("DEFAULT",i,str(pgm_info[i]))
                
            if isinstance(pgm_info[i], list):    
               #print("En:",encodeStr(str(pgm_info[i]))) 
               
               config.set("DEFAULT",i,encode(str(pgm_info[i])))
                

    
    with open(full_configfile, 'w',encoding='utf-8') as configfile:
      config.write(configfile)
    
    
    
    
    del config 






def upload_to_bucket(blob_name, path_to_file, bucket_name):
    # 上傳到Google Bucket空間 (upload_to_bucket)
    if PGM_INFO['GCP金鑰檔名']=="" or PGM_INFO['GCP金鑰檔名'].upper()=="NONE":
        print("無GCP 金鑰檔案，無法傳送,請自行上傳:",path_to_file)
        
        return
    if not os.path.isfile(PGM_INFO['GCP金鑰檔名']):
       print(PGM_INFO['GCP金鑰檔名'])
       print("GCP 金鑰檔案{}不存在，無法傳送,請自行上傳:".format(PGM_INFO['GCP金鑰檔名']),path_to_file)
       return

    print("-上傳 "+path_to_file+" 至 GCP "+blob_name,end=" ")
    #input("wait")
    #return
    """ Upload data to a bucket"""
    ####
    # Explicitly use service account credentials by specifying the private key
    # file.
    
    
    # 讀取GCP憑證檔 -------
    storage_client = storage.Client.from_service_account_json(PGM_INFO['GCP金鑰檔名'])
    buckets = list(storage_client.list_buckets())
    print(buckets)
    #print(buckets = list(storage_client.list_buckets())
    
    bucket = storage_client.get_bucket(bucket_name)
    #print(bucket.blob('109團_指定作業_聖母頌.xlsx'))
    #print(bucket)
    blob = bucket.blob(blob_name)
    blob.upload_from_filename(path_to_file)

    #returns a public url
    #return blob.public_url
    print("-已上傳到URL:\n",blob.public_url)
    
    return blob.public_url







#--- 主程式---------------------------------------
pgm_ver="20211024"
mess_str="莒光國小兒童樂隊,國賽選拔競賽排行程式 Ver:"+pgm_ver
print(mess_str)
print("-"*46)


# 主變數區
PGM_INFO=dict()
CONFIGFILENAME='ini/competition.ini'
LoadConfig(CONFIGFILENAME,PGM_INFO)
PGM_INFO["_UPDATE"]=""


#adj_df=""
song_list=PGM_INFO["曲目資訊"].split(",")
score_info=list()


message("#競賽排行程式:"+pgm_ver+"啟動 ----------")

# 下載入所有檔案-----------------------
ret=list()

# 由資料庫中產生/或網路下傳
print("*下載檔案*****"+"*"*30)

#各團逐一處理,Ex: 111, 112, 113
for group in PGM_INFO["_GRADE"]:
    src_path=PGM_INFO.get("{}_URL".format(group),"")

    if src_path.upper()=="DB":
       #----由資料庫產生 
       load_one_group_from_db(group)
       message("-{}由資料庫產生.".format(group))
    else:
       print("-準備下載{}.xlsx 來源為:{}".format(group,src_path))
       try:
           download_file_name="{}.xlsx".format(group)
           
           # 如果本機有檔案，先看看是否在一小時內
           if os.path.exists(download_file_name)==True:
               file_date_time=get_modification_date(download_file_name)
               print( datetime.now())
               print( file_date_time)
               print( "Time Dif:" ,(datetime.now() -  file_date_time).seconds)
               print( (datetime.now() -  file_date_time).seconds < (60*60*1) )
               
               #input("A")
               #----如果現在-檔案時間小於　1小時,則不下載
               if (datetime.now() -  file_date_time).seconds < (60*60*1) :
                    print(" 上次下載檔案[{}]尚在新鮮期中".format(download_file_name) )
                    message("-{}尚新鮮.".format(download_file_name))
                    continue
           #print("Donload",download_file_name)
           DownLoadFile(src_path,download_file_name) 
              
       except Exception  as e:
           print(" 下載失敗!",e)
           message("-下載{}. 失敗!".format(download_file_name))
       else:
           message("-下載{}.  OK.".format(download_file_name))
        
print("*載入檔案*****"+"*"*30)




# 由本地檔案載入所有曲目
PGM_INFO["_UPDATE"]=""
for group in PGM_INFO["_GRADE"]:
    #先載入最新檔案

    #----- 依照各團載入比賽曲成績
    #message("Load {}.".format(group))
    new_ret=load_one_group(group,song_list)
    if new_ret==None:
        # 改抓舊備份檔案
        new_ret=load_one_group(group,song_list,bak=True)
    else:
        # 檔案正常，備份成備份檔
        #print("OK")
        md5_bef=check_md5( "{}.xlsx".format(group) )
        md5_aft=check_md5( "{}_BAK.xlsx".format(group) )
        #print("MD5_bf:",md5_bef)
        #print("MD5_af:",md5_aft)
        
        #-----檔案不一樣才備份
        if md5_bef!=md5_aft:
            shutil.copy ("{}.xlsx".format(group),"{}_BAK.xlsx".format(group))
            #---更新檔案時間
            update_str=datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            PGM_INFO["{}_UPDATE".format(group) ]=update_str
            SaveConfig(CONFIGFILENAME,PGM_INFO)
            message( "-{}成績更新:{}".format(group,update_str) )
            #PGM_INFO["_UPDATE"]=PGM_INFO("_UPDATE","")+" "+PGM_INFO["{}_UPDATE".format(group) ]
        else:
            print(" 檔案未更新!")
            message( "-{}成績未更新".format(group) )
            
    #----產生更新日期
    group_update=PGM_INFO.get( "{}_UPDATE".format(group),"" )
    if group_update !="":
       PGM_INFO["_UPDATE"]=PGM_INFO.get("_UPDATE","")+"{}更新時間:".format(group)+group_update+"　　　"

    if new_ret!=None:    
       ret=ret+new_ret

# 載入計分調整檔
message("-Load Adj.")
adj_df=load_adj()
if adj_df.empty:
   adj_df=load_adj(bak=True)
   #adj_df=pd.DataFrame(adj_ret)
else:
   #adj_df=pd.DataFrame(adj_ret)
   # 檔案正常，備份成備份檔
   #-----檔案不一樣才備份
    md5_bef=check_md5( "加減分.xlsx" )
    md5_aft=check_md5( "加減分_BAK.xlsx" )
    if md5_bef!=md5_aft:
        shutil.copy ("加減分.xlsx","加減分_BAK.xlsx")
        #---更新檔案時間
        print(" 出勤檔更新")
        PGM_INFO["ADJ_UPDATE" ]=datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        message("-出勤檔更新:".format(PGM_INFO["ADJ_UPDATE" ]))
        SaveConfig(CONFIGFILENAME,PGM_INFO)
    else:
        print(" 檔案未更新!")
   



print("*計算分數*****"+"*"*30)
message("-計算分數")
# 匯入Pandas DataFrame---------------------------
columns =['序號', '班級', '姓名']
for song_id,song_name in enumerate(song_list):
    
    #print(song_id,song_name)
    columns.extend([str(song_id+1)+"_通過方式",
                   str(song_id+1)+"_示範加分",
                   str(song_id+1)+"_很好加分",
                   str(song_id+1)+"_通過時間",
                   str(song_id+1)+"_通過分數"] )
df1=pd.DataFrame(ret,columns=columns)


#計算各首曲目之排名分數 -------
students_count=len(df1.index)
print("-計算各曲目排名")
for song_id,song_name in enumerate(song_list):
    message("-計算曲目:{}".format(song_name))
    print(" {} {}".format(song_id,song_name) )
    sort_field='{}_通過時間'.format(song_id+1)
    score_field='{}_通過分數'.format(song_id+1)
    
    # 未通過不計分，所以拆兩組計算================
    #.copy(deep=True)
    df_none_space=df1[ df1[sort_field ] !=''].copy(deep=True)
    df_space=df1[ df1[sort_field ] ==''].copy(deep=True)
    
    # 通過者依照通過時間計算================
    df_none_space[score_field] =  students_count - df_none_space[ sort_field ].rank(method='min').astype(int)
    
    # 兩組合併回來 ============================
    df1=pd.concat([df_none_space,df_space])

print("-合併分數調整")


# 移除本次不參賽名單
print("-移除本次不參賽名單")
remove_students=PGM_INFO.get("REMOVE_STUDENT","").split(",")
for stu_info in remove_students:
    stu_cls,stu_name = stu_info.split("_")
    print(" ",stu_cls,stu_name)
#print(remove_student)
#pgm_info["REMOVE_STUDENT"]="601_林祈薰,602_陳育芸"
    indexname=df1[ (df1['班級']==stu_cls) & (df1['姓名']==stu_name) ].index
    df1.drop( indexname ,
         inplace=True)


# 合併分數調整檔
df1= pd.merge(df1,
              adj_df,
              how='left',
              #left_on=['班級','姓名'],
              left_on=['班級','姓名',],
              #right_on = ['班級','姓名'] )
              right_on = ['班級','姓名'], )
message("-合併調整檔完成")

#加總分數 -----------------
df1['總積分']=0
#df1['出勤加減分']=0
#df1['老師加減分']=0

df1['出勤扣分'].fillna(value=0, inplace=True)
df1['老師加分'].fillna(value=0, inplace=True)


df1['出勤扣分']=df1['出勤扣分'].astype(int)
df1['老師加分']=df1['老師加分'].astype(int)





sum_list=df1['總積分']+df1['出勤扣分']+df1['老師加分']

print("-加總各曲目分數及調整分數")
for song_id,song_name in enumerate(song_list):
    print(" {} {}".format(song_id,song_name) )
    score_field='{}_通過分數'.format(song_id+1)
    sco_list=df1[score_field]
    sum_list+=sco_list
    
    score_field='{}_示範加分'.format(song_id+1)
    sco_list=df1[score_field]
    sum_list+=sco_list
    
    
    score_field='{}_很好加分'.format(song_id+1)
    sco_list=df1[score_field]
    sum_list+=sco_list
    
df1['總積分']=sum_list




# 計算總分數排行
#print(df1['總積分'].rank(method='min',ascending=False).astype(int))
df1['總排名'] =  df1['總積分'].rank(method='min',ascending=False).astype(int)


print("*產生檔案"+"*"*30)

# 寫入計算用Excel除錯用
try:
  df1.to_excel("Result.xlsx")
except:
    print( "無法寫入參考用Excel (Result.xlsx)" )

# print(df1.head(2))
# df2=df1.sort_values('總排名')
# 排序號產生比賽表 ---

df2=df1.sort_values('序號')
#df1.to_excel("Result.xlsx")

# 產生成績報表並上傳
GenReport(df2)
message("#競賽排行程式:"+pgm_ver+"批次結束 ----------")
#https://bit.ly/2XUKI7y

