#一次寄多个
import tkinter as tk
from tkinter import ttk
import subprocess
import pandas as pd
from tkinter import END, EW, HORIZONTAL, VERTICAL, filedialog,messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import re
import log
import logging
import os
import errorcode
from pathlib import Path
from threading import Thread
import example
from dotenv import load_dotenv

load_dotenv(encoding='utf-8')
logger = logging.getLogger('robot')
example.example()

#上传文件
def upload_file(entry):
    entry.delete(0, END)
    selectFile = tk.filedialog.askopenfilename(filetypes =(("xlsx files","*.xlsx"),("xls files","*.xls"))) 
    entry.insert(0, selectFile)
def uploadappendix_file():
    selectFiles = tk.filedialog.askopenfilenames()
    for i, e in enumerate(selectFiles):          # 使用 enumerate 将串列变成带有索引值的字典
        appendix.insert('end', f'{e}')         # Text 从后方加入内容
        if i != len(selectFiles):
            appendix.insert('end', '\n')   
   
#侦测寄件人excel裡面是否包含前后空白，并做tuple处理
def getupload_fileload_senderentry(upload_fileload_senderentry):
    global sender
    try:
        df= pd.read_excel(upload_fileload_senderentry)
    except Exception as e:
        return messagebox.showerror(title='寄件人档案路径有误',message='请输入正确寄件人档案路径')
    if df.shape[1] != 2:
        messagebox.showerror(title='寄件人档案格式有误',message='寄件人档案格式有误，请确认')
        sendbutton['state']=tk.NORMAL
        val.set('0%')
        return logger.warning(f'寄件人档案格式有误，停止寄送')
    else:
        df.columns=['邮箱','密码']
        sender=[]
        for i in df.values:
            l=[]
            for j in i:
                l.append(str(j).strip())
            sender.append(tuple(l))
        return sender
#收件人与最大寄送数处理
def getupload_fileload_recipient(recipientlist,most):
    global recipient
    try:
        temp=[]
        try:
            for i in range(most):
                temp.append(recipientlist.pop(0))
        except:
            recipient=','.join([str(i) for i in temp])
            return recipient
        recipient=','.join([str(i) for i in temp])
        return recipient
    except:
        return messagebox.showerror(title='收件人档案路径有误',message='请输入正确收件人档案路径')

#收件人路径处理
def checkrecipient(recipientopen):
    checkrecipientlist=[]
    recipientlist=[]
    for i in recipientopen['邮箱']:
        #^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$
        if re.search('^[a-z0-9A-z.]+@(qq|hotmail|yahoo|outlook|gmail|[0-9]+).(com$|com.(ru|tw)$)',str(i).strip()):
            recipientlist.append(i)
            checkrecipientlist.append([str(i).strip(),'待寄送','-'])
        else:
            checkrecipientlist.append([str(i).strip(),'寄送失败','收件人格式有误'])
    checkrecipientlistdf=pd.DataFrame(checkrecipientlist,columns=['收件人','寄送状态','错误原因']).set_index('收件人')
    return recipientlist,checkrecipientlistdf
#最大收件人输入控制
def mostnumber(input):
    if input.isdigit():
        if 0<int(input)<501:
            return True
        else:
            return False
    elif input=="":
        return True
    else:
        return False
#登入
def login(fromm,pas):
    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:
        try:
            smtp.ehlo()  # 验证SMTP伺服器
            smtp.starttls()  # 建立加密传输
            smtp.login(fromm,pas)  # 登入寄件者gmail
            logger.info(f'{fromm}登入成功')

        except Exception as e:
            logger.warning(f'{fromm}无法登入，回传{fromm,"-","寄送失败!",str(e)}')
            return fromm,"-","寄送失败!",str(e)
#寄送
def send(title,fromm,pas,sep,html_text,**test):
    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 设定SMTP伺服器 
        #print(title,fromm,pas,sep,html)
        try:
            smtp.ehlo()  # 验证SMTP伺服器
            smtp.starttls()  # 建立加密传输
            smtp.login(fromm,pas)  # 登入寄件者gmail
            content = MIMEMultipart()  #建立MIMEMultipart物件
            content["subject"] = title  #邮件标题
            content["from"] = fromm  #寄件者
            content['bcc']=sep #密件副本
            content.attach(MIMEText(html_text,'html'))
            for i in appendixload:
                part_attach1 = MIMEApplication(open('%s'%i,'rb').read())   #开启附件
                part_attach1.add_header('Content-Disposition','attachment',filename='%s'%(i.split('/')[-1]))#为附件命名
                content.attach(part_attach1)
            smtp.send_message(content)  # 寄送邮件
            logger.debug(f'{fromm}寄给{sep}已寄送，回传{fromm,sep,"已寄送!","-"}')
            return fromm,sep,"已寄送!","-"
        except Exception as e:
            logger.warning(f'{fromm}寄给{sep}寄送失败，回传{fromm,sep,"寄送失败!",str(e)}')
            return fromm,sep,"寄送失败!",str(e)
#错误码
def errorcodecheck(errorcode1):
    return errorcode.CODE_DICT[errorcode1]
#按下确认后之处理
def confirm():
    global appendixload
    global processbar
    global val
    if not 'xls' in upload_fileload_entry1.get():
        logger.warning(f'confirm寄件人弹窗:{upload_fileload_entry1.get()}--->寄件人档案路径有误')
        return messagebox.showwarning(title='寄件人档案路径有误',message='请输入正确寄件人档案路径')
    elif not 'xls' in upload_fileload_entry2.get():
        logger.warning(f'confirm收件人弹窗:{upload_fileload_entry2.get()}--->收件人档案路径有误')
        return messagebox.showwarning(title='收件人档案路径有误',message='请输入正确收件人档案路径')
    elif subject.get()=='':
        logger.warning(f'confirm主旨弹窗:{subject.get()}--->主旨为空')
        return messagebox.showwarning(title='主旨',message='请输入主旨')
    elif textExample.get(1.0,'end')=='\n':
        logger.warning(f"confirm内文弹窗:{textExample.get(1.0,'end')}--->内文为空")
        return messagebox.showwarning(title='内文',message='请输入内文')
    #判断完寄、收、主、内后判断附件
    else:
        appendixload=appendix.get(1.0,'end').split('\n')[:-2]
        appendixload=[i.strip() for i in list(set(appendixload)) if i.strip()!='']
    result = messagebox.askyesno(title='确认一下',message="---请确认下列资讯是否正确---\n寄件者档案路径:%s\n收件者档案路径:%s\n收件人上限:%s\n主旨:%s\n夹带附件路径:%s"%(upload_fileload_entry1.get(),upload_fileload_entry2.get(),mostsend.get(),subject.get(),appendixload))
    logger.info(f"confirm结果弹窗:{result}，面板资讯，寄件人:{upload_fileload_entry1.get()}、收件人:{upload_fileload_entry2.get()}、最大上限:{mostsend.get()}、主旨:{subject.get()}、附件:{appendixload}")

    if result :
        sendbutton['state']=tk.DISABLED#禁止传送可以用
        val.set(f'寄送中...')
        frm.update()
        if appendix!=[]:#如果附件不为空
            appendixallsize=0
            for i in appendixload:
                if not os.path.isfile(i):
                    sendbutton['state']=tk.NORMAL
                    val.set('0%')
                    logger.warning(f"confirm附件:{appendixload}--->附件档案路径有误")
                    return messagebox.showwarning(title='附件档案路径有误',message='附件路径%s有误，请确认附件路径'%i)
                else:
                    file_size =Path(r'%s'%i).stat().st_size
                    logger.info(f"confirm附件大小:{i,file_size}")
                    appendixallsize+=file_size
            logger.info(f"confirm附件大小总和:{appendixallsize}")
            if appendixallsize>26214400:
                sendbutton['state']=tk.NORMAL
                val.set('0%')
                logger.warning(f"confirm附件大小弹窗:总档案大小超过25MB")
                return messagebox.showerror(title='总档案大小太大',message='总档案大小超过25MB')
        statelist=[]#设一statelist状态list之后做成excel用
        loginstate=[]
        #测试收件者档案路径对不对
        try:
            recipientopen=pd.read_excel(upload_fileload_entry2.get())
        except:
            sendbutton['state']=tk.NORMAL
            val.set('0%')
            logger.warning(f"confirmresult附件大小后收信人弹窗:{upload_fileload_entry2.get()}--->请输入正确收件人档案路径")
            return messagebox.showerror(title='收件人档案路径有误',message='请输入正确收件人档案路径')
        #判断收信人格式
        if recipientopen.shape[1] !=1:
            messagebox.showerror(title='收件人格式有误',message='收件人档案格式有误，请确认')
            logger.warning('收件人档案格式有误，停止寄送')
            sendbutton['state']=tk.NORMAL
            val.set('0%')
        else:
        #将收件者档案的COLUMN重命名并进checkrecipient资料处理
            recipientopen.columns=['邮箱']
            recipientinfo=checkrecipient(recipientopen)
            recipientlist=[]
            for i in recipientinfo[0]:
                if not i in recipientlist:
                    recipientlist.append(i)
            recipientdf=recipientinfo[1]
        
        #执行寄件人寄送
        logger.info('开始进行寄送...')
        processbar_count=0

        if getupload_fileload_senderentry(r'%s'%upload_fileload_entry1.get())==None:#寄件人档案格式有误(抓不到值)，停止寄送时触发
           return 
        recipiented=[]
        for senders in getupload_fileload_senderentry(r'%s'%upload_fileload_entry1.get()):
            if not login(senders[0],senders[1]):#判断可正常登入
                loginstate.append([senders[0],senders[1],'登入成功','-'])
                logger.info(f"回传{[senders[0],'*****','登入成功','-']}")
                if len(recipientlist)==0:#判断收件人LIST是否为0，因会逐一取出越来越短
                    statelist.append((senders[0],'-','寄送失败!','收件人不足'))
                    logger.warning(f'全部收信人已寄送完成!已无收信人!剩馀寄信人:{[i[0] for i in sender[sender.index(senders)+1:]]}')
                    if processbar_count< processbar['maximum']:
                        processbar_count = processbar_count + 100/len(sender)
                        processbar['value'] = processbar_count
                        val.set(f'寄送中...{int(processbar_count)}%')
                        frm.update()
                        time.sleep(0.01)
                    continue
                
                
                now_getupload_fileload_recipient=getupload_fileload_recipient(recipientlist,int(mostsend.get()))
                now_getupload_fileload_recipientlist=now_getupload_fileload_recipient.split(',')
                for i in now_getupload_fileload_recipientlist:
                    recipiented.append(i)
                
                dist={'title':str(subject.get()),
                'pas':senders[1],
                'sep':now_getupload_fileload_recipient,
                'html_text':textExample.get(1.0,'end'),
                'fromm':senders[0]}
                logger.info(f'获取{dist}')
                sendlist=list(send(**dist))
                for i in now_getupload_fileload_recipientlist:
                    recipientdf['寄送状态'][str(i).strip()]='已寄送'
                try:
                    if sendlist[3]!='-':
                        sendlist[3]=errorcode.CODE_DICT[sendlist[3][1:13]]['msg']
                        statelist.append(sendlist)
                    else:
                        statelist.append(sendlist)
                except:
                    statelist.append(sendlist)
                if processbar_count< processbar['maximum']:
                    processbar_count = processbar_count + 100/len(sender)
                    processbar['value'] = processbar_count
                    val.set(f'寄送中...{int(processbar_count)}%') 
                    frm.update()
                    time.sleep(0.01)
            else:
                loginfalselist=list(login(senders[0],senders[1]))
                try:
                    loginfalselist[3]=errorcode.CODE_DICT[loginfalselist[3][1:13]]['msg']
                    loginstate.append([senders[0],senders[1],'登入失败',loginfalselist[3]])
                    statelist.append(loginfalselist)
                except:
                    loginstate.append([senders[0],senders[1],'登入失败','帐号或密码错误'])
                    statelist.append(loginfalselist)
                processbar_count = processbar_count + 100/len(sender)
                processbar['value'] = processbar_count
                val.set(f'寄送中...{int(processbar_count)}%') 
                frm.update()
                time.sleep(0.01)
        if int(processbar_count)== processbar['maximum']:
            val.set(f'所有寄信人已寄送指定数量收信人完成!{int(processbar_count)}%')
            frm.update()
            sendbutton['state']=tk.NORMAL
        elif int(processbar_count)== 99:
            val.set(f'所有寄信人已寄送指定数量收信人完成!{int(processbar_count+1)}%')
            frm.update()
            sendbutton['state']=tk.NORMAL
        logger.info('本次所有寄信人已寄送指定数量收信人完成!')
        if len(recipientlist)!=0:
            logger.debug(f"寄件人都已寄送指定收件人，无剩馀寄件人，剩馀收件人:{recipientlist}")
            for i in recipientlist:
                recipientdf['寄送状态'][str(i).strip()]='寄送失败'
                recipientdf['错误原因'][str(i).strip()]='寄件人不足'
                statelist.append(('-',i,'寄送失败!','寄件人不足'))
        logindf=pd.DataFrame(loginstate,columns=['帐号','密码','登入状态','错误原因']).set_index('帐号')
        statedf=pd.DataFrame(statelist,columns=['寄件人','收件人','寄送状态','错误原因']).set_index('寄件人')
        localtime = time.localtime()
        resulttime = time.strftime("%Y-%m-%d %p%I.%M.%S ", localtime)
        logindf.to_excel('config/state/寄件人登入状态%s.xlsx'%resulttime)
        recipientdf.to_excel('config/state/收件人寄送状态%s.xlsx'%resulttime)    
        statedf.to_excel('config/state/寄送状态总览%s.xlsx'%resulttime)
        logger.info('--------------------------------------分隔线---------------------------------------')
#多执行绪1号
def threading():
    # Call work function
    t1=Thread(target=confirm)
    t1.start()       
        
root = tk.Tk()
root.title('批量发送邮件机器人')
frm = tk.Frame(root)
frm.grid(padx='20', pady='50')

# 版本號顯示
VERSION_NUMBER = os.getenv('VERSION')
lbl_version = tk.Label(text=f'　Version: {VERSION_NUMBER}')
lbl_version.grid(sticky='e')

#邮件伺服器
tk.Label(frm, text='邮件服务器:').grid(row=0,column=0,sticky='e')
tk.Label(frm, text='gmail').grid(row=0,column=1,sticky='w')

#开启范例档
examplebtn = tk.Button(frm, text='开启格式范例档',command=lambda: subprocess.Popen(f'explorer "{Path("寄、收件人excel格式范例档/").absolute()}"'),background='#F0F8FF',width=15,height=1)
examplebtn.grid(row=0,column=3,sticky='w')

#寄件人档案
tk.Label(frm, text='寄件人档案:').grid(row=1,column=0,sticky='e')
upload_fileload_entry1 = tk.Entry(frm, width='60')
upload_fileload_entry1.grid(row=1, column=1)
uploadbtn = tk.Button(frm, text='选择档案', command=lambda:upload_file(upload_fileload_entry1), background='#F0F8FF',width=15,height=1).grid(row=1, column=3,sticky='w')

#收件人档案
tk.Label(frm, text='收件人档案:').grid(row=2,column=0,sticky='e')
upload_fileload_entry2 = tk.Entry(frm, width='60')
upload_fileload_entry2.grid(row=2, column=1)
uploadbtn = tk.Button(frm, text='选择档案', command=lambda:upload_file(upload_fileload_entry2), background='#F0F8FF',width=15,height=1).grid(row=2, column=3,sticky='w')

#收件者上限
tk.Label(frm, text='寄送上限:').grid(row=3,column=0,sticky='e')
mostsend=tk.Spinbox(frm,width=59,from_=1,to_=500,increment=2)
mostsend.grid(row=3,column=1)
reg=frm.register(mostnumber)
mostsend.config(validate='key',validatecommand=(reg,'%P'))
tk.Label(frm, text='(单一寄件人寄送数量)').grid(row=3,column=3,sticky='w')

#主旨
tk.Label(frm, text='主旨:').grid(row=4,column=0,sticky='e')
subject = tk.Entry(frm, width='60')
subject.grid(row=4, column=1)
tk.Label(frm, text='(纯文字)').grid(row=4,column=3,sticky='w')

#内文
tk.Label(frm, text='内文:').grid(row=5,column=0,sticky='e')
scrollbartextExample_y= tk.Scrollbar(frm,orient=VERTICAL)
scrollbartextExample_x= tk.Scrollbar(frm,orient=HORIZONTAL)
textExample=tk.Text(frm, height=20,width=60,yscrollcommand=scrollbartextExample_y.set,xscrollcommand=scrollbartextExample_x.set,wrap="none")
textExample.grid(row=5, column=1)
scrollbartextExample_y.config(command=textExample.yview)
scrollbartextExample_x.config(command=textExample.xview)
scrollbartextExample_y.grid(row=5,column=2,sticky='ns')
scrollbartextExample_x.grid(row=6,column=1,sticky='ew')
tk.Label(frm, text='(可接受HTML)').grid(row=5,column=3,sticky='w')


#夹带附件
tk.Label(frm, text='夹带附件:').grid(row=7,column=0,sticky='e')
scrollbarappendix = tk.Scrollbar(frm,orient=VERTICAL)
appendix=tk.Text(frm, height=5,width=60,yscrollcommand=scrollbarappendix.set)
appendix.grid(row=7,column=1)
scrollbarappendix.config(command=appendix.yview)
scrollbarappendix.grid(row=7,column=2,sticky='ns')
appendixbtn=tk.Button(frm,text='选择档案',command=uploadappendix_file, background='#F0F8FF',width=15,height=1)
appendixbtn.grid(row=7, column=3,sticky='nw')
tk.Label(frm, text='一行一个附件路径，\n总共不得超过25MB',fg='red').grid(row=7,column=3,sticky='ws')



#发送邮件
sendbutton=tk.Button(frm, text='发送邮件',bg='#F0F8FF',command=threading)
sendbutton.grid(row=8,column=1)

#寄送状态条
processbar = ttk.Progressbar(frm, mode='determinate',length=410)
processbar.grid(column=1,row=9)
val = tk.StringVar()
val.set('0%')
processbar_label = tk.Label(frm, textvariable=val)
processbar_label.grid(column=1,row=10)


#开启纪录档
logbtn = tk.Button(frm, text='查看机器人纪录档',command=lambda: subprocess.Popen(f'explorer "{Path("config/log").absolute()}"'),background='#F0F8FF',width=15,height=1)
logbtn.grid(row=10,column=3,sticky='w')


#开启寄送状态资料夹
logbtn = tk.Button(frm, text='查看寄送状态',command=lambda: subprocess.Popen(f'explorer "{Path("config/state").absolute()}"'),background='#F0F8FF',width=15,height=1)
logbtn.grid(row=9,column=3,sticky='w')

root.mainloop()


