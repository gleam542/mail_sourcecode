# 一次寄多个
import tkinter as tk
from tkinter import ttk
import subprocess
import pandas as pd
from tkinter import END, EW, HORIZONTAL, VERTICAL, filedialog, messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
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


class Mail:
    def __init__(self):
        load_dotenv(encoding='utf-8')
        self.logger = logging.getLogger('robot')
        example.example()  # 創造範例XLSX檔
        self.root = tk.Tk()
        self.root.title('批量发送邮件机器人')
        self.frm = tk.Frame(self.root)
        self.frm.grid(padx='20', pady='50')

        # 版本號顯示
        VERSION_NUMBER = os.getenv('VERSION')
        self.lbl_version = tk.Label(text=f'　Version: {VERSION_NUMBER}')
        self.lbl_version.grid(sticky='e')

        # 邮件伺服器
        tk.Label(self.frm, text='邮件服务器:').grid(row=0, column=0, sticky='e')
        self.host_cb = ttk.Combobox(self.frm, values=['gmail.com', '189.cn', 'yeah.net'], state="readonly", width=58)
        self.host_cb.grid(row=0, column=1)
        self.host_cb.current(0)
        self.host_cb.bind('<<ComboboxSelected>>', self.combobox_selecter)

        # 开启范例档
        self.example_btn = tk.Button(self.frm, text='开启格式范例档', command=lambda: subprocess.Popen(
            f'explorer "{Path("寄、收件人excel格式范例档/").absolute()}"'), background='#F0F8FF', width=15, height=1)
        self.example_btn.grid(row=0, column=3, sticky='w')

        # 寄件人档案
        tk.Label(self.frm, text='寄件人档案:').grid(row=1, column=0, sticky='e')
        self.upload_file_load_entry1 = tk.Entry(self.frm, width='60')
        self.upload_file_load_entry1.grid(row=1, column=1)
        tk.Button(self.frm, text='选择档案',
                  command=lambda: self.upload_file(self.upload_file_load_entry1),
                  background='#F0F8FF', width=15, height=1).grid(row=1, column=3, sticky='w')

        # 收件人档案
        tk.Label(self.frm, text='收件人档案:').grid(row=2, column=0, sticky='e')
        self.upload_file_load_entry2 = tk.Entry(self.frm, width='60')
        self.upload_file_load_entry2.grid(row=2, column=1)
        tk.Button(self.frm, text='选择档案',
                  command=lambda: self.upload_file(self.upload_file_load_entry2),
                  background='#F0F8FF', width=15, height=1).grid(row=2, column=3, sticky='w')

        # 收件者上限
        tk.Label(self.frm, text='寄送上限:').grid(row=3, column=0, sticky='e')
        self.most_send = tk.Spinbox(self.frm, width=59, from_=1, to_=500, increment=2)
        self.most_send.grid(row=3, column=1)
        self.reg = self.frm.register(self.most_number)
        self.most_send.config(validate='key', validatecommand=(self.reg, '%P'))
        tk.Label(self.frm, text='(单一寄件人寄送数量)').grid(row=3, column=3, sticky='w')

        # 寄送頻率
        tk.Label(self.frm, text='寄送頻率:').grid(row=4, column=0, sticky='e')
        self.frequency_ = tk.Spinbox(self.frm, width=59, from_=1, to_=3600, increment=2)
        self.frequency_.grid(row=4, column=1)
        self.fre = self.frm.register(self.frequency)
        self.frequency_.config(validate='key', validatecommand=(self.fre, '%P'))
        tk.Label(self.frm, text='單封寄出時間(秒)').grid(row=4, column=3, sticky='w')

        # 主旨
        tk.Label(self.frm, text='主旨:').grid(row=5, column=0, sticky='e')
        self.subject = tk.Entry(self.frm, width='60')
        self.subject.grid(row=5, column=1)
        tk.Label(self.frm, text='(纯文字)').grid(row=6, column=3, sticky='w')

        # 内文
        tk.Label(self.frm, text='内文:').grid(row=6, column=0, sticky='e')
        self.scrollbar_text_Example_y = tk.Scrollbar(self.frm, orient=VERTICAL)
        self.scrollbar_text_Example_x = tk.Scrollbar(self.frm, orient=HORIZONTAL)
        self.textExample = tk.Text(self.frm, height=20, width=60, yscrollcommand=self.scrollbar_text_Example_y.set,
                                   xscrollcommand=self.scrollbar_text_Example_x.set, wrap="none")
        self.textExample.grid(row=6, column=1)
        self.scrollbar_text_Example_y.config(command=self.textExample.yview)
        self.scrollbar_text_Example_x.config(command=self.textExample.xview)
        self.scrollbar_text_Example_y.grid(row=6, column=2, sticky='ns')
        self.scrollbar_text_Example_x.grid(row=7, column=1, sticky='ew')
        tk.Label(self.frm, text='(可接受HTML)').grid(row=6, column=3, sticky='w')

        # 夹带附件
        tk.Label(self.frm, text='夹带附件:').grid(row=8, column=0, sticky='e')
        self.scrollbar_appendix = tk.Scrollbar(self.frm, orient=VERTICAL)
        self.appendix = tk.Text(self.frm, height=5, width=60, yscrollcommand=self.scrollbar_appendix.set)
        self.appendix.grid(row=8, column=1)
        self.scrollbar_appendix.config(command=self.appendix.yview)
        self.scrollbar_appendix.grid(row=8, column=2, sticky='ns')
        self.appendix_btn = tk.Button(self.frm, text='选择档案', command=self.upload_appendix_file, background='#F0F8FF',
                                      width=15, height=1)
        self.appendix_btn.grid(row=8, column=3, sticky='nw')
        tk.Label(self.frm, text='一行一个附件路径，\n总共不得超过25MB', fg='red').grid(row=8, column=3, sticky='ws')

        # 发送邮件
        self.send_button = tk.Button(self.frm, text='发送邮件', bg='#F0F8FF', command=self.threading)
        self.send_button.grid(row=9, column=1)

        # 寄送状态条
        self.processbar = ttk.Progressbar(self.frm, mode='determinate', length=410)
        self.processbar.grid(column=1, row=10)
        self.val = tk.StringVar()
        self.val.set('0%')
        self.processbar_label = tk.Label(self.frm, textvariable=self.val)
        self.processbar_label.grid(column=1, row=11)

        # 开启纪录档
        self.log_btn = tk.Button(self.frm, text='查看机器人纪录档',
                                 command=lambda: subprocess.Popen(f'explorer "{Path("config/log").absolute()}"'),
                                 background='#F0F8FF', width=15, height=1)
        self.log_btn.grid(row=11, column=3, sticky='w')

        # 开启寄送状态资料夹
        self.log_btn = tk.Button(self.frm, text='查看寄送状态',
                                 command=lambda: subprocess.Popen(f'explorer "{Path("config/state").absolute()}"'),
                                 background='#F0F8FF', width=15, height=1)
        self.log_btn.grid(row=10, column=3, sticky='w')

        self.root.mainloop()

    # 上传文件
    def upload_file(self, entry):
        entry.delete(0, END)
        select_file = tk.filedialog.askopenfilename(
            filetypes=(("xlsx files", "*.xlsx"), ("xls files", "*.xls")))  # selectFile是檔案路徑
        entry.insert(0, select_file)

    def upload_appendix_file(self):
        select_files = tk.filedialog.askopenfilenames()  # 這個也是選擇檔案路徑
        for i, e in enumerate(select_files):  # 使用 enumerate 将串列变成带有索引值的字典
            self.appendix.insert('end', f'{e}')  # Text 从后方加入内容
            if i != len(select_files):
                self.appendix.insert('end', '\n')

    # 侦测寄件人excel裡面是否包含前后空白，并做tuple处理
    def getup_load_file_load_sender_entry(self, upload_file_load_sender_entry):
        try:
            df = pd.read_excel(upload_file_load_sender_entry)



        except Exception as e:
            return messagebox.showerror(title='寄件人档案路径有误', message='请输入正确寄件人档案路径')

        if df.shape[1] != 2 and df.shape[1] != 3:
            messagebox.showerror(title='寄件人档案格式有误', message='寄件人档案格式有误，请确认')
            self.send_button['state'] = tk.NORMAL
            self.val.set('0%')
            return self.logger.warning(f'寄件人档案格式有误，停止寄送')
        elif df.shape[1] == 2:
            df.columns = ['邮箱', '密码']
        elif df.shape[1] == 3:
            df.columns = ['邮箱', '密码', '暱称']
        self.sender = []
        for i in df.values:
            l = []
            for j in i:
                l.append(str(j).strip())

            self.sender.append(tuple(l))
        return self.sender

    # 收件人与最大寄送数处理
    def getup_load_file_load_recipient(self, recipient_list, most):
        try:
            temp = []
            try:
                for i in range(most):
                    temp.append(recipient_list.pop(0))
            except:
                recipient = ','.join([str(i) for i in temp])
                return recipient
            recipient = ','.join([str(i) for i in temp])
            return recipient
        except:
            return messagebox.showerror(title='收件人档案路径有误', message='请输入正确收件人档案路径')

    # 收件人路径处理
    def check_recipient(self, recipient_open):
        check_recipient_list = []
        recipient_list = []
        for i in recipient_open['邮箱']:
            # ^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$

            if re.search('^[a-z0-9A-z.]+@(yeah|qq|hotmail|yahoo|outlook|gmail|[0-9]+).(net$|com$|com$|cn$|ru$|tw$)',
                         str(i).strip()):
                recipient_list.append(i)  # 若符合信箱格式 放入待寄送清單
                check_recipient_list.append([str(i).strip(), '待寄送', '-'])
            else:
                check_recipient_list.append([str(i).strip(), '寄送失败', '收件人格式有误'])
        check_recipient_list_df = pd.DataFrame(check_recipient_list, columns=['收件人', '寄送状态', '错误原因']).set_index('收件人')
        return recipient_list, check_recipient_list_df

    # 最大收件人输入控制
    def most_number(self, input):
        if input.isdigit():
            if 0 < int(input) < 501:
                return True
            else:
                return False
        elif input == "":
            return True
        else:
            return False

    # 寄送頻率输入控制
    def frequency(self, input):
        if input.isdigit():
            if 0 < int(input) < 1001:
                return True
            else:
                return False
        elif input == "":
            return True
        else:
            return False

    # 信箱下拉選單監控 如果選到yeah.net自動跳回預設 並顯示暫不支援彈窗
    def combobox_selecter(self, event):
        if self.host_cb.current() == 2:
            messagebox.showinfo('系统提示', '目前暂不支援此邮件服务器')
            self.host_cb.current(0)


    # 登入、寄送
    def send(self, title, fromm, pas, sep, html_text, sender_name, total_recipient, processbar_count, **test):
        # 開始寄送
        self.logger.info('開始寄送')
        try:
            with smtplib.SMTP(host='smtp.' + str(self.host_cb.get()), timeout=120, port=25) as smtp:  # 设定SMTP伺服器
                # smtp.set_debuglevel(1)
                smtp.ehlo()  # 验证SMTP伺服器
                smtp.starttls()  # 建立加密传输
                smtp.login(fromm, pas)  # 登入寄件者mail
                self.logger.info(f"回传{[fromm, pas, 'SMTP SERVER建立成功', '-']}")
                sep = self.getup_load_file_load_recipient(sep[0], sep[1])  # 取得收件者mail清單(算上你能寄的最大數量的清單)
                sep = str(sep).split(',')
                self.logger.info(f"{fromm}準備寄送名單:{sep}")
                add_rate = (100 - processbar_count) / total_recipient  # 計算寄件百分比

                for i in sep:
                    # 建立MIMEMultipart物件並附值
                    content = MIMEMultipart()  # 建立MIMEMultipart物件
                    content["subject"] = title  # 邮件标题
                    content["from"] = formataddr([sender_name, fromm])  # 暱稱加上寄件者
                    content['bcc'] = i  # 密件副本
                    content.attach(MIMEText(html_text, 'html'))  # 內文
                    for j in self.appendix_load:  # 附件
                        part_attach1 = MIMEApplication(open('%s' % j, 'rb').read())  # 开启附件
                        part_attach1.add_header('Content-Disposition', 'attachment',
                                                filename='%s' % (j.split('/')[-1]))  # 为附件命名
                        content.attach(part_attach1)
                    try:

                        smtp.send_message(content)  # 寄送邮件
                        self.logger.info(f'{fromm}寄给{i}已寄送')
                        self.logger.info(f'寄送頻率:{int(self.frequency_.get())}，等待{int(self.frequency_.get())}秒...')
                        time.sleep(int(self.frequency_.get()))
                    except smtplib.SMTPException as e:
                        if i == sep[0]:  # 全寄失敗
                            try:
                                self.logger.warning(f'{fromm}寄给{i}寄送失败，原因:{errorcode.CODE_DICT[str(e)[1:4]]["msg"]}')
                                sep.reverse()
                                return fromm, sep, "寄送失败!", errorcode.CODE_DICT[str(e)[1:4]][
                                    "msg"], "sendfail", pas, "SMTP SERVER建立成功", processbar_count
                            except:
                                self.logger.warning(f'{fromm}寄给{i}寄送失败，原因:{e}')
                                sep.reverse()
                                return fromm, sep, "寄送失败!", str(e), "sendfail", pas, "SMTP SERVER建立成功", processbar_count
                        else:  # 部分寄失敗
                            try:
                                self.logger.warning(f'{fromm}寄给{i}寄送失败，原因:{errorcode.CODE_DICT[str(e)[1:4]]["msg"]}')
                                sep_sucess = sep[:sep.index(i)]
                                sep_fail = sep[sep.index(i):]
                                sep.reverse()
                                if processbar_count < self.processbar['maximum']:  # 進度條顯示
                                    processbar_count += add_rate
                                    self.logger.info(f'寄件進度{processbar_count}%, sep={sep}')
                                    self.processbar['value'] = processbar_count
                                    self.val.set(f'寄送中...{int(processbar_count)}%')
                                    self.frm.update()
                                    time.sleep(0.01)
                                return fromm, sep_fail, "寄送失败!", errorcode.CODE_DICT[str(e)[1:4]][
                                    "msg"], "sendalittlefail", pas, "SMTP SERVER建立成功", processbar_count, sep_sucess
                            except:
                                self.logger.warning(f'{fromm}寄给{i}寄送失败，原因:{e}')
                                sep_sucess = sep[:sep.index(i)]
                                sep_fail = sep[sep.index(i):]
                                sep.reverse()
                                if processbar_count < self.processbar['maximum']:  # 進度條顯示
                                    processbar_count += add_rate
                                    self.logger.info(f'寄件進度{processbar_count}%, sep={sep}')
                                    self.processbar['value'] = processbar_count
                                    self.val.set(f'寄送中...{int(processbar_count)}%')
                                    self.frm.update()
                                    time.sleep(0.01)
                                return fromm, sep_fail, "寄送失败!", str(
                                    e), "sendalittlefail", pas, "SMTP SERVER建立成功", processbar_count, sep_sucess

                    if processbar_count < self.processbar['maximum']:  # 進度條顯示
                        processbar_count += add_rate
                        self.logger.info(f'寄件進度{processbar_count}%, sep={sep}')
                        self.processbar['value'] = processbar_count
                        self.val.set(f'寄送中...{int(processbar_count)}%')
                        self.frm.update()
                        time.sleep(0.01)

                # 寄送完成
                self.logger.info(f'{fromm}寄给{sep}已寄送，回传{fromm, sep, "已寄送!", "-"}')
                return fromm, sep, "已寄送!", "-", "sucess", pas, "SMTP SERVER建立成功", processbar_count
        except smtplib.SMTPException as e:
            try:
                self.logger.warning(f'{fromm}登入失敗，原因:{errorcode.CODE_DICT[str(e)[1:4]]["msg"]}')
                return fromm, pas, "登入失败!", errorcode.CODE_DICT[str(e)[1:4]][
                    "msg"], "loginfail", "-", 'fail', processbar_count
            except:
                self.logger.warning(f'{fromm}登入失敗，原因:{e}')
                return fromm, pas, "登入失敗!", str(e), "loginfail", "-", 'fail', processbar_count

    # 按下确认后之处理
    def confirm(self):
        if 'xls' not in self.upload_file_load_entry1.get():
            self.logger.warning(f'confirm寄件人弹窗:{self.upload_file_load_entry1.get()}--->寄件人档案路径有误')
            return messagebox.showwarning(title='寄件人档案路径有误', message='请输入正确寄件人档案路径')
        elif 'xls' not in self.upload_file_load_entry2.get():
            self.logger.warning(f'confirm收件人弹窗:{self.upload_file_load_entry2.get()}--->收件人档案路径有误')
            return messagebox.showwarning(title='收件人档案路径有误', message='请输入正确收件人档案路径')
        elif self.subject.get() == '':
            self.logger.warning(f'confirm主旨弹窗:{self.subject.get()}--->主旨为空')
            return messagebox.showwarning(title='主旨', message='请输入主旨')
        elif self.textExample.get(1.0, 'end') == '\n':
            self.logger.warning(f"confirm内文弹窗:{self.textExample.get(1.0, 'end')}--->内文为空")
            return messagebox.showwarning(title='内文', message='请输入内文')
        # 判断完寄、收、主、内后判断附件
        else:
            self.appendix_load = self.appendix.get(1.0, 'end').split('\n')[:-2]  # 看不太懂這行
            self.appendix_load = [i.strip() for i in list(set(self.appendix_load)) if i.strip() != '']

        result = messagebox.askyesno(title='确认一下', message=(f'请确认下列资讯是否正确\n寄件者档案路径:{self.upload_file_load_entry1.get()}'
                                                            f'\n收件者档案路径:{self.upload_file_load_entry2.get()}'
                                                            f'\n收件人上限:{self.most_send.get()}\n寄送頻率:{self.frequency_.get()}'
                                                            f'\n主旨:{self.subject.get()}'
                                                            f'\n夹带附件路径:{self.appendix_load}'))

        if result:
            self.send_button['state'] = tk.DISABLED  # 禁止传送可以用 DISABLED按鈕無法按
            self.val.set(f'寄送中...')
            self.frm.update()
            if self.appendix:  # 如果附件不为空
                appendix_all_size = 0
                for i in self.appendix_load:
                    if not os.path.isfile(i):
                        self.send_button['state'] = tk.NORMAL
                        self.val.set('0%')
                        self.logger.warning(f"confirm附件:{self.appendix_load}--->附件档案路径有误")
                        return messagebox.showwarning(title='附件档案路径有误', message=f'附件路径{i}有误，请确认附件路径')
                    else:
                        file_size = Path(r'%s' % i).stat().st_size
                        self.logger.info(f"confirm附件大小:{i, file_size}")
                        appendix_all_size += file_size
                self.logger.info(f"confirm附件大小总和:{appendix_all_size}")
                if appendix_all_size > 26214400:
                    self.send_button['state'] = tk.NORMAL
                    self.val.set('0%')
                    self.logger.warning(f"confirm附件大小弹窗:总档案大小超过25MB")
                    return messagebox.showerror(title='总档案大小太大', message='总档案大小超过25MB')
            state_list = []  # 设一寄送状态list之后做成excel用
            login_state = []  # 设一登入状态list之后做成excel用
            # 测试收件者档案路径对不对
            try:
                recipient_open = pd.read_excel(self.upload_file_load_entry2.get())
            except:
                self.send_button['state'] = tk.NORMAL
                self.val.set('0%')
                self.logger.warning(f"收件人檔案讀取錯誤:{self.upload_file_load_entry2.get()}--->请输入正确收件人档案路径")
                return messagebox.showerror(title='收件人档案路径有误', message='请输入正确收件人档案路径')
            # 判断收信人格式
            if recipient_open.shape[1] != 1:
                self.logger.warning('收件人档案格式有误，停止寄送')
                self.send_button['state'] = tk.NORMAL
                self.val.set('0%')
                return messagebox.showerror(title='收件人格式有误', message='收件人档案格式有误，请确认')
            else:
                # 将收件者档案的COLUMN重命名并进check_recipient资料处理
                recipient_open.columns = ['邮箱']
                recipient_info = self.check_recipient(recipient_open)  # 寄送名單以及寄送狀態list
                self.logger.info(f'recipient_info: {recipient_info}')
                recipient_list = []
                for i in recipient_info[0]:
                    if i not in recipient_list:
                        recipient_list.append(i)
                recipient_df = recipient_info[1]

            # 执行寄件人寄送
            self.logger.info('开始进行寄送...')
            processbar_count = 0

            if self.getup_load_file_load_sender_entry(
                    r'%s' % self.upload_file_load_entry1.get()) is None:  # 寄件人档案格式有误(抓不到值)，停止寄送时触发
                return
            yeah_buffer = 0  # yeah緩衝計數用
            self.logger.info(f'總寄件人數量:{len(self.getup_load_file_load_sender_entry(r"%s" % self.upload_file_load_entry1.get()))}')
            count_sender = len(self.getup_load_file_load_sender_entry(r"%s" % self.upload_file_load_entry1.get()))
            for senders in self.getup_load_file_load_sender_entry(
                    r'%s' % self.upload_file_load_entry1.get()):  # 先挑第一個寄件者寄信
                self.logger.info(f'寄件人:{senders}, 寄件人登入狀態:{login_state}')
                self.logger.info(f'剩餘寄件人數量:{count_sender} 剩餘收件人數量:{len(recipient_list)}')
                if len(recipient_list) == 0:  # 判断收件人LIST是否为0，因会逐一取出越来越短
                    state_list.append((senders[0], '-', '寄送失败!', '收件人已寄完'))
                    login_state.append((senders[0], senders[1], '無須登入', '收件人已寄完'))
                    self.logger.warning(
                        f'全部收信人已寄送完成!已无收信人!剩馀寄信人:{[i[0] for i in self.sender[self.sender.index(senders) + 1:]]}')

                else:  # 有收件人,寄信
                    sender_name = senders[2] if (len(senders) == 3 and senders[2] != 'nan') else ''
                    self.logger.info(f'寄件人資訊{senders}')
                    dist = {'title': str(self.subject.get()),  # 主旨
                            'sender_name': sender_name,
                            'pas': senders[1],  # 寄件者密碼
                            'sep': (recipient_list, int(self.most_send.get())),  # 收件者清單以及寄件最大數量
                            'html_text': self.textExample.get(1.0, 'end'),  # 寄件內容
                            'fromm': senders[0],  # 寄件者帳號
                            'total_recipient': len(recipient_list),
                            'processbar_count': processbar_count}

                    self.logger.info(f'寄件傳入資料:{dist}')
                    while True:

                        try:
                            send_list = list(self.send(**dist))
                            self.logger.info(f'寄件資料回傳列表:{send_list}')
                            if self.host_cb.get() == 'yeah.net':
                                if yeah_buffer > 99 or (send_list[3] == '收件者 SMTP 主機拒絕提供服務，因為已經超過其能提供的最大服務量。稍後再試'):
                                    self.logger.info('因SMTP主機暫時拒絕提供服務，緩衝處理中...請稍後15分鐘')
                                    self.val.set(f'因SMTP主機暫時拒絕提供服務，緩衝處理中...請稍後15分鐘!{int(processbar_count)}%')
                                    time.sleep(900)
                                    yeah_buffer = 0
                                yeah_buffer += 1
                        except UnicodeEncodeError as e:
                            self.logger.warning(f'寄件人密碼有非英文或數字字元--錯誤代碼:{e}')
                        except Exception as e:
                            self.logger.warning(f'錯誤代碼:{e}')
                            self.logger.warning('連線錯誤10秒後嘗試重連...')
                            self.val.set('发生不明原因正在尝试重新连线,请先检察您的网路')
                            time.sleep(10)
                            continue
                        break
                    processbar_count = send_list[7]
                    # 寄送結果分類
                    if send_list[4] == "loginfail":  # 登入失敗
                        login_state.append(list(send_list[:4]))
                        state_list.append([send_list[0], send_list[5], send_list[2], send_list[3]])
                    if send_list[4] == "sendfail":  # 寄送失敗
                        login_state.append([send_list[0], send_list[5], send_list[6], '-'])
                        state_list.append(list(send_list[:4]))
                        for i in send_list[1]:
                            recipient_df['寄送状态'][str(i.strip())] = '寄送失败'
                        recipient_list = recipient_list + send_list[1]
                    if send_list[4] == "sendalittlefail":  # 部分寄送失敗
                        login_state.append([send_list[0], send_list[5], send_list[6], '-'])
                        state_list.append([send_list[0], send_list[8], '已寄送!', '-'])
                        state_list.append(list(send_list[:4]))
                        for i in send_list[8]:
                            recipient_df['寄送状态'][str(i).strip()] = '已寄送'
                        for i in send_list[1]:
                            recipient_df['寄送状态'][str(i.strip())] = '寄送失败'
                        recipient_list = recipient_list + send_list[1]  # 加回寄送失敗
                    if send_list[4] == "sucess":  # 寄送成功
                        for i in send_list[1]:
                            recipient_df['寄送状态'][str(i).strip()] = '已寄送'
                        login_state.append([send_list[0], send_list[5], send_list[6], '-'])
                        state_list.append(send_list[:4])
                    count_sender -= 1
                    # 進度條

            if int(processbar_count) == self.processbar['maximum']:
                self.val.set(f'所有寄信人已寄送指定数量收信人完成!{int(processbar_count)}%')
                self.frm.update()
                self.send_button['state'] = tk.NORMAL
            elif len(recipient_list) == 0:
                processbar_count = 100
                self.val.set(f'所有寄信人已寄送指定数量收信人完成!{int(processbar_count)}%')
                self.frm.update()
                self.send_button['state'] = tk.NORMAL
            self.logger.info('本次所有寄信人已寄送指定数量收信人完成!')
            if len(recipient_list) != 0:
                self.logger.debug(f"寄件人都已寄送指定收件人，无剩馀寄件人，剩馀收件人:{recipient_list}")
                self.val.set(f'寄件人不足,請新增寄件人或調整寄送上限之後重新發送一次   完成進度:{int(processbar_count)}%')
                self.send_button['state'] = tk.NORMAL
                for i in recipient_list:
                    recipient_df['错误原因'][str(i).strip()] = '寄件人不足'
                    state_list.append(('-', i, '寄送失败!', '寄件人不足'))
            login_df = pd.DataFrame(login_state, columns=['帐号', '密码', '登入状态', '错误原因']).set_index('帐号')
            state_df = pd.DataFrame(state_list, columns=['寄件人', '收件人', '寄送状态', '错误原因']).set_index('寄件人')
            localtime = time.localtime()
            result_time = time.strftime("%Y-%m-%d %p%I.%M.%S ", localtime)
            login_df.to_excel('config/state/寄件人登入状态%s.xlsx' % result_time)
            recipient_df.to_excel('config/state/收件人寄送状态%s.xlsx' % result_time)
            state_df.to_excel('config/state/寄送状态总览%s.xlsx' % result_time)
            self.logger.info('--------------------------------------分隔线---------------------------------------')

    # 多执行绪1号
    def threading(self):
        # Call work function
        t1 = Thread(target=self.confirm)
        t1.start()
