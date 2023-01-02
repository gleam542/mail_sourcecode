import tkinter as tk
from tkinter import messagebox
import requests as req
import mail
from pathlib import Path
import pickle
import base64

class Key(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('金鑰驗證')
        self.geometry('700x100')
        self.setupUI()

    
    def setupUI(self):
        row = tk.Frame(self)
        row.pack()
        tk.Label(row, text='請輸入金鑰：').pack(side='left',ipady=50)
        self.keyentry = tk.Entry(row,width=70)
        self.keyentry.pack(side='left',ipadx=20)
        self.keybutton = tk.Button(row, text='確認',command=self.kitapi)
        self.keybutton.pack(side='right',ipadx=10)
    
    def kitapi(self):
        resp=req.post('http://18.163.192.24/emailbatchsend/robot.php',data={'ProductKey':self.keyentry.get()})
        if resp.json()['Result'] !=1:
           messagebox.showerror("错误",str(resp.json()['ErrMsg']))
           return 0
            
        if resp.json()['Result'] ==1 :
            if resp.json()["Used"]==1:
                messagebox.showwarning("警告","此金鑰已被使用，請重新輸入")
                self.destroy()
                Key().mainloop()
                return 0
            else:
                messagebox.showinfo("驗證成功","綁定成功!")
                #生成金鑰檔
                file_name='key.pkl'
                f = open(file_name,'wb')
                pickle.dump(base64.b32encode(self.keyentry.get().encode('utf-8')),f)
                f. close()
                self.destroy()
                mail.Mail()
                return 1

def verify():
    #檢查金鑰檔
    pkl_file = Path('key.pkl')
    if pkl_file.exists():#key存在
        pkl_content=base64.b32decode(pickle.load(open(pkl_file,'rb'))).decode('UTF-8')
        resp=req.post('http://18.163.192.24/emailbatchsend/robot.php',data={'ProductKey':pkl_content})
        if resp.json()['Result']==1:
            if resp.json()['Used']==1:
                mail.Mail()
        else:
            Key().mainloop()
            
    else:#key不存在  
        #註冊金鑰
        Key().mainloop()

if __name__ == '__main__':
    verify()


    
    

        

    