import pandas as pd
from pathlib import Path
def example():
    path = Path('.').absolute()
    if not Path(f'{path}/寄、收件人excel格式范例档').exists():
        Path(f'{path}/寄、收件人excel格式范例档').mkdir()
    df1=pd.DataFrame({'帐号':['123@gmail.com','abc@gmail.com','123abc@gmail.com'],'密码':['asdfqwerzxcvqwer','asdfqwerzxcvqwer','rtyufghjvbnmeidk'], '暱称':['帅哥', '小哥哥', '小姐姐']}).set_index('帐号')
    df2=pd.DataFrame({'帐号':['123@gmail.com','abc@gmail.com','123abc@gmail.com']})
    df1.to_excel(f'{path}/寄、收件人excel格式范例档/寄件人范例档.xlsx')
    df2.to_excel(f'{path}/寄、收件人excel格式范例档/收件人范例档.xlsx',index=False)