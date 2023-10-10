from tkinter import *
import pandas as pd
import numpy as np
import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
import time
import os
import sys
this_version = 0.11
from tkinter import Tk,Button,Frame,Label
import requests, threading, time
pop = Tk()
pop.geometry("300x200")
pop.attributes('-topmost',True)
pop.eval('tk::PlaceWindow . center')
pop.overrideredirect(True)


def te():
    global avl_version
    time.sleep(1)
    avl_version = (requests.get(url='https://raw.githubusercontent.com/Avi2810/test/main/version.txt').text).strip()
    if float(avl_version) > this_version:
        f2.pack_forget()
        f3.pack(fill='both',expand=True,padx=10,pady=10)
    else:
        pop.destroy()

def update_app():
    pop.geometry("300x100")
    pop.config(background='#404040')
    f1.pack_forget()
    f4.pack(padx=5,pady=5,fill='both',expand=True)
    import requests
    import time
    time.sleep(2)
    version = (requests.get(url='https://raw.githubusercontent.com/Avi2810/test/main/version.txt').text).strip()
    with open(f"app_v{version}.py",'w') as updated_file:
        updated_file.write(requests.get(url=f'https://raw.githubusercontent.com/Avi2810/test/main/{version}/app.py').text)
    with open(f"app.bat",'w') as bat_file:
        bat_file.write(f"python app_v{version}.py")
    label.config(text='Updates Installed Successfully')
    time.sleep(3)
    pop.destroy()
    os._exit(1)

f1 = Frame(pop,background='#ccc4a7')
f1.pack(fill='both',expand=True)
f2 = Frame(f1,background='#d9d9d9')
f2.pack(fill='both',expand=True,padx=10,pady=10)
Label(f2,text="Checking for Updates...",font=('Ericsson Hilda',15,'bold'),background='#d9d9d9',foreground='#4a473d').place(relx=0.5, rely=0.5,anchor=CENTER)

f3 = Frame(f1,background='#d9d9d9')
# f3.pack(fill='both',expand=True,padx=10,pady=10)
Label(f3,text='New Update Available.\nDo you want to update?',font=('Ericsson Hilda',15,'bold'),background='#d9d9d9',foreground='#4a473d').pack(pady=20)
Button(f3,text='Yes',font=(None,12),width=9,command=threading.Thread(target=update_app).start).pack(side='left',padx=25)
Button(f3,text='No',font=(None,12),width=9,command=pop.destroy).pack(side='right',padx=25)


f4 = Frame(pop)
# f4.pack(padx=5,pady=5,fill='both',expand=True)
label = Label(f4,text='Installing Updates. \nPlease wait...',font=('Ericsson Hilda',15,'bold'))
label.pack(pady=20)

pop.after_idle(threading.Thread(target=te).start)
pop.mainloop()






import os
from tkinter import *
if 'pass_file' not in os.listdir('./'):
    pas = Tk()
    pas.geometry('300x200')
    pas.config(background='#90959e')
    pas.attributes('-topmost',True)
    pas.overrideredirect(True)
    pas.eval('tk::PlaceWindow . center')

    password = StringVar(pas)
    def save_password():
        from cryptography.fernet import Fernet
        key = Fernet.generate_key()
        fernet = Fernet(key)
        encMessage = fernet.encrypt(password.get().encode())
        with open("pass_file",'wb') as file:
            file.write((key+b'\n'+encMessage))
        pas.destroy()

    Label(pas,text='Enter Your Password',font=('Ericsson Hilda',15,'bold'),background='#90959e').pack(pady=20)
    Entry(pas,textvariable=password,font=('Ericsson Hilda',12),justify='center').pack(pady=10)
    Button(pas,text="Save",font=('Ericsson Hilda',12),width=20,command=save_password).pack(pady=20)
    pas.mainloop()
from cryptography.fernet import Fernet
with open("pass_file",'r') as file:
    r=file.readlines()
    print(r)
fernet = Fernet(r[0].strip())
decMessage = fernet.decrypt(r[1]).decode()
passw = decMessage
print(passw)


with open('CR_Data','r') as cr_old:
    cr_old = cr_old.readlines()
for _ in range(len(cr_old)):
    cr_old[_] = cr_old[_].strip()
old_cr = list(pd.Series(cr_old,name='Change ID*+').drop_duplicates())

cr_old = []

user = os.getenv('username')
auth_code_string = ''


options = webdriver.EdgeOptions()
prefs={"download.default_directory":"C:\CM Automation"}
options.add_experimental_option("prefs",prefs)
# options.add_argument('headless')
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
service=Service(EdgeChromiumDriverManager().install())
# service.creation_flags = CREATE_NO_WINDOW
driver = webdriver.Chrome(service=service,options=options)
driver.maximize_window()
time.sleep(1)
driver.get('https://nextgentm-in.sdt.ericsson.net/arsys/forms/umt-ars-in/SHR%3ALandingConsole/Default+Administrator+View/?cacheid=c4ed3626')



from tkinter import Tk, Button, Entry, Label, StringVar

app = Tk()
app.geometry('600x300')
app.config(background='#709179')
app.attributes('-topmost',True)
code = StringVar(app)
def login(x=None):
    global auth_code_string
    auth_code_string = code.get()
    print(auth_code_string)
    app.destroy()

    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'domain')))
    domain_bar = driver.find_element(By.ID,'domain')
    domain_bar.send_keys('Emlpyee')
    login_btn = driver.find_element(By.ID,'loginBtn')
    login_btn.click()

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'login')))
    login = driver.find_element(By.ID,'login')
    login.send_keys(user)
    passwd = driver.find_element(By.ID,'passwd')
    passwd.send_keys(passw)
    login_btn = driver.find_element(By.ID,'loginBtn')
    login_btn.click()

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'response')))
    aut_response = driver.find_element(By.ID,'response')
    aut_response.send_keys(auth_code_string)
    submit_btn = driver.find_element(By.ID,'ns-dialogue-submit')
    submit_btn.click()
    time.sleep(3)
    alert = Alert(driver)
    try:
        alert.accept()
    except:
        pass
Label(app,text="Enter Microsoft Authentication Code",font=('Ericsson Hilda',15,'bold'),background='#709179').place(x=145,y=80)
Entry(app,font=('Ericsson Hilda',15),justify='center',textvariable=code).place(x=190,y=130)
Button(app,text='Login',font=('Ericsson Hilda',15),width=20,command=login).place(x=190,y=190)
app.bind('<Return>',login)
app.mainloop()



time.sleep(5)
while True:
    time.sleep(1)
    try:
        if driver.find_element(By.LINK_TEXT,'IT Home'):
            break
    except:
        pass


# driver.refresh()

driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[1]/fieldset/div/div[2]/fieldset/a[1]').click()
time.sleep(1)
driver.find_element(By.LINK_TEXT,'Change Management').click()
driver.find_element(By.LINK_TEXT,'Search Change').click()
time.sleep(5)

driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/div[4]/fieldset/div[1]/a').click()
time.sleep(.5)
driver.find_element(By.XPATH,'//table/tbody/tr/td[text()="Request For Change"]').click()
driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a').click()

while True:
    if driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/div[1]/fieldset/div[1]/textarea').get_attribute('value') != '':
        break

driver.find_element(By.LINK_TEXT,'Select All').click()
time.sleep(1)
driver.find_element(By.LINK_TEXT,'Report').click()
time.sleep(5)

driver.switch_to.window(driver.window_handles[1])

driver.find_element(By.XPATH,'//span[text()="Slot_CM1"]').click()
time.sleep(1)
driver.find_element(By.XPATH,'//div[text()="Run"]').click()

driver.switch_to.default_content()
driver.switch_to.frame(driver.find_elements(By.TAG_NAME,'iframe')[1])
while True:
    try:
        if 'none' not in driver.find_element(By.ID,'progressBar').get_attribute('style'):
            break
    except:
        pass

driver.switch_to.default_content()
driver.switch_to.frame(driver.find_elements(By.TAG_NAME,'iframe')[1])
while True:
    try:
        if 'none' in driver.find_element(By.ID,'progressBar').get_attribute('style'):
            break
    except:
        pass

time.sleep(2)
driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[5]/input').click()
driver.find_element(By.XPATH,'//*[@id="exportFormat"]').send_keys('Spudsoft Excel')
driver.find_element(By.XPATH,'//*[@id="exportReportDialogokButton"]/input').click()
time.sleep(5)

driver.close()
driver.switch_to.window(driver.window_handles[0])

driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[1]/fieldset/div/div[2]/fieldset/a[5]').click()
time.sleep(3)
driver.close()
driver.quit()




#....................................................................
cm_data = pd.read_excel('CM_data.xlsx')
dic_bw = dict(zip(cm_data['Signum'],cm_data['BW']))
list_CM = list(dic_bw.keys())
import itertools
iter_CM = itertools.cycle(list_CM)

import glob
import os.path

folder_path = 'C:\CM Automation'
file_type = '\*'
files = glob.glob(folder_path + file_type)
file = max(files, key=os.path.getctime)
print(file)

df = pd.read_excel(file)
previous_data = pd.read_excel('Previous_data.xlsx')
not_completed = pd.merge(previous_data,df['Change ID*+'],on='Change ID*+',how='inner')
outer = df.merge(not_completed['Change ID*+'], how='outer', indicator=True)
anti_join = outer[(outer._merge=='left_only')].drop('_merge', axis=1)
df = anti_join
zero = []
for key in dic_bw.keys():
    if dic_bw[key] == 0:
        zero.append(key)
crs = pd.merge(not_completed,pd.DataFrame(zero,columns=['Signum']),on='Signum',how='inner')['Change ID*+']
df = pd.concat([df, pd.read_excel(file).merge(crs,how='inner',on='Change ID*+')])
df['Scheduled Start Date+'] = pd.to_datetime(df['Scheduled Start Date+'],format="%m/%d/%Y %I:%M:%S %p")
df = df.sort_values(by=['Scheduled Start Date+'])

now = datetime.datetime.now()
today_9 = datetime.datetime(year=datetime.datetime.now().year,month=datetime.datetime.now().month,day=datetime.datetime.now().day,hour=20,minute=59,second=59)
tomorrow = datetime.datetime.now()+datetime.timedelta(days=1)
tomorrow_9 = datetime.datetime(year=tomorrow.year,month=tomorrow.month,day=tomorrow.day,hour=20,minute=59,second=59)

raw_data = df.loc[:]
raw_data['Scheduled Start Date+'] = raw_data['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

expired = df[(df['Scheduled Start Date+'] < now)]
expired['Scheduled Start Date+'] = expired['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

df = df[~(df['Scheduled Start Date+'] < now)]

daytime =  df[(df['Scheduled Start Date+'] <= today_9)]
daytime['Scheduled Start Date+'] = daytime['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

tonight_l1l2 = df[((df['Scheduled Start Date+'] > today_9) & (df['Scheduled Start Date+'] <= tomorrow_9)) & ((df['Impact*'] == '1-Extensive/Widespread') | (df['Impact*'] == '2-Significant/Large'))]
tonight_l1l2['Scheduled Start Date+'] = tonight_l1l2['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

tonight_l3l4 = df[((df['Scheduled Start Date+'] > today_9) & (df['Scheduled Start Date+'] <= tomorrow_9)) & ((df['Impact*'] == '3-Moderate/Limited') | (df['Impact*'] == '4-Minor/Localized'))]
tonight_l3l4['Scheduled Start Date+'] = tonight_l3l4['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

future_l1l2 = df[(df['Scheduled Start Date+'] > tomorrow_9) & ((df['Impact*'] == '1-Extensive/Widespread') | (df['Impact*'] == '2-Significant/Large'))]
future_l1l2['Scheduled Start Date+'] = future_l1l2['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

future_l3l4 = df[(df['Scheduled Start Date+'] > tomorrow_9) & ((df['Impact*'] == '3-Moderate/Limited') | (df['Impact*'] == '4-Minor/Localized'))]
future_l3l4['Scheduled Start Date+'] = future_l3l4['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")


sheets = [expired,daytime,tonight_l1l2,tonight_l3l4,future_l1l2,future_l3l4]
for sheet in sheets:
    sheet['Signum'] = ''
    first_col = sheet.pop('Signum')
    sheet.insert(0,'Signum',first_col)
    for idx in sheet.index:
        if any(dic_bw.values()):
            cm = next(iter_CM)
            while dic_bw[cm] <= 0:
                cm = next(iter_CM)
            sheet.loc[idx,'Signum'] = cm
            dic_bw[cm] -=1
        else:
            sheet.loc[idx,'Signum'] = np.NaN
    sheet.dropna(subset='Signum',inplace=True)
    sheet = sheet.sort_values(by=['Signum'])
previous_data = pd.concat(sheets)
previous_data.to_excel('Previous_data.xlsx',index=False)


#..................slot.........................
import datetime
slot_dict = {7: 1, 8: 1, 9: 2, 10: 3, 11: 4, 12: 5, 13: 6, 14: 7, 15: 7, 16: 8, 17: 9, 18: 9, 19: 10, 20: 10, 21: 11, 22: 12, 23: 13, 0: 14, 1: 15, 2: 15, 3: 15, 4: 15, 5: 15, 6: 15}
now_hrs = datetime.datetime.now().hour
slot = slot_dict[now_hrs]


slot_filename = f'C:/CM Automation/Slot_{slot}_{datetime.datetime.today().strftime("%d%m%Y")}.xlsx'
with pd.ExcelWriter(slot_filename) as writer:
    raw_data.to_excel(writer,sheet_name='Raw data',index=False)
    expired.to_excel(writer,sheet_name='Expired',index=False)
    daytime.to_excel(writer,sheet_name='Daytime',index=False)
    tonight_l1l2.to_excel(writer,sheet_name='Tonight L1, L2',index=False)
    tonight_l3l4.to_excel(writer,sheet_name="Tonight L3, L4",index=False)
    future_l1l2.to_excel(writer,sheet_name='Future L1, L2',index=False)
    future_l3l4.to_excel(writer,sheet_name="Future L3, L4",index=False)
    not_completed.to_excel(writer,sheet_name='Not Completed',index=False)



time.sleep(3)
import win32com.client as win32
to = 'PDLLEASEMA@pdl.internal.ericsson.com'
sub = f'Slot {slot} Data'

body = f"Hi,<br>Please find the attached file for Slot {slot} data<br><br><p>Regards,<br>Change Management Team"

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = " ; ".join(to.split(';'))
mail.Subject = sub
mail.HTMLBody = body
mail.Attachments.Add(slot_filename)
mail.Display()
sys.exit()
