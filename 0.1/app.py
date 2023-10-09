import pandas as pd
import numpy as np
import datetime
# from subprocess import CREATE_NO_WINDOW
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


with open('CR_Data','r') as cr_old:
    cr_old = cr_old.readlines()
for _ in range(len(cr_old)):
    cr_old[_] = cr_old[_].strip()
old_cr = list(pd.Series(cr_old,name='Change ID*+').drop_duplicates())

cr_old = []

user = os.getenv('username')
# passw = os.getenv('pass_word')
passw = ''
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
code = StringVar()
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
print(dic_bw)

import glob
import os.path

folder_path = 'C:\CM Automation'
file_type = '\*'
files = glob.glob(folder_path + file_type)
file = max(files, key=os.path.getctime)
print(file)

df = pd.read_excel(file)
df['Scheduled Start Date+'] = pd.to_datetime(df['Scheduled Start Date+'],format="%m/%d/%Y %I:%M:%S %p")
df = df.sort_values(by=['Scheduled Start Date+'])
import itertools
iter_CM = itertools.cycle(list_CM)
today_9 = datetime.datetime(year=datetime.datetime.now().year,month=datetime.datetime.now().month,day=datetime.datetime.now().day,hour=20,minute=59,second=59)
tomorrow = datetime.datetime.now()+datetime.timedelta(days=1)
tomorrow_9 = datetime.datetime(year=tomorrow.year,month=tomorrow.month,day=tomorrow.day,hour=20,minute=59,second=59)

raw_data = df.loc[:]
raw_data['Scheduled Start Date+'] = raw_data['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

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


duplicate_df = pd.DataFrame()
sheets = [daytime,tonight_l1l2,tonight_l3l4,future_l1l2,future_l3l4]
for sheet in sheets:
    sheet['Signum'] = ''
    first_col = sheet.pop('Signum')
    sheet.insert(0,'Signum',first_col)
    for idx in sheet.index:
        if any(dic_bw.values()):
            cm = next(iter_CM)
            while dic_bw[cm] <= 0:
                cm = next(iter_CM)
            # print(_+1,cm)
            sheet.loc[idx,'Signum'] = cm
            cr_old.append(sheet.loc[idx,"Change ID*+"])
            if (sheet.loc[idx,"Change ID*+"]) in old_cr:
                duplicate_df = pd.concat([duplicate_df,sheet.loc[[idx]]])
            dic_bw[cm] -=1
        else:
            sheet.loc[idx,'Signum'] = np.NaN
    sheet.dropna(subset='Signum',inplace=True)
    sheet = sheet.sort_values(by=['Signum'])


cr_old = '\n'.join(cr_old)
with open('CR_Data','w') as cr_data:
    cr_data.writelines(cr_old)


#..................slot.........................
import datetime
today_date = str(datetime.datetime.today().day)
with open('slot_no','r') as slot_no:
    slot_data = slot_no.readline()
slot_data = slot_data.split()
for _ in range(len(slot_data)):
    slot_data[_] = slot_data[_].strip()

slot = ''
if today_date == slot_data[0]:
    slot = str(int(slot_data[1])+1)
else:
    slot = str(1)

with open('slot_no','w') as slot_no:
    slot_no.write(f'{today_date} {slot}')



slot_filename = f'C:/CM Automation/Slot_{slot}_{datetime.datetime.today().strftime("%d%m%Y")}.xlsx'
with pd.ExcelWriter(slot_filename) as writer:
    raw_data.to_excel(writer,sheet_name='Raw data',index=False)
    daytime.to_excel(writer,sheet_name='Daytime',index=False)
    tonight_l1l2.to_excel(writer,sheet_name='Tonight L1, L2',index=False)
    tonight_l3l4.to_excel(writer,sheet_name="Tonight L3, L4",index=False)
    future_l1l2.to_excel(writer,sheet_name='Future L1, L2',index=False)
    future_l3l4.to_excel(writer,sheet_name="Future L3, L4",index=False)
    duplicate_df.to_excel(writer,sheet_name="Not Completed",index=False)
    # pd.merge(raw_data.drop_duplicates(),old_cr,on='Change ID*+',how='inner').to_excel(writer,sheet_name="Not Completed",index=False)


time.sleep(3)
import win32com.client as win32
to = 'nishkam.kamal@ericsson.com; shivom.tiwari@ericsson.com'
sub = f'Slot {slot} Data'

body = f"Hi,<br>Please find the attached file for Slot {slot} data<br><br><p>Regards,<br>Abhishek Chakraborty"

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = " ; ".join(to.split(';'))
mail.Subject = sub
mail.HTMLBody = body
mail.Attachments.Add(slot_filename)
mail.Display()
sys.exit()
