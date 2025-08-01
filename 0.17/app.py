# from tkinter import *
import pandas as pd
import numpy as np
import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
import time
import os
import sys
import requests
import winreg
import platform
import zipfile
import io


def get_edge_version():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Edge\BLBeacon")
        version, _ = winreg.QueryValueEx(key, "version")
        winreg.CloseKey(key)
        return version
    except:
        return None

def download_edge_driver(version=None, destination_folder="Driver"):
    # Create destination folder if it doesn't exist
    os.makedirs(destination_folder, exist_ok=True)
    
    # Get Edge version if not provided
    if not version:
        version = get_edge_version()
        if not version:
            print("Could not determine Edge version. Please specify manually.")
            return None
    
    # Determine platform
    system = platform.system()
    if system == "Windows":
        platform_name = "win64"  # or win32 for 32-bit
    elif system == "Darwin":
        platform_name = "mac64"
    elif system == "Linux":
        platform_name = "linux64"
    else:
        print(f"Unsupported platform: {system}")
        return None
    
    # Microsoft Edge driver download URL
    url = f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/{version}/edgedriver_{platform_name}.zip"
    
    try:
        print(f"Downloading Edge WebDriver version {version}...")
        response = requests.get(url)
        if response.status_code != 200:
            print(f"Failed to download Edge WebDriver. Status code: {response.status_code}")
            return None
    
        # Extract the zip file
        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_ref:
            zip_ref.extractall(destination_folder)
        
        driver_path = os.path.join(destination_folder, "msedgedriver.exe" if system == "Windows" else "msedgedriver")
        print(f"Edge WebDriver downloaded to: {driver_path}")
        return driver_path
    
    except Exception as e:
        print(f"Error downloading Edge WebDriver: {e}")
        return None

# Download Edge WebDriver
driver_path = download_edge_driver(destination_folder="Driver")

try:
    import cryptography
except:
    os.system("pip install cryptography")
from cryptography.fernet import Fernet
try:
    import jinja2
except:
    os.system("pip install jinja2")


if 'pass_file' not in os.listdir('./'):
    pas = Tk()
    pas.geometry('300x200')
    pas.config(background='#90959e')
    pas.attributes('-topmost',True)
    pas.overrideredirect(True)
    pas.eval('tk::PlaceWindow . center')

    password = StringVar(pas)
    def save_password():
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

with open("pass_file",'r') as file:
    r=file.readlines()
fernet = Fernet(r[0].strip())
decMessage = fernet.decrypt(r[1]).decode()
passw = decMessage



daytime_activities = ["TRX GROW",\
                        "De grow Dummy cells deletion",\
                        "Blocking",\
                        "Name change activities",\
                        "Site swap",\
                        "Site integration",\
                        "POP Migration",\
                        "HOP migration",\
                        "shifting TN sw downloading Microwave swap",\
                        "Mmu migration",\
                        "MMU SWAP",\
                        "New POI AT TDM/IP",\
                        "Expansion of BICC/CMN/SIP/Session Limit Increase in existing POI",\
                        "TG Creation",\
                        "RAN swap related activities",\
                        "CGI Creation",\
                        "MRO Testing as per TRAI/DOT",\
                        "Code Opening activities",\
                        "routing changes DoT compliance Codes",\
                        "New Level Opening",\
                        "Roaming Launch Configuration Code-except GT changes",\
                        "Short Code Opening",\
                        "Emergency Number Routing Changes",\
                        "Traffic change on POI",\
                        "Barring of specific B No to handle fraudulent  cases",\
                        "RA reconciliation cases"\
                        "Revenue Leakage cases",\
                        "CLI Whitelisting/Blacklisting",\
                        "Test Routing",\
                        "Regulatory Compliance",\
                        "Any Emergency number",\
                        "service routing change received as part of Regulatory compliance",\
                        "LB delta correction",\
                        "SIMBOX",\
                        "NIM",\
                        "Radii",\
                        "Short code",\
                        "B Number",\
                        "B No",\
                        "Toll Free"]


user = os.getenv('username')
auth_code_string = ''

options = webdriver.EdgeOptions()
prefs={"download.default_directory":"C:\\CM Automation"}
options.add_experimental_option("prefs",prefs)
service=Service(executable_path=driver_path)
driver = webdriver.Edge(options=options, service=service)
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
    time.sleep(5)
    alert = Alert(driver)
    try:
        alert.accept()
    except:
        pass
Label(app,text="Enter Microsoft Authentication Code",font=('Ericsson Hilda',15,'bold'),background='#709179').place(x=145,y=80)
auth_entry = Entry(app,font=('Ericsson Hilda',15),justify='center',textvariable=code)
auth_entry.place(x=190,y=130)
Button(app,text='Login',font=('Ericsson Hilda',15),width=20,command=login).place(x=190,y=190)
app.bind('<Return>',login)
app.after(1, lambda: auth_entry.focus_force())
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

# driver.close()
# driver.switch_to.window(driver.window_handles[0])
driver.switch_to.default_content()
time.sleep(1)
driver.find_element(By.XPATH, '//a[contains(@class,"ardbnbtnLogout")]').click()

# driver.find_element(By.XPATH,'/html/body/div[1]/div[5]/div[2]/div/div/div[1]/fieldset/div/div[2]/fieldset/a[5]').click()
time.sleep(5)
# driver.close()
driver.quit()




#....................................................................
cm_data = pd.read_excel('./CM_data.xlsx')
dic_bw = dict(zip(cm_data['Signum'],cm_data['BW']))
list_CM = list(dic_bw.keys())
import itertools

import glob
import os.path

folder_path = 'C:\\CM Automation'
file_type = '\\*'
files = glob.glob(folder_path + file_type)
file = max(files, key=os.path.getctime)
print(file)

df = pd.read_excel(file)
df.loc[:,"Activity Duration Remarks"] = "NOK"
df.loc[(((pd.to_datetime(df['Scheduled End Date+'], format="%m/%d/%Y %I:%M:%S %p") - pd.to_datetime(df['Scheduled Start Date+'], format="%m/%d/%Y %I:%M:%S %p")).dt.seconds/60/60) <= 24), "Activity Duration Remarks"] = 'OK'
df.loc[:,"Category T1 Remarks"] = "OK"
df.loc[:,"Category T2 Remarks"] = "OK"
df.loc[:,"Category T3 Remarks"] = "OK"
df.loc[df['Operational Categorization Tier 1+'].isna(),'Category T1 Remarks'] = 'NOK'
df.loc[df['Operational Categorization Tier 2'].isna(),'Category T2 Remarks'] = 'NOK'
df.loc[df['Operational Categorization Tier 3'].isna(),'Category T3 Remarks'] = 'NOK'
df.loc[((df['Operational Categorization Tier 1+']=='VAS') & (df['Operational Categorization Tier 3'].isna())), 'Category T3 Remarks'] = 'OK'
df.loc[((df['Operational Categorization Tier 1+']=='Tech Lan') & (df['Operational Categorization Tier 3'].isna())), 'Category T3 Remarks'] = 'OK'
total_cr_count = len(df)

mor6 = datetime.time(hour=6)
nig23 = datetime.time(hour=23)
def check_day_act(summary):
    for activity in daytime_activities:
        if activity.lower() in summary.lower():
            return True
    return False
d = df.copy()
d['Scheduled Start Date+'] = pd.to_datetime(d['Scheduled Start Date+'],format="%m/%d/%Y %I:%M:%S %p")
d['Scheduled End Date+'] = pd.to_datetime(d['Scheduled End Date+'],format="%m/%d/%Y %I:%M:%S %p")
df.loc[:, "Day Activity Remarks"] = "NOK"
df.loc[((d['Scheduled End Date+'].dt.hour <= 6) & ((d['Scheduled Start Date+'].dt.hour >= 23) | (d['Scheduled Start Date+'].dt.hour <= 6))), "Day Activity Remarks"] = "Night Activity"
df.loc[((d['Summary*'].apply(check_day_act)) & ~((d['Scheduled End Date+'].dt.hour <= 6) & ((d['Scheduled Start Date+'].dt.hour >= 23) | (d['Scheduled Start Date+'].dt.hour <= 6)))), "Day Activity Remarks"] = "OK"



try:
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
except:
    previous_data = pd.DataFrame(columns=['Signum'])
    not_completed = pd.DataFrame(columns=['Signum'])
df['Scheduled Start Date+'] = pd.to_datetime(df['Scheduled Start Date+'],format="%m/%d/%Y %I:%M:%S %p")
df['Scheduled End Date+'] = pd.to_datetime(df['Scheduled End Date+'],format="%m/%d/%Y %I:%M:%S %p")
df = df.sort_values(by=['Scheduled Start Date+'])

now = datetime.datetime.now()
today_9 = datetime.datetime(year=datetime.datetime.now().year,month=datetime.datetime.now().month,day=datetime.datetime.now().day,hour=20,minute=59,second=59)
tomorrow = datetime.datetime.now()+datetime.timedelta(days=1)
tomorrow_9 = datetime.datetime(year=tomorrow.year,month=tomorrow.month,day=tomorrow.day,hour=20,minute=59,second=59)

raw_data = pd.read_excel(file)

expired = df[(df['Scheduled End Date+'] < now)].copy()
expired['Scheduled Start Date+'] = expired['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")
expired['Scheduled End Date+'] = expired['Scheduled End Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

df = df[~(df['Scheduled End Date+'] < now)].copy()
df['Scheduled End Date+'] = df['Scheduled End Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

daytime =  df[(df['Scheduled Start Date+'] <= today_9)].copy()
daytime['Scheduled Start Date+'] = daytime['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

tonight_l1l2 = df[((df['Scheduled Start Date+'] > today_9) & (df['Scheduled Start Date+'] <= tomorrow_9)) & ((df['Impact*'] == '1-Extensive/Widespread') | (df['Impact*'] == '2-Significant/Large'))].copy()
tonight_l1l2['Scheduled Start Date+'] = tonight_l1l2['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

tonight_l3l4 = df[((df['Scheduled Start Date+'] > today_9) & (df['Scheduled Start Date+'] <= tomorrow_9)) & ((df['Impact*'] == '3-Moderate/Limited') | (df['Impact*'] == '4-Minor/Localized'))].copy()
tonight_l3l4['Scheduled Start Date+'] = tonight_l3l4['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

future_l1l2 = df[(df['Scheduled Start Date+'] > tomorrow_9) & ((df['Impact*'] == '1-Extensive/Widespread') | (df['Impact*'] == '2-Significant/Large'))].copy()
future_l1l2['Scheduled Start Date+'] = future_l1l2['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

future_l3l4 = df[(df['Scheduled Start Date+'] > tomorrow_9) & ((df['Impact*'] == '3-Moderate/Limited') | (df['Impact*'] == '4-Minor/Localized'))].copy()
future_l3l4['Scheduled Start Date+'] = future_l3l4['Scheduled Start Date+'].dt.strftime("%m/%d/%Y %I:%M:%S %p")

current_cm = ''
aval_cms = []
for key in dic_bw.keys():
    if dic_bw[key] > 0:
        aval_cms.append(key)
iter_CM1 = itertools.cycle(aval_cms)
sheets = [expired,daytime,tonight_l1l2]
for sheet in sheets:
    sheet.sort_values(by=['Scheduled Start Date+'],inplace=True)
    sheet['Signum'] = ''
    first_col = sheet.pop('Signum')
    sheet.insert(0,'Signum',first_col)
    for idx in sheet.index:
        cm = next(iter_CM1)
        current_cm = cm
        sheet.loc[idx,'Signum'] = cm
        dic_bw[cm] -=1
    sheet.sort_values(by=['Signum'],inplace=True)

for key in dic_bw.keys():
    if dic_bw[key]<=0:
        dic_bw[key] = 0


if current_cm == '':
    iter_CM2 = itertools.islice(itertools.cycle(dic_bw.keys()),(list(dic_bw.keys()).index(aval_cms[0])),None)
else:
    iter_CM2 = itertools.islice(itertools.cycle(dic_bw.keys()),(list(dic_bw.keys()).index(current_cm)+1),None)


sheets = [tonight_l3l4,future_l1l2,future_l3l4]
for sheet in sheets:
    sheet.sort_values(by=['Scheduled Start Date+'],inplace=True)
    sheet['Signum'] = ''
    first_col = sheet.pop('Signum')
    sheet.insert(0,'Signum',first_col)
    for idx in sheet.index:
        if any(dic_bw.values()):
            cm = next(iter_CM2)
            while dic_bw[cm] <= 0:
                cm = next(iter_CM2)
            sheet.loc[idx,'Signum'] = cm
            dic_bw[cm] -=1
        else:
            try:
                sheet.loc[idx,'Signum'] = np.NaN
            except:
                sheet.loc[idx,'Signum'] = np.nan

    sheet.dropna(subset='Signum',inplace=True)
    sheet.sort_values(by=['Signum'],inplace=True)
previous_data = pd.concat([expired,daytime,tonight_l1l2,tonight_l3l4,future_l1l2,future_l3l4])
previous_data.to_excel('Previous_data.xlsx',index=False)


sheets = [expired,daytime,tonight_l1l2,tonight_l3l4,future_l1l2,future_l3l4]
total_given = 0
for s in sheets:
    total_given += len(s)


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



dashboard = pd.DataFrame(columns=['Change Manager','Assigned','Not Completed'])
dashboard['Change Manager'] = list_CM
for cm in list_CM:
    assigned_count = len(previous_data[previous_data['Signum']==cm])
    not_completed_count = len(not_completed[not_completed['Signum']==cm])
    dashboard.loc[dashboard['Change Manager']==cm, 'Assigned'] = assigned_count
    dashboard.loc[dashboard['Change Manager']==cm, 'Not Completed'] = not_completed_count
def red(cell_value):
    if cell_value != 0:
        return "background-color: #e84f4f; color: #000000"
    
def green(cell_value):
    if cell_value != 0:
        return "background-color: #d9e89e; color: #000000"
    
html = dashboard.style\
    .hide(axis='index')\
        .apply_index(lambda c:["background-color: #5ac4a5; color: #000000" for s in c],axis='columns')\
            .applymap(red, subset=['Not Completed'])\
                .applymap(green, subset=['Assigned'])\
                    .to_html()
html = ''.join(html.split('\n'))



import win32com.client as win32
to = 'PDLLEASEMA@pdl.internal.ericsson.com'
sub = f'Slot {slot} Data'

body = f"""<p>Hi,<br />Please find the attached file for Slot {slot} data. Pl check all tabs.<br /><span style="color: #e67e23;"><strong>{total_given}</strong> out of <strong>{total_cr_count}</strong> CRs are allocated for this slot.</span></p>
<p>{html}</p>
<p>Regards,<br />Change Management Team</p>"""

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = " ; ".join(to.split(';'))
mail.Subject = sub
mail.HTMLBody = body
mail.Attachments.Add(slot_filename)
mail.Display()
sys.exit()
