#==================Libraries====================
'''
Libraries for web scrapping the data from the PPA website
'''
from selenium import webdriver
import chromedriver_autoinstaller

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time #Library for creating delays

'''
Libraries for data manipulation
'''
import pandas as pd
import numpy as np
import re

'''
Libraries used to send the emails
'''
from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

#==================Collecting the Data#==================

# Check if the current version of chromedriver exists
# and if it doesn't exist, download it automatically,
# then add chromedriver to path
chromedriver_autoinstaller.install()

#Desabling some of the options
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
wd = webdriver.Chrome(options=options)

wd.get("https://agent.kleinsystems.com/") #Open the website

print('Open Website')

#Xpaths of the forms for log in
username_xpath = '/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]/input'
password_xpath = '/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input'

#Data for log in
username = 'Tiger'
password = 'rivtow29'

#Log in
'''
OBS: I do not know why, but when acessing the website by selenium,
you need to log in two times.
'''
for i in range(2):
    #Switch to frame main
    wd.switch_to.frame("main")
    
    #Filling the form
    wd.find_element(By.XPATH, username_xpath).send_keys(username)
    wd.find_element(By.XPATH, password_xpath).send_keys(password)

    #Pressing the log-in
    wd.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table/tbody/tr[3]/td/input').click()
    
    time.sleep(2) # Sleep for 2 seconds

    #Switch to main frame again
    wd.switch_to.default_content()


#Change for the right frame
wd.switch_to.frame("main")

#Go to the tab that has our data
wd.find_element(By.XPATH, '//*[@id="leftMenu__ctl0_TrafficTD"]').click()

time.sleep(1.5) # Sleep for 3 seconds

#Downloading the table
wd.find_element(By.XPATH, '//*[@id="_ctl4_btnExportToExcel"]').click()

time.sleep(0.5) # Sleep for 3 seconds

print('Downloaded the Table')

#Read the data that we collected
data = pd.read_excel('Agent Current Traffic Report.xls')

print('Read the Table')

#==================Cleaning and subsetting the Data==================

#Function used to filter out the Vessels Operating in Prince Rupert
def pRupertShips(data):
  '''
  Based in a dataframe, it returns the index of the rows that are related to Prince Rupert
  '''
  import re

  r = re.compile('FAIRVIEW CONTAINER|WESTVIEW|PRINCE RUPERT|RIDLEY COAL|ALTAGAS|PEMBINA|TRIPLE ISLAND')
  vmatch = np.vectorize(lambda x:bool(r.search(x)))

  #Returns a list with trues or falses, it checks value by value
  index_from = vmatch(data['From'].values)
  index_to = vmatch(data['To'].values)

  index = []
  for i in range(len(index_from)):
    index.append(index_from[i]==True or index_to[i]==True)

  return index

#Organizing the column names
cols = list(data.columns)
cols[1] = 'Job PO'
cols[26: 28]= ['Hellicopter from', 'Hellicopter to']
data.columns = cols

#Sorting by the order time
data = data.sort_values(by='Order Time')

#Acessing the file with the name of the important vessels
with open('vesselNames.txt') as f:
    rawNames = f.read()

#Creating a list
ships = rawNames.split('\n')

#Concatenating the rows that have vessels that are in our text file
newShips = pd.DataFrame()
for ship in ships:
  row = data[data['Vessel Name'] == ship]
  newShips = pd.concat([newShips,row])

#Subsetting the columns that we want from the data
subsetting = newShips.sort_values(by='Order Time')[['Job #', 'Vessel Name','Order Time', 'From', 'To', 'Agency','Tug From', 'Tug To', 'Vessel Type/Dock']]

#Creating a new colum where we will categorize the rows
subsetting['Observation'] = ['None']*subsetting.shape[0]

#Filling the rows from Prince Rupert
#Add Prince Rupert in the Observation column
index = pRupertShips(subsetting)
subsetting.loc[index, 'Observation'] = 'Prince Rupert'

#Checking if was already given to us
tug_from = subsetting['Tug From'].values
tug_to = subsetting['Tug To'].values
index = []
for i in range(subsetting.shape[0]):
    #It appends the indexes that satisfies the condition into a list
    #That we will use to subset the data that was already assigned to us
  index.append(tug_from[i] == 'Saam Towage' or tug_to[i] == 'Saam Towage')

#Add SAAM Towage in the Observation column
subsetting.loc[index, 'Observation'] = 'Saam Towage'

#Create the file with the Vessels that we are following to send to the email
subsetting.to_excel('vesselsNames.xlsx')

#Subset the vessels that we are looking for that have not been 
obsVessels = subsetting[subsetting['Observation'] == 'None']

#=================Creating a Conditional Formatting for the Table===================
'''
The tugs assigned to Seaspan or Group Ocean will appear with a red background
'''
def highlight(s):
  '''
  Function used to highlight the background of the tugs assigned to other companies
  '''
  if s['Tug From'] == 'Group Ocean' or s['Tug From'] == 'Seaspan' or s['Tug To'] == 'Seaspan' or s['Tug To'] == 'Group Ocean':
    return ['background-color: red; color: white'] * len(s)
  else:
    return [''] * len(s)

#Apply the stilization to the table that we will send by email
obsVessels = obsVessels.style.apply(highlight, axis=1)

#==================Sending the Email==================
'''
It was created an email specially for this project.
For this, we generated an app password for the new google account
'''
#Email Data
me = "automationsaamtowage@gmail.com"
password = 'hvcceumjvvtxrgzh'
server = 'smtp.gmail.com:587'
you = 'dispatch.yvr@saamtowage.com'
#I send to myself when testing
#you = 'gabrielcanuto@my.capilanou.ca'


#Alternative text in the case that the email reader has no support to html
text = """
In order to vizualise the table, you need to be able to render HTML"""

#Html for the Main version
html = """
<html><body><p>Hello Dispatchers.</p>
<p>These are the Vessels that we are monitoring that were not assigned to us.</p>
{table}
<p>Regards,</p>
<p>Gabriel Canuto,</br>Operations Intern</p>
</body></html>
"""

#Transform the data frame into html format and concatenate into the html code
html = html.format(table=obsVessels.render())

message = MIMEMultipart(
    "alternative", None, [MIMEText(text), MIMEText(html,'html')])
message['Subject'] = "Monitoring Saam Vessels"
message['From'] = me
message['To'] = you

#File path for the excel file that contain the vessels
filename = "vesselsNames.xlsx"  # In same directory as script

# Open PDF file in binary mode
with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)
message.attach(part)

server = smtplib.SMTP(server)
server.ehlo()
server.starttls()
server.login(me, password)
server.sendmail(me, you, message.as_string())
server.quit()

