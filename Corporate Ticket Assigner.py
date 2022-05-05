import os
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from pandas import ExcelWriter
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from xgboost import XGBClassifier
from sklearn.metrics import confusion_matrix
from datetime import timedelta, date,datetime
import win32com.client as win32

'''dataset = pd.ExcelFile('Open Tickets.xlsx')'''

outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI") 
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items

received_dt = datetime.now() - timedelta(days=6)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
#messages = messages.Restrict("[SenderEmailAddress] = 'helpdesk_us@smcp.com'")
#messages = messages.Restrict("[Subject] = 'Sample Report'")
outputDir = r"C:\Users\dbouvier\Documents\Ticket Assigner 2.0"

senders = ['Helpdesk US']
file = "Open"
for i in list(messages):
    if i.Class == 43:
        if str(i.sender) in senders:
            for x in i.Attachments:
                #print(str(x))
                #print(file)
                #print(file in str(x.FileName))
                if file in str(x.FileName):
                    x.SaveASFile(os.path.join(outputDir, "Open Tickets.xlsx"))
                    print(f"attachment {x.FileName} from {senders} saved")

new_tickets =pd.read_excel('Open Tickets.xlsx').iloc[:,np.r_[5:7,9]].dropna()
new_tickets.columns = ["Ticket","Summary","Date Entered"]
new_tickets = new_tickets.iloc[1:,:]
new_tickets = new_tickets[new_tickets.Ticket != 'Ticket #']


#for f_name in os.listdir(r"C:\Users\dbouvier\Documents\Ticket Assigner 2.0"):
                    #if f_name.startswith('Pending'):
                        #previous_tickets = pd.read_excel(f_name)'''

for f_name in os.listdir(r"C:\Users\dbouvier\SMCP\Helpdesk USA - Documents\03 - Store Support Tracking"):
                    if f_name.startswith('Pending'):
                        previous_tickets = pd.read_excel(r"C:\Users\dbouvier\SMCP\Helpdesk USA - Documents\03 - Store Support Tracking\{0}".format(f_name))


previous_tickets.rename(columns = {'Ticket #' : 'Ticket'}, inplace = True)

inner_join = pd.merge(new_tickets, 
                      previous_tickets, 
                      on ='Ticket', 
                      how ='left')
inner_join = inner_join.iloc[:,np.r_[0:3,5:11]]
inner_join.rename(columns = {'Summary_x' : 'Summary'}, inplace = True)
inner_join['Person in charge'] = inner_join['Person in charge'].fillna(0)

real_new_tickets = inner_join.loc[inner_join['Person in charge'] == 0]


dataset = pd.read_csv('Tickets Archives for TRAINING.csv',encoding='unicode_escape')


for i in range(0, 506):
   dataset.iloc[[i],0] = dataset.iloc[[i],0].str.lstrip('Store ')

for i in range(1015, len(dataset)):
   dataset.iloc[[i],0] = dataset.iloc[[i],0].str.lstrip('Store ')

MappingFile = pd.ExcelFile('Mapping table ticket type-job title.xlsx')
TicketFamilies = pd.read_excel(MappingFile,'Sheet1').iloc[:,np.r_[0,1]]
EmployeeNames = pd.read_excel(MappingFile,'Sheet2')


for row in range(0,len(dataset)):
    for row1 in range(0,len(TicketFamilies)):
        if dataset["ticket family"][row] == TicketFamilies["Ticket"][row1]:
            dataset["ticket family"][row] = TicketF# Import OS 
# import osamilies["Family"][row1]


import re
import nltk
#nltk.download('stopwords')
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
corpus = []
for i in range(0,len(dataset)):
    review = re.sub('[^a-zA-Z]', ' ', str(dataset['ticket'][i]))
    review = review.lower()
    review = review.split()
    ps = PorterStemmer()
    review = [ps.stem(word) for word in review if not word in set(stopwords.words('english'))]
    review = ' '.join(review)
    corpus.append(review)
    
real_new_tickets = real_new_tickets.reset_index(drop=True)
dataset1 = real_new_tickets["Summary"].to_frame()
#dataset1 = dataset1.reset_index(drop=True)

for i in range(0, len(dataset1)):
   dataset1.iloc[[i],0] = dataset1.iloc[[i],0].str.lstrip('Store ')

corpus1 = []
for i in range(0,len(dataset1)):
    review1 = re.sub('[^a-zA-Z]', ' ', str(dataset1["Summary"][i]))
    review1 = review1.lower()
    review1 = review1.split()
    ps1 = PorterStemmer()
    review1 = [ps1.stem(word) for word in review1 if not word in set(stopwords.words('english'))]
    review1 = ' '.join(review1)
    corpus1.append(review1)

corpus.extend(corpus1)

cv = CountVectorizer(max_features = 1500) 
X = cv.fit_transform(corpus).toarray()
X_train = X[:-len(corpus1),:]
X_test = X[-len(corpus1):,:]
y = dataset.iloc[:,-1].values
y_train = y

sc_X = StandardScaler()
X_train = sc_X.fit_transform(X_train)
X_test = sc_X.transform(X_test)

classifier = XGBClassifier()
classifier.fit(X_train, y_train)

y_pred = classifier.predict(X_test)
y_pred = pd.DataFrame(y_pred)
prediction = y_pred.iloc[:,0]
dataexport = dataset1.join(prediction)
dataexport.columns = ["Summary","Ticket Family"]


JobTitlesTickets = pd.read_excel(MappingFile,'Sheet1').iloc[:,np.r_[1:3]]

JobbiesTitles = EmployeeNames["Job title"].tolist()

dataexport["Person in charge"] = 0

for x in JobbiesTitles:
    globals()[x] = []
    for i in range(0,len(JobTitlesTickets)):
        if x == JobTitlesTickets["Associated Job Title"][i]:
            globals()[x].append(JobTitlesTickets["Family"][i]) 
    for i in range(0,len(EmployeeNames)):
        if x == EmployeeNames["Job title"][i]:
            EmpName = EmployeeNames["Employee name"][i]
            for i in range(0,len(dataexport)):
                if dataexport["Ticket Family"][i] in globals()[x]:
                    dataexport.iloc[i,-1] = EmpName
 
for row in range(0,len(dataexport)):
    for row1 in range(0,len(TicketFamilies)):
        if dataexport["Ticket Family"][row] == TicketFamilies["Family"][row1]:
            dataexport["Ticket Family"][row] = TicketFamilies["Ticket"][row1]


dataexport.to_csv(r'C:\Users\dbouvier\Documents\Ticket Assigner 2.0\tickets generated by algorithm new.csv')


dataexport["Ticket"] = real_new_tickets["Ticket"]


inner_join = pd.merge(inner_join, 
                      dataexport, 
                      on ='Ticket', 
                      how ='left')
inner_join['Person in charge_x'] = np.where(inner_join['Person in charge_x'] == 0, inner_join['Person in charge_y'], inner_join['Person in charge_x'])
#for row in range(0,len(inner_join)):
    #if inner_join["Person in charge"][row] == 0:
        #inner_join["Person in charge"][row] = inner_join.iloc[row,-1]
        
inner_join = inner_join.iloc[:,:-3]

inner_join['Solved?'] = inner_join['Solved?'].fillna(0)
inner_join['Solved?'] = np.where(inner_join['Solved?'] == 0, "No", inner_join['Solved?'])

inner_join.rename(columns = {'Summary_x' : 'Summary','Date Entered_x':'Date entered'
                             ,'Person in charge_x':'Person in charge',
                             'Next Step':'Next step','Current Status':'Current status'}, inplace = True)

writer = ExcelWriter('current_tickets.xlsx')
inner_join.to_excel(writer,'Pending Tickets')
writer.save()

#to_insert = datetime.today().strftime('%m%d')

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 
#mail.CC =
mail.Subject = 'Pending Tickets Tracker for Week {0} Available'.format(datetime.today().strftime('%m%d'))
mail.HtmlBody = """Hi Team,<br><br>
 
The pending tickets tracker for this current week is live on the following SharePoint folder: 
<br>https://smcponline.sharepoint.com/sites/HelpdeskUSA


<br><br>You can sort files name from Z to A to have this week’s latest tracker displayed on top.

<br><br>Please check it out when you have time,

<br><br>Best,
<br>Dimitri Bouvier
<br>IS Systems Analyst, North America

""" 
mail.Display(False)
