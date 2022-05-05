# -*- coding: utf-8 -*-
"""
Created on Mon Apr  4 08:06:28 2022

@author: dbouvier
"""

import os
from datetime import datetime, timedelta
import win32com.client as win32
import pandas as pd
import threading
import sys
import cx_Oracle
import pythoncom
import re
from pandas import ExcelWriter
from tabulate import tabulate
from pathlib import Path
import time
import shutil
import xml.etree.ElementTree as ET
from lxml import etree
import win32net
pd.options.mode.chained_assignment = None  # default='warn'



loc = r"C:\Oracle\instantclient_21_3"

dsn_tns_sandro = 
dsn_tns_maje = 

os.environ["PATH"] = loc + ";" + os.environ["PATH"]
conn_s = cx_Oracle.connect(user=, password=, dsn=)
conn_m = cx_Oracle.connect(user=, password=, dsn=)

c_sandro = conn_s.cursor()
c_maje = conn_m.cursor()
def find_str(s, char):
    index = 0
    
    if char in s:
        c = char[0]
        for ch in s:
            if ch == c:
                if s[index:index+len(char)] == char:
                    return index
            index += 1
    return -1

lock = threading.Lock()

print('start of program')
df_tasks = pd.DataFrame(columns=['task','status'])
main_dic = {}

# bck_df = pd.DataFrame(columns = ['file','cdate'])
# df_store_recap = pd.DataFrame(columns = ['store','tooOld?'])


def wunderkind(task):
    global df_tasks
    print(f'start of {task}')
    maje_columns = ['Codemagasin', 'e-mail', 'mobile', 'nom', 'prenom', 'Codecivilite',
       'codepays', 'Complémentdenom', 'Libellevoie ', 'Complémentdadresse',
       'Codepostal', 'Ville', 'jour', 'mois', 'annee', 'Opt_email', 'Opt_SMS',
       'Opt_postal', 'canalEntree']
    sandro_columns = ['Codemagasin', 'e-mail', 'mobile', 'nom', 'prenom', 'Codecivilite',
           'codepays', 'Complémentdenom', 'Libellevoie ', 'Complémentdadresse',
           'Codepostal', 'Ville', 'jour', 'mois', 'annee', 'Opt_email', 'Opt_SMS',
           'Opt_postal', 'canalEntree']
    maje_daily_df = pd.DataFrame(columns = maje_columns)
    sandro_daily_df = pd.DataFrame(columns = sandro_columns)
    maje_files_went_over = []
    sandro_files_went_over = []
    
    
    for maje_file in os.listdir('///app/CREATION_CLIENTS/MAJE/IN'):
        if maje_file.startswith('maje_subscribes'):
            maje_files_went_over.append(maje_file)
            # print(maje_file)
            extract = list(pd.read_csv(r'///app/CREATION_CLIENTS/MAJE/IN/{0}'.format(maje_file)).iloc[:,0])
            for email_address in extract:
                maje_daily_df = maje_daily_df.append({'Codemagasin':'4249',
                                                      'e-mail':email_address,
                                                      'codepays':'US',
                                                      'Opt_email':'IN',
                                                      'canalEntree':'Wunderkind'},ignore_index=True)
    
            
    for sandro_file in os.listdir('///app/CREATION_CLIENTS/SANDRO/IN'):
        if sandro_file.startswith('sandroparis_subscribes'):
            sandro_files_went_over.append(sandro_file)
            # print(sandro_file)
            extract = list(pd.read_csv(r'///app/CREATION_CLIENTS/SANDRO/IN/{0}'.format(sandro_file)).iloc[:,0])
            for email_address in extract:
                if 'Female' in email_address:
                    code_civ = 2
                elif 'Male' in email_address:
                    code_civ = 1
                else:
                    code_civ = ''
                sandro_daily_df = sandro_daily_df.append({'Codemagasin':'4049',
                                                      'e-mail':email_address.replace(';','').replace('Female','').replace('Male',''),
                                                       'Codecivilite':code_civ,
                                                      'codepays':'US',
                                                      'Opt_email':'IN',
                                                      'canalEntree':'Wunderkind'},ignore_index=True)
    
    if len(maje_daily_df)>0:
        writer = ExcelWriter(r'\\\app\CREATION_CLIENTS\MAJE\IN\STORENEXT_CREATION_CLIENTS_MAJE_IN_{0}.xlsx'.format(datetime.today().strftime('%m%d%y')))
        maje_daily_df.to_excel(writer,'{0}'.format(datetime.today().strftime('%m%d%y')),index = False)
        writer.save()
        writer.close()
        
    if len(maje_daily_df)>0:
        writer = ExcelWriter(r'\\\app\CREATION_CLIENTS\SANDRO\IN\STORENEXT_CREATION_CLIENTS_SANDRO_IN_{0}.xlsx'.format(datetime.today().strftime('%m%d%y')))
        sandro_daily_df.to_excel(writer,'{0}'.format(datetime.today().strftime('%m%d%y')),index = False)
        writer.save()
        writer.close()
    
    if len(maje_files_went_over) >0:
        for maje_file in maje_files_went_over:
            shutil.move("///app/CREATION_CLIENTS/MAJE/IN/{0}".format(maje_file), "///app/CREATION_CLIENTS/MAJE/ARCHIVE/{0}".format(maje_file))
    if len(sandro_files_went_over)>0:
        for sandro_file in sandro_files_went_over:
            shutil.move("///app/CREATION_CLIENTS/SANDRO/IN/{0}".format(sandro_file), "///app/CREATION_CLIENTS/SANDRO/ARCHIVE/{0}".format(sandro_file))
    
    if len(maje_files_went_over) >0 or len(sandro_files_went_over)>0:
        outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
        mail = outlook.CreateItem(0)
        mail.To = 
        mail.CC = 
        mail.Subject = 'Daily SMCP Wunderkind Client Import Process'
        mail.HtmlBody = """***Automated Email***
        <br><br>Hi Team,<br><br>
        
        Today's client files from Wunderkind have been put into creation_clients folders for you to import in Storeland.
            
        <br><br>Best,
        <br>Dimitri Bouvier
        <br>IT Systems Analyst, North America
             
        <br><br>Sandro • Maje • Claudie Pierlot • De Fursac
        <br>44 Wall Street
        <br>New York, NY 10005
        """
            
        mail.Display(False)
    else:
        print('no files with rows from salescycle found today')
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)
def transferFileChecker(task):
    print(f'start of {task}')
    global df_tasks
    number_of_hours = 12

    folder_sandro = r'\\srvmediacontact\storeland\depart'
    folder_maje = r'\\frm-ss-023\STORLAND23\DEPART'
    
    df = pd.DataFrame(columns =['fileName','cDate','storeNumber'])
    
    def file_checker(x,y):
        for file in os.listdir(x):
            if file.startswith('Transfer') and (int(file[-4:]) >=4000 and int(file[-4:]) < 5000):
                # print(file)
                y = y.append({'fileName':file,
                            'cDate':os.path.getctime(r'{0}\{1}'.format(x,file)),
                            'storeNumber':int(file[-4:])}, ignore_index = True)
        return y
    
    df = file_checker(folder_sandro,df)
    df = file_checker(folder_maje,df)
    
    def getTimeStamp(x):
        return datetime.utcfromtimestamp(x)
    
    df['cDate'] = df['cDate'].apply(lambda x : getTimeStamp(x))
    
    df['tooOld?'] = df['cDate']
    
    def tooOldCheck(x):
        if x < (datetime.today() - timedelta(hours =number_of_hours)):
            return 'yes'
        else:
            return 'no'
    
    df['tooOld?'] = df['tooOld?'].apply(lambda x : tooOldCheck(x))
    
    
    outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
    mail = outlook.CreateItem(0)
    mail.To = 
    # mail.CC = 
    mail.Subject = 'Automated Daily Mediacontact Folders Check'
    mail.Body = """***Automated Email***

Hi Team,
 
The following stores may have a Mediacontact issue:
    
{list_of_stores}

Please note that the current threshold to determine whether a file can be considered as 'stuck' in the mediacontact folder is if the file creation date is older than {hours} hours old.    

Best,
Dimitri Bouvier
IT Systems Analyst, North America
     
Sandro • Maje • Claudie Pierlot • De Fursac
44 Wall Street
New York, NY 10005

    """.format(list_of_stores = list(df[df['tooOld?'] == 'yes']['storeNumber']),hours= number_of_hours).replace('[','').replace(']','')
    # mail.Attachments.Add(r'C:/Users/dbouvier/Downloads/SentOrdersWithoutMatch{0}.csv'.format(datetime.today().strftime('%m%d%y')))
    mail.Display(False)
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)
def blmAudStlComparision(task):
    print(f'start of {task}')
    days_back = 45

    start_date_minus_four = (datetime.today() - timedelta(days=(days_back+4))).strftime('%Y-%m-%d')
    start_date = (datetime.today() - timedelta(days=days_back)).strftime('%Y-%m-%d')
    
    end_date = datetime.today().strftime('%Y-%m-%d')
    
    
    sql = """
    select trunc(jourdevente)jourdevente, sum(ca) from (
    select codemagasin, trunc(jourdevente)jourdevente, canetrealise ca from historique_caisses
    where codemagasin in (select codemagasin from stats_magasin
    where codeaxestat=17
    and codeelementstat=8)
    and trunc(jourdevente) >= to_date('{0}','yyyy-mm-dd')
    and trunc(jourdevente) <= to_date('{1}','yyyy-mm-dd')
    and ticketannule=0
    and typeligne=1
    union all
    select codemagasin, trunc(jourdevente), canetrealise*-1 ca from historique_caisses
    where codemagasin in (select codemagasin from stats_magasin
    where codeaxestat=17
    and codeelementstat=8)
    and trunc(jourdevente) >= to_date('{0}','yyyy-mm-dd')
    and trunc(jourdevente) <= to_date('{1}','yyyy-mm-dd')
    and ticketannule=0
    and typeligne=2)
    group by trunc(jourdevente)
    order by trunc(jourdevente)
    """.format(start_date_minus_four,(datetime.today() - timedelta(days=4)).strftime('%Y-%m-%d'))
    
    
    maje_results = pd.read_sql(sql,conn_m)
    sandro_results = pd.read_sql(sql,conn_s)
    
    maje_results.rename(columns = {'JOURDEVENTE':'date'},inplace=True)
    sandro_results.rename(columns = {'JOURDEVENTE':'date'},inplace=True)
    
    date_range = pd.date_range(start_date,end_date, freq='D').strftime('%d%m%Y').to_list()
    
    def scanner(x,y):
        if y['QTY'][x] == 0 and str(y['UPC'][x]).startswith('366160'):
            return 'loyalty correct maje'
        elif y['QTY'][x] == 0 and str(y['UPC'][x]).startswith('360717'):
            return 'loyalty correct sandro'
        elif str(y['UPC'][x]).startswith('366160') and len(str(y['UPC'][x])) ==13:
            return 'MAJE'
        elif str(y['UPC'][x]).startswith('360717') and len(str(y['UPC'][x])) ==13:
            return 'SANDRO'
        else:
            return 'GC/UnknownSKU'
    df_maje = pd.DataFrame(columns=['date','totalAmount','filePath'])
    #extracting Maje files
    for folder in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee'):
        if folder in date_range:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}'.format(folder)):
                if file.startswith('AUDITED'):
                    # print(file)        
                    my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                    my_file = my_file[my_file['UPC'] > 3661600000000]
                    my_file = my_file[my_file['UPC'] < 3661610000000]
                    my_file.reset_index(drop=True,inplace=True)
                    my_file['outcome'] = my_file.index
                    my_file['outcome'] = my_file['outcome'].apply(lambda x: scanner(x,my_file))
                    my_file = my_file[my_file['outcome'] == 'MAJE']
                    df_maje = df_maje.append({'date':(datetime.strptime(folder,'%d%m%Y') - timedelta(days=4)),
                                              'totalAmount':sum(my_file['sum'].astype(float)),#MERCH_AMT
                                              'filePath':r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file)},
                                             ignore_index=True)
    df_sandro = pd.DataFrame(columns=['date','totalAmount','filePath'])
    #extracting Sandro files
    for folder in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee'):
        if folder in date_range:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}'.format(folder)):
                if file.startswith('AUDITED'):
                    # print(file)
                    my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                    my_file = my_file[my_file['UPC'] > 3607170000000]
                    my_file = my_file[my_file['UPC'] < 3607180000000]
                    my_file.reset_index(drop=True,inplace=True)
                    my_file['outcome'] = my_file.index
                    my_file['outcome'] = my_file['outcome'].apply(lambda x: scanner(x,my_file))
                    my_file = my_file[my_file['outcome'] == 'SANDRO']
                    df_sandro = df_sandro.append({'date':(datetime.strptime(folder,'%d%m%Y') - timedelta(days=4)),
                                              'totalAmount':sum(my_file['sum'].astype(float)),#MERCH_AMT
                                              'filePath':r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file)},
                                             ignore_index=True)
    
    maje = df_maje.merge(maje_results, how='left', on ='date').rename(columns={'SUM(CA)':'totalAmountStoreland'}).dropna()
    sandro = df_sandro.merge(sandro_results, how='left', on ='date').rename(columns={'SUM(CA)':'totalAmountStoreland'}).dropna()
    
    # maje['totalAmountStoreland'] = maje['totalAmountStoreland'].astype('int')
    # sandro['totalAmountStoreland'] = sandro['totalAmountStoreland'].astype('int')
    
    maje['auditedFileStorelandDelta'] = round(abs(maje['totalAmount'] - maje['totalAmountStoreland']),2)
    sandro['auditedFileStorelandDelta'] = round(abs(sandro['totalAmount'] - sandro['totalAmountStoreland']),2)
    
    maje = maje[['date','totalAmount','totalAmountStoreland','auditedFileStorelandDelta','filePath']]
    sandro = sandro[['date','totalAmount','totalAmountStoreland','auditedFileStorelandDelta','filePath']]
    
    maje.rename(columns={'totalAmount':'totalAmountAudited'},inplace=True)
    sandro.rename(columns={'totalAmount':'totalAmountAudited'},inplace=True)
    
    
    sandro.sort_values(by=['date'],ascending=False) .to_html(r'C:\Users\dbouvier\Downloads\SandroAUDSTL.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
    maje.sort_values(by=['date'],ascending=False).to_html(r'C:\Users\dbouvier\Downloads\MajeAUDSTL.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
    
    for brand in ['Sandro','Maje']:
        outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
        mail = outlook.CreateItem(0)
        mail.To =
        mail.CC = 
        mail.Subject = '{brand} BLM Audited Sales Files VS Storeland Comparison Last 45 Trailing Days'.format(brand = brand.upper())
        mail.HtmlBody = """***Automated Email***<br><br>
        Hi Team,<br><br>
        
        Here is the recap table of the last 45 trailing days of audited sales files vs storeland sales for BLM {brand}:<br><br>
        {table}
        <br><br>
        Best,<br>
        Dimitri Bouvier<br>
        IT Systems Analyst, North America
        <br><br>
        Sandro • Maje • Claudie Pierlot • De Fursac<br>
        44 Wall Street<br>
        New York, NY 10005<br>
        """.format(brand=brand,table=open(r'C:\Users\dbouvier\Downloads\{brand}AUDSTL.html'.format(brand=brand)).read())
        mail.Display(False)
    global df_tasks
    print(f'end of {task}')

    with lock:
        df_tasks = df_tasks.append({'task':task,
                                    'status':'done'},ignore_index=True)
def blmAudDaiComparison(task):
    global df_tasks
    print(f'start of {task}')
    trailing_days = 15
    percent_thresold = 3
    
    start_date_minus_four = (datetime.today() - timedelta(days=(trailing_days-3))).strftime('%Y-%m-%d')
    start_date = (datetime.today() - timedelta(days=trailing_days)).strftime('%Y-%m-%d')
    end_date_minus_four = (datetime.today() - timedelta(days=3)).strftime('%Y-%m-%d')
    end_date = datetime.today().strftime('%Y-%m-%d')
    
    date_range_daily = pd.date_range(start_date_minus_four,end_date_minus_four, freq='D').strftime('%d%m%Y').to_list()
    date_range_audited = pd.date_range(start_date,end_date, freq='D').strftime('%d%m%Y').to_list()
    
    df_audited_maje, df_audited_sandro, df_daily_sandro, df_daily_maje = pd.DataFrame(columns=['date','totalAmount','filePath','file']),pd.DataFrame(columns=['date','totalAmount','filePath','file']),pd.DataFrame(columns=['date','totalAmount','filePath','file']),pd.DataFrame(columns=['date','totalAmount','filePath','file'])
    
    for folder in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee'):
        if folder in date_range_daily:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}'.format(folder)):
                if file.startswith('SALES') and file not in df_daily_maje['file']:
                    my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                    my_file = my_file[my_file['UPC'] > 3661600000000]
                    my_file = my_file[my_file['UPC'] < 3661610000000]
                    # print(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    df_daily_maje = df_daily_maje.append({'date':(datetime.strptime(folder,'%d%m%Y')-timedelta(days=1)),
                                                          'totalAmount':sum(my_file['sum'].astype('int')),#MERCH_AMT
                                                          'filePath':r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file),
                                                          'file':file},
                                                         ignore_index=True)
                    # sys.exit('')
        if folder in date_range_audited:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}'.format(folder)):
                # for file in os.listdir(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}'.format(folder)):
                    if file.startswith('AUDITED') and file not in df_audited_maje['file']:
                        my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file))
                        my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                        my_file = my_file[my_file['UPC'] > 3661600000000]
                        my_file = my_file[my_file['UPC'] < 3661610000000]
                        # print(r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file))
                        df_audited_maje = df_audited_maje.append({'date':(datetime.strptime(folder,'%d%m%Y')-timedelta(days=4)),
                                                          'totalAmount':sum(my_file['sum'].astype('int')),#MERCH_AMT
                                                          'filePath':r'\\\app\EDI_USA\BLM\maje\SAV-Arrivee\{0}\{1}'.format(folder,file),
                                                          'file':file},
                                                         ignore_index=True)
    
    
    
    for folder in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee'):
        if folder in date_range_daily:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}'.format(folder)):
                if file.startswith('SALES') and file not in df_daily_sandro['file']:
                    my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                    my_file = my_file[my_file['UPC'] > 3607170000000]
                    my_file = my_file[my_file['UPC'] < 3607180000000]
                    # print(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file))
                    df_daily_sandro = df_daily_sandro.append({'date':(datetime.strptime(folder,'%d%m%Y')-timedelta(days=1)),
                                                          'totalAmount':sum(my_file['sum'].astype('int')),#MERCH_AMT
                                                          'filePath':r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file),
                                                          'file':file},
                                                         ignore_index=True)
                    # sys.exit('')
        if folder in date_range_audited:
            for file in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}'.format(folder)):
                # for file in os.listdir(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}'.format(folder)):
                    if file.startswith('AUDITED') and file not in df_audited_sandro['file']:
                        my_file = pd.read_csv(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file))
                        my_file['sum'] = my_file['MERCH_AMT']*my_file['QTY']
                        my_file = my_file[my_file['UPC'] > 3607170000000]
                        my_file = my_file[my_file['UPC'] < 3607180000000]
                        # print(r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file))
                        df_audited_sandro = df_audited_sandro.append({'date':(datetime.strptime(folder,'%d%m%Y')-timedelta(days=4)),
                                                          'totalAmount':sum(my_file['sum'].astype('int')),#MERCH_AMT
                                                          'filePath':r'\\\app\EDI_USA\BLM\sandro\SAV-Arrivee\{0}\{1}'.format(folder,file),
                                                          'file':file},
                                                         ignore_index=True)
    
    df_audited_maje.drop_duplicates(inplace=True)
    df_audited_sandro.drop_duplicates(inplace=True)
    df_daily_sandro.drop_duplicates(inplace=True)
    df_daily_maje.drop_duplicates(inplace=True)
    
    maje = df_audited_maje.merge(df_daily_maje, how='left', on ='date').rename(columns={'filePath_x':'filePathAudited',
                                                                                        'filePath_y':'filePathDaily',
                                                                                        'totalAmount_x':'totalAmountAudited',
                                                                                        'totalAmount_y':'totalAmountDaily'})
    maje['deltaDailyAudited'] = maje['totalAmountDaily'] - maje['totalAmountAudited']
    maje = maje[['date','totalAmountDaily','totalAmountAudited','deltaDailyAudited','filePathDaily','filePathAudited']]
    sandro = df_audited_sandro.merge(df_daily_sandro, how='left', on ='date').rename(columns={'filePath_x':'filePathAudited',
                                                                                        'filePath_y':'filePathDaily',
                                                                                        'totalAmount_x':'totalAmountAudited',
                                                                                        'totalAmount_y':'totalAmountDaily'})#[['date','totalAmountDaily','totalAmountAudited','filePathDaily','filePathAudited']]
    sandro['deltaDailyAudited'] = sandro['totalAmountDaily'] - sandro['totalAmountAudited']
    sandro = sandro[['date','totalAmountDaily','totalAmountAudited','deltaDailyAudited','filePathDaily','filePathAudited']]
        
    
    maje_export = maje.copy()
    sandro_export = sandro.copy()
    
    maje_export['delta_percent_audited'] = round(((maje_export['deltaDailyAudited']/maje_export['totalAmountAudited'])*100).astype(float),2)
    maje_export['delta_percent_daily'] = round(((maje_export['deltaDailyAudited']/maje_export['totalAmountDaily'])*100).astype(float),2)
    maje_export = maje_export[(maje_export['delta_percent_audited']<=-percent_thresold) | (maje_export['delta_percent_audited']>=percent_thresold) | (maje_export['delta_percent_daily']<=-percent_thresold) | (maje_export['delta_percent_daily']>=percent_thresold)]
    
    sandro_export['delta_percent_audited'] = round(((sandro_export['deltaDailyAudited']/sandro_export['totalAmountAudited'])*100).astype(float),2)
    sandro_export['delta_percent_daily'] = round(((sandro_export['deltaDailyAudited']/sandro_export['totalAmountDaily'])*100).astype(float),2)
    sandro_export = sandro_export[(sandro_export['delta_percent_audited']<=-percent_thresold) | (sandro_export['delta_percent_audited']>=percent_thresold) | (sandro_export['delta_percent_daily']<=-percent_thresold) | (sandro_export['delta_percent_daily']>=percent_thresold)]
    
    
    if len(maje_export)>0:
        maje_export[list(maje_export.columns)[:-2]].sort_values(by=['date'],ascending=False).to_html(r'C:\Users\dbouvier\Downloads\MajeAUDDAY.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
        outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
        mail = outlook.CreateItem(0)
        mail.To = 
        mail.CC = 
        mail.Subject = 'MAJE BLM Audited Sales Files VS Daily Sales Files Last {0} Trailing Days'.format(trailing_days)
        mail.HtmlBody = """***Automated Email***<br><br>
        Hi Team,<br><br>
        
        The following {number_of_days_issue} days of sales (subset of the last {trailing} trailing days) have a high discrepancy sales delta between the audited and daily sales file for BLM Maje and may need to be reviewed:<br><br>
        (Note that the criterion to consider a sales discrepancy as a high discrepancy is set to {percent_criterion} %)<br><br><br>
        {html_table}
        <br><br>
        Best,<br>
        Dimitri Bouvier<br>
        IT Systems Analyst, North America
        <br><br>
        Sandro • Maje • Claudie Pierlot • De Fursac<br>
        44 Wall Street<br>
        New York, NY 10005<br>
        """.format(number_of_days_issue = len(maje_export),trailing = trailing_days,percent_criterion=percent_thresold,html_table=open(r'C:\Users\dbouvier\Downloads\MajeAUDDAY.html').read())
        mail.Display(False)
    
    if len(sandro_export)>0:
        sandro_export[list(sandro_export.columns)[:-2]].sort_values(by=['date'],ascending=False).to_html(r'C:\Users\dbouvier\Downloads\SandroAUDDAY.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
        outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
        mail = outlook.CreateItem(0)
        mail.To = 
        mail.CC = 
        mail.Subject = 'SANDRO BLM Audited Sales Files VS Daily Sales Files Last {0} Trailing Days'.format(trailing_days)
        mail.HtmlBody = """***Automated Email***<br><br>
        Hi Team,<br><br>
        
        The following {number_of_days_issue} days of sales (subset of the last {trailing} trailing days) have a high discrepancy sales delta between the audited and daily sales file for BLM Sandro and may need to be reviewed:<br><br>
        (Note that the criterion to consider a sales discrepancy as a high discrepancy is set to {percent_criterion} %)<br><br>
        {html_table}
        <br><br>
        Best,<br>
        Dimitri Bouvier<br>
        IT Systems Analyst, North America
        <br><br>
        Sandro • Maje • Claudie Pierlot • De Fursac<br>
        44 Wall Street<br>
        New York, NY 10005<br>
        """.format(number_of_days_issue = len(sandro_export),trailing = trailing_days,percent_criterion=percent_thresold,html_table=open(r'C:\Users\dbouvier\Downloads\SandroAUDDAY.html').read())
        mail.Display(False)
        
        print(f'end of {task}')
        with lock:
            df_tasks = df_tasks.append({'task':task,
                                       'status':'done'},ignore_index=True)
def backUpMagCheck(task):
    def getTimeStamp(x):
        return datetime.utcfromtimestamp(x)

    def tooOldCheck(x):
        if x < (datetime.today() - timedelta(hours =number_of_hours)):
            return 'yes'
        else:
            return 'no'
    df_store_recap = pd.DataFrame(columns = ['store','tooOld?','tooSmall?','bckpSize','bckpDate'])#,'tooOldLocal?'
    # global df_store_recap
    print(f'start of {task}')
    global df_tasks
    number_of_hours = 48
    df_IPs = pd.read_csv(r'C:\Users\dbouvier\Downloads\ipaddresses.csv')#[['Name','Operating System','Private IP']]
    df_IPs.reset_index(inplace=True)
    df_IPs.columns = df_IPs.iloc[0,:]
    df_IPs = df_IPs.iloc[1:,:]
    df_IPs.reset_index(drop=True,inplace=True)
    df_IPs = df_IPs[['Name','Operating System','Private IP']]
    
    df_IPs['to delete'] = df_IPs.index
    
    def checker(x,y):
        if ('MAG' in y['Name'][x]) and ('C01' in y['Name'][x]):#('Windows 10' in y['Operating System'][x]) and ('172.23' in y['Private IP'][x]) and 
            return 'no'
        else:
            return 'yes'
    def nameShorter(x):
        return x[5:9]
    
    df_IPs['to delete'] = df_IPs['to delete'].apply(lambda x: checker(x,df_IPs))
    df_IPs['Name'] = df_IPs['Name'].apply(lambda x: nameShorter(x))
    
    df_IPs = df_IPs[df_IPs['to delete'] != 'yes']
    
    df_IPs.set_index(['Name'],inplace=True)
    
    # sys.exit('abc')
    
    stores_list_sql = """select * from magasins m
    where m.codepays in ('US','CA')
    and m.codetypegestionmag =1
    and m.datefermeture is null
    and m.dateouverture<sysdate
    and m.numeromodem is not null
    and (m.codemagasin >= 4000 and m.codemagasin <=4999)
    order by m.codemagasin
    """
    
    list_of_stores = []
    list_of_stores.extend(list(pd.read_sql(stores_list_sql,conn_m)['CODEMAGASIN']))
    list_of_stores.extend(list(pd.read_sql(stores_list_sql,conn_s)['CODEMAGASIN']))
    
    # list_of_stores = list_of_stores[:10]
    
    dic_brands = {}
    dic_brands['sandro'] = r'\\frx-d1-sto-02\BACKUP-MAGASINS\BCKMAGSANDRO'
    dic_brands['maje'] = r'\\frx-d1-sto-02\BACKUP-MAGASINS\BCKMAGMAJE'
    regex = re.compile('4\d\d\d')
    
    for brand in ['maje','sandro']:
        for folder in os.listdir(dic_brands[brand]):
            if regex.match(folder) and int(folder) in list_of_stores:
                # print(folder)
                # global bck_df
                bck_df = pd.DataFrame(columns = ['file','cdate','size'])
                backup_exists = 0
                for file in os.listdir(f'{dic_brands[brand]}\{folder}'):
                    # print(file)
                    if 'Data' in file:
                        backup_exists = 1
                        bck_df = bck_df.append({'file':file,
                                                'cdate':os.path.getctime(f'{dic_brands[brand]}\{folder}\{file}'),
                                                'size':int(os.path.getsize(f'{dic_brands[brand]}\{folder}\{file}'))/1000},
                                                                         ignore_index=True)
                if backup_exists == 1:
                    bck_df['size'] = bck_df['size'].astype(int)
                    bck_df['cdate'] = bck_df['cdate'].apply(lambda x : getTimeStamp(x))
                    bck_df = bck_df.sort_values(by=['cdate'], ascending=False).reset_index(drop=True)
                    bck_df['tooOld?'] = bck_df['cdate']
                    bck_df['tooOld?'] = bck_df['tooOld?'].apply(lambda x : tooOldCheck(x))
                    
                    # if bck_df['tooOld?'][0] == 'yes':
                    df_store_recap = df_store_recap.append({'store':folder,
                                                            'tooOld?':'yes' if bck_df['tooOld?'][0] == 'yes' else 'no',
                                                            'tooSmall?':'yes' if bck_df['size'][0] <20 else 'no',
                                                            'bckpSize':bck_df['size'][0],
                                                            'bckpDate':bck_df['cdate'][0]}
                                                           ,ignore_index=True)
                else:
                    df_store_recap = df_store_recap.append({'store':folder,
                                                            'tooOld?':'no backup'}
                                                           ,ignore_index=True)
                # for line_number in range(0,len(df_store_recap)):
                #     if df_store_recap['tooOld?'][line_number] == 'yes':
    
                    
    
    username = 'CAISSE'
    password = 'CAISSE'
    use_dict={}
    use_dict['password']=password
    use_dict['username']=username
    
    df_store_recap.set_index('store',drop=False,inplace=True)
    df_store_recap['tooOldLocal?'] = 'not checked'
    
    # sys.exit('abc')
    
    # print('starting to remote in local stores to check backups')
    
    # for ip in df_IPs['Private IP']:
    for i in df_store_recap[df_store_recap['tooOld?'] == 'yes']['store']:
        if i in list(df_IPs.index):
            # print(i)
            # ip = df_IPs['Private IP'][i]
            # print(r'\\{IP}\Partage\Data'.format(IP = df_IPs['Private IP'][i]))
            use_dict['remote']=r'\\{IP}\Partage\Data'.format(IP = df_IPs['Private IP'][i])
            try:
                win32net.NetUseAdd(None, 2, use_dict)
                if 'Data' in os.listdir(r'\\{IP}\Partage'.format(IP = df_IPs['Private IP'][i])):
                    if 'Data.zip' in os.listdir(r'\\{IP}\Partage\Data'.format(IP = df_IPs['Private IP'][i])):
                        un_essai_dt = datetime.fromtimestamp(os.path.getmtime(r'\\{IP}\Partage\Data\Data.zip'.format(IP = df_IPs['Private IP'][i])))
                        df_store_recap['bckpDate'][i] = un_essai_dt
                        df_store_recap['bckpSize'][i] = int(int(os.path.getsize(r'\\{IP}\Partage\Data\Data.zip'.format(IP = df_IPs['Private IP'][i])))/1000)
                        if un_essai_dt < (datetime.today() - timedelta(hours =number_of_hours)):
                            df_store_recap['tooOldLocal?'][i] = 'yes'
                        else:
                            df_store_recap['tooOldLocal?'][i] = 'no'
                win32net.NetUseDel(None,r'\\{IP}\Partage\Data'.format(IP = df_IPs['Private IP'][i]))
            except:
                print(f"couldn't remote in {i}")
                
    # sys.exit('abc')   
    # print(type(df_store_recap['bckpSize']))
                 
    df_store_recap['issue'] = df_store_recap.index
    
    def summary(x,y):
        string = ''
        if y['tooOldLocal?'][x] == 'no':
            string = 'MC Issue'
        elif (y['tooOld?'][x] == 'yes') and (y['tooOldLocal?'][x] != 'no'):
            string = 'Zip File Too Old'
        elif y['bckpSize'][x] >20 and (y['tooOld?'][x] == 'no'):
            string = 'All good!'
        if y['bckpSize'][x] <20:
            if string == '':
                string = 'No BackUp in file'
            else:
                string = string + ' + No BackUp in file'
        if (y['tooOld?'][x] == 'yes') and (y['tooOldLocal?'][x] == 'not checked'):
            string = string + ' + could not remote in'
        return string
    
    df_store_recap['issue'] = df_store_recap['issue'].apply(lambda x: summary(x,df_store_recap))
    
    df_IPs.reset_index(inplace=True,drop=False)

    df_IPs.rename({'Name':'store'},inplace=True)
    df_IPs.columns = ['store','operating system','ip','to delete']
    
    df_store_recap = pd.merge(left=df_store_recap.iloc[:,1:],right=df_IPs, 
                                on='store',how='left' )
    
    df_store_recap.drop_duplicates(subset=['store'],inplace=True, keep='last')
    
    # bla = abcd['store'].value_counts()
    
    df_store_recap[['store','issue','bckpSize', 'bckpDate','operating system','ip']].sort_values(by=['issue'],ascending=False).to_html(r'C:\Users\dbouvier\Downloads\bckpMagReport.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
    
    # sys.exit('abc')
    abc = df_store_recap['issue'].value_counts()
    abc = pd.DataFrame(abc)
    abc.reset_index(inplace=True,drop=False)
    abc.columns = ['issue','count']
    
    abc.to_html(r'C:\Users\dbouvier\Downloads\bckpMagReportKPIs.html'
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
    
    outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
    mail = outlook.CreateItem(0)
    mail.To = 
    # mail.CC = 
    mail.Subject = 'Automated Daily Winstore Backups Report'
    mail.HtmlBody = """***Automated Email***<br><br>
    
    Hi Team,<br><br>
     
    Please find below the latest Winstore BackUps report and related KPIs:<br><br>
    
    {KPI}<br><br>
    
    {report}<br><br>
    
    
    Best,<br>
    Dimitri Bouvier<br>
    IT Systems Analyst, North America<br><br>
     
    Sandro • Maje • Claudie Pierlot • De Fursac<br>
    44 Wall Street<br>
    New York, NY 10005<br>
    """.format(KPI = (open(r'C:\Users\dbouvier\Downloads\bckpMagReportKPIs.html').read()),report = (open(r'C:\Users\dbouvier\Downloads\bckpMagReport.html').read()))
    # """.format(list_of_stores = list(df_store_recap[df_store_recap['tooOld?'] == 'yes']['store']),hours= number_of_hours,list_of_no_backups= list(df_store_recap[df_store_recap['tooOld?'] == 'no backup']['store'])).replace('[','').replace(']','').replace("'",'')
    # mail.Attachments.Add(r'C:/Users/dbouvier/Downloads/SentOrdersWithoutMatch{0}.csv'.format(datetime.today().strftime('%m%d%y')))
    mail.Display(False)
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)
def dailyChecklistCheck(task):
    global df_tasks
    print(f'start of {task}')
    global dsn_tns_sandro
    global dsn_tns_maje
    folder_path = r'C:\Users\dbouvier\Documents\output file'
    connection_dic = {'maje':cx_Oracle.connect(user='STORELAND', password='STORELAND', dsn=dsn_tns_maje),
                  'sandro':cx_Oracle.connect(user='STORELAND', password='STORELAND', dsn=dsn_tns_sandro)}
    #--------------------------------------parameters section------------------------------------------------

    season_code = ['E22','H22']
    
    #-------------------------------------------------------------------------------------------------------
    
    queries_dic = {'replenishment':"""select distinct lp.CODELISTEPICKING, llp.CODEMAGASIN, m.nommagasin,to_char(lp.DATECREATION,'MM/DD/YYYY HH24:MM') DATECREATION, lp.QTETOTALEAPRELEVER, 
                    decode(lp.CODEREGROUPEMENT,3, 'Replenishment',11,'Wholesale',5,'Direct feed',2, 'Optimal stock adjustment','','E-com','NA') type from LISTES_PICKING lp, LIGNES_LISTE_PICKING llp,
                    magasins m
                    where lp.codedepot=9989
                    and lp.CODEETATLISTEPICKING<>3
                    and llp.CODELISTEPICKING=lp.CODELISTEPICKING
                    and lp.QTETOTALEPRELEVEE=0
                    and lp.coderegroupement=3
                    and trunc(lp.DATECREATION)>= sysdate-1
                    and llp.codemagasin=m.codemagasin
                    order by datecreation desc
                    """,
                    'partial_short':"""select * from lignes_liste_picking
                    where codelivraisonmag<>0 and qteprelevee=0
                    and codelistepicking in (select codelistepicking from listes_picking
                    where codedepot=9989
                    and trunc(datevalidation)>to_date('{date}','yyyy-mm-dd'))
                    -- datevalidation should be 14 days before.
                    """.format(date = (datetime.today()-timedelta(days=14)).strftime('%Y-%m-%d')),
                    'dropship':"""select * From SMCP_ECOM_DROPSHIP_PO
                    where PO_NUMBER not in (select codecommande from SMCP_COMMANDES_CLIENTS_WEB
                    where codemagasin in (4096,4095,4296,4957,4831,4196,4596))
                    and trunc(CREATION_DATE)>to_date('{date}','yyyy-mm-dd')
                    -- CREATION_DATE is 2 weeks before current date.
                    """.format(date = (datetime.today()-timedelta(days=14)).strftime('%Y-%m-%d')),
                    'hbc':"""select swpc.po_number, swpc.codelistepicking, to_char(lp.datecreation,'MM/DD/YYYY'), to_char(swpc.dateassignpo,'MM/DD/YYYY'), to_char(lp.datevalidation,'MM/DD/YYYY'),
                    to_char(swpc.dateassignship,'MM/DD/YYYY'), swpc.shipload ASN, swpc.tosend, retourne_codebarre(llp.codeinternearticle) CodeBarres, retourne_codesaisonarticle(codeinternearticle,1) codeSaison,
                    llp.qteaprelever, llp.qteprelevee, llp.codelivraisonmag from smcp_ws_pick_2_cde swpc
                    left join lignes_liste_picking llp on llp.codelistepicking=swpc.codelistepicking
                    inner join listes_picking lp on llp.codelistepicking=lp.codelistepicking
                    where trunc(swpc.dateassignpo) >to_date('{date}','yyyy-mm-dd')
                    order by swpc.dateassignpo desc
                    -- Best practice is to select approximately 3 months earlier when asking from the bind of the query. 
                    """.format(date = (datetime.today()-timedelta(days=90)).strftime('%Y-%m-%d'))
    }
    for season in season_code:
        queries_dic['nrf_'+season] = """select c.codecoloris,'' NRF_Code, cl.libcoloris ColorEN, c.libcoloris ColorFR from coloris c
                        left outer join coloris_langue cl on c.codecoloris=cl.codecoloris and cl.codelangue=2
                        where c.codecoloris not in (select codecoloris from smcp_nrf_colorcodes)
                        and c.codecoloris in (select distinct codecoloris from articles
                        where retourne_codesaisonarticle(codeinternearticle,1) in ('{season}'))
                        """.format(season = season)#.replace('[','').replace(']','')
    
    results_dic = {}
    
    for brand in ['maje','sandro']:
        for query in queries_dic:
            results_dic[query+'_'+brand] = pd.read_sql(queries_dic[query],connection_dic[brand])
            
            
    df = pd.DataFrame(columns=['query_brand','outcome'])      
    
    
      
    for query_brand in results_dic:
        if 'nrf' in query_brand:
            if len(results_dic[query_brand]) == 0:
                df = df.append({'query_brand':query_brand,
                                'outcome':'Green'},ignore_index=True)
            else:
                df = df.append({'query_brand':query_brand,
                    'outcome':'Red'},ignore_index=True)
        elif 'replen' in query_brand:
            if len(results_dic[query_brand]) > 0:
                df = df.append({'query_brand':query_brand,
                    'outcome':'Green'},ignore_index=True)
            else:
                df = df.append({'query_brand':query_brand,
                    'outcome':'Red'},ignore_index=True)
        elif 'partial' in query_brand:
            if len(results_dic[query_brand]) == 0:
                df = df.append({'query_brand':query_brand,
                    'outcome':'Green'},ignore_index=True)
            else:
                temoin = 0
                for line in range(0,len(results_dic[query_brand])):
                    if (int(results_dic[query_brand]['QTEAPRELEVER'][line]) != int(results_dic[query_brand]['QTEPRELEVEE'][line])) or pd.isnull(results_dic[query_brand]['CODELIVRAISONMAG'][line]):
                        temoin = 1
                        break
                if temoin == 0:
                    df = df.append({'query_brand':query_brand,
                        'outcome':'Green'},ignore_index=True)
                else:
                    df = df.append({'query_brand':query_brand,
                        'outcome':'Red'},ignore_index=True)
        elif 'dropship' in query_brand:
            if len(results_dic[query_brand]) == 0:
                 df = df.append({'query_brand':query_brand,
                     'outcome':'Green'},ignore_index=True)
            else:
                temoin = 0
                for line in range(0,len(results_dic[query_brand])):
                    if pd.isnull(results_dic[query_brand]['STATUS'][0]) == False:
                        temoin = 1
                        break
                if temoin == 0:
                    df = df.append({'query_brand':query_brand,
                        'outcome':'Green'},ignore_index=True)
                else:
                    df = df.append({'query_brand':query_brand,
                        'outcome':'Red'},ignore_index=True)
        elif 'hbc' in query_brand:
            if len(results_dic[query_brand]) > 0:
                 df = df.append({'query_brand':query_brand,
                     'outcome':'Green'},ignore_index=True)
            else:
                 df = df.append({'query_brand':query_brand,
                     'outcome':'Red'},ignore_index=True)      
    
    df.sort_values(by=['query_brand'],ascending=True).to_html(r'{folder_path}\outcomes.html'.format(folder_path=folder_path)
                            ,index=False
                            ,index_names=False
                            ,na_rep='')
    
    writer = pd.ExcelWriter(r'{folder_path}\Daily Queries Results {today}.xlsx'.format(folder_path=folder_path, today= datetime.today().strftime("%m-%d-%Y")))
    
    
    for query_brand in sorted(results_dic):
        results_dic[query_brand].to_excel(writer,query_brand,index = False)
    
    writer.save()
    outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
    mail = outlook.CreateItem(0)
    mail.To =
    mail.CC = 
    mail.Subject = 'Automated Daily Checklist Queries Report Generation'
    mail.HtmlBody = """***Automated Email***<br><br>
    
    Good morning Bolin,<br><br>
     
    Today's report on the outcome of the daily checklist queries:<br>
    (Please note that the result of each query has been attached as an excel file to this email)<br><br>
        
    {report}<br><br>
    
    Best,<br>
    Dimitri Bouvier<br>
    IT Systems Analyst, North America<br><br>
         
    Sandro • Maje • Claudie Pierlot • De Fursac<br>
    44 Wall Street<br>
    New York, NY 10005<br>
    """.format(report = open(r'{folder_path}\outcomes.html'.format(folder_path=folder_path)).read())
    mail.Attachments.Add(r'{folder_path}\Daily Queries Results {today}.xlsx'.format(folder_path=folder_path, today= datetime.today().strftime("%m-%d-%Y")))
    mail.Display(False)
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)
def all850s(task):
    def mapping850(brand):
        print(f'start of {brand} 850 mapping')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\{brand}'.format(brand = brand)
        df = pd.read_excel(r'{0}\850_mapping.xlsx'.format(path), engine = 'openpyxl')
        path_list = list(set(df['path']))
        
        increment = 0
        for folder in os.listdir(r"\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee".format(brand = brand)):   
            if datetime.strptime(folder, "%d%m%Y") > datetime.today() - timedelta(days=7):
                for file in os.listdir(r"\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee\{folder}".format(brand = brand,folder=folder)):
                    increment +=1
                    if increment % 10000 == 0:
                        print(increment)
                    if file.startswith('850_') and r"\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee\{folder}\{file}".format(brand=brand,folder=folder,file=file) not in path_list: #list(set(df['Path'])):
                        # print(r"Opening \\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee\{folder}\{file}".format(brand=brand,folder=folder,file=file))
                        list_of_SKUs = []
                        for line in open(r"\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee\{folder}\{file}".format(brand=brand,folder=folder,file=file)).readlines():
                            if line.startswith('DETAILS'):
                                list_of_SKUs.append(line.split('|')[5])
                
                        df = df.append({'path':r"\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\SAV-Arrivee\{folder}\{file}".format(brand=brand,folder=folder,file=file),
                                                'SKU':list_of_SKUs},ignore_index=True)
                    # sys.exit('')
        
        # writer = ExcelWriter(r'{0}\850_mapping{1}.xlsx'.format(path,datetime.today().strftime('%m%d%y')))
        writer = ExcelWriter(r'{0}\850_mapping.xlsx'.format(path))
        df.to_excel(writer,'{0}'.format(datetime.today().strftime('%m%d%y')),index = False)
        writer.save()
        writer.close()
        print(f'end of {brand} 850 mapping')

    global df_tasks
    print(f'start of {task}')
    mapping_dic = {}
    for brand in ['Maje','Sandro']:
        mapping_dic[brand] = threading.Thread(target = mapping850, args =[brand])
    
    for thread in mapping_dic:
        mapping_dic[thread].start()
    
    for thread in mapping_dic:
        mapping_dic[thread].join()

    date_check = datetime.today() - timedelta(days = 5)

    maje_mapping = pd.read_excel(r"C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Maje\850_mapping.xlsx", engine = 'openpyxl')
    sandro_mapping = pd.read_excel(r"C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Sandro\850_mapping.xlsx", engine = 'openpyxl')
    

    maje_file_dic = {}
    orders_resent_summary = pd.DataFrame(columns = ['orderNumber','SKUs'])#,'SKUs'])
    
    # print("--------------------------------------------------The following orders have mis-matching SKUs:--------------------------------------------------------------------")
    
    orderDicMaje = {}
    orderDicSandro = {}

    for line in range(0,len(maje_mapping)):
        if datetime.strptime(maje_mapping['path'][line][59:67], '%d%m%Y') >= date_check:
            if '360717' in maje_mapping['SKU'][line]:
                maje_dic = {}
                sandro_dic = {}
                SKU_error = []
                D_line = {}
                D_line_san = {}
                count = 0
                count_san = 0
                orderDicMaje[maje_mapping['path'][line][89:101]] = ''
                orderDicSandro[maje_mapping['path'][line][89:101]] = ''
    
                maje_dic['ALLOW_CHARGE'] = 'ALLOW_CHARGE|0||0|0|'
                for line_str in open(maje_mapping['path'][line]).readlines():
                    if line_str.startswith('HEADING'):
                        maje_dic['HEADING'] = line_str

                    elif line_str.startswith('TERMS'):
                        maje_dic['TERMS'] = line_str
                    elif line_str.startswith('REF'):
                        maje_dic['REF'] = line_str
                    elif line_str.startswith('DETAILS') and '360717' in line_str:
                        count+=1
                        SKU_error.append(line_str.split('|')[-2])
                        D_line[f'{count}'] = line_str.split('|')[3:-1]
                        '''maje_file_dic[count] ="""{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}{DETAILS}|1|{DETAILS2}""".format(HEADING = maje_dic['HEADING'],ALLOW_CHARGE = maje_dic['ALLOW_CHARGE'],TERMS = maje_dic['TERMS'],REF = maje_dic['REF'],DETAILS = line_str[:20],DETAILS2 = line_str[23:].strip())'''
                    elif line_str.startswith('DETAILS') and '366160' in line_str:
                        count_san+=1
                        D_line_san[f'{count_san}'] = line_str.split('|')[3:-1]
    
    
                orderDicMaje[maje_mapping['path'][line][89:101]] = """{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}""".format(HEADING = maje_dic['HEADING'],ALLOW_CHARGE = maje_dic['ALLOW_CHARGE'],TERMS = maje_dic['TERMS'],REF = maje_dic['REF'])
                orderDicSandro[maje_mapping['path'][line][89:101]] = """{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}""".format(HEADING = maje_dic['HEADING'],ALLOW_CHARGE = maje_dic['ALLOW_CHARGE'],TERMS = maje_dic['TERMS'],REF = maje_dic['REF'])
    
                for line_number in D_line:
                    D_line[line_number] = 'DETAILS|{orderNumber}|{lineNumber}|{qty1}|{qty2}|{SKU}|'.format(orderNumber = maje_mapping['path'][line][89:101],
                                                                                                                  lineNumber = line_number,
                                                                                                                  qty1 =D_line[line_number][0],
                                                                                                                  qty2 = D_line[line_number][1],
                                                                                                                  SKU = D_line[line_number][2])
                    if line_number == '1':
                        orderDicMaje[maje_mapping['path'][line][89:101]] = """{initial}{detail}""".format(initial = orderDicMaje[maje_mapping['path'][line][89:101]], detail = D_line[line_number])
                    else:
                        orderDicMaje[maje_mapping['path'][line][89:101]] = """{initial}
{detail}""".format(initial = orderDicMaje[maje_mapping['path'][line][89:101]], detail = D_line[line_number])
                
                for line_number in D_line_san:
                    D_line_san[line_number] = 'DETAILS|{orderNumber}|{lineNumber}|{qty1}|{qty2}|{SKU}|'.format(orderNumber = maje_mapping['path'][line][89:101],
                                                                                                                  lineNumber = line_number,
                                                                                                                  qty1 =D_line_san[line_number][0],
                                                                                                                  qty2 = D_line_san[line_number][1],
                                                                                                                  SKU = D_line_san[line_number][2])
                    if line_number == '1':
                        orderDicSandro[maje_mapping['path'][line][89:101]] = """{initial}{detail}""".format(initial = orderDicSandro[maje_mapping['path'][line][89:101]], detail = D_line_san[line_number])
                    else:
                        orderDicSandro[maje_mapping['path'][line][89:101]] = """{initial}
{detail}""".format(initial = orderDicSandro[maje_mapping['path'][line][89:101]], detail = D_line_san[line_number])
                
    
                # sys.exit('bla')
                # print(maje_mapping['path'][line])--------------------------------------------------------------------------------------------
                # print(maje_mapping['SKU'][line])--------------------------------------------------------------------------------------------
                orders_resent_summary = orders_resent_summary.append({'orderNumber':maje_mapping['path'][line][89:101],#line_str.split('|')[1],
                                                                      'SKUs':SKU_error},ignore_index=True)
                                                                                                    # 'fileName':maje_mapping['path'][line][68:],

    sandro_file_dic = {}

    for line in range(0,len(sandro_mapping)):
        dt_object_str = sandro_mapping['path'][line][61:69]
        if datetime.strptime(sandro_mapping['path'][line][61:69], '%d%m%Y') >= date_check:
            if '366160' in sandro_mapping['SKU'][line]:
                sandro_dic = {}
                maje_dic = {}
                SKU_error = []
                D_line = {}
                D_line_maj = {}
                count = 0
                count_maj = 0
                sandro_dic['ALLOW_CHARGE'] = 'ALLOW_CHARGE|0||0|0|'
                orderDicSandro[sandro_mapping['path'][line][93:105]] = ''
                orderDicMaje[sandro_mapping['path'][line][93:105]] = ''
                # sys.exit('bla')
                for line_str in open(sandro_mapping['path'][line]).readlines():
                    if line_str.startswith('HEADING'):
                        sandro_dic['HEADING'] = line_str
                        # sandro_orders_resent_summary.append(line_str.split('|')[3])
                        # sandro_orders_resent_summary[line_str.split('|')[3]] = sandro_mapping['path'][line]
                        
                    elif line_str.startswith('TERMS'):
                        sandro_dic['TERMS'] = line_str
                    elif line_str.startswith('REF'):
                        sandro_dic['REF'] = line_str
                    elif line_str.startswith('DETAILS') and '366160' in line_str:
                        count+=1
                        SKU_error.append(line_str.split('|')[-2])
                        D_line[f'{count}'] = line_str.split('|')[3:-1]
                        '''sandro_file_dic[count] ="""{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}{DETAILS}|1|{DETAILS2}""".format(HEADING = sandro_dic['HEADING'],ALLOW_CHARGE = sandro_dic['ALLOW_CHARGE'],TERMS = sandro_dic['TERMS'],REF = sandro_dic['REF'],DETAILS = line_str[:20],DETAILS2 = line_str[23:].strip())'''
                    elif line_str.startswith('DETAILS') and '360717' in line_str:
                        count_maj+=1
                        D_line_maj[f'{count_maj}'] = line_str.split('|')[3:-1]
                        
                orderDicSandro[sandro_mapping['path'][line][93:105]] = """{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}""".format(HEADING = sandro_dic['HEADING'],ALLOW_CHARGE = sandro_dic['ALLOW_CHARGE'],TERMS = sandro_dic['TERMS'],REF = sandro_dic['REF'])
                orderDicMaje[sandro_mapping['path'][line][93:105]] = """{HEADING}{ALLOW_CHARGE}                    
{TERMS}{REF}""".format(HEADING = sandro_dic['HEADING'],ALLOW_CHARGE = sandro_dic['ALLOW_CHARGE'],TERMS = sandro_dic['TERMS'],REF = sandro_dic['REF'])
    
                for line_number in D_line:
                    D_line[line_number] = 'DETAILS|{orderNumber}|{lineNumber}|{qty1}|{qty2}|{SKU}|'.format(orderNumber = sandro_mapping['path'][line][93:105],
                                                                                                                  lineNumber = line_number,
                                                                                                                  qty1 =D_line[line_number][0],
                                                                                                                  qty2 = D_line[line_number][1],
                                                                                                                  SKU = D_line[line_number][2])
                    if line_number == '1':
                        orderDicSandro[sandro_mapping['path'][line][93:105]] = """{initial}{detail}""".format(initial = orderDicSandro[sandro_mapping['path'][line][93:105]], detail = D_line[line_number])
                    else:
                        orderDicSandro[sandro_mapping['path'][line][93:105]] = """{initial}
{detail}""".format(initial = orderDicSandro[sandro_mapping['path'][line][93:105]], detail = D_line[line_number])
    
    
                for line_number in D_line_maj:
                    D_line_maj[line_number] = 'DETAILS|{orderNumber}|{lineNumber}|{qty1}|{qty2}|{SKU}|'.format(orderNumber = sandro_mapping['path'][line][93:105],
                                                                                                                  lineNumber = line_number,
                                                                                                                  qty1 =D_line_maj[line_number][0],
                                                                                                                  qty2 = D_line_maj[line_number][1],
                                                                                                                  SKU = D_line_maj[line_number][2])
                    if line_number == '1':
                        orderDicMaje[sandro_mapping['path'][line][93:105]] = """{initial}{detail}""".format(initial = orderDicMaje[sandro_mapping['path'][line][93:105]], detail = D_line_maj[line_number])
                    else:
                        orderDicMaje[sandro_mapping['path'][line][93:105]] = """{initial}
{detail}""".format(initial = orderDicMaje[sandro_mapping['path'][line][93:105]], detail = D_line_maj[line_number])
    
    
    
                #sys.exit('bla')
                # print(sandro_mapping['path'][line])--------------------------------------------------------------------------------------------
                # print(sandro_mapping['SKU'][line])--------------------------------------------------------------------------------------------
                orders_resent_summary = orders_resent_summary.append({'orderNumber':sandro_mapping['path'][line][93:105],#line_str.split('|')[1],
                                                                      'SKUs':SKU_error},ignore_index=True)
                                                                                  # 'fileName':sandro_mapping['path'][line][70:],
    alread_resent_maje = list(pd.read_excel(r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Maje\already resent.xlsx', engine = 'openpyxl')['order'])
    alread_resent_sandro = list(pd.read_excel(r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Sandro\already resent.xlsx', engine = 'openpyxl')['order'])
    

    
    
    print('-----------------------------------------Creation of files to be re-resent-------------------------------------------------')
    
    newly_resent_maje = []
    newly_resent_sandro = []
    
    if len(orderDicMaje) >0:
        for file_number in orderDicMaje:
                # order_number = maje_file_dic[file_number].split('|')[3]
                if str(orderDicMaje[file_number].split('|')[3]) not in alread_resent_sandro:
                    print(str(orderDicMaje[file_number].split('|')[3]))
                    Path(r"C:\Users\dbouvier\Documents\output file\Sandro\{0}".format(datetime.today().strftime('%m%d%y'))).mkdir(parents=True, exist_ok=True)
    

                    with open(r'C:\Users\dbouvier\Documents\output file\Sandro\{date}\850_Sandro_SAKSDS_4196_{L1}_{L2}.txt'.format(L1=orderDicMaje[file_number].split('|')[3]
                                                                                                                                   ,L2=datetime.today().strftime('%d%H%M%S')
                                                                                                                                   , date = datetime.today().strftime('%m%d%y')),'w') as text_file:
                        text_file.write(orderDicMaje[file_number].replace('|MAJ|','|SAN|').replace('|4596|','|4196|'))
                    time.sleep(1)
                    alread_resent_sandro.append(str(orderDicMaje[file_number].split('|')[3]))
                    newly_resent_sandro.append(str(orderDicMaje[file_number].split('|')[3]))
    
    if len(orderDicSandro)>0:
        for file_number in orderDicSandro:
            # order_number = sandro_file_dic[file_number].split('|')[3]
            if str(orderDicSandro[file_number].split('|')[3]) not in alread_resent_maje:
                print(str(orderDicSandro[file_number].split('|')[3]))
                Path(r"C:\Users\dbouvier\Documents\output file\Maje\{0}".format(datetime.today().strftime('%m%d%y'))).mkdir(parents=True, exist_ok=True)

                with open(r'C:\Users\dbouvier\Documents\output file\Maje\{date}\850_Maje_SAKSDS_4596_{L1}_{L2}.txt'.format(L1=orderDicSandro[file_number].split('|')[3]
                                                                                                                    ,L2=datetime.today().strftime('%d%H%M%S')
                                                                                                                    ,date = datetime.today().strftime('%m%d%y')),'w') as text_file:
                    text_file.write(orderDicSandro[file_number].replace('|SAN|','|MAJ|').replace('|4196|','|4596|'))
                alread_resent_maje.append(str(orderDicSandro[file_number].split('|')[3]))
                newly_resent_maje.append(str(orderDicSandro[file_number].split('|')[3]))
                time.sleep(1)
                
    for brand in ['Maje','Sandro']:
        if str(datetime.today().strftime('%m%d%y')) in os.listdir(r'C:\Users\dbouvier\Documents\output file\{brand}'.format(brand=brand)):
          for file in os.listdir(r'C:\Users\dbouvier\Documents\output file\{brand}\{date}'.format(brand=brand, date = datetime.today().strftime('%m%d%y'))):
              shutil.copyfile(r'C:\Users\dbouvier\Documents\output file\{brand}\{date}\{file}'.format(brand=brand, date = datetime.today().strftime('%m%d%y'), file=file), r'\\\app\EDI_USA\DROPSHIP\SAKS\{brand}\Arrivee\{file}'.format(brand=brand,file=file))
    
    to_compare = orders_resent_summary.copy()
    orders_resent_summary = pd.DataFrame(columns = ['orderNumber','SKUs'])#,'SKUs'])
    
    to_compare.drop_duplicates(subset ="orderNumber",
                          keep = "first", inplace = True)
    to_compare.set_index('orderNumber',inplace=True)
    
    for order_number in newly_resent_maje:
        orders_resent_summary = orders_resent_summary.append({'orderNumber':order_number,
                                                              'SKUs':to_compare['SKUs'][order_number]},ignore_index=True)
    for order_number in newly_resent_sandro:
        orders_resent_summary = orders_resent_summary.append({'orderNumber':order_number,
                                                              'SKUs':to_compare['SKUs'][order_number]},ignore_index=True)

    orders_resent_summary.drop_duplicates(subset=['orderNumber'],inplace=True)
    alread_resent_sandro = pd.DataFrame(alread_resent_sandro, columns = ['order'])
    alread_resent_maje = pd.DataFrame(alread_resent_maje, columns = ['order'])
    
    
    #----------------------comment these if you want your script not to generte previously sent files------------------------
    writer = ExcelWriter(r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Sandro\already resent.xlsx')
    alread_resent_sandro.to_excel(writer,'{0}'.format(datetime.today().strftime('%m%d%y')),index = False)
    writer.save()
    writer.close()
    
    
    writer = ExcelWriter(r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\850\Maje\already resent.xlsx')
    alread_resent_maje.to_excel(writer,'{0}'.format(datetime.today().strftime('%m%d%y')),index = False)
    writer.save()
    writer.close()
    #----------------------------------------------------------------------------------------------------------------------------
    
    
    print(orders_resent_summary)                                                                      

    if len(orders_resent_summary) >0:
        outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize())
        mail = outlook.CreateItem(0)
        mail.To = 
        mail.CC = 
        mail.Subject = 'Automated re-integration of 850 files'
        mail.Body = """***Automated Email***
        
Good Morning Team,
 
The following list of {0} orders had a brand SKU discrepancy in the 850 file which have been taken care of:

{1}

Best,
Dimitri Bouvier
IT Systems Analyst, North America
     
Sandro • Maje • Claudie Pierlot • De Fursac
44 Wall Street
New York, NY 10005
""".format(len(orders_resent_summary),tabulate(orders_resent_summary, headers = ['orderNumber','SKUs'],  tablefmt='psql'))
        mail.Display(False)    

    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                    'status':'done'},ignore_index=True)
def all945s(task):
    def retPath(brand,folder_2):
        # print(f'start of 945 RET mapping for {brand} {folder_2}')
        print(f'start of 945 RET path mapping for {brand}')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\945\{brand}\arrivee'.format(brand=brand.upper())
        # print(r'{path}\{folder_2}_945_mapping.xlsx'.format(path = path, folder_2=folder_2))
        df = pd.read_excel(r'{path}\{folder_2}_945_mapping.xlsx'.format(path = path, folder_2=folder_2), engine = 'openpyxl')
        increment = 0
        for folder in os.listdir(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE".format(brand=brand.upper())): 
            if folder != "ECOM" and datetime.strptime(folder, "%d%m%Y") > datetime.today() - timedelta(days=14):
                # print(folder)
                for file in os.listdir(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}".format(brand=brand,folder=folder)):
                    increment +=1
                    # if increment % 10000 == 0:
                        # print(increment)
                    if file.startswith('SMCP') and '945_RET' in file and r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand.upper(),folder=folder,file=file) not in list(df['path'].astype(str)):
                        # paths_list.append(r"\\\app\PANALPINA\MAJE\SAV-ARRIVEE\{0}\{1}".format(folder,file))
                        # print(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand.upper(),folder=folder,file=file))
                        try:
                            for a in open(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand.upper(),folder=folder,file=file)).readlines():
                                if a.startswith('SMCP945H2') and a[find_str(a, '{brand}-'.format(brand=brand.upper()[:4])):find_str(a, '{brand}-'.format(brand=brand.upper())[:4])+12] not in list(df["PT"]):
                                    # print('{brand}-'.format(brand=brand.upper()[:4]))
                                    df = df.append({'PT':a[find_str(a, '{brand}-'.format(brand=brand.upper()[:4])):find_str(a, '{brand}-'.format(brand=brand.upper()[:4]))+12]
                                                    ,'path':r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand.upper(),folder=folder,file=file)}
                                                    ,ignore_index=True)
                        except UnicodeDecodeError:
                            print(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file} IS IN ERROR".format(brand=brand.upper(),folder=folder,file=file))
                
               
        writer = pd.ExcelWriter(r'{path}\{folder_2}_945_mapping.xlsx'.format(path=path,folder_2=folder_2))
        df.to_excel(writer,'{brand}_pt'.format(brand=brand.lower()),index = False)
        writer.save()
        print(f'end of 945 RET path mapping for {brand}')
    def dropLines(brand):
        print(f'start of 945 DROP lines mapping for {brand}')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\945\{brand}\arrivee\DROP'.format(brand=brand.upper())
        df = pd.read_excel(r'{0}\945_mapping_lines.xlsx'.format(path), engine = 'openpyxl')
        PTs_lines_discrepency = []
        path_list = list(set(df['Path']))
        
        for folder in os.listdir(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE".format(brand=brand.upper())):   
            if folder != "ECOM" and datetime.strptime(folder, "%d%m%Y") > datetime.today() - timedelta(days=7):
                # print(folder)
                for file in os.listdir(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}".format(brand=brand,folder=folder)):
                    if file.startswith('SMCP') and '945_DROP' in file and r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand,folder=folder,file=file) not in path_list: #list(set(df['Path'])):
                        path_dictionnary,d_dictionnary,d_dictionnary,h2_dictionnary,h1_dictionnary,h1_map,h2_map = {},{},{},{},{},{},{}
                        # print(r"Opening \\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand,folder=folder,file=file))           
                        for line in open(r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand,folder=folder,file=file)):
                        
                            if line.split('|')[0] == 'SMCP945D':
                                if line.split('|')[-1].strip() not in list(d_dictionnary.keys()):
                                    d_dictionnary[line.split('|')[-1].strip()] = [line.split('|')[4]]#[:-2]]
                                    path_dictionnary[line.split('|')[-1].strip()] = r"\\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand,folder=folder,file=file)
                                else:
                                    if line.split('|')[4] not in d_dictionnary[line.split('|')[-1].strip()]:
                                        d_dictionnary[line.split('|')[-1].strip()].append(line.split('|')[4])
                        
                            elif line.split('|')[0] == 'SMCP945H2':
                                if line.split('|')[8].strip() not in list(h2_dictionnary.keys()):
                                    h2_dictionnary[line.split('|')[8].strip()] = [line.split('|')[6]]
                                else:
                                    if line.split('|')[6] not in h2_dictionnary[line.split('|')[8].strip()]:
                                        h2_dictionnary[line.split('|')[8].strip()].append(line.split('|')[6])
                                if line.split('|')[8] not in list(h2_map.keys()):
                                    h2_map[line.split('|')[8]] = [int(line.split('|')[1])]
                                else:
                                    h2_map[line.split('|')[8]].append(int(line.split('|')[1]))  
                            elif line.split('|')[0] == 'SMCP945H1':
                                h1_map[int(line.split('|')[1])] = line.split('|')[5]
                                
                    try:
                        for pt in h2_map:
                            for line_number in h2_map[pt]:
                                if pt not in list(h1_dictionnary.keys()):
                                    h1_dictionnary[pt] = [h1_map[line_number]]
                                else:
                                    h1_dictionnary[pt].append(h1_map[line_number])
        
                
                        for pt in h2_map:
                            if pt not in list(df['PT']):
                                try:
                                    df = df.append({'PT':pt,
                                                    'H1':h1_dictionnary[pt],
                                                    'H2':h2_dictionnary[pt],
                                                    'D':d_dictionnary[pt],
                                                    'Path':path_dictionnary[pt]}, ignore_index=True)
                                except KeyError:
                                    PTs_lines_discrepency.append(pt)
                                    print(f"PT in error - {0} from the following path: \\\app\PANALPINA\{brand}\SAV-ARRIVEE\{folder}\{file}".format(brand=brand,folder=folder,file=file))
                    except NameError:
                        pass
        writer = pd.ExcelWriter(r'{0}\945_mapping_lines.xlsx'.format(path))
        df.to_excel(writer, '{brand}_drop'.format(brand=brand.lower()),index=False)  
        writer.save()
        print(f'end of 945 DROP lines mapping for {brand}')
    def retLines(brand,folder_2):
        print(f'start of 945 RET lines mapping for {brand}')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\945\{brand}\arrivee'.format(brand=brand)
        path_mapping_df = pd.read_excel(f"{path}\{folder_2}_945_mapping.xlsx", engine = 'openpyxl')
        line_mapping_df = pd.read_excel(f"{path}\945_mapping_lines.xlsx", engine = 'openpyxl')
        h1_dictionnary = {}
        h2_dictionnary = {}
        d_dictionnary = {}
        path_dictionnary = {}
        PTs_lines_discrepency = []
        counter = 0
        for a in range(0,len(path_mapping_df)):
            # if (path_mapping_df['PT'][a] not in list(line_mapping_df['PT']))  and '945_RET' in path_mapping_df['path'][a]:
            if '945_RET' in path_mapping_df['path'][a] and (path_mapping_df['PT'][a] not in list(line_mapping_df['PT'])):#'945_RET' in path_mapping_df['path'][a] and 
                h1_map,h2_map = {},{}
                counter += 1
                for line in open(path_mapping_df['path'][a]):
                    if line.split('|')[0] == 'SMCP945D':
                        if line.split('|')[-1].strip() not in list(d_dictionnary.keys()):
                            d_dictionnary[line.split('|')[-1].strip()] = [line.split('|')[4]]#[:-2]]
                            path_dictionnary[line.split('|')[-1].strip()] = path_mapping_df['path'][a]
                        else:
                            if line.split('|')[4] not in d_dictionnary[line.split('|')[-1].strip()]:
                                d_dictionnary[line.split('|')[-1].strip()].append(line.split('|')[4])
                
                    elif line.split('|')[0] == 'SMCP945H2':
                        if line.split('|')[8].strip() not in list(h2_dictionnary.keys()):
                            h2_dictionnary[line.split('|')[8].strip()] = [line.split('|')[6]]
                        else:
                            if line.split('|')[6] not in h2_dictionnary[line.split('|')[8].strip()]:
                                h2_dictionnary[line.split('|')[8].strip()].append(line.split('|')[6])
                        if line.split('|')[8] not in list(h2_map.keys()):
                            h2_map[line.split('|')[8]] = [int(line.split('|')[1])]
                        else:
                            h2_map[line.split('|')[8]].append(int(line.split('|')[1]))  
                    elif line.split('|')[0] == 'SMCP945H1':
                        h1_map[int(line.split('|')[1])] = line.split('|')[5]

                for pt in h2_map:
                    if pt == path_mapping_df['PT'][a]:
                        try:
                            for line_number in h2_map[pt]:
                                if pt not in list(h1_dictionnary.keys()):
                                    h1_dictionnary[pt] = [h1_map[line_number]]
                                else:
                                    h1_dictionnary[pt].append(h1_map[line_number])
                        except KeyError:
                            print(path_mapping_df['PT'][a]+" - "+path_mapping_df['path'][a]+'is in error please check it out')
                try:
                    line_mapping_df = line_mapping_df.append({'PT':path_mapping_df['PT'][a],
                                                              'H1':h1_dictionnary[path_mapping_df['PT'][a]],
                                                              'H2':h2_dictionnary[path_mapping_df['PT'][a]],
                                                              'D':d_dictionnary[path_mapping_df['PT'][a]],
                                                             'Path':path_mapping_df['path'][a]}, ignore_index=True)
                except KeyError:
                    PTs_lines_discrepency.append(path_mapping_df['PT'][a])
                    print(f"PT in error - {path_mapping_df['PT'][a]}")
        
        writer = pd.ExcelWriter(r'{0}\945_mapping_lines.xlsx'.format(path))
        line_mapping_df.to_excel(writer, datetime.today().strftime('%m%d%Y'),index=False)  
        writer.save()
        print(f'end of 945 RET lines mapping for {brand}')

    global df_tasks
    print(f'start of {task}')
    folder_dic_brand = {'Maje':'MA','Sandro':'SA'}
    mapping_dic_ret = {}
    mapping_lines_drop = {}
    mapping_dic_ret_lines = {}
    for brand in folder_dic_brand:
        mapping_dic_ret[brand] = threading.Thread(target = retPath, args =[brand,folder_dic_brand[brand]])
        mapping_lines_drop[brand] = threading.Thread(target = dropLines, args =[brand])
        mapping_dic_ret_lines[brand] = threading.Thread(target = retLines, args =[brand,folder_dic_brand[brand]])
        
    for thread in mapping_dic_ret:
        mapping_dic_ret[thread].start()
        mapping_lines_drop[thread].start()
        
    for thread in mapping_dic_ret:
        mapping_dic_ret[thread].join()
    
    for brand in folder_dic_brand:
        mapping_dic_ret_lines[brand].start()
    
    for thread in mapping_lines_drop:
        mapping_lines_drop[thread].join()
        mapping_dic_ret_lines[thread].join()
    
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)
def allEcom(task):
    global df_tasks
    print(f'start of {task}')
    brands_list = ['Maje','Sandro']
    def departOrders(brand):
        print(f'start of USSentOrders mapping for {brand}')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\Demandware\{brand}\depart'.format(brand=brand)
        df = pd.read_excel("{path}\{brand}_DMW_orders_mapping.xlsx".format(path=path,brand=brand), engine = 'openpyxl')
        paths_list = list(set(df['filePath']))
        
        for folder_or_file in os.listdir(r"\\\app\DMW\{brand}\SAV-Depart".format(brand=brand)):
            if folder_or_file.startswith('USSentOrders') and r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file) not in paths_list:
                paths_list.append(r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file))
                root = ET.parse(r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file)).getroot();
                for returnnode in root.findall(".//order"):
                    for orderNumFile in returnnode.findall(".//orderNumber"):
                        # if orderNumFile.text not in list(df['orderNumber']):
                        if int(orderNumFile.text) not in list(df['orderNumber']) and str(orderNumFile.text) not in list(df['orderNumber']):
                            # print(orderNumFile.text)
                            # print(type(orderNumFile.text))
                            df = df.append({'orderNumber':orderNumFile.text
                                            ,'filePath':r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file)
                                            ,'xmlStr':ET.tostring(returnnode, encoding='utf-8').decode('utf-8')},ignore_index=True)
            
            elif re.match(r"[0-9]+", folder_or_file) and len(folder_or_file) == 8:
                # print(folder_or_file)
                for file in os.listdir(r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file)):
                    if file.startswith('USSentOrders') and r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file) not in paths_list:
                        paths_list.append(r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file))
                        root = ET.parse(r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file)).getroot()
                        for returnnode in root.findall(".//order"):
                            for orderNumFile in returnnode.findall(".//orderNumber"):
                                # if orderNumFile.text not in list(df['orderNumber']):
                                if int(orderNumFile.text) not in list(df['orderNumber']) and str(orderNumFile.text) not in list(df['orderNumber']):
                                    # print(orderNumFile.text)
                                    # print(type(orderNumFile.text))
                                    df = df.append({'orderNumber':orderNumFile.text
                                                    ,'filePath':r"\\\app\DMW\{brand}\SAV-Depart\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file)
                                                    ,'xmlStr':ET.tostring(returnnode, encoding='utf-8').decode('utf-8')},ignore_index=True)
        writer = pd.ExcelWriter(r'{path}\{brand}_DMW_orders_mapping.xlsx'.format(path=path,brand=brand))
        df.to_excel(writer, "{brand}_DMW_orders_mapping".format(brand=brand),index=False)    
        writer.save()
        print(f'end of USSentOrders mapping for {brand}')
    def arriveeTickets(brand):
        print(f'start of USTickets mapping for {brand}')
        path = r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\Demandware\{brand}\arrivee'.format(brand=brand)
        df = pd.read_excel(r'{path}\{brand}_DMW_tickets_mapping.xlsx'.format(path=path,brand=brand), engine = 'openpyxl')
        parser = etree.XMLParser(recover=True)
        
        paths_list = list(set(df['filePath'].astype('str')))
        paths_list_new = []
        for string in paths_list:
            paths_list_new.append(string[find_str(string,'USTicket_'):find_str(string,'.xml')+4])
        
        for folder_or_file in os.listdir(r"\\\app\DMW\{brand}\SAV-Arrivee".format(brand=brand)):
            if folder_or_file.startswith('USTicket') and folder_or_file[find_str(folder_or_file,'USTicket_'):find_str(folder_or_file,'.xml')+4] not in paths_list_new:
                # print(r"\\\app\DMW\MAJE\SAV-Arrivee\{0}".format(folder_or_file))
                root =  open(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file), "r").read()
                doc_file = etree.fromstring(root, parser=parser)
                paths_list.append(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file))
                for returnnode in doc_file.findall(".//Ticket"):
                    try:
                        orderNumber = returnnode.find(".//Paiement/NoDePieceIdentite").text
                    except AttributeError:
                        orderNumber = 'null'
                    if str(returnnode.find(".//NoTicket").text) not in list(df['NoTicket'].astype('str')):
                        # print(returnnode.find(".//NoTicket").text)
                        df = df.append({'NoTicket':returnnode.find(".//NoTicket").text
                                        ,'orderNumber':orderNumber
                                        ,'filePath':r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file)
                                        ,'xmlStr':etree.tostring(returnnode, encoding='UTF-8').decode('utf-8')},ignore_index=True)
        
                   
                            
            elif re.match(r"[0-9]+", folder_or_file) and len(folder_or_file) == 8:
                for file in os.listdir(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}".format(brand=brand,folder_or_file=folder_or_file)):
                    if file.startswith('USTicket') and file[find_str(file,'USTicket_'):find_str(file,'.xml')+4] not in paths_list_new:
                        # print(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file))
                        root =  open(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file), "r").read()
                        doc_file = etree.fromstring(root, parser=parser)
                        paths_list.append(r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file))
                        for returnnode in doc_file.findall(".//Ticket"):
                            for orderNumFile in returnnode.findall(".//NoTicket"):
                                try:
                                    orderNumber = returnnode.find(".//Paiement/NoDePieceIdentite").text
                                except AttributeError:
                                    orderNumber = 'null'
                                if str(returnnode.find(".//NoTicket").text) not in list(df['NoTicket'].astype('str')):
                                    # print(returnnode.find(".//NoTicket").text)
                                    df = df.append({'NoTicket':returnnode.find(".//NoTicket").text
                                        ,'orderNumber':orderNumber
                                        ,'filePath':r"\\\app\DMW\{brand}\SAV-Arrivee\{folder_or_file}\{file}".format(brand=brand,folder_or_file=folder_or_file,file=file)
                                        ,'xmlStr':etree.tostring(returnnode, encoding='UTF-8').decode('utf-8')},ignore_index=True)
        writer = pd.ExcelWriter(r'C:\Users\dbouvier\OneDrive - SMCP\11 - IT - Operations\02 - Applications Operations\Mapping Tables\Demandware\{brandUpper}\arrivee\{brand}_DMW_tickets_mapping.xlsx'.format(brandUpper=brand.upper(),brand=brand))
        df.to_excel(writer, "{brand}_DMW_tickets_mapping",index=False)    
        writer.save()

        print(f'end of USTickets mapping for {brand}')
    ecom_tasks_dic = {}
    
    for brand in brands_list:
        ecom_tasks_dic[brand+'orders'] = threading.Thread(target = departOrders, args =[brand])
        ecom_tasks_dic[brand+'tickets'] = threading.Thread(target = arriveeTickets, args =[brand])
    
    for thread in ecom_tasks_dic:
        ecom_tasks_dic[thread].start()

    for thread in ecom_tasks_dic:
        ecom_tasks_dic[thread].join()
    
    print(f'end of {task}')
    with lock:
        df_tasks = df_tasks.append({'task':task,
                                   'status':'done'},ignore_index=True)

main_dic['wunderkind'] = threading.Thread(target = wunderkind, args =['wunderkind'])
main_dic['transferFileChecking'] = threading.Thread(target = transferFileChecker, args =['transferFileCheck'])
# main_dic['blmAudStlComparison'] = threading.Thread(target = blmAudStlComparision, args =['blmAudStlComparison'])
main_dic['blmAudDaiComparison'] = threading.Thread(target = blmAudDaiComparison, args =['blmAudDaiComparison'])
main_dic['backUpMagCheck'] = threading.Thread(target = backUpMagCheck, args =['backUpMagCheck'])
main_dic['dailyChecklistQueriesCheck'] = threading.Thread(target = dailyChecklistCheck, args =['dailyChecklistQueriesCheck'])
main_dic['all850s'] = threading.Thread(target = all850s, args =['all850s'])
main_dic['all945s'] = threading.Thread(target = all945s, args =['all945s'])
main_dic['allEcom'] = threading.Thread(target = allEcom, args =['allEcom'])

for thread in main_dic:
    main_dic[thread].start()

for thread in main_dic:
    main_dic[thread].join()

# sys.exit('abc')

df_tasks.sort_values(by=['task'],ascending=True).to_html(r'{folder_path}\DailyTasksStatus.html'.format(folder_path=r'C:\Users\dbouvier\Downloads')
                            ,index=False
                            ,index_names=False
                            ,na_rep='')


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 
# mail.CC = 
mail.Subject = 'Daily automated tasks report IS Team'
mail.HtmlBody = """***Automated Email***<br><br>

Good morning IS Team,<br><br>
 
Here's today's report on all automated IS tasks ran this morning and their outcome:<br><br>

{report}<br><br>
    
Best,<br>
Dimitri Bouvier<br>
IT Systems Analyst, North America<br><br>
     
Sandro • Maje • Claudie Pierlot • De Fursac<br>
44 Wall Street<br>
New York, NY 10005<br>
""".format(report = open(r'{folder_path}\DailyTasksStatus.html'.format(folder_path=r'C:\Users\dbouvier\Downloads')).read())
# mail.Attachments.Add(r'C:/Users/dbouvier/Downloads/SentOrdersWithoutMatch{0}.csv'.format(datetime.today().strftime('%m%d%y')))
mail.Display(False)


print('end of program')
