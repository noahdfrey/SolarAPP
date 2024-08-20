# -*- coding: utf-8 -*-
"""
Created on Wed May  3 15:42:15 2023

@author: nfrey
"""

import webbrowser
import os
import glob
import shutil
import time
from datetime import datetime, timedelta
import pandas as pd

# This Python script aims to download and summarize all of the current 
# registrant info from the SolarAPP website to create tables for the weekly 
# AHJ Digest.

# You will need to download the module named 'xlsxwriter'

# to run this code, you will need four inputs:

# 1. date of the last AHJ Digest (written in 'YYYY_MM_DD' format).

last_digest_date = '2024_03_11'

# 2. the newest AHJ ID (found in the admin portal of the SolarAPP website).

newest_ahj_id = 11992

# 3. path to your downloads folder.

download = 'C:/Users/nfrey/Downloads'

# 4. path to a folder holding AHJ Digest information.

digest_folder = 'C:/Users/nfrey/OneDrive - NREL/AHJ Digest'

# in addition to these four inputs, you will need to do three tasks:
    # 1. log in to your SolarAPP admin account in your default browser.
    # 2. clear out your downloads folder.
    # 3. make sure this python file is in your digest_folder.

# with these four inputs and the three tasks complete, the following code 
# should run from your computer without issues:

## STEP 1: Create registrant_info file of all current registrant info    

# read in the registrant_info csv file from the last digest.

old_registrants = pd.read_csv(digest_folder + '/' + last_digest_date + '/' + 
                              last_digest_date + '_Registrant_Info.csv')

# create a list of all AHJ IDs found in the old registrant_info file.

ahj_ids = old_registrants['AHJ_ID'].tolist()

# sort the list of ahj_ids (ascending order).

ahj_ids.sort()

# create a list of the possible AHJ IDs that may be in use since the last AHJ
# digest. This is a list ranging from the number after the final AHJ ID in the
# last digest and the number after the newest AHJ ID.

possible_ids = list(range((ahj_ids[-1] + 1), (newest_ahj_id + 1)))

# add these additional possible AHJ IDs to the list of old AHJ IDs.

ahj_ids.extend(possible_ids)

# if downloading the files stops for some reason, you can plug in x as the 
# last downloaded file number (not AHJ ID) to the next line to pick up where 
# you left off.

#ahj_ids = ahj_ids[x:]

# create an empty list to store all of the download urls.

urls = []

# loop through the list of all possible ids to create a list of unique urls.

for ids in ahj_ids:
    link = 'https://solarapp.nrel.gov/admin/ahjs/' + str(ids) + '/export'
    urls.append(link)
    
# set your web browser to the default (I use chrome).

browser = webbrowser.get()

# call todays date

today = datetime.today().strftime('%Y_%m_%d')

# if you need yesterdays date, you can use this line

#today = (datetime.today() - timedelta(days=1)).strftime('%Y_%m_%d')

# path to destination folder

destination = digest_folder + '/' + today + '/' + today + '_Registration_Files'

# loop through the list of download urls to open each one in your browser.
# This loop is the equivalent of selecting each AHJ and clicking the "Export
# Onboarding Info" prompt to download that AHJs registrant info. It also moves
# the downloaded file into your AHJ Digest folder. This section of code is the 
# most intensive and takes ~40 minutes to run

for url in urls:
    browser.open(url)
    time.sleep(5)
    shutil.copytree(download, destination, copy_function = shutil.move,
    dirs_exist_ok=True)
    
# read all downloaded csv files into Python
    
all_files = glob.glob(os.path.join(destination, '*.csv'))

# create an empty dataframe to hold all relevant registration information

registrant_info = pd.DataFrame()

# loop through all downloaded files to pull all relevant registration
# information and add it to the "registrant_info" dataframe

for file in all_files:
    df = pd.read_csv(file, usecols=['description', 'value'])
    df = df[df['value'] != 'complete']
    df = df.transpose()
    df.columns = df.iloc[0]
    df['completed1'] = 'completed'
    df['completed2'] = 'completed'
    df['completed3'] = 'completed'
    df['completed4'] = 'completed'
    df['completed5'] = 'completed'
    column_names = df.columns.values.tolist()
    where_registrant_stopped = column_names[32]
    df.iloc[1, 32] = where_registrant_stopped
    df = df.rename(columns = {where_registrant_stopped:
                              'where_registrant_stopped'})
    df = df[['AHJ ID', 'AHJ Name', 'Onboarding completed', '"Mode"',
             'Progress', 'integration_mode', 'local_settings',
             'integration_setup', 'terms_and_conditions',
             'where_registrant_stopped']]
    df = df[df['AHJ ID'] != 'AHJ ID']
    registrant_info = pd.concat([registrant_info, df], ignore_index=True)

# rename ahj_id column in the registrant_info dataframe.

registrant_info = registrant_info.rename(columns={'AHJ ID': 'AHJ_ID'})

# specify int data type for ahj_id column

registrant_info['AHJ_ID'] = registrant_info['AHJ_ID'].astype(int)
    
# specify the link to the AHJ Report from SolarAPP website.

report_link = 'https://solarapp.nrel.gov/admin/ahjs/export-all'

# specify the destination for the AHJ Report.

report_destination = digest_folder + '/' + today + '/' + today + '_Report'

# download the AHJ Report and move it to the destination

browser.open(report_link)
time.sleep(20)
shutil.copytree(download, report_destination, copy_function = shutil.move,
                dirs_exist_ok=True)

# create a variable holding the new path to the ahj_report

report_file = glob.glob(os.path.join(report_destination, '*.csv'))
report_file = report_file[0]

# read the AJH Report and pull all relevant information.

ahj_report = pd.read_csv(report_file, 
                         usecols=['user_name', 'email', 'id', 'name', 'city', 
                                  'project_count', 'verified', 'signed_up_at'],
                         dtype={'id': 'int64'})

# rename id column in ahj_report dataframe
                                                   
ahj_report = ahj_report.rename(columns={'id': 'AHJ_ID'})

# merge the registrant_info and ahj_report dataframes on the ahj_id column                                                   

registrant_info = pd.merge(registrant_info, ahj_report, on=['AHJ_ID'], 
                           how='left')

# set AHJ_ID as the index of the registrant_info dataframe and sort by it

registrant_info = registrant_info.set_index('AHJ_ID').sort_index()

# format columns to match AHJ Digest

registrant_info['"Mode"'] = registrant_info['"Mode"'].str.replace('"', '')
registrant_info['user_name'] = registrant_info['user_name'].str.replace('"', '')
registrant_info['name'] = registrant_info['name'].str.replace('"', '')
registrant_info['city'] = registrant_info['city'].str.replace('"', '')
registrant_info['"Mode"'] = registrant_info['"Mode"'].str.capitalize()

# change column names to match AHJ Digest

registrant_info = registrant_info.rename(columns={'Progress': 
                                                      'Overall Progress',
                                                  'integration_mode':
                                                      'Integration Mode',
                                                  'local_settings': 
                                                      'Local Settings',
                                                  'integration_setup': 
                                                      'Integration Setup',
                                                  'terms_and_conditions':
                                                      'Terms and Conditions',
                                                  '"Mode"':
                                                      'Integration Selection',
                                                  'Onboarding completed':
                                                      'Onboarding Completed'})
    
# strip the % symbol from progress columns, convert them to floats, and round

registrant_info['Overall Progress'] = registrant_info['Overall Progress'].str.\
    strip('%').astype(float).round(2)
    
registrant_info['Local Settings'] = registrant_info['Local Settings'].str.\
    strip('%').astype(float).round(2)
    
registrant_info['Integration Setup'] = registrant_info['Integration Setup'].str.\
    strip('%').astype(float).round(2)
    
registrant_info['Terms and Conditions'] = registrant_info['Terms and Conditions'].str.\
    strip('%').astype(float).round(2)
    
registrant_info['Integration Mode'] = registrant_info['Integration Mode'].str.\
    strip('%').astype(float).round(2)
    
# remove the time from the two columns that have datetime information

registrant_info['signed_up_at'] = pd.to_datetime(registrant_info['signed_up_at']).dt.date

registrant_info['Onboarding Completed'] = pd.to_datetime(registrant_info['Onboarding Completed'], 
                                                         errors='coerce').dt.date

registrant_info['Onboarding Completed'] = registrant_info['Onboarding Completed'].fillna('Not yet completed')
    
# export the registrant_info dataframe to your AHJ Digest folder

registrant_info.to_csv(digest_folder + '/' + today + '/' + today +
                       '_Registrant_Info.csv')

## STEP 2: Create new_registrants file for all new registrants

# change the data type of the AHJ_ID column to int for the old_registrants dateframe.

old_registrants['AHJ_ID'] = old_registrants['AHJ_ID'].astype(int)

# merge the registrant_info and old_registrant dataframes using the 'AHJ_ID'
# column.

new_registrants = pd.merge(registrant_info, old_registrants, on = 'AHJ_ID',
                           how = 'left',
                           indicator = True)

# remove the AHJ IDs that are found in both dataframes.

new_registrants = new_registrants.loc[lambda x : x['_merge'] != 'both']

# drop the indicator column.

new_registrants = new_registrants.drop(columns = ['_merge'])

# drop all empty columns.

new_registrants = new_registrants.dropna(axis=1, how='all')

# remove the _x suffix from the column headers

new_registrants.columns = new_registrants.columns.str.replace('_x', '')

# set AHJ_ID as the index of the new_registrants dataframe and sort by it

new_registrants = new_registrants.set_index('AHJ_ID').sort_index()

# export the new_registrants dataframe to your AHJ Digest folder.

new_registrants.to_csv(digest_folder + '/' + today + '/' + today +
                       '_New_Registrants.csv')

## STEP 3: Create registrant_progress file showing all AHJs that have
##         progressed in registration

registrant_progress = pd.merge(registrant_info, old_registrants, 
                               on = 'AHJ_ID', 
                               suffixes=['_' + today, '_' + last_digest_date])

# make sure the Onboarding Completed columns are in string format

registrant_progress[[('Onboarding Completed_' + today), 
                     ('Onboarding Completed_' + last_digest_date)]] = registrant_progress[[('Onboarding Completed_' + today), 
                                          ('Onboarding Completed_' + last_digest_date)]].astype('str')

# select rows in which the current progress is more than old progress

registrant_progress1 = registrant_progress[registrant_progress['Overall Progress_' + today] > 
                                          registrant_progress['Overall Progress_' + last_digest_date]]

# select rows in which the new date of onboarding completed is different from
# the old date of onboarding complete

registrant_progress2 = registrant_progress[registrant_progress['Onboarding Completed_' + today] != 
                                           registrant_progress['Onboarding Completed_' + last_digest_date]]

# combine the two selected dataframes

registrant_progress = pd.concat([registrant_progress1, registrant_progress2])

# remove duplicate rows

registrant_progress = registrant_progress.drop_duplicates()

# set AHJ_ID as the index of the registrant_progress dataframe and sort by it

registrant_progress = registrant_progress.set_index('AHJ_ID').sort_index()

# export the registrant_progress dataframe to your AHJ Digest folder.

registrant_progress.to_csv(digest_folder + '/' + today + '/' + today +
                       '_Registrant_Progress.csv')

## This is the end of the current code. Everything else is still in progress

## STEP 5: Create a registrant_stage file showing a count of how many AHJs have
##         completed each registration stage

#registrant_stage = pd.DataFrame()

#int_stage = registrant_info[registrant_info['integration_mode'] == '100%']

#local_settings_stage =  registrant_info[registrant_info['local_settings'] ==
#                                        '100%']

#int_setup_stage =  registrant_info[registrant_info['integration_setup'] ==
#                                   '100%']

#terms_stage =  registrant_info[registrant_info['terms_and_conditions'] ==
#                               '100%']



# format new_registrants dataframe to fit digest format

new_registrants = new_registrants.drop(columns = ['Onboarding Completed',
                                                  'user_name',
                                                  'email',
                                                  'name',
                                                  'city',
                                                  'project_count',
                                                  'signed_up_at'])

new_registrants['2021 Rank'] = ''

new_registrants = new_registrants[['AHJ Name', '2021 Rank', 
                                  'Integration Selection', 'Integration Mode',
                                  'Local Settings', 'Integration Setup',
                                  'Terms and Conditions', 'Overall Progress']]

with pd.ExcelWriter(digest_folder + '/' + today + '/' + today +
                       '_New_Registrants.xlsx', engine = 'xlsxwriter') as writer:

    new_registrants.to_excel(writer, sheet_name='New_Registrants',
                             startrow = 1, header=False)

    workbook = writer.book
    worksheet = writer.sheets['New_Registrants']

    header_format = workbook.add_format({'bold': True,
                                         'text_wrap': True,
                                         'align': 'center',
                                         'valign': 'center',
                                         'bottom': 1,
                                         'top': 1})
    
    center_columns = workbook.add_format()
    center_columns.set_align('center')
    
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(3, 8, 11, center_columns)

    for col_num, value in enumerate(new_registrants.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
        
# format registrant_progress dataframe to fit digest format

registrant_progress['2021 Rank'] = ''

registrant_progress['Last Interaction in PD'] = ''

registrant_progress = registrant_progress[['AHJ Name_' + today, '2021 Rank', 
                                           'Overall Progress_' + last_digest_date,
                                           'Overall Progress_' + today,
                                           'Onboarding Completed_' + today,
                                           'Last Interaction in PD']]

registrant_progress = registrant_progress.rename(columns = {'AHJ Name_' + today: 
                                                            'AHJ Name',
                                                            'Onboarding Completed_' + today:
                                                                'Onboarding Completed',
                                                            'Overall Progress_' + last_digest_date:
                                                            'Overall Progress ' + last_digest_date,
                                                            'Overall Progress_' + today: 
                                                            'Overall Progress ' + today})

with pd.ExcelWriter(digest_folder + '/' + today + '/' + today +
                       '_Registrant_Progress.xlsx', engine = 'xlsxwriter') as writer:

    registrant_progress.to_excel(writer, sheet_name='Registrant_Progress',
                                 startrow = 1, header=False)

    workbook = writer.book
    worksheet = writer.sheets['Registrant_Progress']

    header_format = workbook.add_format({'bold': True,
                                         'text_wrap': True,
                                         'align': 'center',
                                         'valign': 'center',
                                         'bottom': 1,
                                         'top': 1})
    
    center_columns = workbook.add_format()
    center_columns.set_align('center')
    
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(3, 4, 14, center_columns)
    worksheet.set_column(5, 6, 18, center_columns)

    for col_num, value in enumerate(registrant_progress.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)




