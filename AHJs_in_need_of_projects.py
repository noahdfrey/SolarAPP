# -*- coding: utf-8 -*-
"""
Created on Tue May  9 12:53:59 2023

@author: nfrey
"""

import webbrowser
import pandas as pd
import numpy as np
import os
import time
import shutil
import glob
import datetime

# This Python script aims to download and identify the most recent approved 
# projects for each AHJ that is either piloting or adopting SolarAPP+.

# to run this code you will need two inputs:
    
# 1. Path to your downloads folder

download = 'C:/Users/nfrey/Downloads'

# 2. path to a folder holding most recent project information and this code

update_folder = 'C:/Users/nfrey/OneDrive - NREL/Most Recent Project'

# in addition to these two inputs, you will need to do four tasks:
    # 1. log in to your SolarAPP+ admin account in your default browser.
    # 2. clear out your downloads folder
    # 3. make sure that the "SolarAPP AHJ ID" column of the "Most Recent 
    #       Project Update" pipedrive filter is filled in completely.
    # 4. download the "Most Recent Project Update" Pipedrive filter as a csv.
    #       Don't open or move this file, it should be the only file in your 
    #       downloads folder.

# with these two inputs and the four tasks complete, the following code should 
# run from your computer without issues:
    
# specify today's date to name the file

today = datetime.datetime.today().strftime('%Y_%m_%d')

# create a new folder in your update folder with today's date.

os.makedirs(update_folder + '/' + today)

# specify the destination for the pipedrive filter csv file.

pipedrive_file_destination = update_folder + '/' + today + '/' + today + '_Pipedrive_File'

# move the file from your downloads folder to the pipedrive filter destination

shutil.copytree(download, pipedrive_file_destination, 
                copy_function = shutil.move, dirs_exist_ok = True)

# create a variable holding the new path to the pipedrive file

pipedrive_file = glob.glob(os.path.join(pipedrive_file_destination, '*.csv'))
pipedrive_file = pipedrive_file[0]

# specify the link to the csv file of all approved projects    

url = 'https://solarapp.nrel.gov/admin/projects/export/file/all_approved_Projects_Export_.csv'

# download the csv file of all approved projects. Note: You can either do this
# manually by copying and pasting the url, or by running this section of code

browser = webbrowser.get()
browser.open(url, autoraise=False)

# give the file 60 seconds to download

time.sleep(60)

# move the file from your downloads folder to your update folder and rename it 
# using today's date

os.replace(download + '/all_approved_Projects_Export_.csv', 
           update_folder + '/' + today + '/' + today +
           '_all_approved_Projects_Export_.csv')

# read the approved projects csv file into Python

most_recent = pd.read_csv(update_folder + '/' + today + '/' + today +
                           '_all_approved_Projects_Export_.csv',
                           usecols=['id', 'project_created_at', 'ahj_id',
                                    'ahj'])

# convert the "project_created_at" column to be datetime format (no time)

most_recent['project_created_at'] = pd.to_datetime(most_recent['project_created_at']).dt.date

# sort the projects from newest to oldest

most_recent.sort_values('project_created_at', inplace=True, ascending=False)

# get the first project (most recent) from each AHJ

most_recent_date = most_recent.groupby(by=['ahj_id'], as_index=False).first()

# get the number of total approved projects from each AHJ

most_recent_count = most_recent.groupby(by=['ahj_id'], as_index=False).count()

most_recent_count = most_recent_count.rename(columns={'id': 'Approved Projects'})

most_recent_count = most_recent_count[['ahj_id', 'Approved Projects']]

# merge the two dataframes holding the date of most recent project and the 
# total number of approved projects

most_recent = pd.merge(most_recent_date, most_recent_count, on=['ahj_id'], 
                       how = 'left')

# specify the link to the AHJ Report from SolarAPP website.

report_link = 'https://solarapp.nrel.gov/admin/ahjs/export-all'

# specify the destination for the AHJ Report.

report_destination = update_folder + '/' + today + '/' + today + '_Report'

# download the AHJ Report

browser.open(report_link)

# give the file 60 seconds to download

time.sleep(60)

# move the file from your downloads folder to the report destination

shutil.copytree(download, report_destination, copy_function = shutil.move,
                dirs_exist_ok = True)

# create a variable holding the new path to the ahj_report

report_file = glob.glob(os.path.join(report_destination, '*.csv'))
report_file = report_file[0]

# read the AJH Report and pull all relevant information.

ahj_report = pd.read_csv(report_file, 
                         usecols=['id', 'project_count'],
                         dtype={'id': 'int64'})

# rename id column in ahj_report dataframe
                                                   
ahj_report = ahj_report.rename(columns={'id': 'ahj_id'})

# merge the most_ recent dataframe with the ahj_report dataframe

most_recent = pd.merge(most_recent, ahj_report, on=['ahj_id'], how = 'left')

# read the pipedrive_filter csv file into Python

pipedrive = pd.read_csv(pipedrive_file)

# rename columns

pipedrive = pipedrive.rename(columns={'Organization - SolarAPP AHJ ID': 'ahj_id',
                                      'Organization - 2022 Rank': '2022 Rank',
                                      'Deal - Adopted Features': 'Adopted Features',
                                      'Deal - Title': 'AHJ Name',
                                      'Deal - Pilot Type': 'Pilot Type',
                                      'Deal - Last stage change': 'Pilot Start Date',
                                      'Deal - Stage': 'Stage'})

# merge the most_recent table with the pipedrive_filter csv file

most_recent = pd.merge(pipedrive, most_recent, on = ['ahj_id'], how = 'left')

# export the file as a csv to your update folder with the date in the name

most_recent.to_csv(update_folder + '/' + today + '/' + today + 
                   '_Most_Recent_Projects.csv')

# keep only columns needed for report table

most_recent = most_recent[['AHJ Name', '2022 Rank', 'project_created_at',
                           'Adopted Features', 'Pilot Type', 
                           'Pilot Start Date', 'Stage']]

# rename project_created_at column to the date of last issued permit

most_recent = most_recent.rename(columns={'project_created_at':
                                          'Date of Last Issued Permit'})
    
# create a variable holding the date two weeks ago

tod = datetime.datetime.now()
two_weeks = datetime.timedelta(days=14)
two_weeks_ago = tod - two_weeks
two_weeks_ago = two_weeks_ago.date()

# create a variable holding the date two weeks ago

four_weeks = datetime.timedelta(days=28)
four_weeks_ago = tod - four_weeks
four_weeks_ago = four_weeks_ago.date()

# replace all empty values with np.nan

most_recent = most_recent.replace(r'^\s*$', np.nan, regex=True)
    
# identify high volume AHJs that have not had a project in the past two weeks

high_vol = most_recent[(most_recent['2022 Rank'] < 501)]
high_vol = high_vol[(high_vol['Date of Last Issued Permit'].isnull()) | 
                    (high_vol['Date of Last Issued Permit'] < two_weeks_ago)]

# identify low volume AHJs that have not had a project in the past four weeks

low_vol = most_recent[(most_recent['2022 Rank'] > 500)]
low_vol = low_vol[(low_vol['Date of Last Issued Permit'].isnull()) | 
                  (low_vol['Date of Last Issued Permit'] < four_weeks_ago)]

# identify AHJs with no rank that have not had a project in the past four
# weeks

no_rank = most_recent[most_recent['2022 Rank'].isnull()]
no_rank = no_rank[(no_rank['Date of Last Issued Permit'].isnull()) | 
                  (no_rank['Date of Last Issued Permit'] < four_weeks_ago)]

# combine the high_volume and low_volume dataframes

most_recent = pd.concat([high_vol, low_vol, no_rank])

# sort the dataframe based on the date of last issued permit

most_recent.sort_values('Date of Last Issued Permit', inplace=True, 
                        ascending=False)                          

# format the table

with pd.ExcelWriter(update_folder + '/' + today + '/' + today + 
                   '_Most_Recent_Projects.xlsx', engine = 'xlsxwriter') as writer:

    most_recent.to_excel(writer, sheet_name='Most_Recent',
                             startrow = 1, header=False)

    workbook = writer.book
    worksheet = writer.sheets['Most_Recent']

    header_format = workbook.add_format({'bold': True,
                                         'text_wrap': True,
                                         'align': 'center',
                                         'valign': 'center',
                                         'bottom': 1,
                                         'top': 1})
    
    center_columns = workbook.add_format({'text_wrap': True,
                                          'align': 'center'})
    
    
    worksheet.set_column(1, 1, 23)
    worksheet.set_column(2, 2, 5, center_columns)
    worksheet.set_column(3, 6, 13, center_columns)

    for col_num, value in enumerate(most_recent.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
