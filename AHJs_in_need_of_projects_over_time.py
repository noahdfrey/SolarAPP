# -*- coding: utf-8 -*-
"""
Created on Tue Jan 16 19:18:49 2024

@author: nfrey
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime

over_time_csv = pd.read_csv('C:/Users/nfrey/OneDrive - NREL/Most Recent Project/Most Recent Project Over Time.csv')

over_time_csv['Percentage1'] = 1 - over_time_csv['Total Percentage']

over_time_csv['Total Percentage'] = over_time_csv['Total Percentage']*100

over_time_csv['High Volume Percentage'] = over_time_csv['High Volume Percentage']*100

over_time_csv['Low Volume Percentage'] = over_time_csv['Low Volume Percentage']*100

over_time_csv['Percentage1'] = over_time_csv['Percentage1']*100

over_time_csv['Date'] = pd.to_datetime(over_time_csv['Date'])

print(over_time_csv)

plt.stackplot(over_time_csv['Date'], over_time_csv['High Volume Percentage'], 
              over_time_csv['Low Volume Percentage'], over_time_csv['Percentage1'], 
              labels=('High Volume AHJs in Need of Projects', 
                      'Low Volume AHJs in Need of Projects', 'AHJs with Projects'),
              colors=('#d7191c', '#fdae61', '#1a9641'))
plt.legend(loc='upper left')
handles, labels = plt.gca().get_legend_handles_labels()
order = [2, 1, 0]
plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order])
plt.margins(0,0)
plt.xticks(rotation=40, ha='right')
plt.tick_params(right=True, labelright=True)
plt.xlabel("Date")
plt.ylabel("Percentage of AHJs (%)")
plt.tight_layout()
plt.savefig('2024_6_25_Most_Recent_Figure.png', dpi=300)
plt.show()


