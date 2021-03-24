import requests
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import urllib.request
import time
from bs4 import BeautifulSoup

# worldometerCorona

WorldCorona_Page = 'https://www.worldometers.info/coronavirus/'
response = requests.get(WorldCorona_Page)
# soup = BeautifulSoup(response.text, 'html.parser')
soup = BeautifulSoup(response.text, 'html.parser')
table_data = soup.find('table', attrs={'id': 'main_table_countries_today'})
# print(table_data)
html_Data = pd.read_html(str(table_data))
# print(html_Data[0])
df = pd.DataFrame(html_Data[0])
df2 = df.rename(columns={'Country,Other': 'Country', 'ActiveCases': 'Active Cases', 'TotalRecovered': 'Total Recovered',
                         'TotalDeaths': 'Total Deaths', 'NewCases': 'New Cases', 'TotalCases': 'Total Cases'})
df2['Active Cases'] = df2['Active Cases'].astype(str).str.replace(',', '')
df2['Active Cases'] = df2['Active Cases'].astype(int)
print(df2['Active Cases'][:50])
#df2 = df2['Active Cases']
df2.to_csv('C:/Users/kriti/Documents/Corona.csv')
filename = 'C:/Users/kriti/Documents/Corona.csv'
df = pd.read_csv(filename, encoding='latin-1')
df = df.loc[df['Country'].isin(['Canada', 'India'])]
# df = df.loc[df['Country'].isin(['China', 'Mexico'])]
#print(df)
# Total Cases
fig = plt.figure(1)
plt.bar(df['Country'], df['Total Cases'], color='g', width=0.5)
plt.xlabel('Country')
plt.ylabel('Total Cases')
plt.title('Total Cases - Canada Vs India')
sht = xw.Book('C:/Users/kriti/Documents/CanadaIndiaDashboard.xlsx').sheets[0]
SS_TotalCases = sht.pictures.add(fig, name='MyPlot', update=True, left=sht.range('B27').left, top=sht.range('B27').top)
SS_TotalCases.height = 250
SS_TotalCases.width = 300

# Active Cases
fig2 = plt.figure(2)
plt.bar(df['Country'], df['Active Cases'], color='b', width=0.5)
plt.xlabel('Country')
plt.ylabel('Active Cases')
plt.title('Active Cases - Canada Vs India')
SS_ActiveCases = sht.pictures.add(fig2, name='MyPlot2', update=True, left=sht.range('L27').left,
                                  top=sht.range('L27').top)
SS_ActiveCases.height = 250
SS_ActiveCases.width = 300

# Recovered Cases
fig3 = plt.figure(3)
plt.bar(df['Country'], df['Total Recovered'], width=0.5, color='g')
plt.xlabel('Country')
plt.ylabel('Total Recovered')
plt.title('Total Recovered - Canada Vs India')
SS_TotalRecovered = sht.pictures.add(fig3, name='MyPlot3', update=True,
                                     left=sht.range('B49').left, top=sht.range('B49').top)
SS_TotalRecovered.height = 250
SS_TotalRecovered.width = 300

# Total Deaths
fig4 = plt.figure(4)
plt.bar(df['Country'], df['Total Deaths'], width=0.5, color='b')
plt.xlabel('Country')
plt.ylabel('Total Deaths')
plt.title('Total Deaths - Canada Vs India')
SS_TotalRecovered = sht.pictures.add(fig4, name='MyPlot4', update=True,
                                     left=sht.range('L49').left, top=sht.range('L49').top)
SS_TotalRecovered.height = 250
SS_TotalRecovered.width = 300

# New Cases
fig5 = plt.figure(5)


plt.barh(df['Country'], df['New Cases'], align='center', alpha=0.5)
plt.xlabel('Country')
plt.ylabel('New Cases')
plt.title('New Cases - Canada Vs India')
#plt.ylim(0, 3000)

SS_TotalRecovered = sht.pictures.add(fig5, name='MyPlot5', update=True,
                                     left=sht.range('F5').left, top=sht.range('F5').top)
SS_TotalRecovered.height = 250
SS_TotalRecovered.width = 450

xw.Book('C:/Users/kriti/Documents/CanadaIndiaDashboard.xlsx').save()
