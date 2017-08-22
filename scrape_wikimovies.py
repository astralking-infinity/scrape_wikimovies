#!/usr/bin/python

import requests
import os
import xlsxwriter
from bs4 import BeautifulSoup

url = 'http://en.wikipedia.org/wiki/2016_in_film'
filename = 'movies_2016.xlsx'
directory = 'wikipedia'

count = 0

# months = ['January', 'February', 'March', 'April', 'May', 'June',
#           'July', 'August', 'September', 'October', 'November', 'December']

# Create directory/folder for the data extracted and move to that directory/folder
if not os.path.exists(directory):
    os.makedirs(directory)
os.chdir(os.path.join(os.getcwd(), directory))

# Create spreadsheet file
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet('Movies')

# Fill in the spreadsheet headers
worksheet.write(0, 0, 'Title')
worksheet.write(0, 1, 'Studio')
worksheet.write(0, 2, 'Cast and crew')
worksheet.write(0, 3, 'Genre')

# Download webpage
print('Dowloading webpage...')
response = requests.get(url)
response.raise_for_status
print('Done.\n')

# Parse html or webpage
soup = BeautifulSoup(response.content, 'html.parser')

tables = soup.find_all('table', {'class': 'wikitable sortable'})

print('Extracting movie data...')
for table in tables:
    movies = table.find_all('tr')
    for row in range(1, len(movies)):
        count += 1
        details = movies[row].find_all('td')

        if len(details) == 6:
            starter = 1

        if len(details) == 5:
            starter = 0

        print('Movie:', details[0 + starter].text,)
        for col in range(4):
            worksheet.write(count, col, details[col + starter].text)
        print('...Done.')

workbook.close()
print('Total movies extracted...', count)
