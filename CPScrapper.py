##Written by Ethan Fowler
##Modified by D3ltaT1me
##Github: https://github.com/D3ltaT1me/CyberPatriot
##Email: fortnut695@gmail.com

from bs4 import BeautifulSoup as BS
import requests
import re
import xlsxwriter as xl
import time
import sys
##used regular expressions to remove the tags from each line
def removeTag(raw_text):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', raw_text)
    return cleantext

##Pull in the scoreboard and parse through it
try:
    page = requests.get('http://scoreboard.uscyberpatriot.org')
except:
    print("[!]Error: Webpage is unavailable...")
    sys.exit()

html = BS(page.content, 'html.parser')


##Set up for the Excel file
book = xl.Workbook()
sheetName = str(input("What round of competition is it?(ex. round1): "))
sheet = book.add_worksheet(sheetName)

cols = ["Placement", "Team Number", "Location", "Division", "Teir", "Scored Images", "Play Time", "Current Score"]
for index, col in enumerate(cols):
    sheet.write(0, index, col)

print("~"*15  + "Starting program" + "~"*15)
##Starts at 8 and ends at 15 in order to skip the labels at the top of the webpage
start = 8
end = 15
placement = 1
R = 1
start_time = time.time()

while True:
    ##Take out the table with the scores
    test = html.find_all('td')[start:end]

    ##make sure the line has a value
    if not len(test) == 0:
        ##insert a placement
        test.insert(0,placement)

        ##Created a new list for the newly formatted elements in the table
        L = []
        for x in test:
            x = str(x)
            x = removeTag(x)
            #print(removeTag(x))
            if x.isdigit():
                x = float(x)
            else:
                pass
            L.append(x)

        ##Adds the elements of the List to each column in the spreadsheet    
        row = sheet.row(R)
        for index, col in enumerate(cols):
            val = L[index]
            row.write(index, val)

        start += 8
        end += 8
        placement += 1
        R += 1

    else:
        break

elapsed_time = time.time() - start_time
elapsed_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
print("Time Elapsed: ", elapsed_time)
fileName = str(input("Please enter a fielname(ex. round1Scores.xls): "))
if not ".xls" in fileName:
    book.save(fileName +".xls")
else:
    book.save(fileName)