import urllib
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
count=1
row=1
url = "https://www.imdb.com/chart/top?ref_=nv_mv_250"
page = urllib.request.urlopen(url)
soup = BeautifulSoup(page, "html.parser")
soup1 = soup.find("div", {"class","lister"})
workbook = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0,"Name of the movie")
for movie in soup1.findAll("tr"):
    if(count!=1):
        string=(movie.find("td", {"class" : "titleColumn"}).text).replace("\n", "")
        string1 = string.split('.')
        string2 = string1[1].split('(')
        worksheet.write(row, 0, string2[0])
        row += 1
    count += 1

workbook.close()
