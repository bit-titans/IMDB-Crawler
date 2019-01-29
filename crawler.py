import urllib
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter


def internal (url , workSheet , count):

    urlPage=urllib.request.urlopen("https://www.imdb.com/title/tt0111161/?ref_=nv_sr_1")
    soup=BeautifulSoup(urlPage, "html.parser")
    imdb=soup.find("div", {"class": "subtext"}).text
    # print(imdb)
    details=imdb.split("|")
    censor=details[0]
    time=details[1]
    genres=details[2].split(",")
    genre1=genres[0]
    try:
        genre2=genres[1]
    except:
        genre2='-'
    try:
        genre3=genres[2]
    except:
        genre3='-'
    try:
        genre4=genres[3]
    except:
        genre4='-'
    release_date=details[3]
    summary_text=soup.find("div", {"class":"summary_text"})
    summary = soup.findAll("div",{"class":"credit_summary_item"})
    director=summary[0].text.split(":")[1]
    writers=summary[1].text.split(",")
    writer1=writers[0]
    try:
        writer2=genres[1]
    except:
        writer2='-'
    stars_cast=summary[2].text.split("|")[0]
    stars=stars_cast.split(",")
    star1=stars[0]
    try:
        star2=stars[1]
    except:
        star2='-'
    try:
        star3=stars[2]
    except:
        star3='-'
    workSheet.write(count, 5, str(censor))
    workSheet.write(count, 6, time)
    workSheet.write(count, 7, genre1)
    workSheet.write(count, 8, genre2)
    workSheet.write(count, 9, genre3)
    workSheet.write(count, 10, release_date)
    workSheet.write(count, 11, str(summary_text))
    workSheet.write(count, 12, director)
    workSheet.write(count, 13, writer1)
    workSheet.write(count, 14, writer2)
    workSheet.write(count, 15, star1)
    workSheet.write(count, 16, star2)
    workSheet.write(count, 17, star3)
    return

count=1
row=1
url = "https://www.imdb.com/chart/top?ref_=nv_mv_250"
page = urllib.request.urlopen(url)
soup = BeautifulSoup(page, "html.parser")
soup1 = soup.find("div", {"class", "lister"})
workbook = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0,"Name of the movie")
for movie in soup1.findAll("tr"):
    if(count!=1):
        string=(movie.find("td", {"class" : "titleColumn"}).text).replace("\n", "")
        string1 = string.split('.')
        string2 = string1[1].split('(')
        worksheet.write(row, 0, string2[0])
        internal("Abhi", worksheet, row)

        row += 1

    count += 1

workbook.close()

