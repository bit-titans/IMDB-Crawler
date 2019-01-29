import urllib
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
count = 1
row = 1
"""
    Created by Abhishek Koushik B N
                       AND
               Akash R
    On 29/Jan/2019
"""


def internal(url, work_sheet, counts):
    print(str(counts) + "/250")
    urlPage=urllib.request.urlopen(url)
    soup=BeautifulSoup(urlPage, "html.parser")
    imdb=soup.find("div", {"class": "subtext"}).text
    details = imdb.split("|")
    length = len(details)
    if length == 4 :
            censor = details[0]
            time = details[1]
            genres = details[2].split(",")
            genre1 = genres[0]
            try:
                genre2 = genres[1]
            except:
                genre2 = '-'
            try:
                genre3 = genres[2]
            except:
                genre3 = '-'
            try:
                genre4 = genres[3]
            except:
                genre4 = '-'
            release_date=details[3]
    else:
        censor = "Not Rated"
        time = details[0]
        genres = details[1].split(",")
        genre1 = genres[0]
        try:
            genre2 = genres[1]
        except:
            genre2 = '-'
        try:
            genre3 = genres[2]
        except:
            genre3 = '-'
        try:
            genre4 = genres[3]
        except:
            genre4 = '-'
        release_date = details[2]
    summary_text = soup.find("div", {"class":"summary_text"}).text
    summary = soup.findAll("div",{"class":"credit_summary_item"})
    director = summary[0].text.split(":")[1].split("|")[0]
    writers = summary[1].text.split(":")
    writer1 = writers[1].split(",")[0]
    try:
        writer2 = writers[1].split(",")[1]
        writer2 = writer2.split("|")[0]
    except:
        writer2 = '-'
    stars_cast = summary[2].text.split("|")[0]
    stars = stars_cast.split(",")
    star1 = stars[0].split(":")[1]
    try:
        star2 = stars[1]
    except:
        star2 = '-'
    try:
        star3 = stars[2]
    except:
        star3 = '-'
    work_sheet.write(counts, 5, str(censor))
    work_sheet.write(counts, 6, time)
    work_sheet.write(counts, 7, genre1)
    work_sheet.write(counts, 8, genre2)
    work_sheet.write(counts, 9, genre3)
    work_sheet.write(counts, 10, release_date)
    work_sheet.write(counts, 11, str(summary_text))
    work_sheet.write(counts, 12, director)
    work_sheet.write(counts, 13, writer1)
    work_sheet.write(counts, 14, writer2)
    work_sheet.write(counts, 15, star1)
    work_sheet.write(counts, 16, star2)
    work_sheet.write(counts, 17, star3)
    return


count = 1
row = 1
url = "https://www.imdb.com/chart/top?ref_=nv_mv_250"
page = urllib.request.urlopen(url)
soup = BeautifulSoup(page, "html.parser")
soup1 = soup.find("div", {"class", "lister"})
workbook = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Name of the movie")
worksheet.write(0, 1, "Link")
worksheet.write(0, 2, "Year Released")
worksheet.write(0, 3, "IMDB Rating")
worksheet.write(0, 4, "Number of Ratings")
worksheet.write(0, 5, "Censor Board Rating")
worksheet.write(0, 6, "Length of the movie")
worksheet.write(0, 7, "Genre 1")
worksheet.write(0, 8, "Genre 2")
worksheet.write(0, 9, "Genre 3")
worksheet.write(0, 10, "Release Date")
worksheet.write(0, 11, "Story Summary")
worksheet.write(0, 12, "Director Name")
worksheet.write(0, 13, "Writer 1")
worksheet.write(0, 14, "Writer 2")
worksheet.write(0, 15, "Star 1")
worksheet.write(0, 16, "Star 2")
worksheet.write(0, 17, "Star 3")
for movie in soup1.findAll("tr"):
    if count != 1:
        string = movie.find("td", {"class" : "titleColumn"}).text.replace("\n", "")
        string1 = string.split('.')
        string2 = string1[1].split('(')
        worksheet.write(row, 0, string2[0])
        string3 = string.split('(')
        worksheet.write(row, 2, string3[1][0:4])
        rating = movie.find("td", {"class": "ratingColumn imdbRating"}).text
        worksheet.write(row, 3, rating)
        link = str(movie.find("a"))
        url_string1 = link.split("\"")
        url_string2 = url_string1[1].split("\"")
        worksheet.write(row, 1, "https://www.imdb.com"+url_string2[0])
        rating_temp = str(movie.find("td",{"class":"ratingColumn imdbRating"}))
        ratings = rating_temp.split("on ")
        ratings = ratings[1].split(" ")[0]
        worksheet.write(row, 4, ratings)
        internal("https://www.imdb.com" + url_string2[0], worksheet, row)
        row += 1
    count += 1
workbook.close()

