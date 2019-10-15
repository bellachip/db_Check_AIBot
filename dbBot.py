#!/bin/python
from bs4 import BeautifulSoup as soup
import os
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup

# directory
os.chdir('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot')
driver = webdriver.Chrome('C:\\Users\\yangb\\Desktop\\chromedriver.exe')
first_name_arr = []
last_name_arr = []
ex_name = 'db_check.xlsx'
ex = openpyxl.load_workbook(ex_name)
sheet = ex["Sheet1"]


# get name of one row
def get_name():
    for i in range(3):
        if i >= 2:
            ex_last = sheet['A' + str(i)].value
            first_name_arr.append(ex_last)
            ex_first = sheet['B' + str(i)].value
            last_name_arr.append(ex_first)
            # return last_name_arr, first_name_arr
            return ex_last, ex_first


# returns last name  and first name of first row
l, f = get_name()

# will iterate through these urls
# store_urls = ['https://exclusions.oig.hhs.gov/', 'https://ori.hhs.gov/case_summary']

r_url = 'https://exclusions.oig.hhs.gov/'


# does the job of clearnace checking for the first url
def clearance_check(url=r_url, last_name_search=l, first_name_search=f):
    # opening url and do the search
    driver.get(url)
    print(driver.current_url)
    # finds the sesarch bars for last name and first name
    last_name = driver.find_element_by_id('ctl00_cpExclusions_txtSPLastName')
    first_name = driver.find_element_by_id('ctl00_cpExclusions_txtSPFirstName')
    last_name.send_keys(last_name_search)
    first_name.send_keys(first_name_search)
    last_name.send_keys(Keys.ENTER)
    # first_name.send_keys(Keys.ENTER)

    # opening up connection, grabbing the search results page
    page_html = driver.page_source

    # lxml parsing through the current results page
    page_soup = soup(page_html, "lxml")

    # Name = page_soup.find("div", {"id": "ctl00_cpExclusions_pnlEmpty"}).ul.li.text

    # gets the no results text
    # if "No Results" is in the page then return no results variable
    noe = page_soup.find("div", {"id": "ctl00_cpExclusions_pnlEmpty"})

    if noe is not None:
        no_results = "No Results"
    else:  # else then that means there is results for the person so scrape the info of the person
        # gets the rows
        rows = page_soup.find("table", {"class": "leie_search_results"}).find("tbody").findAll("tr")

        i = 1

        # iterate through each row
        for row in rows:
            cols = row.findAll("td")
            for col in cols:
                cell = col.text
        if cell.find(last_name_search.upper()):
            no_results = "Yes"
    return no_results

    # return no_results
    # elif url == 'https://ori.hhs.gov/case_summary':
    #     # opening up connection, grabing the page
    #     u_client = uReq(url)
    #     page_html = u_client.read()
    #     u_client.close()
    #
    #     # html parsing
    #     page_soup = soup(page_html, "html.parser")
    #     years = page_soup.find_all("h3")
    #
    #     table_rows = page_soup.findAll("div", {"class": "views-field views-field-title"})
    #     for table_row in table_rows:
    #         names = table_row.a.text
    #
    #         if names.find(first_name_search):
    #             no_results = "Yes"


# checks if the person exists in the gov data base.
def clr_check(c_check=clearance_check(r_url, l, f)):
    # for word in c_check.split():
    c2 = sheet['C2']  # should switch to the row number
    print(c_check)
    if c_check == 'No Results':
        no_results_cell = 'No Results'
        c2.value = no_results_cell
    elif c_check == 'Yes':
        c2.value = 'Yes'

    return c2.value


clr_check()

ex.save('db_check.xlsx')
driver.close()

# # # grabs each row
# # rows = content.findAll("tr")
# # # conts = page_soup.findAll("tr", {"class" :"bg-lighter-yellow"})
# #
# # # header for the csv file
# # headers = "title, author, abstract_number, poster\n"
# #
# # # each link
# # link = my_url + '\n'
# #
# # f.write(headers)
# # f.write(link)
# #
# # # iterates through each row
# # for row in rows:
# #     # removes the header row in the table
# #     if row.find("a") is None:
# #         continue
# #     if row.find("p") is None:
# #         continue
# #     if row.find("td", {"valign": "top"}) is None:
# #         continue
# #
# #         # grabs the title
# #     title = row.a.text
# #
# #     print(title)
# #     # grabs the author
# #     author = row.p.text
# #
# #     # grabs all the abstract container
# #     abstract_con = row.findAll("td", {"valign": "top"})
# #     # only grabs the text in the abstrac_con
# #     abstract_number = abstract_con[2].p.text
# #
# #     # for identifying if its a poster or not
# #     if abstract_number[0] == 'T':
# #         poster = "yes"
# #     else:
# #         poster = "no"
# #
# ex = openpyxl.load_workbook('db_check.xlsx')
# sheet = ex.get_sheet_by_name('Sheet1')

# for rowOfCellObjects in sheet['A2':'B2']:
#     for cellObj in rowOfCellObjects:
# cellCord = cellObj.coordinate
# cellVal = cellObj.value
# first_name = sheet['A2']
# last_name = sheet['B2']
# driver.get(my_url)
# driver.findElement(By.xpath("(//input[@id='ctl00_cpExclusions_txtSPLastName'])")).sendKeys(last_name)
# driver.findElement(By.xpath("(//input[@id='ctl00_cpExclusions_txtSPFirstName'])")).sendKeys(last_name)


# filename = "test.csv"
# f = open(filename, "w")  # csvfile open
#
# # iterates through each link
# #for i in range(len(url_nums)):
# my_url = 'https://exclusions.oig.hhs.gov/'
#
# # opening up connection, grabing the page
# uClient = uReq(my_url)
# page_html = uClient.read()
# uClient.close()
#
# # html parsing
# page_soup = soup(page_html, "html.parser")
#
# print(page_soup)
#
# # # grabs table
# # table = page_soup.find("table", {"class": "t4"})
# # # grabs each row
# # rows = table.findAll("tr")
# # # conts = page_soup.findAll("tr", {"class" :"bg-lighter-yellow"})
# #
# # # header for the csv file
# # headers = "title, author, abstract_number, poster\n"
# #
# # # each link
# # link = my_url + '\n'
# #
# # f.write(headers)
# # f.write(link)
# #
# # # iterates through each row
# # for row in rows:
# #     # removes the header row in the table
# #     if row.find("a") is None:
# #         continue
# #     if row.find("p") is None:
# #         continue
# #     if row.find("td", {"valign": "top"}) is None:
# #         continue
# #
# #         # grabs the title
# #     title = row.a.text
# #
# #     print(title)
# #     # grabs the author
# #     author = row.p.text
# #
# #     # grabs all the abstract container
# #     abstract_con = row.findAll("td", {"valign": "top"})
# #     # only grabs the text in the abstrac_con
# #     abstract_number = abstract_con[2].p.text
# #
# #     # for identifying if its a poster or not
# #     if abstract_number[0] == 'T':
# #         poster = "yes"
# #     else:
# #         poster = "no"
# #
# #     # write to the csv file
# #     # strng handlers and concatenation
# #     f.write(title.replace(",", "|") + "," + author.replace(",", "|") + "," + abstract_number + "," + poster + "\n")
# f.close()  # file close
#
