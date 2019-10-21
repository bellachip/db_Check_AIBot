from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup
import os
import openpyxl

# directory
os.chdir('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot')

first_name_arr = []
last_name_arr = []
new_arr_l = []
new_arr_f = []
ex_name = 'db_check.xlsx'
ex = openpyxl.load_workbook(ex_name)
sheet = ex["Sheet1"]
j = 0


# same function
def get_name(r):
    ex_last = str(r[0].value)
    first_name_arr.append(ex_last)
    ex_first = str(r[1].value)
    last_name_arr.append(ex_first)
    # return last_name_arr, first_name_arr
    return ex_last, ex_first


# opening up connection, grabing the page
def b_url_check():
    url = 'https://ori.hhs.gov/case_summary'
    u_client = uReq(url)
    page_html = u_client.read()
    u_client.close()

    # html parsing
    page_soup = soup(page_html, "html.parser")
    years = page_soup.find_all("h3")

    table_rows = page_soup.findAll("div", {"class": "views-field views-field-title"})
    # print(table_rows)

    for table_row in table_rows:
        names = table_row.a.text
        new_name = names.replace('Case Summary: ', '').replace(',', '')
        new_arr_f.append(new_name)
        # new_arr_f.append(new_name.split()[0])
        # new_arr_l.append(new_name.split()[1])
    return new_arr_f


# check if the scriped string is in the excel data
def is_name(a, fir):
    if a in fir:
        rv = "Yes"
    else:
        rv = "No Results"
    return rv


# checks if the person exists in the gov data base.
def clr_check(c_check, r):
    # for word in c_check.split():
    d2 = r[3]  # should switch to the row number
    print(c_check)
    if c_check == 'No Results':
        no_results_cell = 'No Results'
        d2.value = no_results_cell
    elif c_check == 'Yes':
        d2.value = 'Yes'

    return d2.value


def main():
    for i, row in enumerate(sheet.iter_rows()):
        if i == 0:
            continue
        l, f = get_name(row)
        str_cat = l + ' ' + f
        first = b_url_check()
        clr_check(is_name(str_cat, first), row)
        ex.save('db_check.xlsx')


main()
