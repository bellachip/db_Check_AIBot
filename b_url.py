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


def get_name():
    for i in range(3):
        if i >= 2:
            ex_last = str(sheet['A' + str(i)].value).strip()
            first_name_arr.append(ex_last)
            ex_first = str(sheet['B' + str(i)].value).strip()
            last_name_arr.append(ex_first)
            # return last_name_arr, first_name_arr
    return ex_last, ex_first


# # returns last name array and first name array
l, f = get_name()
print(l)


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
    print(table_rows)

    for table_row in table_rows:
        names = table_row.a.text
        new_name = names.replace('Case Summary: ', '').replace(',', '')
        new_arr_f.append(new_name)
        # new_arr_f.append(new_name.split()[0])
        # new_arr_l.append(new_name.split()[1])

    return new_arr_f


# c_arr = []
# for i in range(len(l)):
#     str_cat = str(f[i]) + " " + str(l[i])
#     c_arr.append(str_cat)

# returns the array of the excel strings

str_cat = l + " " + f
first = b_url_check()
print(str_cat)

print(first)
print()


# check if the scriped string is in the excel data
def is_name(a):
    if str_cat in first:
        rv = "Yes"
    else:
        rv = "No Results"

    return rv


# print(first)


# checks if the person exists in the gov data base.
def clr_check(c_check=is_name(str_cat)):
    # for word in c_check.split():
    d2 = sheet['D2']  # should switch to the row number
    print(c_check)
    if c_check == 'No Results':
        no_results_cell = 'No Results'
        d2.value = no_results_cell
    elif c_check == 'Yes':
        d2.value = 'Yes'

    return d2.value


clr_check()

ex.save('db_check.xlsx')
