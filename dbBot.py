#!/bin/python
from bs4 import BeautifulSoup as soup
import os
from docx import Document
from docx.shared import Inches
import openpyxl
from openpyxl.cell import cell
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup
from docx.oxml.shared import OxmlElement, qn
import datetime
from datetime import timedelta

# directory
os.chdir('C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot')

# define the name of the directory to be created
# directory structure for screenshots and files
debarment_file_path = 'C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Debarment_files'
screenshots_path = 'C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBot\\Screenshots'

try:
    os.mkdir(screenshots_path)
except OSError:
    print("Creation of the directory %s failed" % screenshots_path)
else:
    print("Successfully created the directory %s " % screenshots_path)

try:
    os.mkdir(debarment_file_path)
except OSError:
    print("Creation of the directory %s failed" % debarment_file_path)
else:
    print("Successfully created the directory %s " % debarment_file_path)

driver = webdriver.Chrome('C:\\Users\\yangb\\Desktop\\chromedriver.exe')

first_name_arr = []
last_name_arr = []
new_arr_f = []
ex_name = 'db_check.xlsx'
ex = openpyxl.load_workbook(ex_name)
sheet = ex["Sheet1"]
b_sheet = ex['sheet2']


def create_doc(first_name_docx, last_name_docx, a_res_value, b_res_value):
    document = Document()
    document.add_heading('Debarment Check', 0)
    p = document.add_paragraph(
        'Prior to being invited to participate in development/authoring of a publication sponsored '
        'by Genzyme/Sanofi, a debarment check must be completed for each US author.')

    records = (
        ('Author Name', first_name_docx + ' ' + last_name_docx),
        ('Name of Institution', ' '),
        ('City, State', '')
    )

    debarment_list = (
        ('Office of Inspectors General LIst of Excluded Individuals.', a_res_value),
        ('System for Award Management', b_res_value),
        ('Office of Research Integrity', '')
    )

    table = document.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Information'
    hdr_cells[1].text = 'Id'

    for qty, id in records:
        row_cells = table.add_row().cells
        row_cells[0].text = qty
        row_cells[1].text = id

    document.add_paragraph('\n')

    b_table = document.add_table(rows=1, cols=2)
    b_table.style = 'Table Grid'
    b_hdr_cells = b_table.rows[0].cells
    b_hdr_cells[0].text = 'Debarment List'
    b_hdr_cells[1].text = 'Findings'

    for d_list, findings in debarment_list:
        b_row_cells = b_table.add_row().cells
        b_row_cells[0].text = d_list
        b_row_cells[1].text = findings

    document.add_paragraph(
        '\nIf the potential author is listed on any of the above, they may not be invited to author a '
        'publication sponsored by Genzyme or Sanofi; advise publication lead of findings of this '
        'search.\n')
    document.add_paragraph(
        'Once the debarment check has been completed, upload this document to the appropriate record '
        'in Datavision.\n')
    completion = document.add_paragraph('Debarment check completed by:\n')
    completion.add_run('Debarment').bold = True
    document.add_paragraph('Date check completed:\n')

    document.add_picture(first_name_docx + '.png', width=Inches(6))

    # Set a cell background (shading) color to RGB D9D9D9.
    a_cell_1 = table.cell(0, 0)
    a_co = a_cell_1._tc.get_or_add_tcPr()
    a_cell_2 = table.cell(0, 1)
    a_ct = a_cell_2._tc.get_or_add_tcPr()

    a_cell_color_1 = OxmlElement('w:shd')
    a_cell_color_1.set(qn('w:fill'), '#94C167')

    a_cell_color_2 = OxmlElement('w:shd')
    a_cell_color_2.set(qn('w:fill'), '#94C167')

    a_co.append(a_cell_color_1)
    a_ct.append(a_cell_color_2)

    b_cell_1 = b_table.cell(0, 0)
    b_co = b_cell_1._tc.get_or_add_tcPr()
    b_cell_2 = b_table.cell(0, 1)
    b_ct = b_cell_2._tc.get_or_add_tcPr()

    b_cell_color_1 = OxmlElement('w:shd')
    b_cell_color_1.set(qn('w:fill'), '#94C167')

    b_cell_color_2 = OxmlElement('w:shd')
    b_cell_color_2.set(qn('w:fill'), '#94C167')

    b_co.append(b_cell_color_1)
    b_ct.append(b_cell_color_2)

    document.add_page_break()

    document.save(first_name_docx + '.docx')


# get name of one row
# same function
def get_name(r):
    ex_last = str(r[0].value)
    first_name_arr.append(ex_last)
    ex_first = str(r[1].value)
    last_name_arr.append(ex_first)
    # return last_name_arr, first_name_arr
    return ex_last, ex_first


# will iterate through these urls
# store_urls = ['https://exclusions.oig.hhs.gov/', 'https://ori.hhs.gov/case_summary']

r_url = 'https://exclusions.oig.hhs.gov/'


# does the job of clearnace checking for the first url
def clearance_check(url, last_name_search, first_name_search):
    driver.get(url)
    # opening url and do the search
    # driver = webdriver.Chrome('C:\\Users\\yangb\\Desktop\\chromedriver.exe')
    # driver.get(url)
    # print(driver.current_url)
    # finds the sesarch bars for last name and first name
    last_name = driver.find_element_by_id('ctl00_cpExclusions_txtSPLastName')
    first_name = driver.find_element_by_id('ctl00_cpExclusions_txtSPFirstName')
    last_name.send_keys(last_name_search)
    first_name.send_keys(first_name_search)
    last_name.send_keys(Keys.ENTER)
    # first_name.send_keys(Keys.ENTER)
    driver.save_screenshot("screenshot.png")
    # opening up connection, grabbing the search results page
    page_html = driver.page_source
    # driver.close()
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

        # iterate through each row
        for row in rows:
            cols = row.findAll("td")
            for col in cols:
                cell = col.text
        if cell.find(last_name_search.upper()):
            no_results = "Yes"

    return no_results


# opening up connection, grabing the page
def b_url_check():
    url_b = 'https://ori.hhs.gov/case_summary'
    u_client = uReq(url_b)
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
def clr_check(c_check, r, url_col):
    # for word in c_check.split():
    c2 = r[url_col]  # should switch to the row number
    # print(c_check)
    if c_check == 'No Results':
        no_results_cell = 'No, individual is not listed'
        c2.value = no_results_cell
    elif c_check == 'Yes':
        c2.value = 'Yes, individual appears on this list'

    return c2.value


counter = 0


def insert_date(r, date_col):
    d = datetime.datetime.today()

    r[date_col].value = d
    r[date_col + 1].value = d + timedelta(days=366)
    return r[date_col].value, r[date_col + 1].value


# counter = 1


# def create_rejected_sheet(r, url_col, rejected_sheet):
#     if str(r[url_col].value) == 'No, individual is not listed':
#         for cell.value in rejected_sheet['A']:
#             if cell.value is None:
#                 cell.value = r[url_col].value
#
#             else:
#                 i + 1
#                 break;
#     else:
#         return None


k = 1


# def copyRange(startCol, startRow, endCol, endRow, sheet):
# rangeSelected=[]
# for i in range(startRow, endRow +1, 1):
#     rowSelected =[]
#     for j in range(startCol, endCol+1, 1):
#         rowSelected.append(sheet.cell(row = i, column =j).value)
#         rangeSelected.append(rowSelected)
# return rangeSelected

def check_row_isempty(r):
    # for num, r in enumerate(b_sheet.iter_rows()):
    # for row_cells in b_sheet.iter_rows(min_col=1, max_col=1):
    #     for cell in row_cells:
    #         if cell.value is None:
    #             return row_cells
    #         else:
    #             break

    b_sheet.cell(row=b_sheet.max_row + 1, column=1).value = str(r[0].value)

    # print(row_cells.index)

            # for cell in row_cells:
        #     print('%s: cell.value=%s' % (cell, cell.value))








def cr(r, r2):
      # column numbersr
    # print(r)

        # sheet.c(row=i, column=j).value = sheet_b.c(row=i, colmn=j).value
    ex.save(ex_name)


for i, j in enumerate(sheet.iter_rows()):
    if i == 0:
        continue

    print(i)

    # gets the last name first name string
    l, f = get_name(j)
    chk = clearance_check(r_url, l, f)
    driver.save_screenshot(f + ".png")
    clr_check(chk, j, 6)

    check_row_isempty(j)
    # if check_row_isempty() != -1:
    #
    #     cr(j, check_row_isempty())
    # create_rejected_sheet(j, 7, ex['Not_Debared'])

    str_cat = l + ' ' + f
    first = b_url_check()
    # print(is_name(str_cat, first))
    clr_check(is_name(str_cat, first), j, 7)

    insert_date(j, 2)
    c_cell = str(sheet['G' + str(i)].value)
    d_cell = str(sheet['H' + str(i)].value)
    create_doc(f, l, c_cell, d_cell)

    ex.save('db_check.xlsx')

driver.close()
