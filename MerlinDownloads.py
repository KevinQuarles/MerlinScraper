#! python3
# MerlinDownloads.py - downloads Merlin files from the web
# puts the files in their folder and unzipes the files


import shutil, zipfile, os, time, datetime, re, pandas as pd, openpyxl, calendar, xlrd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait

# gets the current username to pull from the downloads folder
path = os.path.abspath('C:\\Users')
pathFolders = os.listdir(path)
for i in pathFolders:
    if i == 'chehem':
        userName = i
    elif i == 'Kevqua':
        userName = i
    elif i == 'keleng':
        userName = i
    elif i == 'hughes.speer':
        userName = i


print('Downloads Folder file path created.')


# creates a list of clean invoice numbers based on user input
invNums = []
invCount = input('How many invoices to search?:\n')
while not invCount.isdigit():
    invCount = input('How many invoices to search?:\n')


for i in range(int(invCount)):
    invNum = input('Please enter a valid invoice number:\n')
    invNum = invNum.upper()
    while invNum[:6] != 'CONUS-' or not invNum[11:].isdigit or len(invNum) != 16:
        invNum = input('Please enter a valid invoice number (ex:CONUS-2019-49586):\n')
    else:
        invNums.append(invNum)
    print(invNums)

#############for each invoice number, start a file download and migration based on date########################

for invNum in invNums:

# open chrome browser and go to site
    browser = webdriver.Chrome()
    browser.get('http://wizard.merlinnetwork.org/members.php')
    browser.maximize_window()
    ##signInElem = browser.find_element_by_class_name('sign-in-btn').click()
    ##
    ### switch tabs
    ##for handle in browser.window_handles:
    ##    browser.switch_to.window(handle)
    # logging in
    time.sleep(2)
    emailElem = browser.find_element_by_id('signin_username')
    time.sleep(2)
    emailElem.send_keys('username')
    time.sleep(2)
    passwordElem = browser.find_element_by_id('signin_password')
    passwordElem.send_keys('password')
    time.sleep(2)
    signInButtonElem = browser.find_element_by_xpath('/html/body/div[4]/form/table/tfoot/tr/td/input').click()
    time.sleep(5)
    reportElem = browser.find_element_by_link_text('Reporting').click()
    time.sleep(2)
    royaltyElem = browser.find_element_by_link_text('Royalty').click()
    time.sleep(2)
    allElem = browser.find_element_by_link_text('All').click()
    time.sleep(2)
    advFilterElem = browser.find_element_by_link_text('Advanced Filter').click()
    time.sleep(2)
    invElem = browser.find_element_by_id('merlin_statement_filters_invoice_number').click()
    time.sleep(2)
    invElem = browser.find_element_by_id('merlin_statement_filters_invoice_number')
    time.sleep(2)
    invElem.send_keys(invNum)
    time.sleep(2)
    filterElem = browser.find_element_by_xpath('//*[@id="member_sales_reports_sfAdminFilterDialog"]/div[3]/button[1]').click()
    time.sleep(2)
    invElem = browser.find_element_by_link_text(invNum).click()
    time.sleep(2)
    filterElem = browser.find_element_by_link_text('Download').click()
    time.sleep(2)
    browser.back()
    iconList = browser.find_elements_by_class_name('icon-download')
    time.sleep(2)

# loop through icons and files for the invoice
    fileCount = 0
    for i,icon in enumerate(iconList):
        browser.switch_to.window(browser.window_handles[0])
        newIconList = browser.find_elements_by_class_name('icon-download')
        newIcon = newIconList[i]
        newIcon.click()
        time.sleep(2)
        txtList = browser.find_elements_by_class_name('sf_row.txt')
        tsvList = browser.find_elements_by_class_name('sf_row.tsv')
        gzList = browser.find_elements_by_class_name('sf_row.gz')
        csvList = browser.find_elements_by_class_name('sf_row.csv')
        zipList = browser.find_elements_by_class_name('sf_row.zip')
        fileList = txtList + tsvList + csvList + zipList + gzList
        for file in fileList:
            file.click()
            time.sleep(2)
            fileCount += 1
        browser.back()

    print(str(fileCount) + ' Files Downloaded')

# get year and month from invoice
    print('Please enter the four-digit invoice year of ' + invNum + ': ')
    year = input()
          
    while len(year) != 4 or not year.isdigit:
        print('Please enter correct year of ' + invNum + '(ex:2019): ')
        year = input()
          
    print('Please enter the two-digit invoice month of ' + invNum + ': ')
    month = input()
          
    while len(month) !=2 or not month.isdigit:
        print('Please enter correct month of ' + invNum + '(ex:08): ')
        month = input()


    yearMonth = (year + '-' + month)
    fileDate = (yearMonth + '-' + '15')

    print('Dates Gathered ' + fileDate)

# move files, create destination folder and unzip the files
    sourcePath = os.path.abspath('C:\\Users\\' + userName + '\Downloads')
    sourceFiles = os.listdir(sourcePath)
    os.makedirs(r'\\\Wind Up Merlin\\'+ str(year) +'\\' + str(yearMonth) + '\\' +invNum) 
    destinationPath = os.path.abspath(r'\\\Wind Up Merlin\\'+ str(year) +'\\' + str(yearMonth) + '\\' +invNum)

    fileCount = 0

    for file in sourceFiles:
        if file.endswith('.xls'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.tsv'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.txt'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.csv'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.gz'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.pdf'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
        elif file.endswith('.zip'):
            shutil.move(os.path.join(sourcePath, file), os.path.join(destinationPath, file))
            fileCount += 1
            
# give this time to think then close the current browser

    time.sleep(15)        
    browser.close()
            
    print('-------------------------------------------------------------------------------------------')    
    print(fileDate + ' ' + invNum + ':\n' + str(fileCount) + ' files have been downloaded!')
    print('-------------------------------------------------------------------------------------------')    

#TODO    
##Create a folder to drop all needed pdf invoices 
##list contents of that folder
##count how many are in the folder
##strip out the invoice numbers
##make input list and loop range count variable based on length
##when downloading make inv folder in downloads and move to that folder immediately upon download
##also download statements csv or read pdf
##put to pandas df
##from pandas df, pull inv date and tuple with inv list item (invNum)
##use dates from tupled list to make new folder paths
##put all downloading and moving into a function that takes a data and num from a tuples list as input
##once all files in folders build formatting cases with pandas for each file and dsp
##may need specific dsp functions
##get filedate column from pandas df
##get use_fildate column from folder name
##get filename column from filename
##get inv number column from folder name
##clean upcs
##delete unneeded rows and totals
##combine multiple similiar df's
##put in imports folder inside invoice folder
##all complete


