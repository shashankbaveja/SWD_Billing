## The following program automates the procedure of billing students for t-shirts by different assocs
## the program is tested on win 8.1 with python 2.7

## download splinter to navigate through web, xlrd to read your excel, xlwt to write in excel files
## works perfectly fine on Google Chrome given you have downloaded chrome driver

## excel file must contain 'only' IDs of students to be billed in the first column without any breaks (empty columns in between)
## it will write IDs of all the failed billing in an excel file names 'failed.xls' in the directory containing your program

## file path you enter must be the absolute path ie with double backslashes ('\\')
## example file path - C:\\Users\\shashank baveja\\Desktop\\swd_billing\\Tshirt.xlsx

## please report any bugs you find

from splinter import Browser
import xlrd
import xlwt
import sys

def goto_billCode(browser):
     bill_code_btn = browser.find_by_id('Wizard1_SideBarContainer_SideBarList_ctl00_SideBarButton')
     bill_code_btn.click()

def fill_billID(browser,bill_id):
     browser.fill('Wizard1$codeTxt', bill_id)
     next_step = browser.find_by_id('Wizard1_StartNavigationTemplateContainerID_StartNextButton')
     next_step.click()

def goto_finish(browser):
     finish = browser.find_by_id('Wizard1_StepNavigationTemplateContainerID_StepNextButton')
     finish.click()
     finish2 = browser.find_by_id('Wizard1_FinishNavigationTemplateContainerID_FinishButton')
     finish2.click()


     
user_name = raw_input("enter users name ")
user_pass = raw_input("enter users password ")
bill_id = raw_input("enter Bill Identifier ")
bill_amt = raw_input("enter Billing Amount ")
file_location = raw_input("enter full path of your excel file ")
total_id_count = input("enter total number of students to be billed")

browser= Browser('chrome')
try:
     browser.visit('http://swd/Login.aspx')
except:
     browser.visit('http://universe.bits-pilani.ac.in:12349/Login.aspx')
     
browser.fill('TextBox1', user_name)
browser.fill('TextBox2', user_pass)

button = browser.find_by_id('loginBtn')
button.click()

bill_btn = browser.find_by_id('Button1')
bill_btn.click()

wb = xlwt.Workbook()
wsheet = wb.add_sheet('Failed')

try:
     workbookread = xlrd.open_workbook(file_location)
except:
     print "file location invalid"
     print >> sys.stderr, "Exception: %s" % str(e)
     sys.exit(1)

sheetread = workbookread.sheet_by_index(0)


count = 0
fail_count = 0
bill_count = 0
while True:
     if(total_id_count >= 50):
          bill_count = 50
     else:
          bill_count = total_id_count
          
     goto_billCode(browser)
     fill_billID(browser,bill_id)
     i = 0
     while i < bill_count :
          idno = sheetread.cell(count,0).value
          browser.fill('Wizard1$idnoTxt', idno)
          browser.fill('Wizard1$amountTxt', bill_amt)
          add = browser.find_by_id('Wizard1_addBtn')
          add.click()
          status = browser.find_by_id('Wizard1_idnoStatusLbl')
          count = count + 1
          i = i + 1
          if(status.text == 'Please Enter Valid IDNO'): 
               wsheet.write(fail_count,0,idno)
               fail_count = fail_count + 1

     goto_finish(browser)
     
     if(total_id_count<50):
          break
     
     total_id_count = total_id_count - 50


print ("total unsuccefull billings")
print (fail_count)
wb.save('failed.xls')
