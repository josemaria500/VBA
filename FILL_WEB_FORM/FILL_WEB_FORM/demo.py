import sys
from easygui import *
from win32com.client import Dispatch
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
   
def check_open_excel(excel_name):
    try:
        xlApp = Dispatch("Excel.Application")   
        wb = xlApp.Workbooks.Item(excel_name)
        return wb
    except:
        msgbox(msg = "This file must be called by the WorkBook", title = "Error")
        sys.exit()
        
def read_excel_data(ws):
    allData = ws.cells(1,1).CurrentRegion
    numrows = allData.Rows.Count
    data = []
    for i in range(2, numrows):
        dict = {}
        dict["FirstName"] = str(ws.cells(i,1))
        dict["LastName"] = str(ws.cells(i,2))
        dict["Address"] = str(ws.cells(i,3))
        dict["Email"] = str(ws.cells(i,4))
        dict["CardType"] = str(ws.cells(i,5))
        dict["CardNumber"] = int(ws.cells(i,6))
        dict["VerifCode"] = int(ws.cells(i,7))
        data.append(dict)
    return data
   
def writewebform(data):
    url = "http://www.practiceselenium.com/check-out.html"
    caps = DesiredCapabilities().CHROME
    #caps["pageLoadStrategy"] = "normal"  #  complete
    caps["pageLoadStrategy"] = "eager"  #  interactive
    #caps["pageLoadStrategy"] = "none"
    driver = webdriver.Chrome(desired_capabilities=caps, executable_path= "chromedriver.exe")
    driver.get(url)
    assert "Check Out" in driver.title
    driver.maximize_window()
    for i in range(len(data)):
        elem = driver.find_element_by_id("email")
        elem.send_keys(data[i]["Email"])

        elem = driver.find_element_by_id("name")
        elem.send_keys(data[i]["FirstName"])

        elem = driver.find_element_by_id("address")
        elem.send_keys(data[i]["Address"])

        sel = Select(driver.find_element_by_id("card_type"))
        for opt in sel.options:
            sel.select_by_visible_text(data[i]["CardType"])

        elem = driver.find_element_by_id("card_number")
        elem.send_keys(data[i]["CardNumber"])

        elem = driver.find_element_by_id("cardholder_name")
        elem.send_keys(data[i]["FirstName"] + " " + data[i]["LastName"])

        elem = driver.find_element_by_id("verification_code")
        elem.send_keys(data[i]["VerifCode"])

        boton = driver.find_element_by_xpath("//button[contains(text(),'Place Order')]")
        boton.click()
        
        driver.get("http://www.practiceselenium.com/check-out.html")
    driver.close()
     

# 1st check: if "demo.xlsm" is open     
wb = check_open_excel("demo.xlsm")

# 2nd check, in a cell on the "Hide" sheet, appears "demo".
# We eliminate the possibility of having another file with the same name with the wrong information.
wb.Sheets("Hide").visible = True
ws = wb.Sheets("Hide")
if ws.cells(1,2).value != "Demo":
    msgbox(msg = "This file must be called by the WorkBook: Demo.xlsm", title = "Error")
    sys.exit()     
wb.Sheets("Hide").visible = 2 # xlSheetVeryHidden

# Load data from Excel
ws=wb.Sheets("Hoja1")
data = read_excel_data(ws)
# Fill web form
writewebform(data)

# Changing the cell value to "Finished" returns control to Excel
wb.Sheets("Hide").visible = True
ws = wb.Sheets("Hide")
ws.cells(2,2).value = "Finished"
wb.Sheets("Hide").visible = 2 # xlSheetVeryHidden
sys.exit()  