#Things to do:
#
# fix EXE conversion
#

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime
import pandas as pd
import openpyxl
import time


def fileWrite(sheetTitle, col1, col2, col3, col4, col5):
    data = {'col1': col1, 'col2': col2, 'col3': col3, 'col4': col4, 'col5': col5}
    df = pd.DataFrame(data)
    with pd.ExcelWriter('data.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists = 'overlay') as writer:
        df.to_excel(writer, sheet_name = sheetTitle, header = None, startrow = writer.sheets[sheetTitle].max_row, index = False)


def grabTextFromClassname(url : str, key : int):

    findString = "//div[@class='info-prices']/span[" + key + "]"
    webpage = webdriver.Chrome()
    webpage.get(url)
    time.sleep(2)
    petrolData = webpage.find_element(By.CLASS_NAME, "selected-price")
    dieselData = webpage.find_element(By.XPATH, findString)
    text = [petrolData.text, dieselData.text]
    webpage.quit()
    return(text)


def updateSheet(ampolLink : str, ampolIdentifier : str, ampolE10: bool, bpLink : str, bpIdentifier :str, bpE10: bool,  sheetTitle: str):
    ampolPrice = grabTextFromClassname(ampolLink, ampolIdentifier)
    bpPrice = grabTextFromClassname(bpLink, bpIdentifier)

    if(ampolE10):
        ampolPrice[0] = str(float(ampolPrice[0]) + 2)

    if(bpE10):
        bpPrice[0] = str(float(bpPrice[0]) + 2)

    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    col1 = [dt_string]
    col2 = [ampolPrice[0]]
    col3 = [ampolPrice[1]]
    col4 = [bpPrice[0]]
    col5 = [bpPrice[1]]

    fileWrite(sheetTitle, col1, col2, col3, col4, col5)

print("Lytton")
updateSheet("https://petrolspy.com.au/map/station/5212b1660364706598e39864", "3", False, 
            "https://petrolspy.com.au/map/station/574566b6e4b072e48718f237", "4", True,
            "Lytton")

print("\n\nGoodna")
updateSheet("https://petrolspy.com.au/map/station/5212b1660364706598e39b92", "4", False, 
            "https://petrolspy.com.au/map/station/5212adc90364fe88b9785a17", "3", False,
            "Goodna")

print("\n\nWindsor")
updateSheet("https://petrolspy.com.au/map/station/63980755b9914209d956f6a2", "4", False, 
            "https://petrolspy.com.au/map/station/5965a35eb99142726c4dd1a9", "3", False,
            "Windsor")

print("\n\nKenmore")
updateSheet("https://petrolspy.com.au/map/station/5212b1660364706598e39770", "4", False, 
            "https://petrolspy.com.au/map/station/5212adc90364fe88b97859d4", "3", True,
            "Kenmore")

print("\n\nJimboomba")
updateSheet("https://petrolspy.com.au/map/station/5212b1660364706598e395ca", "3", False, 
            "https://petrolspy.com.au/map/station/53665f0503648cbfe539926f", "3", False,
            "Jimboomba")

print("\n\nAnnerley")
updateSheet("https://petrolspy.com.au/map/station/5212b1660364706598e396e7", "4", False, 
            "https://petrolspy.com.au/map/station/5a249c15b991420497aa77a4", "4", False,
            "Annerley")