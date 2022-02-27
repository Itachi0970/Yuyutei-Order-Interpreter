from urllib import request
import webbrowser
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
import time
import requests
from requests.utils import dict_from_cookiejar
from bs4 import BeautifulSoup
import PySimpleGUI as sg
import getpass
from tkinter import Tk
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import wget
import zipfile
import os
from os import path

def last_filled_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list))

email_to_share_to = 'yuyutei-order-interpreter-501@yuyutei-order-interpreter.iam.gserviceaccount.com'
layout = [[sg.Text('Downloading ChromeDriver...')]]
window = sg.Window('ChromeDriver', layout).Finalize()
window.read(timeout=0)

if path.exists("chromedriver.exe"):
    os.remove("chromedriver.exe")
chromedriver_url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
response = requests.get(chromedriver_url)
version_number = response.text
download_url = "https://chromedriver.storage.googleapis.com/" + version_number + "/chromedriver_win32.zip"
# download the zip file using the url built above
latest_driver_zip = wget.download(download_url,'chromedriver.zip')

# extract the zip file
with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
    zip_ref.extractall() # you can specify the destination folder path here
# delete the zip file downloaded above
os.remove(latest_driver_zip)

window.close()


while True:
    layout = [  [sg.Text('Share the spreadsheet with this email before continuing (Even if your sheet is public):'), sg.InputText(email_to_share_to, use_readonly_for_disable=True, disabled=True, key='-IN-'), sg.Button('Copy')],
                [sg.Text('Enter the name of the spreadsheet: '), sg.InputText()],
                [sg.Text('Enter your Google email to see a list of discrepancies (optional):'), sg.InputText()],
                [sg.Button('Ok'), sg.Button('Quit')],
                [sg.Button('Submit', visible=False, bind_return_key=True)]]

    window = sg.Window('Enter information', layout).Finalize()

    while True:
        event, values = window.read()
        if event in (None, 'Quit'):
            quit()
        elif event in ('Copy'):
            r = Tk()
            r.withdraw()
            r.clipboard_append(email_to_share_to)
        elif event in ('Ok', 'Submit'):
            spreadsheetName = values[0]
            email = values[1]
            break

    window.close()

    # define the scope
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('yuyutei-order-interpreter-19e9d666de35.json', scope)

    # authorize the clientsheet 
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    try:
        sheet = client.open(spreadsheetName)
        break
    except Exception:
        layout = [[sg.Text('Spreadsheet not found. Please make sure you entered the name correctly and shared it with the correct email.')],
                [sg.Button('Ok')]]
                
        window = sg.Window('Spreadsheet not found', layout).Finalize()
        while True:
            event, values = window.read()
            if event in (None, 'Ok'):
                break


layout = [[sg.Text('YYT Interpreter is starting...')]]             
window = sg.Window('YYT Interpreter', layout).Finalize()
window.read(timeout=0)


order_form = sheet.worksheet("Order Form")

header_list = order_form.row_values(1)
game_column = header_list.index("Game")
buyer_column = header_list.index("Buyer Name")
card_number_column = header_list.index("Set/Card #")
card_name_column = header_list.index("Card Name (EN/JP)")
card_amount_column = header_list.index("Amount")
card_price_column = header_list.index("Listed price")
total_price_column = header_list.index("Total")
comments_column = header_list.index("Comments")
url_column = header_list.index("URL")

list_of_urls = []

for row in order_form.get_all_values()[1:]:
    url = [row[card_amount_column], row[url_column], row[card_number_column], row[card_name_column], row[card_price_column], row[buyer_column], row[total_price_column], row[comments_column]]
    list_of_urls.append(url)

session = requests.Session()
API_ENDPOINT = "https://yuyu-tei.jp/api/cart/add.php"

discrepancies_to_write = []
discrepancies_to_write.append(["Buyer Name", "Set/Card #", "Card Name (EN/JP)", "Amount", "Buyer entered price", "Buyer entered Total", "Actual price", "Actual total", "URL", "Comments", "Discrepancy"])

discrepancy_sheet = client.create(spreadsheetName + " Discrepancies")
discrepancy_sheet_url = "https://docs.google.com/spreadsheets/d/%s" % discrepancy_sheet.id
discrepancy_worksheet = discrepancy_sheet.worksheet("Sheet1")
email_entered = True
try:
    discrepancy_sheet.share(email, perm_type='user', role='writer', notify=False)
except Exception:
    email_entered = False
    pass


window.close()

progressbar = [
    [sg.ProgressBar(len(list_of_urls), orientation='h', size=(66, 10), key='progressbar')]
]
outputwin = [
    [sg.Output(size=(100,20))]
]

layout = [
    [sg.Frame('Progress',layout= progressbar)],
    [sg.Frame('Output', layout = outputwin)]
]

# url[0] is qty
# url[1] is URL
# url[2] is card ID
# url[3] is card name
# url[4] is card price
# url[5] is buyer name
# url[6] is total price
# url[7] is comments

window = sg.Window('Progress Meter', layout, keep_on_top=True)
progress_bar = window['progressbar']
window.read(timeout=0)

for url in list_of_urls:

    window.read(timeout=0)

    try:
        game = url[1].split("game_",1)[1].split("/",1)[0]
        gid = 0
        if game == "ws":
            gid = 7
        elif game == "ygo":
            gid = 32
        elif game == "wx":
            gid = 21
        elif game == "yrd":
            gid = 42
        else:
            gid = 0

        #if gid is still 0, game was not found
        if gid == 0:
            discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], " ", " ", url[1], url[7], "Game \"" + game + "\" is currently not supported. Add item manually and tell David to include the game in the next version."])
            window['progressbar'].update_bar(list_of_urls.index(url)+1)
            continue
        
    except IndexError:
        print("Item " + url[2] + " " + url[3] + " NOT added to cart.")
        window['progressbar'].update_bar(list_of_urls.index(url)+1)
        discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], " ", " ", url[1], url[7], "URL invalid or nonexistent. Add item manually."])
        continue

    try:
        page = session.get(url[1])
        soup = BeautifulSoup(page.content, 'html.parser')
        price = soup.find('p', class_='price').find('b').extract().contents[0]
        price = price.split("円",1)[0]
        if url[4] != price:
            discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], price, int(price)*int(url[0]), url[1], url[7], "Price does not match. Item was still added to cart automatically."])

    except Exception:
        print("Item " + url[2] + " " + url[3] + " NOT added to cart.")
        window['progressbar'].update_bar(list_of_urls.index(url)+1)
        discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], price, int(price)*int(url[0]), url[1], url[7], "Price somehow not found in webpage. This should never happen, please notify David."])
        continue

    kizu = 0
    try:
        if url[1].split("kizu=",1)[1] == "1":
            kizu = 1
    except IndexError:
        pass

    try:
        payload = "{\"quantity\":\"" + str(url[0]) + "\",\"mode\":\"sell\",\"item\":{\"gid\":\"" + str(gid) + "\",\"ver\":\"" + str(url[1].split("VER=",1)[1].split("&",1)[0]) + "\",\"cid\":\"" + str(url[1].split("CID=",1)[1].split("&",1)[0]) + "\",\"kizu\":\"" + str(kizu) + "\"}}"
        post = session.post(API_ENDPOINT, data=payload)

    except IndexError:
        print("Item " + url[2] + " " + url[3] + " NOT added to cart.")
        window['progressbar'].update_bar(list_of_urls.index(url)+1)
        continue

    print("Item " + url[2] + " " + url[3] + " added to cart.")
    window['progressbar'].update_bar(list_of_urls.index(url)+1)

cookies = dict_from_cookiejar(session.cookies)

print("Added items to cart!")
print("Please wait for page to refresh.")
useoptions = webdriver.ChromeOptions()
useoptions.add_experimental_option("detach", True)
useoptions.add_experimental_option('excludeSwitches', ['enable-logging'])
useoptions.add_argument("--log-level=3")
username = getpass.getuser()

useoptions.add_experimental_option('useAutomationExtension', False)

driver = webdriver.Chrome(options=useoptions)
driver.get("https://yuyu-tei.jp/sell_cart/cart.php")
for key, value in cookies.items():
    driver.add_cookie({'name': key, 'value': value})

driver.refresh()

window.close()

discrepancy_worksheet.update('A1', discrepancies_to_write)

layout = [[sg.Text('Keep note of items that did not have enough stock, as these are not detected in Discrepancies.')],
            [sg.Button('Ok')]]

window = sg.Window('Items added to cart', layout, keep_on_top=True).Finalize()
while True:
    event, values = window.read()
    if event in (None, 'Ok'): # if user closes window or clicks cancel
        break

driver.service.stop()

if email_entered:
    webbrowser.open(discrepancy_sheet_url)
