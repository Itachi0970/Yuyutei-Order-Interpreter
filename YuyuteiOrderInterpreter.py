from types import NoneType
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

email_to_share_to = 'yuyutei-order-interpreter-501@yuyutei-order-interpreter.iam.gserviceaccount.com'
layout = [[sg.Text('Downloading ChromeDriver...')]]
window = sg.Window('ChromeDriver', layout).Finalize()
window.read(timeout=0)

try: 
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

except PermissionError:
    pass

window.close()

while True:
    layout = [  [sg.Text('Share the spreadsheet with this email before continuing (Even if your sheet is public):'), sg.InputText(email_to_share_to, use_readonly_for_disable=True, disabled=True, key='-IN-'), sg.Button('Copy')],
                [sg.Text('Enter the name of the spreadsheet:'), sg.InputText()],
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
                window.close()
                break


layout = [[sg.Text('YYT Interpreter is starting...')]]             
window = sg.Window('YYT Interpreter', layout).Finalize()
window.read(timeout=0)


order_form = sheet.worksheet("Order Form")

try:
    header_list = order_form.row_values(1)
    game_column = header_list.index("Game")
    buyer_column = header_list.index("Buyer Name")
    card_number_column = header_list.index("Set/Card #")
    card_name_column = header_list.index("Card Name (EN/JP)")
    card_amount_column = header_list.index("Amount")
    card_price_column = header_list.index("Listed price")
    total_price_column = header_list.index("Total")
    comments_column = header_list.index("Comments")
except ValueError:
    pass

list_of_urls = []

for row in order_form.get_all_values()[1:]:
    url = [row[card_amount_column], row[game_column], row[card_number_column], row[card_name_column], row[card_price_column], row[buyer_column], row[total_price_column], row[comments_column]]
    list_of_urls.append(url)

session = requests.Session()
API_ENDPOINT = "https://yuyu-tei.jp/api/cart/add.php"

discrepancies_to_write = []
discrepancies_to_write.append(["Buyer Name", "Set/Card #", "Card Name (EN/JP)", "Amount", "Buyer entered price", "Buyer entered Total", "Closest actual price", "Closest actual total", "Game", "Comments", "Discrepancy"])

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
# url[1] is game
# url[2] is card ID
# url[3] is card name
# url[4] is card price
# url[5] is buyer name
# url[6] is total price
# url[7] is comments

window = sg.Window('Progress Meter', layout)
progress_bar = window['progressbar']
window.read(timeout=0)

for url in list_of_urls:

    # Get the game codes to input into the URL based on the game in the spreadsheet
    game_code = ""
    gid = 0
    if "weiss" in url[1].lower():
        game_code = "ws"
        gid = 7
    elif url[1].lower() == "ygo" or url[1].lower() == "yugioh" or url[1].lower() == "yu-gi-oh":
        game_code = "ygo"
        gid = 32
    elif url[1].lower() == "wixoss":
        game_code = "wx"
        gid = 21
    elif url[1].lower() == "rush" or "rush duel" in url[1].lower():
        game_code = "yrd"
        gid = 42
    elif "fgo" in url[1].lower() or "fate" in url[1].lower():
        game_code = "fgoac"
        gid = 38
    elif "chaos" in url[1].lower():
        game_code = "chaos"
        gid = 8
    elif "vg" in url[1].lower() or "vanguard" in url[1].lower():
        game_code = "vg"
        gid = 13
    elif "prememo" in url[1].lower() or "precious" in url[1].lower():
        game_code = "pm"
        gid = 11
    elif "rebirth" in url[1].lower() or "re:birth" in url[1].lower():
        game_code = "re"
        gid = 41
    elif "z/x" in url[1].lower() or "zillions" in url[1].lower():
        game_code = "zx"
        gid = 20
    elif "lyce" in url[1].lower():
        game_code = "lo"
        gid = 34
    elif "emblem" in url[1].lower() or "FE" in url[1]:
        game_code = "fe"
        gid = 36
    elif "pok" in url[1].lower():
        game_code = "poc"
        gid = 31
    elif "masters" in url[1].lower():
        game_code = "dm"
        gid = 35
    elif "spirits" in url[1].lower():
        game_code = "bs"
        gid = 39
    elif "digi" in url[1].lower():
        game_code = "digi"
        gid = 43
    elif "kan" in url[1].lower():
        game_code = "kan"
        gid = 25
    elif "gab" in url[1].lower() or "gundam" in url[1].lower():
        game_code = "gab"
        gid = 44
    elif "dragon" in url[1].lower() or "sdbh" in url[1].lower():
        game_code = "dcd"
        gid = 37

    # If game was not found, make a discrepancy entry
    if gid == 0:
            discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], " ", " ", url[1], url[7], "Game \"" + url[1] + "\" is currently not supported. Add item manually and tell David to include the game in the next version."])
            window['progressbar'].update_bar(list_of_urls.index(url)+1)
            continue

    kizu = 0
    page = session.get("https://yuyu-tei.jp/game_" + game_code + "/sell/sell_price.php?name=" + url[2])
    soup = BeautifulSoup(page.content, 'html.parser')
    prices = soup.find_all('p', class_='price')
    cardFound = False
    someCardsFound = False
    price_string = ""
    if len(prices) >= 1:
        someCardsFound = True
    for price in prices:
        price_string = price.find('b').contents[0]
        price_string = price_string.split("円",1)[0]

        if str(price_string) == str(url[4]):
            cardFound = True
            full_card = price.parent.parent.parent
            card_link = full_card.find('div', class_='image_box').find('a').get('href')
            ver = card_link.split('VER=',1)[1].split('&',1)[0]
            cid = card_link.split('CID=',1)[1].split('&',1)[0]
            payload = "{\"quantity\":\"" + str(url[0]) + "\",\"mode\":\"sell\",\"item\":{\"gid\":\"" + str(gid) + "\",\"ver\":\"" + str(ver) + "\",\"cid\":\"" + str(cid) + "\",\"kizu\":\"" + str(kizu) + "\"}}"
            post = session.post(API_ENDPOINT, data=payload)
            print("Item " + url[2] + " " + url[3] + " added to cart.")
            window['progressbar'].update_bar(list_of_urls.index(url)+1)
            break

    if cardFound == False:
        prices_extracted = []
        for price in prices:
            price_string = price.find('b').contents[0]
            price_string = price_string.split("円",1)[0]
            prices_extracted.append(price_string)
        closest_price = 0
        if len(prices_extracted) >= 1:
            closest_price = min(prices_extracted, key=lambda list_value : abs(int(list_value) - int(url[4])))
            

        # Card could not be found, find the closest price, record it, then check the damaged section.
        kizu = 1
        page = session.get("https://yuyu-tei.jp/game_" + game_code + "/sell/sell_price.php?name=" + url[2] + "&kizu=1")
        soup = BeautifulSoup(page.content, 'html.parser')
        prices = soup.find_all('p', class_='price')
        cardFound = False
        if len(prices) >= 1:
            someCardsFound = True
        for price in prices:
            price_string = price.find('b').contents[0]
            price_string = price_string.split("円",1)[0]

            if str(price_string) == str(url[4]):
                cardFound = True
                full_card = price.parent.parent.parent
                card_link = full_card.find('div', class_='image_box').find('a').get('href')
                ver = card_link.split('VER=',1)[1].split('&',1)[0]
                cid = card_link.split('CID=',1)[1].split('&',1)[0]
                payload = "{\"quantity\":\"" + str(url[0]) + "\",\"mode\":\"sell\",\"item\":{\"gid\":\"" + str(gid) + "\",\"ver\":\"" + str(ver) + "\",\"cid\":\"" + str(cid) + "\",\"kizu\":\"" + str(kizu) + "\"}}"
                post = session.post(API_ENDPOINT, data=payload)
                print("Item " + url[2] + " " + url[3] + " added to cart.")
                window['progressbar'].update_bar(list_of_urls.index(url)+1)
                break

        if cardFound == False:
            # Card still not found, find out why.
            print("Item " + url[2] + " " + url[3] + " NOT added to cart.")
            window['progressbar'].update_bar(list_of_urls.index(url)+1)
            if someCardsFound:
                discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], int(closest_price), int(closest_price)*int(url[0]), url[1], url[7], "Card not found at this price. Add item manually."])
            elif someCardsFound == False:
                discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], 0, 0, url[1], url[7], "Card not found. Add item manually."])
            continue

    # Check cart page for items that didn't have enough stock
    cart_page = session.get("https://yuyu-tei.jp/sell_cart/cart.php")
    soup = BeautifulSoup(cart_page.content, 'html.parser')
    newAmounts = soup.find('div', class_='feedback_box message_block error_message')
    if type(newAmounts) != NoneType:
        newAmounts = newAmounts.find_all('li')
        for amount in newAmounts:
            new_amount = amount.contents[0]
            new_amount = new_amount.split('上限',1)[1].split('枚',1)[0]
            discrepancies_to_write.append([url[5], url[2], url[3], url[0], url[4], url[6], 0, 0, url[1], url[7], "Not enough stock. " + new_amount + " were added to cart instead of " + url[0]])

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

if email_entered:
    webbrowser.open(discrepancy_sheet_url)
