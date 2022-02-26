from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from openpyxl import Workbook

import time


MAX_WAIT = 10

with webdriver.Firefox() as browser:

    def wait_for_javascript():
        start_time = time.time()
        while True:
            try:
                welcome = browser.find_elements(By.XPATH, "//store-selector/section/header/div/h1")
                assert("Welcome to Fareway Stores" in welcome[0].text)
                return
            except (AssertionError, WebDriverException, IndexError) as e:
                if time.time() - start_time > MAX_WAIT:
                    raise e
                time.sleep(0.5)

    def wait_for_store_list(storeName):
        start_time = time.time()
        while True:
            try:
                stores = browser.find_elements(By.XPATH,  "//span[@class='store-select-store-name']")
                storeText = [store.text for store in stores]
                assert(storeName in storeText)
                return
            except (AssertionError, WebDriverException, IndexError) as e:
                if time.time() - start_time > MAX_WAIT:
                    browser.find_element(By.XPATH, "//a[@click.delegate='showFulfillmentTimes($event)']").click()
                time.sleep(0.5)

    browser.get("https://shop.fareway.com")
    wait_for_javascript()
    stores = browser.find_elements(By.XPATH,  "//span[@class='store-select-store-name']")
    storeText = [store.text for store in stores]
    buttons = browser.find_elements(By.XPATH, "//compose/ul/li/button")

    storesList = [list(stores) for stores in zip(storeText, buttons)]

    storeResults = {}

    for index, store in enumerate(storesList):
        wait_for_store_list(store[0])
        time.sleep(2)
        store[1] = browser.find_elements(By.XPATH, "//compose/ul/li/button")[index]
        print("after resetting button")
        button = store[1]
        button.location_once_scrolled_into_view
        button.click()
        time.sleep(2)
        
        try:
            if browser.find_element(By.XPATH, "//a[@click.delegate='showFulfillmentTimes($event)']").is_displayed():
                browser.find_element(By.XPATH, "//a[@click.delegate='showFulfillmentTimes($event)']").click()
                time.sleep(0.5)
        except ElementClickInterceptedException:
            pass
        
        
        time_slots = browser.find_elements(By.XPATH, "//div[@class='radio-box-text-title']")
        time_slots = [slot.text.rstrip() for slot in time_slots]
        slots_left = browser.find_elements(By.XPATH, "//div[@class='radio-box-text-subtitle']/span")
        slots_left = [slot.text.rstrip() for slot in slots_left]
        for i in range(len(slots_left)):
            if slots_left[i] == "3 Slots Left":
                slots_left[i] = 0
            elif slots_left[i] == "2 Slots Left":
                slots_left[i] = 1
            elif slots_left[i] == "1 Slot Left":
                slots_left[i] = 2
            elif slots_left[i] == "Slot full":
                slots_left[i] = 3
            else:
                slots_left[i] = "Error"

            
        storeResults[store[0]] = [list(data) for data in zip(time_slots, slots_left)]
        print(storeResults)

        # click "change store" button
        browser.find_element(By.XPATH, "//button[@click.trigger='controller.cancel()']").click()
        if browser.find_element(By.XPATH, "//a[@click.delegate='changeFulfillment($event)']").is_displayed():
            browser.find_element(By.XPATH, "//a[@click.delegate='changeFulfillment($event)']").click()

    wb = Workbook()
    initial_sheet = wb.active
    initial_sheet.title = "Fareway Orders"

    nextRow = 2


    for key, value in storeResults.items():
        initial_sheet.cell(row=nextRow, column=1, value=key)
        nextRow += 1 
        for item in value:
            initial_sheet.cell(row=nextRow, column=2, value=item[0])
            initial_sheet.cell(row=nextRow, column=3, value=item[1])
            nextRow += 1
        
        nextRow += 1

    wb.save('example.xlsx')