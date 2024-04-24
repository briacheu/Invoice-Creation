
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import time
import pandas as pd

from datetime import date



#initialize Selenium
options = Options()
options.add_experimental_option('detach', True)
browser = webdriver.Chrome()
#browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
#System.setProperty("webdriver.chrome.driver", "C:\\Users\\briacheu_a\\Desktop\\chromedriver-win64\\chromedriver.exe")
#browser = webdriver.Chrome(ChromeDriverManager(version='117.0.5938.88').install())




##0-------LOGIN

#load Mirakl
browser.get("https://marketplace.bestbuy.ca/")

#fill out login info (username > next, password > sign in)
user = browser.find_element(By.ID, "username")
user.send_keys("briacheung@bestbuycanada.ca")
next = browser.find_element(By.ID, "submitButton")
next.click()

time.sleep(1) #time to allow the page to load
passw = browser.find_element(By.ID, "password")
passw.send_keys("anim$T234234")

next = browser.find_element(By.NAME, "action")
next.click()

time.sleep(40)
#signin = browser.find_element(By.CLASS_NAME, "c3fc7dc62")
#signin.click()

#time period for password and 2FA
time.sleep(40)
#NOW YOU DO 2FA

##1-------SETUP
#setup dataframes. inv -- invoice details / tax -- province to tax document

df_invoice = pd.read_excel("H:\eCommerce\Business Team\Operations\Operations\Mirakl Invoice Automation\invoice_details.xlsx", sheet_name = "Invoice")

df_tax = pd.read_csv("C:\\Users\\briacheu_a\\Desktop\\Python\\Invoice Creation\\tax.csv")

df_invoice['Processing Note'] = 'Not processed'


##2-------FUNCTION TO FILL OUT INVOICE


def create_invoice(entry):

    #open up invoice creation page
    browser.get("https://marketplace.bestbuy.ca/sellerpayment/operator/accounting-document/manual/create")
    time.sleep(3) #time to allow the page to load

    #invoice or credit
    if df_invoice.iloc[entry,8] == 'Credit': 
        dropdown_invcred = Select(browser.find_element(By.ID, "type"))
        dropdown_invcred.select_by_value('MANUAL_CREDIT')

    #open up menu
    menu = browser.find_element(By.ID, "input-search-single-shop")
    menu.click()

    #fill in form with name, select store based on dataframe
    mirakl_store = browser.find_element(By.XPATH, value='//*[@id="autocomplete-shop-menu"]/div[1]/div/span/input')
    store = df_invoice.iloc[entry, 2]
    mirakl_store.send_keys(store)
    time.sleep(2) #time delay to load the radio button
    #browser.find_element(By.CLASS_NAME,"mui-radio-icon").click() #fails if there are two entries and the one you want isn't first
    xpath_search = '//*[@class="mui-suggestions-container"]//*[text()="' + store + '"]//div[@class="mui-radio-icon"]'

    browser.find_element(By.XPATH, value = xpath_search).click()
    #browser.find_element(By.XPATH, value='//*[@class="mui-suggestions-container"]//label[text()="Tech Solutions"]').click

    #date
    browser.find_element(By.ID, "items[0].operationDate").click()
    time.sleep(2)
    browser.find_element(By.XPATH, value="//span[@class='flatpickr-day today']").click()

    #description
    desc = df_invoice.iloc[entry, 10]
    browser.find_element(By.ID, "items[0].description").send_keys(desc)

    #qty
    browser.find_element(By.ID, "items[0].quantity").send_keys(1)

    #price
    price = df_invoice.iloc[entry, 9]
    browser.find_element(By.ID, "items[0].amount").send_keys(str(price))

    #add tax -- if taxable
    if df_invoice.iloc[entry,7] == 'Y': 

        #get province for tax
        province = df_invoice.iloc[entry, 5]
        print(province)
        x=0

        #add in 1st tax 
        tax_type = df_tax[df_tax['Province'] == province].iloc[0,1]
        dropdown_tax = Select(browser.find_element(By.ID, "items[0].taxes[" + str(x) + "].taxCode"))
        dropdown_tax.select_by_value(tax_type)

        #add in 2nd tax (if applicable)
        tax_type2 = df_tax[df_tax['Province'] == province].iloc[0,2]
        if pd.isna(tax_type2) == False:
            browser.find_element(By.XPATH, value="//button[@class='mui-button-padding-vertical-only btn btn-link btn-link-primary']").click()
            x+=1
            dropdown_tax = Select(browser.find_element(By.ID, "items[0].taxes[" + str(x) + "].taxCode"))
            dropdown_tax.select_by_value(tax_type2)
        
        
        #add BC PST, skip if province is BC -- removed 7/4 due to new legislation
        #if province != 'BC':
        #    browser.find_element(By.XPATH, value="//button[@class='mui-button-padding-vertical-only btn btn-link btn-link-primary']").click()
        #    x+=1
        #    dropdown_tax = Select(browser.find_element(By.ID, "items[0].taxes[" + str(x) + "].taxCode"))
        #    dropdown_tax.select_by_value('BC_PST')

    #add in TAX ZERO if not taxable
    elif df_invoice.iloc[entry,7] == 'N': #'Taxable?' column
        dropdown_tax = Select(browser.find_element(By.ID, "items[0].taxes[0].taxCode"))
        dropdown_tax.select_by_value('TAXZERO')

    #submit as draft
    if df_invoice.iloc[entry,6] == 'Draft': #'Submit or Draft' column
        time.sleep(1)
        browser.find_element(By.XPATH, value="/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/button[1]").click() 

        confirmation = store + " for " + str(price) + "(pre-tax) has been successfully created"
        df_invoice.at[entry, 'Processing Note'] = confirmation
    
    #submit and confirm -- log the invoice number
    if df_invoice.iloc[entry,6] == 'Submit': #'Submit or Draft' column
        #submit fully
        time.sleep(1)
        #browser.find_element(By.XPATH, value="/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/button[1]").click() 
        browser.find_element(By.XPATH, value="/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/button[2]").click() 
        
        browser.find_element(By.XPATH, value="//button[@class='btn btn-solid btn-solid-danger']").click()

        #grab invoice number (pulls the most recent invoice number)
        browser.get("https://marketplace.bestbuy.ca/sellerpayment/operator/accounting-document/list/to-sellers?limit=25")
        time.sleep(2)
        confirmation = "done"
        try: 
            confirmation = browser.find_element(By.XPATH, value='/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div/table/tbody/tr[1]/td[1]/div/div/a').text
        except:
            confirmation = "invoice number can't be pulled"
        df_invoice.at[entry, 'Processing Note'] = confirmation


    #confirm

    
    print(confirmation)

##3-------RUN FUNCTION THROUGH ENTIRE DATAFRAME

for entry in range(len(df_invoice)):
    create_invoice(entry)

time.sleep(20)

#save invoice file 
today = str(date.today())
filepath = "H:\\eCommerce\\Business Team\\Operations\\Operations\\Mirakl Invoice Automation\\invoice_details_" + today + ".csv"
df_invoice.to_csv(filepath)

print("Submission complete")

#rewrite import file
import shutil

source_file = r'H:\eCommerce\Business Team\Operations\Operations\Mirakl Invoice Automation\invoice_details_template.xlsx'
destination_file = r'H:\eCommerce\Business Team\Operations\Operations\Mirakl Invoice Automation\invoice_details.xlsx'


#shutil.copy2(source_file, destination_file)