from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
import pandas as pd
import time
import sys
import warnings
import os
from multiprocessing import freeze_support
from datetime import datetime
import smtplib as smtp
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    #chrome_options  = webdriver.ChromeOptions()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--disable-gpu')
    #chrome_options.add_argument('--headless')
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-first-run")
    chrome_options.add_argument("--no-service-autorun")
    chrome_options.add_argument("--no-default-browser-check")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--disable-plugins-discovery")
    chrome_options.add_argument("--incognito")
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(version_main=108, options=chrome_options)
    driver.maximize_window()
    driver.set_page_load_timeout(60)
    return driver

def login(driver, username, password):

    url = 'https://www.google.com/alerts?hl=en-GB'
    driver.get(url)
    buttons = wait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
    for button in buttons:
        if button.get_attribute('textContent') == 'Sign in':
            driver.execute_script("arguments[0].click();", button)
            time.sleep(3)
            break

    user_form = wait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@type='email']")))
    user_form.send_keys(username)
    time.sleep(2)

    buttons = wait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "button")))
    for button in buttons:
        if button.get_attribute('textContent') == 'Next':
            driver.execute_script("arguments[0].click();", button)
            time.sleep(3)
            break    
        
    pass_form = wait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
    pass_form.send_keys(password)
    time.sleep(2)

    buttons = wait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "button")))
    for button in buttons:
        if button.get_attribute('textContent') == 'Next':
            driver.execute_script("arguments[0].click();", button)
            time.sleep(5)
            break

def change_settings(driver, setting, details):

    setting_id = {'How many':1, 'Sources':2, 'Language':3, 'Region':4,'How often':0}
    # opening the settings menu
    button = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.show_options")))
    driver.execute_script("arguments[0].click();", button)
    time.sleep(2)
    # configuring the alert settings
    div = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@id='create-alert-options']")))
    trs = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
    for tr in trs:
        text = tr.get_attribute('textContent')
        if setting in text:
            button = wait(tr, 5).until(EC.presence_of_element_located((By.TAG_NAME, "div"))).click()
            time.sleep(2)
            break

    if setting != 'Deliver to':
        divs = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.goog-menu.goog-menu-noicon")))
        for i, div in enumerate(divs):
            if i == setting_id[setting]:
                if setting == 'Sources':
                    tag = "span"
                else:
                    tag = 'div.goog-menuitem-content'
                options = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, tag)))
                for option in options:
                    text = option.get_attribute('textContent')
                    for elem in details:
                        if elem in text:
                            if tag == "span":
                                parent = wait(option, 2).until(EC.presence_of_element_located((By.XPATH, '..')))
                                parent.click()
                            else:
                                option.click()
                            time.sleep(1)
    else:
        # handling deliver to setting
        feed_set = False
        div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.goog-menu.goog-menu-vertical')))
        options = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.goog-menuitem-content')))
        for option in options:
            text = option.get_attribute('textContent')
            if feed_set: break
            for elem in details:
                if elem in text:
                    option.click()
                    time.sleep(1)
                    feed_set = True
                    break

    wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Create an alert about...']"))).click()

def get_alerts(driver, brands, settings, mail):

    #driver2 = initialize_bot()
    url = 'https://www.google.com/alerts' 
    driver.get(url)
    data = pd.DataFrame()
    print('-'*75)
    alerts = []
    try:
        # getting the current alerts
        div = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@id='manage-alerts-div']")))
        lis = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.alert_instance")))
        for li in lis:
            alerts.append(li.get_attribute('textContent').strip())
    except:
        pass
    
    # creating alerts
    for brand in brands:
        keywords = brands[brand]
        for keyword in keywords:
            # skip alerts already created
            if keyword in alerts: continue
            print(f'Creating an alert for keyword: "{keyword}" related to brand: "{brand}"')  
            try:
                form = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Create an alert about...']")))
                form.clear()
                form.send_keys(keyword)
                time.sleep(5)
                # static settings
                for _ in range(3):
                    try:
                        change_settings(driver, 'How often', ['As-it-happens'])
                        change_settings(driver, 'Deliver to', ['RSS feed'])
                        break
                    except:
                        continue
                # user defined settings
                for setting in settings:
                    for _ in range(3):
                        try:
                            change_settings(driver, setting, settings[setting])
                            break
                        except:
                            continue

                time.sleep(2)

                # add the alert
                wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//span[@id='create_alert']"))).click()
                time.sleep(3)
            except Exception as err:
                print(f'The below error occurred while getting the alerts for: {brand}')
                print(str(err))

    # getting the current alerts
    alerts = {} 
    div = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[@id='manage-alerts-div']")))
    lis = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.alert_instance")))
    for li in lis:
        alert = li.get_attribute('textContent').strip()
        link = wait(li, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
        alerts[alert] = link

    # scraping the alerts
    data = pd.DataFrame()
    for brand in brands:
        keywords = brands[brand]
        for keyword in keywords:
            print(f'scraping the alerts for keyword: "{keyword}" related to brand: "{brand}"')  
            try:
                url = alerts[keyword]
                df = get_feed(driver, url, brand, keyword)
                data = data.append(df)
            except:
                pass

    print('Outputting scraped data ...')
    # output the daily alerts separately
    data.drop_duplicates(inplace=True)
    stamp = datetime.now().strftime("%m_%d_%Y")
    filename = f'Google_Alerts_{stamp}.xlsx'
    data.to_excel(filename, index=False)
    if mail != '':
        try:
            send_mail(filename, mail)
        except Exception as err:
            print('The following error occurred while sending the alerts mail:\n')
            print(str(err))
            print('-'*75)
    # merge the new alerts with the database
    path = os.getcwd()
    if '//' in path:
        path += '//Google_Alerts.xlsx'
    else:
        path += '\Google_Alerts.xlsx'

    if not os.path.isfile(path):
        data.to_excel('Google_Alerts.xlsx', index=False)
    else:
        try:
            df = pd.read_excel('Google_Alerts.xlsx')
            df['Article Date'] = pd.to_datetime(df['Article Date'])
            data['Article Date'] = pd.to_datetime(data['Article Date'])
            data = data.append(df)
            data.drop_duplicates(inplace=True)
            data.sort_values(by=['Brand', 'Article Date'], ascending=True, inplace=True)
            data.to_excel('Google_Alerts.xlsx', index=False)
        except Exception as err:
            print('Failed to merge the scraped data with the sheet "Google_Alerts.xlsx" due to the below error:')
            print(str(err))
            stamp = datetime.now().strftime("%m_%d_%Y")
            data.to_excel(f'Google_Alerts_{stamp}.xlsx', index=False)

    print('-'*75)
    print('Process Completed Successfully!')
    driver.quit()
    #driver2.quit()
    sys.exit()

def get_feed(driver, url, brand, keyword):
    """Return a Pandas dataframe containing the RSS feed contents.

    Args: 
        url (string): URL of the RSS feed to read.

    Returns:
        df (dataframe): Pandas dataframe containing the RSS feed contents.
    """
    
    driver.get(url)
    
    df = pd.DataFrame(columns = ['Brand', 'Keyword', 'Article Title', 'Article Link', 'Article Summary', 'Article Publisher', 'Article Date'])

    items = wait(driver, 1).until(EC.presence_of_all_elements_located((By.TAG_NAME, "entry")))

    for item in items:        

        title = wait(item, 2).until(EC.presence_of_element_located((By.TAG_NAME, "title"))).get_attribute('textContent').replace('H&amp;M', '').replace('<b>', '').replace('</b>', '').strip()    
        link = wait(item, 2).until(EC.presence_of_element_located((By.TAG_NAME, "link"))).get_attribute('href')
        publisher = ''    
        if ' - ' in title:
            elems = title.split(' - ')
            publisher = elems[-1].strip()
        elif ' | ' in title:
            elems = title.split(' | ')
            publisher = elems[-1].strip()
        else:
            publisher = link.replace('www.', '').replace('http://', 'https://').split('https://')[-1].split('.')[0]
        try:
            publisher = publisher.title()
        except:
            pass
        date = wait(item, 2).until(EC.presence_of_element_located((By.TAG_NAME, "published"))).get_attribute('textContent').split('T')[0].strip()
        summary = wait(item, 2).until(EC.presence_of_element_located((By.TAG_NAME, "content"))).get_attribute('textContent').replace('&nbsp;...', '').replace('H&amp;M', '').replace('<b>', '').replace('</b>', '').strip()


        row = {'Brand': brand, 'Keyword': keyword, 'Article Title': title, 'Article Link': link, 'Article Summary': summary, 'Article Publisher': publisher, 'Article Date': date}
        df = df.append(row, ignore_index=True)

    return df

def send_mail(filename, mail):

    stamp = datetime.now().strftime("%m/%d/%Y")
    connection = smtp.SMTP_SSL('smtp.gmail.com', 465)  
    email_addr = 'scraper.notification1@gmail.com'
    email_passwd = 'fvkeaxrfcurewyou'
    connection.login(email_addr, email_passwd)
    subject = f"Google Alerts For {stamp}"
    body = "Hi,\n\nKindly check the attached sheet for the daily alerts.\n\nRegards."
    #message = f'Subject: {subject}\n\n{body}'
    msg = MIMEMultipart()
    msg['From'] = email_addr
    msg['To'] = mail
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(body))

    with open(filename, "rb") as f:
        part = MIMEApplication(
            f.read(),
            Name=basename(filename))
    # After the file is closed
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(filename)
    msg.attach(part)
    connection.sendmail(from_addr=email_addr, to_addrs=mail, msg=msg.as_string())
    connection.close()

def get_inputs():

    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '//' in path:
        path += '//openrice_settings.csv'
    else:
        path += '\openrice_settings.csv'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "openrice_settings.csv"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        df = pd.read_csv(path)      
    except:
        print('Error: Failed to process the input sheet')
        input('Press any key to exit')
        sys.exit(1)

    cols = df.columns
    for col in cols:
        df[col] = df[col].astype(str)

    df_brands = df[['Brand', 'Keyword']]
    inds = df_brands.index
    brands = {}
    for ind in inds:
        row = df_brands.loc[ind]
        brand = row['Brand']
        keyword = row['Keyword']
        if keyword == 'nan':
            keyword = brand
        if brand == 'nan':
            continue
        if brand in brands:
            brands[brand].append(keyword)
        else:
            brands[brand] = [keyword]    
            
    df_settings = df[['Sources', 'Language', 'Region', 'How many']]
    inds = df_settings.index
    cols = df_settings.columns
    settings = {}
    for ind in inds:
        row = df_settings.loc[ind]
        for col in cols:
            value = row[col]
            if value == 'nan':
                continue
            if col in settings:
                settings[col].append(value)
            else:
                settings[col] = [value]
    
    try:
        mail = df.loc[0, 'Deliver To']
        if mail == 'nan':
            print('Warning: No valid E-mail address is provided for sending the scraped alerts, no mails will be sent!')
            print('-'*75)
            mail = ''
    except:
        print('Warning: No valid E-mail address is provided for sending the scraped alerts, no mails will be sent!')
        print('-'*75)
        mail == ''    
        
    try:
        username = df.loc[0, 'Gmail username']
        if username == 'nan':
            print('Warning: No valid Gmail account is provided for the alerts')
            print('-'*75)
            input('press any key to exit')
            sys.exit()
    except:
            print('Warning: No valid Google account is provided for the alerts')
            print('-'*75)
            input('press any key to exit')
            sys.exit()    
            
    try:
        password = df.loc[0, 'Gmail Password']
        if password == 'nan':
            print('Warning: No valid Google account password is provided for the alerts')
            print('-'*75)
            input('press any key to exit')
            sys.exit()
    except:
            print('Warning: No valid Google account password is provided for the alerts')
            print('-'*75)
            input('press any key to exit')
            sys.exit()

    return brands, settings, mail, username, password

def main():

    freeze_support()

    print('Processing Inputs...')
    print('-'*75)
    brands, settings, mail, username, password = get_inputs()
    print('Initializing the bot...')
    print('-'*75)
    driver = initialize_bot()
    try:
        login(driver, username, password)
    except Exception as err:
        driver.quit()
        print('Error - Failed to login to the Google account due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    try:
        get_alerts(driver, brands, settings, mail)
    except Exception as err:
        driver.quit()
        print('The following error occurred:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()
if __name__ == "__main__":

    main()