import os
import pickle
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
import time
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import requests

Message = "Thank you"
excel_file = "replied_profile.xlsx"
cookies_file = "linkedin_cookies.pkl"
visit_records = "visit_profile.xlsx"

# service = Service(executable_path=r"C:\\Users\\mli91\\Desktop\\Kahan Chudasama\\chromedriver.exe")

chrom_opt = ChromeOptions()

browser = webdriver.Chrome()


def main_linkedin(driver):
    driver.get("https://www.linkedin.com/feed/")
    driver.maximize_window()
    wait = WebDriverWait(driver, 5)
    wait.until(EC.presence_of_element_located((By.ID, "global-nav")))

    messaging_link = driver.find_element(By.XPATH,
                                         '//*[@id="global-nav"]/div/nav/ul/li[4]/a/span')
    messaging_link.click()
    time.sleep(2)
    scroll_container = driver.find_element(By.CSS_SELECTOR, 'ul.msg-conversations-container__conversations-list')
    processed_ids = set()
    last_height = 0
    counter = 0

    while True:
        try:
            if check_internet():
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                conversation_items = soup.find_all('li', class_='msg-conversations-container__convo-item')

                for item in conversation_items:
                    try:
                        if check_internet():
                            item_id = item.get('id')

                            if item_id in processed_ids:
                                # script = f"document.getElementById('{item_id}').scrollIntoView();"
                                # driver.execute_script(script)
                                continue
                            if 'msg-conversation-card--occluded' in item.get('class', []):
                                script = f"document.getElementById('{item_id}').scrollIntoView();"
                                driver.execute_script(script)
                                # item = BeautifulSoup(driver.page_source, 'html.parser').find('li', id=item_id)
                                time.sleep(1)
                                soup = BeautifulSoup(driver.page_source, 'html.parser')
                                item = soup.find('li', id=item_id)
                            if 'msg-conversation-card--occluded' in item.get('class', []):
                                    continue
                            counter += 1
                            name_tag = item.find('h3', class_='msg-conversation-listitem__participant-names')
                            if not name_tag:
                                continue

                            name = name_tag.get_text(strip=True)
                            if os.path.exists('replied_profile.xlsx'):
                                df = pd.read_excel('replied_profile.xlsx')
                                if name in df['User Name'].values:
                                    processed_ids.add(item_id)
                                    driver.execute_script("arguments[0].scrollBy(0, 70);", scroll_container)
                                    continue

                            if os.path.exists('visit_profile.xlsx'):
                                df = pd.read_excel('visit_profile.xlsx')
                                if name in df['Profile Name'].values:
                                    processed_ids.add(item_id)
                                    driver.execute_script("arguments[0].scrollBy(0, 70);", scroll_container)
                                    continue

                            unread_found = item.find(
                                'div',
                                class_='msg-conversation-card__convo-item-container--unread msg-conversation-card msg-conversations-container__pillar'
                            )
                            if unread_found:
                                processed_ids.add(item_id)  # Mark as processed
                                driver.execute_script("arguments[0].scrollBy(0, 70);", scroll_container)

                                continue

                            link_xpath = f"//*[@id='{item_id}']//div[contains(@class, 'msg-conversation-listitem__link') and @role='button']"

                            # Locate and click the conversation link
                            link_tag = driver.find_element(By.XPATH, link_xpath)
                            driver.execute_script("arguments[0].scrollIntoView();", link_tag)  # Scroll into view
                            ActionChains(driver).move_to_element(link_tag).click().perform()

                            soup = BeautifulSoup(driver.page_source, 'html.parser')
                            message_box = soup.find(
                                'div',
                                {
                                    'aria-label': 'Write a message…',
                                    'aria-multiline': 'true',
                                    'contenteditable': 'true',
                                    'role': 'textbox',
                                    'class': 'msg-form__contenteditable t-14 t-black--light t-normal flex-grow-1 full-height notranslate'
                                }
                            )
                            if message_box:
                                messages = soup.find_all('li',
                                                             class_='msg-s-message-list__event')
                                if messages:
                                    last_message = messages[-1]
                                    sender_name_tag = last_message.find('span', class_='msg-s-message-group__name')
                                    sender_name = sender_name_tag.get_text(strip=True) if sender_name_tag else None

                                    if sender_name == name:
                                        send_message(driver, f"{Message}")
                                        profile_link_tag = last_message.find('a', class_='msg-s-event-listitem__link')
                                        user_profile_url = profile_link_tag['href'] if profile_link_tag else None
                                        action = f"Message already sent {name}"

                                        context = {
                                            'User Name': name,
                                            'User Profile': user_profile_url,
                                            'Action': action
                                        }
                                        data = [context]
                                        check_excel(data)
                                        time.sleep(1)
                                        driver.execute_script("arguments[0].scrollBy(0, 70);", scroll_container)

                                    else:

                                        profile_link_tag = last_message.find('a', class_='msg-s-event-listitem__link')
                                        if profile_link_tag:
                                            link = profile_link_tag['href'] if profile_link_tag else None
                                            record = {
                                                'Profile Name': name,
                                                'Profile Link': link
                                            }
                                            record = [record]
                                            visit_excel(record)
                            processed_ids.add(item_id)
                    except Exception as e:
                        check_internet()

            new_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
            if new_height == last_height:
                break
            last_height = new_height
        except Exception as e:
            check_internet()


def check_excel(data):
    if os.path.exists(excel_file):
        existing_df = pd.read_excel(excel_file)
        new_df = pd.DataFrame(data)
        updated_df = pd.concat([existing_df, new_df], ignore_index=False)
    else:
        updated_df = pd.DataFrame(data)

    updated_df.to_excel(excel_file, index=False)


def visit_excel(record):
    if os.path.exists(visit_records):
        existing_df = pd.read_excel(visit_records)
        new_df = pd.DataFrame(record)
        updated_df = pd.concat([existing_df, new_df], ignore_index=False)
    else:
        updated_df = pd.DataFrame(record)

    updated_df.to_excel(visit_records, index=False)


def send_message(driver, message):
    try:
        clear_textbox = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.msg-form__contenteditable"))
        )
        time.sleep(1)

        actions = ActionChains(driver)
        actions.move_to_element(clear_textbox).click().perform()

        clear_textbox.send_keys(Keys.CONTROL + "a")
        clear_textbox.send_keys(Keys.BACKSPACE)

        clear_textbox.send_keys(message)

        clear_textbox.send_keys(Keys.RETURN)
        time.sleep(1)

    except Exception as a:
        print(f"Restart Script: {a}")


def login_manually(driver):
    driver.get('https://www.linkedin.com/login')

    time.sleep(40)
    with open(cookies_file, "wb") as file:
        pickle.dump(driver.get_cookies(), file)


def login_with_cookies(driver):
    time.sleep(1)
    driver.get('https://www.linkedin.com')
    if os.path.exists(cookies_file):
        with open(cookies_file, "rb") as file:
            cookies = pickle.load(file)
            for cookie in cookies:
                driver.add_cookie(cookie)

        driver.refresh()

        main_linkedin(driver)
    else:
        login_manually(driver)
        main_linkedin(driver)


def check_internet():
    while True:
        try:
            response = requests.head('https://www.google.com/', timeout=2)
            if response.status_code == 200:
                return True
        except requests.RequestException:
            continue

        time.sleep(.2)


is_internet = check_internet()

if is_internet:
    try:
        login_with_cookies(browser)
    except:
        check_internet()
    finally:
        browser.quit()
        print("Driver closed.")
