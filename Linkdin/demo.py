import os
import pickle
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChormOptions
import time
from bs4 import BeautifulSoup
import random
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import requests


start_time = time.time()

message_file = "messages.txt"
excel_file = "profile_records.xlsx"
cookies_file = "linkedin_cookies2.pkl"
log_file = 'log.txt'

# Set up logging
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()


#  ChromeDriver Executable Path
service = Service(executable_path=r"G:\\kelectron\\NDA\\chrome-win64\\")
chrom_opt = ChormOptions()

def check_internet():
    """
        Check internet connectivity by sending a request to Google.
        Retries every 5 seconds until a successful connection is established.
    """
    while True:
        try:
            response = requests.head('https://www.google.com/', timeout=2)
            if response.status_code == 200:
                return True
        except requests.RequestException:
            print("loading")
        time.sleep(5)


def load_messages(file_name):
    if not os.path.exists(file_name):
        with open(f'{file_name}', 'w+'):
            logger.info(f"File {file_name} created.")
            print(f"file {file_name} created")
    try:
        with open(file_name, 'r') as file:
            messages = file.readlines()
            return [message.strip() for message in messages if message.strip()]
    except FileNotFoundError:
        logger.error(f"File not found / Add Some Massages In File: {file_name}")
        print("Add Some Messages In File")
        return []


def check_excel(sent_msg):
    """
    Check and update the Excel file with sent messages.

    """
    if os.path.exists(excel_file):
        existing_df = pd.read_excel(excel_file)
        new_df = pd.DataFrame(sent_msg)
        updated_df = pd.concat([existing_df, new_df], ignore_index=False)
        # print(updated_df)
    else:
        updated_df = pd.DataFrame(sent_msg, columns=['Profile Name', 'Message'])

    updated_df.to_excel(excel_file, index=False)


def login_manually(driver):
    """Handle manual login and save cookies."""

    driver.get('https://www.linkedin.com/login')
    logger.info("Please log in manually.")
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.feed-identity-module'))
    )
    with open(cookies_file, "wb") as file:
        pickle.dump(driver.get_cookies(), file)

    logger.info("Cookies saved.")


def send_message(driver, message):
    try:
        clear_textbox = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.msg-form__contenteditable"))
        )
        time.sleep(3)

        actions = ActionChains(driver)
        actions.move_to_element(clear_textbox).click().perform()
        clear_textbox.send_keys(Keys.CONTROL + "a")  # Select all text
        clear_textbox.send_keys(Keys.BACKSPACE)
        clear_textbox.send_keys(message)

        WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.msg-form__send-button"))
        ).click()
        time.sleep(1)
        logger.info("Message sent successfully.")
    except Exception as e:
        logger.error(f"Error sending message: {e}")


def main_linkedin(driver):
    """
       Navigate to LinkedIn connections and send messages to connections.
    """
    try:
        driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.mn-connection-card__details")))

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        connections = soup.find(attrs={'class': "scaffold-finite-scroll__content"})
        connection_cards = connections.find_all('li', class_='mn-connection-card')

        sent_msg_conn = []
        sent_msg = []
        for card in connection_cards:
            time.sleep(2)
            name_tag = card.find('span', class_='mn-connection-card__name')
            name = name_tag.get_text(strip=True)

            # if sent_msg_conn and name == sent_msg_conn[0]:
            #     continue
            # sent_msg_conn.append(name)

            if os.path.exists('profile_records.xlsx'):
                df = pd.read_excel('profile_records.xlsx')
                if name in df['Profile Name'].values:
                    print(f" All-ready sent message {name} ")
                    continue

            logger.info(f"Processing connection: {name}")
            profile_link = card.find('a', class_='mn-connection-card__picture')['href']
            driver.execute_script("window.open(arguments[0], '_blank');", f'https://www.linkedin.com{profile_link}')
            driver.switch_to.window(driver.window_handles[-1])
            # driver.implicitly_wait(3)
            time.sleep(3)
            # driver.execute_script("window.scrollTo(0, Y)")
            driver.execute_script("window.scrollTo(0,300);")
            time.sleep(3)
            try:
                driver.implicitly_wait(2)
                msg_button = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         "//button[contains(@class, 'artdeco-button') and contains(@class, 'artdeco-button--2') and contains(@class, 'artdeco-button--primary') and contains(@class, 'ember-view') and contains(@class, 'pvs-profile-actions__action')]")
                    )
                )
                driver.implicitly_wait(3)
                msg_button.click()
                driver.implicitly_wait(3)

                message = load_messages(message_file)
                random_msg = random.choice(message)
                send_message(driver, random_msg)

                driver.implicitly_wait(3)
                close_button = driver.find_element(By.XPATH,
                                                   "//button[contains(@class, 'msg-overlay-bubble-header__control') and contains(@class, 'artdeco-button--circle') and contains(@class, 'artdeco-button--muted') and contains(@class, 'artdeco-button--1') and contains(@class, 'artdeco-button--tertiary') and contains(@class, 'ember-view') and contains(., 'Close your conversation')]")
                close_button.click()
                time.sleep(2)

                context = {
                    'Profile Name': name,
                    'Message': random_msg
                }
                # print(context)4
                # sent_msg.append(context)
                time.sleep(1)
                data = [context]
                check_excel(data)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(2)

            except Exception as e:
                time.sleep(1)
                # check_excel(sent_msg)
                logger.error(f"Error interacting with connection: {e}")

        # check_excel(sent_msg)
        # df = pd.DataFrame(sent_msg)
        # df.to_excel("profile_records.xlsx", index=False)

    except Exception as e:
        logger.error(f"Error during cookie login: {e}")


def login_with_cookies(driver):
    """
    Load cookies and attempt to log in automatically. If no cookies are found, perform manual login.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
    """
    driver.get('https://www.linkedin.com')
    if os.path.exists(cookies_file):
        with open(cookies_file, "rb") as file:
            cookies = pickle.load(file)
            for cookie in cookies:
                driver.add_cookie(cookie)

        time.sleep(2)
        driver.refresh()

        main_linkedin(driver)
    else:
        login_manually(driver)
        main_linkedin(driver)


# Main script execution loop
while 1:
    driver = webdriver.Chrome()
    is_internet = check_internet()
    if is_internet:
        try:

            # run_duration = 60 * 60
            # while (time.time() - start_time) < run_duration:
                login_with_cookies(driver)
                # time.sleep(2)
            # time.sleep(300)
        except:
            logger.error("lost internet connection")
            check_internet()

        finally:
            driver.quit()
            logger.info("Driver closed.")

