import os
import pickle
import time

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from msedge.selenium_tools import Edge, EdgeOptions
from json_reader import *

meeting_status = {
    'title': '',
    'joined': False,
}


def load_config():
    with open('config.json', encoding='utf-8') as json_data_file:
        config = json.load(json_data_file)
        return config


def configure_driver():
    chrome_options = EdgeOptions()
    chrome_options.use_chromium = True
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--use-fake-ui-for-media-stream')
    chrome_options.add_experimental_option('prefs', {
        'credentials_enable_service': False,
        'profile.default_content_setting_values.media_stream_mic': 1,
        'profile.default_content_setting_values.media_stream_camera': 1,
        'profile.default_content_setting_values.geolocation': 1,
        'profile.default_content_setting_values.notifications': 1,
        'profile': {
            'password_manager_enabled': False
        }
    })
    chrome_options.add_argument('--no-sandbox')

    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])

    config_file = load_config()
    if 'headless' in config_file and config_file['headless']:
        chrome_options.add_argument('--headless')
        print("Enabled headless mode")

    if 'mute_audio' in config_file and config_file['mute_audio']:
        chrome_options.add_argument("--mute-audio")
    # chrome_options.use_chromium = True
    _driver = Edge(EdgeChromiumDriverManager().install(), options=chrome_options)
    # _driver = webdriver.Edge(EdgeChromiumDriverManager().install())
    return _driver


def wait_for_element(_driver, _element_id, _timeout, _mode='id'):
    try:
        if _mode == 'id':
            element = WebDriverWait(_driver, _timeout).until(
                EC.visibility_of_element_located((By.ID, _element_id)))
            return element
        elif _mode == 'xpath':
            element = WebDriverWait(_driver, _timeout).until(
                EC.visibility_of_element_located((By.XPATH, _element_id)))
            return element
        elif _mode == 'class':
            element = WebDriverWait(_driver, _timeout).until(
                EC.visibility_of_element_located((By.CLASS_NAME, _element_id)))
            return element
        elif _mode == 'name':
            element = WebDriverWait(_driver, _timeout).until(
                EC.visibility_of_element_located((By.NAME, _element_id)))
            return element
        elif _mode == 'css':
            element = WebDriverWait(_driver, _timeout).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, _element_id)))
            return element
    except TimeoutException:
        print("Timed out waiting for element")
        return None


def change_view():
    calendar_button = wait_for_element(driver, 'button[aria-label="Calendar Toolbar"]', 120, 'css')
    if not calendar_button:
        print("Calendar button not found")
        return False
    time.sleep(10)
    calendar_button.click()
    ###############################
    view_button = wait_for_element(driver,
                                   ".ms-CommandBar-secondaryCommand > div > button[class*='__topBarContent']", 30,
                                   'css')
    if not view_button:
        print("View button not found")
        return False
    time.sleep(3)
    view_button.click()
    ###############################
    day_button = wait_for_element(driver,
                                  "li[role='presentation'].ms-ContextualMenu-item>button[aria-posinset='1']",
                                  30, 'css')
    if not day_button:
        print("Day button not found")
        return False
    time.sleep(2)
    day_button.click()
    return True


def get_meetings_list():
    meeting_list = []
    meeting_list_temp = wait_for_element(driver,
                                         "div[class*='calendar-multi-day-renderer__cardHolder']", 30, 'css')
    if meeting_list_temp:
        # meetings = meeting_list_temp.find_elements_by_css_selector("div[class*='renderer__eventCard']")
        meetings = meeting_list_temp.find_elements(By.CSS_SELECTOR, "div[class*='event-card-renderer__eventCard']")
        for _ in meetings:
            meeting = {
                'title': _.get_attribute('title'),
                'id': _.get_attribute('id'),
                'full title': _.get_attribute('aria-label')
            }
            meeting_list.append(meeting)
        # write meeting_list to file
        with open('meetings.json', 'w') as outfile:
            json.dump(meeting_list, outfile)
        return meeting_list


def save_cookies(_driver):
    try:
        cookies = _driver.get_cookies()
        with open('cookies.pkl', 'wb') as file:
            pickle.dump(cookies, file)
    except Exception as e:
        print(e)


def load_cookies(_driver):
    try:
        if 'cookies.pkl' in os.listdir():
            with open('cookies.pkl', 'rb') as file:
                cookies = pickle.load(file)
            for cookie in cookies:
                driver.add_cookie(cookie)
    except Exception as e:
        print(e)


def check_if_join(meeting):
    overdue = datetime.now() > meeting['time_start']
    ended = datetime.now() > meeting['time_end']
    if not ended:
        if overdue:
            delay_offset = 0
            return True, delay_offset
        else:
            delay_offset = meeting['time_start'] - datetime.now()
            delay_offset = int(delay_offset.total_seconds())
            if delay_offset > 300:
                return False, delay_offset
            else:
                return True, delay_offset
    return False, 0


def wait_for_meeting(_driver):
    while True:
        try:
            # meeting_list = get_meetings_list()
            get_meetings_list()
            format_meeting_details()
            meeting_list = get_list_from_json()
            if meeting_list:
                print("[{}]Searching for meetings.".format(datetime.now()))
                for meeting in meeting_list:
                    result = check_if_join(meeting)
                    joinNow, delay = result[0], result[1]
                    if joinNow:
                        join_meeting(_driver, meeting, delay)
            else:
                print("[Error]meeting_list is empty")
        except Exception as e:
            print(e)
        time.sleep(30)


def join_meeting(_driver, _meeting, _delay=0):
    global meeting_status
    if _meeting['title'] == meeting_status['title']:
        return

    meeting_box = wait_for_element(_driver, _meeting['id'], 30)
    timer = datetime.now()
    if not meeting_box:
        print("[Error]Meeting box not found")
        return
    meeting_box.click()
    time.sleep(2)
    RSVP_button = wait_for_element(_driver,
                                   "button[data-tid='calv2-peek-rsvp-button'] > span[class*='ms-Button-flexContainer'] > span[class*='ms-Button-textContainer'] > span[class*='ms-Button-label']",
                                   30, 'css')
    if not RSVP_button:
        edit_button = wait_for_element(_driver, "button[data-tid='calv2-peek-edit-button']", 30, 'css')
        if not edit_button:
            print("[Error]RSVP and edit button not found")
            return
        else:
            RSVP_status = 'Started'
    else:
        RSVP_status = RSVP_button.text
    if RSVP_status == 'Accepted' or 'Started':
        print("[{}]Meeting found: ".format(datetime.now()) + _meeting['title'])
        join_button = wait_for_element(_driver,
                                       "button[data-tid='calv2-peek-join-button']",
                                       30, 'css')
        if not join_button:
            print("[Error]Join button not found")
            return
        join_button.click()
        time.sleep(4)

        iframe = driver.find_element(By.CSS_SELECTOR, "iframe[id*='experience-container']")
        _driver.switch_to.frame(iframe)
        mute_button = wait_for_element(_driver,
                                       "div[data-tid='toggle-mute']",
                                       10, 'css')
        prejoin_button = wait_for_element(_driver,
                                          "button[data-tid='prejoin-join-button']",
                                          10, 'css')
        if not mute_button:
            print("[Error]Mute button not found")
            _driver.switch_to.default_content()
            return
        if not prejoin_button:
            print("[Error]Prejoin button not found")
            _driver.switch_to.default_content()
            return
        if mute_button.get_attribute('data-cid') == 'toggle-mute-true':
            mute_button.click()
            time.sleep(1)

        timer = datetime.now() - timer
        if _delay > timer.total_seconds():
            total_delay = _delay - timer.total_seconds()
            print("[{0}]Waiting for meeting to start.({1}s)".format(datetime.now(), total_delay))
            time.sleep(total_delay)

        print("[{0}]Joining meeting now...".format(datetime.now()))

        prejoin_button.click()
        meeting_status['title'] = _meeting['title']
        wait_for_meeting_end(_driver)
        _driver.switch_to.default_content()
    else:
        return


def wait_for_meeting_end(_driver):
    global meeting_status
    time.sleep(10)
    try:
        people_button = wait_for_element(_driver, "button[aria-label='People']", 10, 'css')
        if not people_button:
            print("[Error]People button not found")
            return
        people_button.click()
        time.sleep(2)

        count = 0
        while True:
            participants = wait_for_element(_driver, "span[id*='roster-title-section-2'] > span", 5, 'css')
            if not participants and count >= 2:
                print("[Error]Participants list not found")
                return
            number_of_participants = participants.text
            number_of_participants = int(number_of_participants.split('(')[1].split(')')[0].strip())
            if number_of_participants <= 1 and count >= 4:
                leave_button = wait_for_element(_driver, "button[id='hangup-button']", 10, 'css')
                if not leave_button:
                    print("[Error]Leave button not found")
                    return
                leave_button.click()
                print(f"[{str(datetime.now())}]Meeting ended: " + meeting_status['title'])
                return
            else:
                print(f"[{str(datetime.now())}]Waiting for meeting to end: " + meeting_status['title'] + f" ({number_of_participants} remaining members)")
                count += 1
                time.sleep(15)
    except Exception as e:
        print(e)


def check_if_meeting_is_terminated(_driver):
    return False


if __name__ == '__main__':
    driver = configure_driver()
    driver.get("https://teams.microsoft.com")
    load_cookies(driver)
    driver.maximize_window()
    # calendar_button = wait_for_element(driver, 'app-bar-ef56c0de-36fc-4ef8-b417-3d82ba9d073c', 60)
    if change_view():
        save_cookies(driver)
        wait_for_meeting(driver)

# driver.implicitly_wait(5)
# driver.find_element_by_id("i0116").send_keys("email")
# driver.find_element_by_id("idSIButton9").click()
# driver.implicitly_wait(5)
# driver.find_element_by_id("i0118").send_keys("password")
# driver.find_element_by_id("idSIButton9").click()
# driver.implici