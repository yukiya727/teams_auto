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
from colorama import init, Fore, Back, Style
from json_reader import *

meeting_status = {
    'title': [],
    'joined': False,
}
init(autoreset=True)


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
        print(
            Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + f"Timed out waiting for element ({_element_id})")
        return None


def change_view():
    enter_web_app = wait_for_element(driver, 'a[class="use-app-lnk"]', 10, 'css')
    if enter_web_app:
        enter_web_app.click()
    calendar_button = wait_for_element(driver, 'button[aria-label="Calendar Toolbar"]', 120, 'css')
    if not calendar_button:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Calendar button not found")
        return False
    time.sleep(10)
    calendar_button.click()
    ###############################
    iframe = driver.find_element(By.CSS_SELECTOR, "iframe[id*='experience-container']")
    driver.switch_to.frame(iframe)

    view_button = wait_for_element(driver, 'div[title="Switch your calendar view"] > div', 30, 'css')
    if not view_button:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "View button not found")
        return False
    time.sleep(3)
    view_button.click()
    ###############################
    day_button = wait_for_element(driver,
                                  "li[aria-label='Day view'] > div",
                                  30, 'css')
    if not day_button:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Day button not found")
        return False
    time.sleep(2)
    day_button.click()
    return True


def get_meetings_list():
    meeting_list = []
    # 0.1.2
    # meeting_list_temp = wait_for_element(driver,
    #                                      "div[class*='calendar-multi-day-renderer__cardHolder']", 30, 'css')
    meeting_list_temp = wait_for_element(driver,
                                         '//*[@id="app"]/div/div/div/div/div[5]/div/div/div[2]/div[2]/div/div[2]/div/div[2]',
                                         30, 'xpath')
    if meeting_list_temp:
        # meetings = meeting_list_temp.find_elements_by_css_selector("div[class*='renderer__eventCard']")

        # 0.1.2
        # meetings = meeting_list_temp.find_elements(By.CSS_SELECTOR, "div[class*='event-card-renderer__eventCard']")
        meetings = meeting_list_temp.find_elements(By.CSS_SELECTOR, "div[class*='fui-Primitive'] > div")
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
    # print('Time start:' + str(meeting['time_start']))
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
            # print('Meeting list: found')
            format_meeting_details()
            # print('Meeting list: formatted')
            meeting_list = get_list_from_json()
            # print('Meeting list: retrieved')
            if meeting_list:
                print(Fore.YELLOW + Style.DIM + "[{}]".format(
                    datetime.now()) + Fore.WHITE + Style.NORMAL + "Searching for meetings.".format(datetime.now()))
                for meeting in meeting_list:
                    if meeting['title'] in meeting_status['title']:
                        continue
                    result = check_if_join(meeting)
                    joinNow, delay = result[0], result[1]
                    print(f'[Debug]{meeting["title"]} Join result: ' + str(joinNow))

                    if joinNow:
                        join_meeting(_driver, meeting, delay)
            else:
                print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "meeting_list is empty")
        except Exception as e:
            print(e)
        time.sleep(30)


def join_meeting(_driver, _meeting, _delay=0):
    global meeting_status
    if _meeting['title'] in meeting_status['title']:
        return

    meeting_box = wait_for_element(_driver, _meeting['id'], 30)
    timer = datetime.now()
    if not meeting_box:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Meeting box not found")
        return
    meeting_box.click()
    time.sleep(2)
    # RSVP_button = wait_for_element(_driver,
    #                                'button[id*="menubutton-trigger"] > span[class*="ui-button__content"]',
    #                                30, 'css')
    # if not RSVP_button:
    #     print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "RSVP or edit button not found")
    #     meeting_status['title'].append(_meeting['title'])
    #     return
    # else:
    #     RSVP_status = RSVP_button.text
    # if RSVP_status != 'Tentative' or 'Declined' or 'RSVP':
    #     print(Fore.GREEN + "[{}]Meeting found: ".format(datetime.now()) + _meeting['title'])
    join_button = wait_for_element(_driver,
                                   'button[data-track-module-name="calendarEventPeekViewMeetingJoinButton"]',
                                   30, 'css')
    if not join_button:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Join button not found")
        return

    RSVP_button = join_button.find_element_by_xpath("./following-sibling::*")
    if not RSVP_button:
        print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "RSVP or edit button not found")
        meeting_status['title'].append(_meeting['title'])
        return
    else:
        RSVP_status = RSVP_button.text

    print("RSVP status: " + RSVP_status)
    if RSVP_status != 'Tentative' or RSVP_status != 'Declined' or RSVP_status != 'RSVP':
        join_button.click()
        time.sleep(10)

        # wait = WebDriverWait(driver, 20)
        # iframe = wait.until(EC.presence_of_element_located((By.XPATH, '//iframe[contains(@id, "experience-container")]')))
        # # iframe = _driver.execute_script("return document.querySelector('iframe[id*=experience-container]')")

        # iframe = _driver.execute_script("return document.getElementsByTagName('iframe')[0];")
        # print(iframe)
        # _driver.switch_to.frame(iframe)
        iframe = _driver.find_element_by_xpath('/html/body/div[5]/div')
        # _driver.switch_to.frame(iframe)
        print(iframe
        )

        # Check if there is another iframe within this iframe
        # try:
        #     nested_iframe = _driver.find_element_by_xpath('//iframe')
        #     _driver.switch_to.frame(nested_iframe)
        # except Exception as e:
        #     print(e)
        #     pass


        # _driver.switch_to.frame(0)
        # print(iframe)

        mute_button = wait_for_element(_driver,
                                       "div[data-tid='toggle-mute']",
                                       10, 'css')
        prejoin_button = wait_for_element(_driver,
                                          "button[data-tid='prejoin-join-button']",
                                          10, 'css')
        if not mute_button:
            print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Mute button not found")
            _driver.switch_to.default_content()
            return
        if not prejoin_button:
            print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Prejoin button not found")
            meeting_status['title'].append(_meeting['title'])
            _driver.switch_to.default_content()
            return
        if mute_button.get_attribute('data-cid') == 'toggle-mute-true':
            mute_button.click()
            time.sleep(1)

        timer = datetime.now() - timer
        if _delay > timer.total_seconds():
            total_delay = _delay - timer.total_seconds()
            print(Fore.YELLOW + Style.DIM + "[{0}]".format(
                datetime.now()) + Fore.WHITE + Style.NORMAL + "Waiting for meeting to start.({}s)".format(total_delay))
            time.sleep(total_delay)

        print(Fore.YELLOW + Style.DIM + "[{0}]".format(
            datetime.now()) + Fore.WHITE + Style.NORMAL + "Joining meeting now...")

        prejoin_button.click()
        meeting_status['title'].append(_meeting['title'])
        wait_for_meeting_end(_driver)
        _driver.switch_to.default_content()
    else:
        print(print("RSVP is not 'Accepted', RSVP status: " + RSVP_status))
        meeting_status['title'].append(_meeting['title'])
        return


def wait_for_meeting_end(_driver):
    global meeting_status
    time.sleep(10)
    try:
        people_button = wait_for_element(_driver, "button[aria-label='People']", 10, 'css')
        if not people_button:
            print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "People button not found")
            return
        people_button.click()
        time.sleep(2)

        count = 0
        while True:
            participants = wait_for_element(_driver, "span[id*='roster-title-section-2'] > span", 5, 'css')
            if not participants and count >= 2:
                print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Participants list not found")
                return
            number_of_participants = participants.text
            number_of_participants = int(number_of_participants.split('(')[1].split(')')[0].strip())
            if number_of_participants <= 1 and count >= 4:
                leave_button = wait_for_element(_driver, "button[data-tid='hangup-main-btn']", 10, 'css')
                if not leave_button:
                    leave_button = wait_for_element(_driver, "button[id='hangup-button']", 10, 'css')
                if not leave_button:
                    print(Fore.YELLOW + Back.BLUE + "[Error]" + Fore.YELLOW + Back.BLACK + "Leave button not found")
                    return
                leave_button.click()
                print(
                    Fore.YELLOW + Style.DIM + f"[{str(datetime.now())}]" + Fore.YELLOW + Style.NORMAL + "Meeting ended: " +
                    meeting_status['title'][-1])
                return
            else:
                print(
                    Fore.YELLOW + Style.DIM + f"[{str(datetime.now())}]" + Fore.WHITE + Style.NORMAL + "Waiting for meeting to end: " +
                    meeting_status['title'][-1] + f" ({number_of_participants} remaining members)")
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
# driver.find_element_by_id("i0118
