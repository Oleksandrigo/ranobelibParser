import json
import os
from re import findall
from time import sleep

from docx import Document
from docx.shared import Pt, RGBColor
from selenium.common import exceptions as sel_exeptions
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager
import undetected_chromedriver as u


def acp_api_send_request(browser, message_type, data=None) -> None:
    if data is None:
        data = {}
    message = {
        'receiver': 'antiCaptchaPlugin',
        'type': message_type,
        **data
    }
    browser.execute_script("""return window.postMessage({});""".format(json.dumps(message)))


def create_browser(anti_captcha_api_key: str) -> WebDriver:
    options = ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("ignore-certificate-errors")
    options.add_argument("--no-sandbox")
    options.add_extension(os.path.join(os.getcwd(), 'anticaptcha-plugin_v0.61.zip'))
    service = Service(ChromeDriverManager().install())
    browser = Chrome(options=options, service=service)
    browser.delete_all_cookies()
    browser.get('https://antcpt.com/blank.html')
    print("Сейчас произошла установка вашего ключа в расширение. Все в порядке =)")
    acp_api_send_request(
        browser,
        'setOptions',
        dict(options={'antiCaptchaApiKey': anti_captcha_api_key})
    )
    return browser


def find_max_values(browser: WebDriver):
    # Сейчас мы на странице первой главы
    # Найти кнопку с оглавлением
    # Кнопка оглавления
    try:
        but = browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[3]/div/div[3]/div[2]")
    except sel_exeptions.NoSuchElementException:
        try:
            but = browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[3]/div/div[2]/div")
        except:
            return "FAIL|find_max_values second_try"
    except:
        return "FAIL|find_max_values first_try"
    print("Открываем список глав дла парсинга...")
    but.click()
    sleep(3)
    div = browser.find_element(by=By.XPATH, value="/html/body/div[6]/div/div/div/div/div[2]")
    print("В зависимости от количества глав, сейчас может слегка зависнуть. Ничего не трогайте!")
    dict_all_chapters = {}
    for i in div.find_elements(by=By.TAG_NAME, value="a"):
        if type(dict_all_chapters.get(i.text.split()[1])) is list:
            dict_all_chapters[i.text.split()[1]].append(i.text.split()[3])
            continue
        dict_all_chapters[i.text.split()[1]] = [i.text.split()[3]]
    return dict_all_chapters


BASE_URL = "https://ranobelib.me"
VOLUME_URL = "/v"
CHAPTER_URL = "/c"


def parse_and_save():
    title_url = ""
    input_url = input("Введите ссылку на ранобе в ranobelib.me для парсинга.\n")
    if BASE_URL in input_url:
        s = input_url.split("ranobelib.me")[1]
        if len(s.split("/")) > 2:
            title_url = "/" + s.split("/")[1]
        if "?" in s:
            title_url = s.split("?")[0].strip()

    url = BASE_URL + title_url
    print("Окно хрома НЕ ТРОГАТЬ! НЕ ЗАКРЫВАТЬ! НЕ МЕНЯТЬ ЕГО РАЗМЕР! Можно только спрятать в панель задач! ЭТО ВАЖНО!")
    sleep(3)
    print("Введите 32 значный ключ, который вы получили после регистрации на сайте")
    api_key_anti = input("https://anti-captcha.com/clients/settings/apisetup\n")
    if not len(api_key_anti) == 32:
        print("Длина ключа не равна 32.")
        return
    if not api_key_anti.isalnum():
        print("В ключе присутствуют спецсимволы. Это плохо!")
        return

    browser = create_browser(api_key_anti)
    print("Запущен браузер. (chromedriver)")
    browser.get(url)
    print("Открывается ссылка, введенная ползьователем")
    sleep(2)
    title = browser.find_element(by=By.XPATH, value="/html/body/div[3]/div/div/div/div[2]/div[1]/div[1]/div[1]").text
    print(title)
    # browser.find_element(by=By.XPATH, value="/html/body/div[3]/div/div/div/div[1]/div[2]/a").click()
    print("Переходим на страницу первой главы")
    browser.find_element(by=By.LINK_TEXT, value="Начать читать").click()
    sleep(3)
    print("Пытаемся получить кол-во глав и томов")
    all_chapters = find_max_values(browser)
    if type(all_chapters) is str:
        if "FAIL" in all_chapters:
            print("Во время выполнения программы произошёл сильный баг. Отпишите разрабочику на гитхаб.")
            print(all_chapters)
        return
    print("Получилось!")
    browser.set_page_load_timeout(3)
    print("Начинаем парсинг всех глав!")
    doc_file = Document()
    title_head = doc_file.add_heading(title, 0)
    title_head.style.font.color.rgb = RGBColor.from_string("000000")
    volumes = list(all_chapters.keys())
    volumes.sort(key=float)
    for volume in volumes:
        chapters: list = all_chapters[volume]
        chapters.sort(key=float)
        for chapt in chapters:
            try:
                new_url = url + VOLUME_URL + str(volume) + CHAPTER_URL + str(chapt)
                if browser.current_url != new_url:
                    browser.get(new_url)
            except sel_exeptions.TimeoutException:
                pass
            finally:
                name_chapt = browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[3]/div/a/div[3]")
                head = doc_file.add_heading(name_chapt.text)
                head.style.font.size = Pt(14)
                head.style.font.color.rgb = RGBColor.from_string('000000')

                text = browser.find_element(by=By.XPATH, value="/html/body/div[1]/div[4]").text
                text = text.replace("»", "\"")
                text = text.replace("«", "\"")

                # Создаем пробелы после . и ,
                for i in findall(r"\.\S", text):
                    text = text.replace(i, i[0] + " " + i[-1])
                for i in findall(r",\S", text):
                    text = text.replace(i, i[0] + " " + i[-1])
                # Добавляем параграфы
                for i in text.split("\n"):
                    p = doc_file.add_paragraph(i)
                    p.style.font.name = 'Calibri'

                os.system('cls||clear')
                print(f"Том {volume} / {max(volumes, key=float)}")
                print(f"Главы {chapt} / {max(chapters, key=float)}")

    browser.quit()
    while True:
        try:
            doc_file.save(f"{title}.docx")
            print("ГОТОВО!")
            break
        except PermissionError:
            print(f"Close the {title}.docx!!")
            sleep(1)
        except:
            print("Что-то странное произошло во время сохранения файла...")
            sleep(1)


if __name__ == "__main__":
    parse_and_save()
