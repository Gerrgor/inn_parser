import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from tkinter import messagebox

class Parser:
    def __init__(self):
        pass

    def process_data(self, inn_file, save_file, source):
        try:
            # Загружаем данные из файла с ИНН
            df = pd.read_excel(inn_file, dtype={'ИНН\n\n': str})
            inn_list = df['ИНН\n\n'].astype(str).str.strip().tolist()

            # Обрабатываем ИНН
            results = self.process_inns(inn_list, source)

            # Сохраняем результаты в Excel
            self.save_results_to_excel(results, save_file)
            return True

        except Exception as e:
            print(f"Ошибка при обработке данных: {e}")
            return False

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.page_load_strategy = 'eager'
        driver = webdriver.Chrome(options=chrome_options)
        return driver

    def is_valid_inn(self, inn):
        return len(inn) == 10 and inn.isdigit()

    def wait_for_captcha(self, driver):
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'вы не робот')]"))
            )
            messagebox.showinfo("Проверка на робота", "Пожалуйста, пройдите проверку и нажмите ОК.")
        except TimeoutException:
            print("Проверка на робота не обнаружена.")

    def search_inn(self, driver, inn, source):
        driver.get(source)  # Переходим на выбранный сайт
        wait = WebDriverWait(driver, 5)

        # Пример для List-org
        if source == "https://www.list-org.com/":
            search_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/form/div[1]/input')))
            search_input.clear()
            search_input.send_keys(inn)

            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/form/div[1]/button')))
            search_button.click()
            self.wait_for_captcha(driver)

            try:
                result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[1]/div/p[1]/label/a')))
            except:
                result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[1]/div/p/a')))
            
            result_link.click()
            self.wait_for_captcha(driver)
        else:
            # Добавьте обработку для других источников
            print(f"Источник {source} пока не поддерживается.")

    def get_contact_info(self, driver):
        email = None
        phone_numbers = set()
        sites = set()

        # Получение email
        try:
            email_element = driver.find_element(By.XPATH, '//a[contains(@href, "mailto:")]')
            email = email_element.text if '@' in email_element.text else None
        except NoSuchElementException:
            email = None
        
        # Получение сайта
        for i in range(6, 8):
            for k in range(1, 10):
                try:
                    site_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div/div/div/p/a[{k}]'
                    site_element = driver.find_element(By.XPATH, site_xpath)
                    site = site_element.text.strip()
                    check_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div/div/div/p/i'
                    check_element = driver.find_element(By.XPATH, check_xpath)
                    check = check_element.text.strip()
                    if check == 'Сайт:' and site and ('www' in site or 'http' in site or '.ru' in site or '.org' in site or '.com' in site):
                        sites.add(site)
                except NoSuchElementException:
                    continue

        # Получение телефонов
        for i in range(6, 8):
            for j in range(2, 6):
                for k in range(1, 10):
                    try:
                        phone_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div/div/p[{j}]/a[{k}]'
                        phone_element = driver.find_element(By.XPATH, phone_xpath)
                        phone = phone_element.text.strip()
                        check_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div/div/p[{j}]/i'
                        check_element = driver.find_element(By.XPATH, check_xpath)
                        check = check_element.text.strip()
                        if check == 'Телефон:' and phone and ('+7' in phone or '-' in phone):
                            phone_numbers.add(phone)
                    except NoSuchElementException:
                        continue

        return email, ', '.join(sites) if sites else '', ', '.join(phone_numbers) if phone_numbers else ''

    def process_inns(self, inn_list, source):
        results = []
        seen_results = set()
        
        driver = self.setup_driver()

        for inn in inn_list:
            if not self.is_valid_inn(inn):
                print(f"ИНН {inn} некорректен. Добавление пустой строки...")
                results.append({'ИНН': inn, 'E-mail': '', 'Сайт': '', 'Телефон': ''})
                continue
            
            self.search_inn(driver, inn, source)
            email, sites, phones = self.get_contact_info(driver)

            result_key = (inn, email, sites, phones)
            if result_key not in seen_results:
                seen_results.add(result_key)           
                results.append({'ИНН': inn, 'E-mail': email or "", 'Сайт': sites or '', 'Телефон': phones or ""})

            # Возврат на главную страницу для нового поиска
            driver.get(source)

        driver.quit()  # Закрываем драйвер после завершения работы
        return results

    def save_results_to_excel(self, results, save_file):
        results_df = pd.DataFrame(results)
        results_df.to_excel(save_file, index=False)