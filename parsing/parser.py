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
        self.selected_data = []  # Список выбранных данных для парсинга
        self.inn_column = 1  # По умолчанию используется первый столбец

    def process_data(self, inn_file, save_file, source):
        try:
            # Загружаем данные из файла с ИНН
            df = pd.read_excel(inn_file)

            # Определяем имя столбца с ИНН
            inn_column_index = int(self.inn_column) - 1  # Преобразуем в индекс (начиная с 0)
            inn_column_name = df.columns[inn_column_index]  # Получаем имя столбца по индексу

            # Загружаем ИНН из указанного столбца
            df[inn_column_name] = df[inn_column_name].astype(str).str.strip()
            inn_list = df[inn_column_name].tolist()

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
        wait = WebDriverWait(driver, 3)

        # Пример для List-org
        if source == "https://www.list-org.com/":
            # Вводим ИНН в поисковую строку
            search_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/form/div[1]/input')))
            search_input.clear()
            search_input.send_keys(inn)

            # Нажимаем кнопку поиска
            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/form/div[1]/button')))
            search_button.click()
            self.wait_for_captcha(driver)

            # Проверяем, есть ли сообщение "Найдено 0 организаций"
            try:
                result_message = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[1]/p')))
                if "Найдено 0 организаций" in result_message.text:
                    return False  # Возвращаем False, чтобы пропустить этот ИНН
            except TimeoutException:
                pass  # Сообщение не найдено, продолжаем обработку

            # Если сообщение не найдено, переходим к результату
            try:
                result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[1]/div/p[1]/label/a')))
            except:
                result_link = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[1]/div/p/a')))
            
            result_link.click()
            self.wait_for_captcha(driver)
            return True  # Возвращаем True, если информация найдена

        else:
            # Добавьте обработку для других источников
            print(f"Источник {source} пока не поддерживается.")
            return False

    def get_contact_info(self, driver):
        result = {}

        # Получение полного юридического наименования
        if 'Полное юридическое наименование' in self.selected_data:
            try:
                fullname_element = driver.find_element(By.XPATH, '//a[contains(@href, "/search?type=name")]')
                result["Полное юридическое наименование"] = fullname_element.text
            except NoSuchElementException:
                result["Полное юридическое наименование"] = ""
        
        # Получение руководителя
        if 'Руководитель' in self.selected_data:
            dirnames = set()
            for i in range(3, 6):
                try:
                    dirname_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/table/tbody/tr[2]/td[2]'
                    dirname_element = driver.find_element(By.XPATH, dirname_xpath)
                    dirname = dirname_element.text.strip()
                    dirnames.add(dirname)
                except NoSuchElementException:
                    continue
            result['Руководитель'] = ', '.join(dirnames) if dirnames else ""
        
        # Получение уставного капитала
        if 'Уставной капитал' in self.selected_data:
            capital_text = ""  # Инициализация переменной
            for i in range(3, 6):
                try:
                    check_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/table/tbody/tr[4]/td[1]/i'
                    check_element = driver.find_element(By.XPATH, check_xpath)
                    check = check_element.text.strip()
                    if check == 'Уставной капитал:':
                        capital_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/table/tbody/tr[4]/td[2]'
                        capital_element = driver.find_element(By.XPATH, capital_xpath)
                        capital_text = capital_element.text
                        break  # Выходим из цикла, если нашли значение
                except NoSuchElementException:
                    continue
            result['Уставной капитал'] = capital_text

        # Получение численности персонала
        if 'Численность персонала' in self.selected_data:
            staff_text = ""  # Инициализация переменной
            for i in range(3, 6):
                try:
                    check_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/table/tbody/tr[5]/td[1]/i'
                    check_element = driver.find_element(By.XPATH, check_xpath)
                    check = check_element.text.strip()
                    if check == 'Численность персонала:':
                        staff_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/table/tbody/tr[5]/td[2]'
                        staff_element = driver.find_element(By.XPATH, staff_xpath)
                        staff_text = staff_element.text
                        break  # Выходим из цикла, если нашли значение
                except NoSuchElementException:
                    continue
            result['Численность персонала'] = staff_text

        # Получение статуса
        if 'Статус' in self.selected_data:
            try:
                status_element = driver.find_element(By.XPATH, '//td[contains(@class, "status")]')
                result["Статус"] = status_element.text
            except NoSuchElementException:
                result["Статус"] = ""
        
        # Получение адреса
        if 'Адрес' in self.selected_data:
            index_text = ""  # Инициализация переменной
            address_text = ""  # Инициализация переменной
            for i in range(5, 8):
                try:
                    check_index_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[1]/i'
                    check_index_element = driver.find_element(By.XPATH, check_index_xpath)
                    check_index = check_index_element.text.strip()
                    if check_index == 'Индекс:':
                        index_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[1]'
                        index_element = driver.find_element(By.XPATH, index_xpath)
                        index_text = index_element.text.strip()
                        if index_text.startswith("Индекс: "):
                            index_text = index_text.replace("Индекс: ", "").strip()
                except NoSuchElementException:
                    continue
                try:
                    check_address_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[2]/i'
                    check_address_element = driver.find_element(By.XPATH, check_address_xpath)
                    check_address = check_address_element.text.strip()
                    if check_address == 'Адрес:':
                        address_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[2]/span'
                        address_element = driver.find_element(By.XPATH, address_xpath)
                        address_text = address_element.text
                except NoSuchElementException:
                    continue
            result['Адрес'] = f"{index_text}, {address_text}" if index_text or address_text else ""

        # Получение юридического адреса
        if 'Юридический адрес' in self.selected_data:
            legadd_text = ""  # Инициализация переменной
            for i in range(5, 8):
                for j in range(1,5):
                    try:
                        check_legadd_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[{j}]/i'
                        check_legadd_element = driver.find_element(By.XPATH, check_legadd_xpath)
                        check_legadd = check_legadd_element.text.strip()
                        if check_legadd == 'Юридический адрес:':
                            legadd_xpath = f'/html/body/div[1]/div[2]/div[1]/div[{i}]/div/div[1]/div/p[{j}]/span'
                            legadd_element = driver.find_element(By.XPATH, legadd_xpath)
                            legadd_text = legadd_element.text.strip()
                            if legadd_text.startswith("Юридический адрес: "):
                                legadd_text = legadd_text.replace("Юридический адрес: ", "").strip()
                    except NoSuchElementException:
                        continue
            result['Юридический адрес'] = legadd_text


        # Получение телефонов
        if "Телефон" in self.selected_data:
            phone_numbers = set()
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
            result["Телефон"] = ', '.join(phone_numbers) if phone_numbers else ""

        # Получение email
        if "E-mail" in self.selected_data:
            try:
                email_element = driver.find_element(By.XPATH, '//a[contains(@href, "mailto:")]')
                result["E-mail"] = email_element.text if '@' in email_element.text else ""
            except NoSuchElementException:
                result["E-mail"] = ""

        # Получение сайта
        if "Сайт" in self.selected_data:
            sites = set()
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
            result["Сайт"] = ', '.join(sites) if sites else ""

        return result

    def process_inns(self, inn_list, source):
        results = []
        seen_results = set()
        
        driver = self.setup_driver()

        for inn in inn_list:
            if not self.is_valid_inn(inn):
                print(f"ИНН {inn} некорректен. Добавление пустой строки...")
                results.append({'ИНН': inn, **{key: "" for key in self.selected_data}})
                continue
            
            try:
                # Выполняем поиск ИНН
                info_found = self.search_inn(driver, inn, source)
                if not info_found:
                    # Если информация не найдена, добавляем пустую строку
                    print(f"По ИНН {inn} информация не найдена. Добавление пустой строки...")
                    results.append({'ИНН': inn, **{key: "" for key in self.selected_data}})
                    continue

                # Если информация найдена, извлекаем её
                contact_info = self.get_contact_info(driver)
                if not contact_info:
                    print(f"По ИНН {inn} информация не найдена. Добавление пустой строки...")
                    results.append({'ИНН': inn, **{key: "" for key in self.selected_data}})
                    continue

                # Формируем результат
                result = {'ИНН': inn, **contact_info}
                if tuple(result.items()) not in seen_results:
                    seen_results.add(tuple(result.items()))
                    results.append(result)

            except Exception as e:
                print(f"Ошибка при обработке ИНН {inn}: {e}. Добавление пустой строки...")
                results.append({'ИНН': inn, **{key: "" for key in self.selected_data}})

            # Возврат на главную страницу для нового поиска
            driver.get(source)

        driver.quit()  # Закрываем драйвер после завершения работы
        return results

    def save_results_to_excel(self, results, save_file):
        results_df = pd.DataFrame(results)
        results_df.to_excel(save_file, index=False)