import csv
import time
import random
import logging
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from bs4 import BeautifulSoup
import pandas as pd

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_driver():
    try:
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=options)
        return driver
    except Exception as e:
        logger.error(f"Ошибка при настройке драйвера: {e}")
        raise

def get_description(driver, url):
    try:
        driver.get(url)
        time.sleep(random.uniform(5, 10))
        
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[itemprop='description']"))
        )
        
        description_element = driver.find_element(By.CSS_SELECTOR, "[itemprop='description']")
        return description_element.text.strip()
    except Exception as e:
        logger.warning(f"Не удалось получить описание для {url}: {e}")
        return "Описание не найдено"

def scrape_avito_garages(pages=1):
    base_url = 'https://www.avito.ru/kirovskaya_oblast_kirov/garazhi_i_mashinomesta/sdam/garazh-ASgBAgICAkSYA~YQ5gj2Wg?cd=1'
    driver = setup_driver()
    data = []
    
    try:
        for page in range(1, pages + 1):
            page_url = f"{base_url}&p={page}"
            logger.info(f"Обрабатываем страницу {page}: {page_url}")
            
            driver.get(page_url)
            
            # Имитация человеческого поведения
            time.sleep(random.uniform(10, 15))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/4);")
            time.sleep(random.uniform(3, 5))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(random.uniform(3, 5))
            driver.execute_script("window.scrollTo(0, 3*document.body.scrollHeight/4);")
            time.sleep(random.uniform(3, 5))
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(random.uniform(5, 10))
            
            try:
                WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-marker='item']"))
                )
            except Exception as e:
                logger.warning(f"Элементы не найдены на странице {page}. Возможно, сработала защита от ботов: {e}")
                continue
            
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            items = soup.select("[data-marker='item']")
            
            logger.info(f"Найдено элементов на странице {page}: {len(items)}")
            
            for item in items:
                try:
                    title_element = item.select_one("[itemprop='name']")
                    price_element = item.select_one("[itemprop='price']")
                    link_element = item.select_one("[data-marker='item-title']")
                    location_element = item.select_one("[class*='geo-geo']") or item.select_one("[class*='location-'] span")
                    date_element = item.select_one("[data-marker='item-date']")
                    
                    if all([title_element, price_element, link_element, date_element]):
                        title = title_element.text.strip()
                        price = price_element.get('content')
                        if price is None:
                            logger.warning("Атрибут 'content' отсутствует в элементе цены")
                            price = price_element.text.strip()
                        link = 'https://www.avito.ru' + link_element.get('href', 'N/A')
                        location = location_element.text.strip() if location_element else "Неизвестно"
                        date = date_element.text.strip()
                        
                        # Получаем описание, открывая страницу объявления
                        description = get_description(driver, link)
                        
                        data.append({
                            'title': title,
                            'price': price,
                            'link': link,
                            'location': location,
                            'date': date,
                            'description': description
                        })
                    else:
                        missing_elements = [name for name, element in zip(
                            ['title', 'price', 'link', 'date'],
                            [title_element, price_element, link_element, date_element]
                        ) if element is None]
                        logger.warning(f"Отсутствуют элементы: {', '.join(missing_elements)} в одном из объявлений")
                        logger.debug(f"HTML объявления: {item}")
                except Exception as e:
                    logger.warning(f"Не удалось извлечь данные из элемента: {e}")
                    logger.debug(f"HTML объявления: {item}")
            
            logger.info(f"Обработана страница {page}, извлечено {len(data)} объявлений")
            
            if page < pages:
                # Случайная задержка между страницами
                time.sleep(random.uniform(45, 90))
    
    except Exception as e:
        logger.error(f"Произошла ошибка при скрапинге: {e}")
    
    finally:
        driver.quit()
    
    return data

def save_to_excel(data, filename):
    if not data:
        logger.error("Нет данных для сохранения.")
        return
    
    df = pd.DataFrame(data)
    
    # Проверяем, что столбец 'price' существует перед выполнением операций
    if 'price' not in df.columns:
        logger.error("Столбец 'price' отсутствует в данных.")
        return
    
    # Базовая статистика
    stats = pd.DataFrame({
        'Метрика': ['Общее количество объявлений', 'Средняя цена', 'Минимальная цена', 'Максимальная цена', 'Средняя длина описания'],
        'Значение': [
            len(df),
            df['price'].astype(float).mean(),
            df['price'].astype(float).min(),
            df['price'].astype(float).max(),
            df['description'].str.len().mean()
        ]
    })
    
    # Создаем Excel-writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Данные', index=False)
        stats.to_excel(writer, sheet_name='Статистика', index=False)
        
        # Автоматическая настройка ширины столбцов
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column[0].column_letter].width = adjusted_width

    logger.info(f"Данные и статистика сохранены в файл: {filename}")

if __name__ == "__main__":
    try:
        pages_to_scrape = 1  # Укажите желаемое количество страниц
        data = scrape_avito_garages(pages=pages_to_scrape)
        save_to_excel(data, 'kirov_garages_rent.xlsx')
        logger.info(f"Скрапинг завершен. Собрано {len(data)} объявлений об аренде гаражей в Кирове.")
    except Exception as e:
        logger.error("Произошла ошибка при выполнении скрапера:", exc_info=True)