from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time

review_data = []

try:
    url = input("Masukkan URL Google Maps: ").strip()
except:
    print("❌ Error: MASUKKAN URL YANG BENER LAH!")

try: 
    max_scroll = int(input("Masukkan jumlah maksimal scroll: ").strip())
except:
    print("❌ Error: MASUKKAN ANGKA YANG VALID!")

try:
    file_name = input("Masukkan nama file Excel: ").strip()
except:
    print("❌ Error: MASUKKAN NAMA FILE YANG BENER LAH!")

try:
    # Setup options untuk Chrome
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-notifications')
    options.add_argument('--lang=id')
    driver = webdriver.Chrome(options=options)

    # Buka URL
    driver.get(url)
    
    # Tunggu dan terima cookie jika muncul
    try:
        cookie_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Terima semua') or contains(., 'Accept all')]"))
        )
        cookie_button.click()
    except TimeoutException:
        pass
    
    # Klik tombol ulasan
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Ulasan') or contains(@aria-label, 'Reviews')]"))
    ).click()
    
    # Tunggu panel ulasan muncul
    time.sleep(3)
    
    # Cari div scrollable dengan class yang lebih umum
    scrollable_div = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'm6QErb') and contains(@class, 'DxyBCb')]"))
    )
    
    # Scroll untuk memuat ulasan
    for _ in range(max_scroll):
        driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', scrollable_div)
        time.sleep(2)
    
    # Ambil semua elemen ulasan
    reviews = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'jftiEf')]"))
    )
    
    # Ekstrak data
    for review in reviews:
        try:
            name = review.find_element(By.XPATH, ".//div[contains(@class, 'd4r55')]").text
            comment = review.find_element(By.XPATH, ".//span[@class='wiI7pd']").text
            total_stars = review.find_element(By.XPATH, ".//span[@class='kvMYJc']").get_attribute('aria-label')
            time_comment = review.find_element(By.XPATH, ".//span[@class='rsqaWe']").text
            review_data.append({"Nama": name,"Rating": total_stars, "Waktu": time_comment, "Ulasan": comment })
        except NoSuchElementException:
            continue

except Exception as e:
    print(f"Terjadi error: {str(e)}")

finally:
    driver.quit()

# Konversi ke DataFrame pandas
df = pd.DataFrame(review_data)

# Simpan ke file Excel
excel_filename = f"{file_name}.xlsx"
df.to_excel(excel_filename, index=False, engine='openpyxl')

print(f"Data ulasan berhasil disimpan ke {excel_filename}")
print(f"Total ulasan yang berhasil diambil: {len(df)}")