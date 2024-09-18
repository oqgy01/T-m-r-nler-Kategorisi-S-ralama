import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from io import BytesIO
import numpy as np
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from tqdm import tqdm
import warnings
from colorama import init, Fore, Style
import threading
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import chromedriver_autoinstaller
from concurrent.futures import ThreadPoolExecutor
import subprocess
from selenium.common.exceptions import TimeoutException, WebDriverException
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from datetime import datetime
from datetime import datetime, timedelta
import shutil
import os
from openpyxl import load_workbook
from openpyxl import Workbook
from pathlib import Path
import re
import http.client
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
warnings.filterwarnings("ignore")
pd.options.mode.chained_assignment = None


print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print(Fore.RED + "<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")
print(" ")


#region Ürün Listesi İndirme ve Gereksiz Sütunları Silme

def get_excel_data(url):
    response = requests.get(url)

    if response.status_code == 200:
        # Excel dosyasını oku
        df = pd.read_excel(BytesIO(response.content))
        return df
    else:
        return None

url1 = "https://task.haydigiy.com/FaprikaXls/BIHXQC/1/"
url2 = "https://task.haydigiy.com/FaprikaXls/BIHXQC/2/"
url3 = "https://task.haydigiy.com/FaprikaXls/BIHXQC/3/"

data1 = get_excel_data(url1)
data2 = get_excel_data(url2)
data3 = get_excel_data(url3)

# Üç veriyi birleştir
if data1 is not None and data2 is not None and data3 is not None:
    merged_data = pd.concat([data1, data2, data3], ignore_index=True)

    # İlk olarak Excel dosyasını kaydedelim
    original_file = "Stabil Ürün Listesi.xlsx"
    merged_data.to_excel(original_file, index=False)

    # Stabil Ürün Listesi dosyasındaki gereksiz sütunları silelim
    columns_to_keep = ["UrunAdi", "AramaTerimleri"]
    merged_data = merged_data[columns_to_keep]

    # Gereksiz sütunları sildikten sonra dosyayı yeniden kaydet
    merged_data.to_excel(original_file, index=False)

#endregion

#region XML'den ID'leri Alma

# XML URL'si
xml_url = "https://task.haydigiy.com/FaprikaXml/S4PP8G/1/"

# XML'den Ürün Bilgilerini Çekme ve Temizleme
response = requests.get(xml_url)
xml_data = response.text
soup = BeautifulSoup(xml_data, 'xml')

product_data = []
for item in soup.find_all('item'):
    title = item.find('title').text
    # ' - H' ile başlayan tüm kısımları kaldırmak için düzenli ifade kullanıyoruz
    title_cleaned = re.sub(r' - H.*', '', title)
    
    product_id = item.find('g:id').text if item.find('g:id') else None
    product_data.append({'UrunAdi': title_cleaned, 'ID': product_id})

df_xml = pd.DataFrame(product_data)

# Stabil Ürün Listesi ile Birleştirme
df_calisma_alani = pd.read_excel('Stabil Ürün Listesi.xlsx')
df_merged_stabil = pd.merge(df_calisma_alani, df_xml, how='left', on='UrunAdi')
df_merged_stabil.to_excel('Stabil Ürün Listesi.xlsx', index=False)


#endregion

#region AramaTerimleri Tarihleri Tespit Edip Çıkarma ve Güne Çevirme




# Exceli Oku
df_calisma_alani = pd.read_excel('Stabil Ürün Listesi.xlsx')

# Tarihleri çıkar
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'

# "AramaTerimleri" sütunundaki tarihleri temizle
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)

# Exceli Kaydet
with pd.ExcelWriter('Stabil Ürün Listesi.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')








# Ekceli Okuma
df = pd.read_excel('Stabil Ürün Listesi.xlsx')

# AramaTerimleri Sütununda Tarih Olanları İşleme Alma
def calculate_days_to_today(row):
    arama_terimi = row['AramaTerimleri']

    # Eğer hücre boşsa veya tarih içermiyorsa 0 döndür
    if pd.isna(arama_terimi) or not any(char.isdigit() for char in str(arama_terimi)):
        return 0

    # Tarihi çıkartma
    tarih = datetime.strptime(arama_terimi.split(';')[0], '%d.%m.%Y')
    
    # Bugünkü tarihten uzaklık hesapla
    bugun = datetime.today()
    uzaklik = (bugun - tarih).days

    return uzaklik

# "AramaTerimleri" Sütununu Güncelleme
df['AramaTerimleri'] = df.apply(calculate_days_to_today, axis=1)

# Exceli Kaydet
df.to_excel('Stabil Ürün Listesi.xlsx', index=False)





# Excel dosyasını oku
dosya_adi = 'Stabil Ürün Listesi.xlsx'
df = pd.read_excel(dosya_adi)

# "AramaTerimleri" sütununda 0 olan satırları filtrele
df_temiz = df[df['AramaTerimleri'] != 0]

# Temizlenmiş veriyi aynı dosyaya ya da yeni bir dosyaya kaydet
df_temiz.to_excel('Stabil Ürün Listesi.xlsx', index=False)




#endregion

#region Geri Kalan Düzenlemeler

# Excel dosyasını oku
file_path = "Stabil Ürün Listesi.xlsx"
df = pd.read_excel(file_path)

df = df.sort_values(by="AramaTerimleri", ascending=True)


# Yenilenen satırları kaldıralım, tekilleştirelim
df = df.drop_duplicates()



# "Kategori ID" adında yeni bir sütun oluşturalım ve 109 ile dolduralım
df["Kategori ID"] = 109

# Dolu satırları say
dolu_satir_sayisi = df.notna().all(axis=1).sum()

# Numara sütununu oluştur ve sadece dolu satırlara numaraları ata
df["Numara"] = pd.Series(range(-dolu_satir_sayisi, 0), index=df[df.notna().all(axis=1)].index)

# Son olarak dosyayı yeniden kaydedelim
output_file = "Stabil Ürün Listesi.xlsx"
df.to_excel(output_file, index=False)

#endregion

#region Kategorileri Sıralama


# Global token değişkeni
_auth_token = None

# Token alma fonksiyonu
def get_auth_token():
    global _auth_token
    if _auth_token is None:  
        login_url = "https://siparis.haydigiy.com/api/customer/login"
        login_payload = {
            "apiKey": "MypGcaEInEOTzuYQydgDHQ",
            "secretKey": "jRqliBLDPke76YhL_WL5qg",
            "emailOrPhone": "mustafa_kod@haydigiy.com",
            "password": "123456"
        }
        login_headers = {
            "Content-Type": "application/json"
        }

        response = requests.post(login_url, json=login_payload, headers=login_headers)
        if response.status_code == 200:
            _auth_token = response.json().get("data", {}).get("token")
            if not _auth_token:
                raise Exception("TOKEN ALINAMADI")
        else:
            raise Exception(f"GİRİŞ BAŞARISIZ: {response.text}")
    return _auth_token

# API isteğini gönderen fonksiyon
def send_request(row, token):
    conn = http.client.HTTPSConnection("siparis.haydigiy.com")
    
    category_id = row['Kategori ID']
    display_order = row['Numara']
    product_id = str(row['ID']).replace(".0", "")

    payload = json.dumps({
        "CategoryId": int(category_id),  
        "IsFeaturedProduct": False, 
        "DisplayOrder": int(display_order)
    })

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}',
        'Cookie': '.Application.Customer=64684894-1b54-488d-bd59-76b94842df65'
    }

    conn.request("PUT", f"/adminapi/product/product-categories?productId={product_id}", payload, headers)
    res = conn.getresponse()
    data = res.read()
    conn.close()

    return res.status, data.decode('utf-8')

# Token alma işlemi
token = get_auth_token()
df = pd.read_excel("Stabil Ürün Listesi.xlsx")

# İstekleri 3'erli gruplar halinde paralel olarak göndermek için ThreadPoolExecutor kullanımı
with ThreadPoolExecutor(max_workers=8) as executor:
    futures = []
    for index, row in df.iterrows():
        futures.append(executor.submit(send_request, row, token))
    
    for future in tqdm(as_completed(futures), total=len(futures), desc="API İstekleri Gönderiliyor"):
        try:
            status, response = future.result()
            pass
        except Exception as e:
            print(f"İstek hatası: {e}")

#endregion

#region Gereksiz Excel Dosyalarını Silme

# Silinecek dosyaların isimlerini tanımla
dosyalar = [
    "Stabil Ürün Listesi.xlsx"

]

for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)

#endregion

