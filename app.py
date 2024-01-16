import tkinter as tk
import threading
import os
import re
import sqlite3
from tkinter import messagebox
import cv2
import pytesseract
from PIL import Image, ImageTk
from pytesseract import Output
import json
import pyodbc

# Fonksiyonlar
def load_config(config_file):
    with open(config_file, 'r') as file:
        config = json.load(file)
    return config

def isControlCheck(cursor):
    selectQuery = "SELECT [File] , [Id] , [CompanyId] FROM CompanyDataFiles WHERE IsChecked = ?"
    cursor.execute(selectQuery, True)
    resultsOfPDF = cursor.fetchall()
    return resultsOfPDF
"""
def fileToImages(filePath):
    doc = fitz.open(filePath)
    images = []

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        image = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
        pil_image = Image.frombytes("RGB", [image.width, image.height], image.samples)
        images.append(np.array(pil_image))
    return images[0]
"""
def preprocess_image(image):
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresholded_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    denoised_image = cv2.GaussianBlur(thresholded_image, (5, 5), 0)
    return denoised_image

def processPDF(filePath):
    # PDF işleme fonksiyonu
    int_number_pattern = r'\b\d+\b'
    number_pattern = r'\b\d{1,3}(?:,\d{3})*(?:\.\d{1,3})?\b'
    date_pattern = r'\b\d{2}\.\d{2}\.\d{4}\b'
    date_pattern2 = r'\d{2}-\d{2}-\d{4}'

    dataArray = []
    dateCount = 0
    tüketimVerisi = 0
    tüketimVerisiDG = 0
    date = ''

    image = fileToImages(filePath)
    width = image.shape[1]

    preprocessed_image = preprocess_image(image)
    d = pytesseract.image_to_data(preprocessed_image, output_type=Output.DICT)

    for i, conf in enumerate(d['conf']):
        # Geçerli veri d['text'][i]

        # FATURANIN TARİHİNİN BULUNMASI
        if conf > 60 and re.match(date_pattern, d['text'][i]) and dateCount <= 0:
            date = d['text'][i]
            dateCount += 1
            dataArray.append(date)

        if conf > 60 and re.match(date_pattern2, d['text'][i]) and dateCount <= 0:
            date = d['text'][i]
            dateCount += 1
            dataArray.append(date)
            dataArray.reverse()

        # VERİLERİN BULUNMASI
        if "(M3)" in d['text'][i]:
            (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
            (xM3,yM3,wM3,hM3) = (x,y,width,h)

            for j,conf in enumerate(d['conf']):
                if conf > 60 and (re.match(number_pattern, d['text'][j]) or re.match(int_number_pattern, d['text'][j])):
                    (x, y, w, h) = (d['left'][j], d['top'][j], d['width'][j], d['height'][j])
                    if (yM3 <= y+5 and yM3 >= y-5):
                        tüketimVerisiDG = d['text'][j]
                        dataArray.append(tüketimVerisiDG)

        if conf > 60 and (re.match(number_pattern, d['text'][i]) or re.match(int_number_pattern, d['text'][i])):
            if "(Sm3)" in d['text'][i-1]:
                tüketimVerisiDG += float(d['text'][i].replace(",", "."))
                dataArray.append(tüketimVerisiDG)

        if conf > 60 and (re.match(number_pattern, d['text'][i]) or re.match(int_number_pattern, d['text'][i])) and (d['text'][i - 1] == '(E.T.B)' or d['text'][i - 1] == '(ETB)'):
            tüketimVerisi += float(d['text'][i].replace(",", "."))
            dataArray.append(tüketimVerisi)

        if conf > 60 and (re.match(number_pattern, d['text'][i]) or re.match(int_number_pattern, d['text'][i])) and (d['text'][i-1] == 'Kad' or d['text'][i-1] == 'Kad.'):
            tüketimVerisi += float(d['text'][i].replace(",", "."))
            dataArray.append(tüketimVerisi)

        if conf > 60 and (re.match(number_pattern, d['text'][i]) or re.match(int_number_pattern, d['text'][i])) and d['text'][i-1] == 'TOPLAM':
            tüketimVerisi += float(d['text'][i].replace(",", "."))
            dataArray.append(tüketimVerisi)

    return dataArray

def updateDB(cursor, connectionDB, date, tüketimVerisi, result, musteri, faturaTürü):
    if (date is None or tüketimVerisi is None):
        updateQuery = "UPDATE Fatura SET isControl = ? WHERE PDF = ? AND Musteri = ? AND Tür = ?"
        cursor.execute(updateQuery, (0, result, musteri, faturaTürü))

        updateQuery = "UPDATE Fatura SET KanıtTüketimVerisi = ? WHERE PDF = ? AND Musteri = ? AND Tür = ?"
        cursor.execute(updateQuery, ("0", result, musteri, faturaTürü))

        updateQuery = "UPDATE Fatura SET DurumKontrol = ? WHERE PDF = ? AND Musteri = ? AND Tür = ?"
        cursor.execute(updateQuery, ("Veri Bulunamadı.", result, musteri, faturaTürü))

        connectionDB.commit()
    else:
        month = date[3:5]
        year = date[6:]

        updateQuery = "UPDATE Fatura SET isControl = ? WHERE Ay = ? AND Yıl = ? AND Musteri = ? AND Tür = ?"
        cursor.execute(updateQuery, (0, month, year, musteri, faturaTürü))

        if faturaTürü == "E":
            updateQuery = "UPDATE Fatura SET KanıtTüketimVerisi = ? WHERE Ay = ? AND Yıl = ? AND Musteri = ? AND Tür = ?"
            cursor.execute(updateQuery, (tüketimVerisi, month, year, musteri, faturaTürü))

            updateQuery = "UPDATE Fatura SET DurumKontrol = ? WHERE Ay = ? AND Yıl = ? AND Musteri = ? AND Tür = ?"
            cursor.execute(updateQuery, ("Elektrik Faturası Güncelleme Başarılı.", month, year, musteri, faturaTürü))
        elif faturaTürü == "DG":
            updateQuery = "UPDATE Fatura SET KanıtTüketimVerisi = ? WHERE Ay = ? AND Yıl = ? AND Musteri = ? AND Tür = ?"
            cursor.execute(updateQuery, (tüketimVerisi, month, year, musteri, faturaTürü))

            updateQuery = "UPDATE Fatura SET DurumKontrol = ? WHERE Ay = ? AND Yıl = ? AND Musteri = ? AND Tür = ?"
            cursor.execute(updateQuery, ("Doğalgaz Faturası Güncelleme Başarılı.", month, year, musteri, faturaTürü))

        connectionDB.commit()

class PDFProcessingApp:
    def __init__(self, root, config):
        self.root = root
        self.root.title("ADASO Fatura İşleme Uygulaması")
        self.root.geometry("500x500")
        self.config = config

        # Arkaplan resmini ekleyin
        self.image = Image.open("adaso.jpg")
        self.resized_image = self.image.resize((500, 500))
        self.converted_image = self.resized_image.convert("RGBA")
        self.final_image = self.adjust_opacity(self.converted_image, 0.3)
        self.background_image = ImageTk.PhotoImage(self.final_image)
        self.background_label = tk.Label(root, image=self.background_image)
        self.background_label.place(relwidth=1, relheight=1)

        self.tesseract_cmd = config['tesseract_cmd']
        self.processing = False

        self.total_islem_sayisi = 0
        self.yapilan_islem_sayisi = 0
        self.kalan_islem_sayisi = 0

        # Yeşil başla butonu
        self.basla_button = tk.Button(root, text="Başla", command=self.basla_islem, font=('Times', 14), width=7, height=1)
        self.basla_button.configure(background='green', foreground='white', activebackground='white')

        # Sarı durdur butonu
        self.dur_button = tk.Button(root, text="Durdur", command=self.dur_islem, font=('Times', 14), width=6, height=1)
        self.dur_button.configure(background='orange', foreground='white')

        # Kırmızı kapat butonu
        self.kapat_button = tk.Button(root, text="Kapat", command=self.kapat, font=('Times', 14), width=6, height=1)
        self.kapat_button.configure(background='red', foreground='white')

        self.total_islem_label = tk.Label(root, text=f"İşlenecek Veri Sayısı: {self.total_islem_sayisi}", font=('Times', 12, 'bold'))
        self.yapilan_islem_label = tk.Label(root, text=f"Yapılan İşlem Sayısı: {self.yapilan_islem_sayisi}", font=('Times', 12, 'bold'))
        self.kalan_islem_label = tk.Label(root, text=f"Kalan İşlem Sayısı: {self.kalan_islem_sayisi}", font=('Times', 12, 'bold'))

        self.basla_button.pack(side='left', pady=(0, 420), padx=(150, 0))
        self.dur_button.pack(side='left', pady=(0, 420))
        self.kapat_button.pack(side='left', pady=(0, 420))

        self.total_islem_label.place(relx=0.5, rely=0.87, anchor='center')
        self.yapilan_islem_label.place(relx=0.5, rely=0.92, anchor='center')
        self.kalan_islem_label.place(relx=0.5, rely=0.97, anchor='center')

    def adjust_opacity(self, image, opacity):
        image = image.copy()
        alpha = image.split()[3]
        alpha = alpha.point(lambda p: p * opacity)
        image.putalpha(alpha)
        return image

    def guncelle_labels(self):
        self.total_islem_label.config(text=f"Toplam İşlem Sayısı: {self.total_islem_sayisi}")
        self.yapilan_islem_label.config(text=f"Yapılan İşlem Sayısı: {self.yapilan_islem_sayisi}")
        self.kalan_islem_label.config(text=f"Kalan İşlem Sayısı: {self.kalan_islem_sayisi}")
        if not self.processing:
            messagebox.showinfo("INFO", "İşlem durduruldu. Tekrar başlatmak için 'Başlat' butonuna basınız.")

    def basla_islem(self):
        if not self.processing:
            self.processing = True
            self.basla_button.config(state=tk.DISABLED)
            self.thread = threading.Thread(target=self.main)
            self.thread.start()

    def dur_islem(self):
        if self.processing:
            messagebox.showinfo("INFO", "İşlem sürüyor. İşlem tamamlandığında program duracaktır.")
            self.processing = False

    def kapat(self):
        if self.processing:
            messagebox.showerror("INFO", "Lütfen işlem tamamlanmadan kapatmayın.")
        else:
            self.root.destroy()

    def main(self):
        pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd
        connectionDB = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={self.config['server']};"
            f"DATABASE={self.config['database']};"
            f"UID={self.config['username']};"
            f"PWD={self.config['password']};"
        )
        cursor = connectionDB.cursor()
        resultsOfPDF = isControlCheck(cursor)

        self.total_islem_sayisi = len(resultsOfPDF)
        self.kalan_islem_sayisi = len(resultsOfPDF)
        self.guncelle_labels()

        if self.total_islem_sayisi == 0:
            self.kalan_islem_sayisi = 0
            messagebox.showinfo("INFO", "İşlem başlatılamadı. İşlenecek bir veri bulunamadı.")

        for result in resultsOfPDF:


            if not self.processing:
                break

            try:
                dataArray = processPDF(result[0])
                Id = result[1]
                CompanyId = result[2]

                if not dataArray:
                    updateDB(cursor, connectionDB, None, None, result[0], Id, CompanyId)
                else:
                    date = dataArray[0]
                    tüketimVerisi = dataArray[-1]
                    updateDB(cursor, connectionDB, date, tüketimVerisi, None, Id, CompanyId)

                self.yapilan_islem_sayisi += 1
                self.kalan_islem_sayisi -= 1
                self.guncelle_labels()

            except Exception as e:
                print("An error occurred:", str(e))

        self.processing = False
        self.basla_button.config(state=tk.NORMAL)

if __name__ == '__main__':
    root = tk.Tk()
    root.iconbitmap(default="adasoo.ico")
    config = load_config('config.json')

    app = PDFProcessingApp(root, config)
    root.mainloop()