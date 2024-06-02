import pytesseract
import numpy as np
import cv2
import pandas as pd
import time
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox 

#Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

# Sprawdza czy obszar jest pojedyncza komorka
# Po processingu obrazu komorka jest czarna w srodku
def isDark(cell):
    avg = np.mean(cell)
    if avg < 50:
        return True
    return False

# Preprocessing obrazu np. zrzut ekranu
def preProcess(image):
    n_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return n_image

# Odczytanie struktury tabeli, szukanie linii poziomych i pionowych
def detectTable(image):
    image = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 7, 2)
    # start = time.time()
    image = cv2.bitwise_not(image)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (image.shape[1] // 10, 1))
    h_lines = cv2.erode(image, h_kernel, iterations=1)
    h_lines = cv2.dilate(h_lines, h_kernel, iterations=1)

    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, image.shape[0] // 10))
    v_lines = cv2.erode(image, v_kernel, iterations=1)
    v_lines = cv2.dilate(v_lines, v_kernel, iterations=1)

    combined_lines = cv2.addWeighted(h_lines, 0.5, v_lines, 0.5, 0.0)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    combined_lines = cv2.dilate(combined_lines, kernel, iterations=2)
    cv2.imshow("combined_lines",combined_lines)
    cv2.waitKey(0)
    
    grid, _ = cv2.findContours(combined_lines, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    counter = 0
    for contour in grid:
        x, y, w, h = cv2.boundingRect(contour)
        if w > 20 and h > 20:
            cell = image[y:y+h, x:x+w]
            if isDark(cell):
                counter += 1
        if counter > 3:
            # end = time.time()
            return True, grid

    return False, grid

def extractTable(input_path):
    image = cv2.imread(input_path)
    if image is None:
        print(f"Blad wczytywania pliku: {input_path}")
        return
    image = preProcess(image)
    edges = cv2.Canny(image, 30, 255)
    x, y, w, h = cv2.boundingRect(edges)
    table = image[y:y+h, x:x+w]
    isTable, contours = detectTable(table)
    
    if isTable:
        messagebox.showinfo("Result.", "Table is detected.") 
        data = []
        for i in range(len(contours)-1):
            x, y, w, h = cv2.boundingRect(contours[i])
            if w > 20 and h > 20:
                cell = table[y:y+h, x:x+w]
                text = pytesseract.image_to_string(cell, config='--psm 6')
                data.append((y, x, text.strip()))

        rows = {}
        for y, x, text in data:
            if y not in rows:
                rows[y] = {}
            rows[y][x] = text

        table_list = []
        for y in sorted(rows.keys()):
            row = rows[y]
            table_list.append([row[x] if x in row else '' for x in sorted(row.keys())])

        df = pd.DataFrame(table_list)
        saveXlsx(df)
    else:
        messagebox.showinfo("Result.", "Image without table.") 

# wczytanie pliku
def openFile():
    filepath = filedialog.askopenfilename(title="Open image")
    if filepath.endswith(('.png', '.jpg', '.jpeg')):
        with open(filepath, 'r') as file:
            extractTable(filepath)
    else:
        messagebox.showwarning("Result.", "Invalid file extension.") 

# zapis tabelki z obrazu do pliku xlsx
def saveXlsx(df):
    filepath = filedialog.asksaveasfilename(title="Save Excel file", defaultextension=".xlsx", filetypes=[("Excel files", ".xlsx .xls")])
    if filepath:
        df.to_excel(filepath, index=False, header=False)
        messagebox.showinfo("Result.", "Table extracted.") 

# glowne okno table extractora
def windowExtrTab():
    # konfiguracja glownego okna table extractora
    window = Tk()
    window.geometry('400x200')
    window.title("Table extraction")
    window.configure(bg='black')
    window.resizable(False,False)

    tab_ext = StringVar()
    tab_ext.set("Table extractor")
    Label(window,textvariable=tab_ext,anchor=CENTER, bg="orange", justify=CENTER, height=2, width=30, font=("Arial", 18),relief='raised', borderwidth=2).pack(pady=10)

    Button(window, text="Extract table from image",command=openFile, 
           bg='orange', fg='black', font=("Arial", 12), relief='raised', borderwidth=2, activebackground='darkorange', width=20, cursor="hand2").pack(pady=10)
    Button(window, text="Quit", command=window.destroy, 
           bg='orange', fg='black', font=("Arial", 12), relief='raised', borderwidth=2, activebackground='darkorange', width=20, cursor="hand2").pack(pady=10) 

    window.mainloop()

if __name__ == '__main__':
    windowExtrTab()