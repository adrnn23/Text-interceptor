import pytesseract
import numpy as np
import cv2
import pandas as pd
import os
import time

#Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

input_path = 'path/input'
output_path = 'path/output/'

def isDark(cell):
    avg = np.mean(cell)
    if avg < 50:
        return True
    return False

#Preprocessing obrazu
def preProcess(image):
    n_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    kernel = np.ones((2,2), np.uint8)
    n_image = cv2.erode(n_image, kernel, iterations=1)
    n_image = cv2.dilate(n_image, kernel, iterations=1)
    return n_image

#Odczytanie struktury tabeli
def detectTable(image):
    start = time.time()
    image = cv2.threshold(image, 190, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY)[1]
    image = cv2.bitwise_not(image)
    
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image).shape[1] // 10, 1))
    h_lines = cv2.erode(image, h_kernel, iterations=2)
    h_lines = cv2.dilate(h_lines, h_kernel, iterations=2)

    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, np.array(image).shape[0] // 10))
    v_lines = cv2.erode(image, v_kernel, iterations=2)
    v_lines = cv2.dilate(v_lines, v_kernel, iterations=2)

    #Połączenie linii
    combined_lines = cv2.addWeighted(h_lines, 0.5, v_lines, 0.5, 0.0)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    combined_lines = cv2.dilate(combined_lines, kernel, iterations=1)
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
            end = time.time()
            print("Detection time:", end-start)
            return True, grid

    return False, grid

def extractTable(input_path, output_path, iter):
    image = cv2.imread(input_path)
    image = preProcess(image)
    cv2.imshow("combined_lines",image)
    cv2.waitKey(0)
    edges = cv2.Canny(image, 30, 255)
    x, y, w, h = cv2.boundingRect(edges)
    table = image

    isTable, contours = detectTable(table)

    if isTable:
        print("Table is detected.")
        data = []
    #     for contour in contours:
    #         x, y, w, h = cv2.boundingRect(contour)
    #         if w > 20 and h > 20:
    #             cell = table[y:y+h, x:x+w]
    #             text = pytesseract.image_to_string(cell, config='--psm 6')
    #             data.append((y, x, text.strip()))

    #     rows = {}
    #     for y, x, text in data:
    #         if y not in rows:
    #             rows[y] = {}
    #         rows[y][x] = text

    #     table_list = []
    #     for y in sorted(rows.keys()):
    #         row = rows[y]
    #         table_list.append([row[x] if x in row else '' for x in sorted(row.keys())])

    #     df = pd.DataFrame(table_list)
    #     xlsx_path = 'path/output/xlsx' + str(iter) + '.xlsx'
    #     df.to_excel(xlsx_path, index=False, header=False)
    else:
        print("Image without table.")

#Przetwarzanie calej zawartosci folderu
def processFolder(input_path, output_path):
    iter = 0
    for filename in os.listdir(input_path):
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            iter += 1
            image_path = os.path.join(input_path, filename)
            extractTable(image_path, output_path, iter)

processFolder(input_path, output_path)