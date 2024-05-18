import pytesseract
import numpy as np
import cv2
import pandas as pd

#Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

#Preprocessing obrazu
def preProcess(image):
    kernel = np.ones((1, 1), np.uint8) 

    n_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    n_image = cv2.bitwise_not(n_image)
    dilate_table = cv2.dilate(n_image, kernel, iterations=1)
    erode_table = cv2.erode(dilate_table, kernel, iterations=1)
    n_image = cv2.bitwise_not(erode_table)
    n_image = cv2.morphologyEx(n_image, cv2.MORPH_OPEN, kernel)
    return n_image

#Odczytanie struktury tabeli
def getGrid(table):

    table = cv2.threshold(table, 150, 255, cv2.THRESH_BINARY_INV)[1]
    #Detekcja poziomych linii
    kernel = np.ones((2, 2), np.uint8) 
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (50, 1))
    detect_horizontal = cv2.morphologyEx(table, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    detect_horizontal = cv2.dilate(detect_horizontal, kernel, iterations=1)

    # Detekcja pionowych linii
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))
    detect_vertical = cv2.morphologyEx(table, cv2.MORPH_OPEN, vertical_kernel, iterations=1)
    detect_vertical = cv2.dilate(detect_vertical, kernel, iterations=1)

    #Polaczenie linii
    combined_lines = cv2.addWeighted(detect_horizontal, 0.5, detect_vertical, 0.5, 0.0)
    # cv2.imshow("c",combined_lines)
    # cv2.waitKey(0)
    grid, _ = cv2.findContours(combined_lines, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    return grid

def extractTable():
    image_path = 'path'

    image = cv2.imread(image_path)
    image = preProcess(image)

    edges = cv2.Canny(image, 50, 255)
    x, y, w, h = cv2.boundingRect(edges)
    table = image[y:y+h, x:x+w]
    cv2.imwrite('path', table)

    contours = getGrid(table)

    data = []

    #Wyodrebnienie tekstu z kazdej komorki
    for contour in range(len(contours)-1):
        x, y, w, h = cv2.boundingRect(contours[contour])
        cell = table[y:y+h, x:x+w]
        text = pytesseract.image_to_string(cell, config='--psm 6')
        data.append((y,x,text.strip()))

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
    df.to_excel('path.xlsx', index=False, header=False)

extractTable()