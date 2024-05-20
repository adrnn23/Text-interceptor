import pytesseract
import numpy as np
import cv2

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

imgPath = 'path/text.png'

def resizeImage(image):
    n_image = image.copy()
    if(n_image.shape[0]<512 or n_image.shape[1]<512):
        x = int(n_image.shape[1]*1.6)
        y = int(n_image.shape[0]*1.6)
        n_image = cv2.resize(n_image,(x,y))
    elif(image.shape[0]>1024 or image.shape[1]>1024):
        x = int(n_image.shape[1]/2)
        y = int(n_image.shape[0]/2)
        n_image = cv2.resize(n_image,(x,y))
    return n_image

def removeNoise(image):
    kernel = np.ones((1,1), np.uint8)
    n_image = image.copy()
    n_image = cv2.dilate(n_image, kernel, iterations=1)
    n_image = cv2.erode(n_image, kernel, iterations=1)
    n_image = cv2.morphologyEx(n_image, cv2.MORPH_CLOSE, kernel)
   #n_image = cv2.medianBlur(n_image, 3)
    return n_image

def preProcess():
    image = cv2.imread(imgPath)

    if((image.shape[0]<512 or image.shape[1]<512) or (image.shape[0]>1024 or image.shape[1]>1024)):
        image = resizeImage(image)

    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    #for photos
    ###
    #image = cv2.GaussianBlur(image,(7,7),0)
    image = cv2.adaptiveThreshold(image,255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY,11,2)
    ###
    image = removeNoise(image)

    cv2.imwrite("path/textOpening.png", image)
    cv2.imshow("image", image)
    cv2.waitKey(0)

    text = pytesseract.image_to_string(image)
    with open("path/text.txt", "w") as f:
        f.write(text)
preProcess()