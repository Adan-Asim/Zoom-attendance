#!/usr/bin/env python
# coding: utf-8


# Import required packages
import cv2
import re
from pytesseract import *
import xlsxwriter
from datetime import datetime
from turtle import textinput
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
import time as t
from PIL import Image
from datetime import *
from pathlib import Path


#try:
ws = Tk()
ws.title('Smart Attendance System - FCIT')
ws.geometry('400x410') 
ws['bg']='DarkSlateGray'

# function that lets user select a 'jpeg'/'png' image and saves its path
path = "hi"
def open_file():
    global path
    path = filedialog.askopenfilename(initialdir='/Downloads', title='Select Photo', filetypes=(('JPEG files', '*.jpeg'), ('PNG files', '*.png')))
    if path is not None:
        pass


#a function that displays progress bar while image is uploading and then displays a successful message
def uploadFiles():
    pb1 = Progressbar(
        ws, 
        orient=HORIZONTAL, 
        length=200, 
        mode='determinate'
        )
    pb1.grid(row=8, columnspan=4, padx=10,pady=10)
    for i in range(5):
        ws.update_idletasks()
        pb1['value'] += 20
        t.sleep(1)
    pb1.destroy()
    Label(ws, text='Screenshot Uploaded Successfully!', foreground='white',background='green',font='Helvetica 10 bold').grid(row=9, columnspan=4, pady=10)

Label(ws, text='WELCOME TO SMART ATTENDANCE SYSTEM !',foreground='white',background='Teal',font='Helvetica 12 bold').grid(row=0, columnspan=4, pady=10)    
Label(ws, text='Instructions:',foreground='white',background='DarkSlateGray', font='Helvetica 11 bold').grid(row=1, columnspan=4, pady=10)
Label(ws, text='1. Select mode of screenshot Mobile/Desktop to upload image' ,foreground='white',background='DarkSlateGray',font='Helvetica 9 bold').grid(row=2, columnspan=4, pady=10)
Label(ws, text='2. Click on Upload button, and wait for the screenshot to upload' ,foreground='white',background='DarkSlateGray',font='Helvetica 9 bold').grid(row=3, columnspan=4, pady=10)
Label(ws, text='3. The date and time for the excel attendance file creation is displayed' ,foreground='white',background='DarkSlateGray',font='Helvetica 9 bold').grid(row=4, columnspan=4, pady=10)


#choose image button that let's user select the image from their computer
#choose image button that let's user select the image from their computer

type = 'N'

def Mobile(event):
    global type
    type = 'M'

adhar = Label(
    ws, 
    text='Mobile Screenshot: ',
    background='Teal',
     foreground='white',
    font='Helvetica 9 bold'
    )
adhar.grid(row=5, column=1, padx=15, pady=15)

adharbtn = Button(
    ws, 
    text ='Choose Image', 
    command = lambda:open_file()
    ) 
adharbtn.grid(row=5, column=2)
adharbtn.bind( "<Button>", Mobile)


def Desktop(event):
    global type
    type = 'D'

dadhar = Label(
    ws, 
    text='Desktop Screenshot: ',
    background='Teal',
    foreground='white',
    font='Helvetica 9 bold'
    
    )
dadhar.grid(row=6, column=1, padx=15, pady=15)


ddharbtn = Button(
    ws, 
    text ='Choose Image', 
    command = lambda:open_file()
    ) 
ddharbtn.grid(row=6, column=2)
ddharbtn.bind( "<Button>", Desktop)

#uploaf file button
upld = Button(
    ws, 
    text='Upload File', 
    command=uploadFiles
    )
upld.grid(row=7, column=1, columnspan=2, padx=10)

#display date and time in format
dt = datetime.date(datetime.now())
dt_f = dt.strftime("%A, %d %B %Y, %I:%M%p")
Label(ws, text=dt_f,background='DarkSlateGray',foreground='white',font='Helvetica 9 bold').grid(row=10, columnspan=5, pady=15)

ws.mainloop()
 
# Read image from which text needs to be extracted

string_path =str(path)
img = cv2.imread(string_path)


# Preprocessing the image starts

if type == 'D':
    x = 950
    y = 130
    h = 700
    w = 330
    img = img[y:y + h, x:x + w]
    

# # Convert the image to gray scale
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# # Performing OTSU threshold
ret, thresh1 = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)

# Specify structure shape and kernel size.
# Kernel size increases or decreases the area
# of the rectangle to be detected.
# A smaller value like (10, 10) will detect
# each word instead of a sentence.
rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 18))

# Appplying dilation on the threshold image
dilation = cv2.dilate(thresh1, rect_kernel, iterations = 1)

# Finding contours
contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_NONE)

# Creating a copy of image
im2 = img.copy()

# A text file is created and flushed
dt = datetime.now()
fname = 'Attendance ' + dt.strftime("%m-%d-%Y %H-%M-%S") + '.xlsx'
workbook = xlsxwriter.Workbook(fname)
worksheet = workbook.add_worksheet()


row = 0
col = 0
rollNo = ["init", "init2"]
worksheet.write(row, col, "Present Students")

def exists(text):
    global rollNo
    for x in rollNo:
        if(text == x):
            return True
    return False

# Looping through the identified contours
# Then rectangular part is cropped and passed on
# to pytesseract for extracting text from it
# Extracted text is then written into the text file
for cnt in contours:
	x, y, w, h = cv2.boundingRect(cnt)
	
	# Drawing a rectangle on copied image
	rect = cv2.rectangle(im2, (x, y), (x + w, y + h), (0, 255, 0), 2)
	
	# Cropping the text block for giving input to OCR
	cropped = im2[y:y + h, x:x + w]
	
	# Apply OCR and further processing on the cropped image
	text = pytesseract.image_to_string(cropped)

	text = re.findall("[B][a-z]{3}\d{2}[a-z]\d{3}",text)
	text = str(text)
	text = text[text.find('B'):]
	text = text[0:10]
    
	# Writing the text into file
	if text.startswith('B') and not exists(text):
	     row+=1
	     worksheet.write(row, col, text)
	     rollNo.append(text)	     

	# Close the file
workbook.close()
