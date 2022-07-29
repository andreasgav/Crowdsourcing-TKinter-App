# Loading tkinter libraries to create the software

import tkinter as tk
from tkinter import *

# loading Python Imaging Library
from PIL import Image
from PIL import ImageTk, Image

# To get the dialog box to open when required
from tkinter import filedialog
from tkinter import font as tkFont

# loading walk library to load paintings in the image directory
import os
from os import walk

# loading libraries to handle dataframes and save the sentiment annotations to an excel file
import openpyxl
import pandas as pd

# Create a window
root = Tk()

# Set Title as Image Loader
root.title("Sentiment analysis crowd-sourcing software")

# Set the resolution of window

root.geometry("1920x1080")

# get the working directory

parent_directory = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
parent_directory = parent_directory.replace("\\", "/")
mypath = parent_directory + "/" + "Image_folder_Set/"

# set the targetted variable

emotion_var = tk.StringVar()

# create a list for the names of paintings that will be distributed to the participants of the survey

onlyfiles = []

for (dirpath, dirnames, filenames) in walk(mypath):
    onlyfiles.extend(filenames)
    break

print(onlyfiles)

# setting up of the iterator
item_num = 0

# Allow Window to be resizable
root.resizable(width=True, height=True)

# setting up the path for the image
im_path = mypath + filenames[item_num]

# creating of the function that will take place once the buttton is hit
def open_img():

    # setting global variables to be executed outside the function
    global item_num
    global filenames
    global im_path
    global path

    # Select the Imagename from a folder

    item_num += 1

    im_path = mypath + filenames[item_num]

    # opens the image
    img = Image.open(im_path)

    width, height = img.size

    # resize the image and apply a high-quality down sampling filter
    img = img.resize((width, height), Image.ANTIALIAS)

    # PhotoImage class is used to add image to widgets, icons etc
    img = ImageTk.PhotoImage(img)

    # create a label
    
    panel = Label(root, image=img)

    # set the image as img
    panel.image = img
    panel.grid(row=14, column=3)

    # get the emotion value
    emotion = emotion_var.get()

    # reading the excel file
    xfile = openpyxl.load_workbook(parent_directory + "/Emotion_Labels_Set.xlsx")
    sheet = xfile.get_sheet_by_name("Sheet_1")
    df = pd.read_excel(parent_directory + "/Emotion_Labels_Set.xlsx", engine="openpyxl")

    # setting where the sentiment value will be stored
    a_num_row = df.shape[0] + 2
    b_num_row = df.shape[0] + 1

    # Naming the sheet in the excel
    A_new_value = str("A" + str(a_num_row))
    B_new_value = str("B" + str(b_num_row))
    C_new_value = str("C" + str(b_num_row))

    sheet[A_new_value] = str(filenames[item_num])
    sheet[B_new_value] = emotion

    if emotion == "1":
        sheet[C_new_value] = "Joy"
    elif emotion == "2":
        sheet[C_new_value] = "Sadness"
    elif emotion == "3":
        sheet[C_new_value] = "Disgust"
    elif emotion == "4":
        sheet[C_new_value] = "Surprise"
    elif emotion == "5":
        sheet[C_new_value] = "Anger"
    elif emotion == "6":
        sheet[C_new_value] = "Fear"
    else:
        sheet[C_new_value] = "Wrong button pressed"

    # saving the data to excel
    xfile.save(parent_directory + "/Emotion_Labels_Set.xlsx")

    # erase the entry box after the value is passed
    emotion_entry.delete(0, END)
    
    
# set the text to be displayed on the software
emotion_label = Label(root, text="Sentiment selection", font=("calibre", 10, "bold"))

emotion_label_joy = Label(root, text="Joy=1", font=("calibre", 10, "bold"))
emotion_label_surpise = Label(root, text="Sadness=2", font=("calibre", 10, "bold"))
emotion_label_anger = Label(root, text="Disgust=3", font=("calibre", 10, "bold"))
emotion_label_sadness = Label(root, text="Surprise=4", font=("calibre", 10, "bold"))
emotion_label_fear = Label(root, text="Anger=5", font=("calibre", 10, "bold"))
emotion_label_disgust = Label(root, text="Fear=6", font=("calibre", 10, "bold"))

emotion_entry = Entry(root, textvariable=emotion_var, font=("calibre", 10, "normal"))


# placing the label and entry in
# the required position using grid
# method

emotion_label.grid(row=1, column=3)
emotion_label_joy.grid(row=2, column=3)
emotion_label_surpise.grid(row=3, column=3)
emotion_label_anger.grid(row=4, column=3)
emotion_label_sadness.grid(row=5, column=3)
emotion_label_fear.grid(row=6, column=3)
emotion_label_disgust.grid(row=7, column=3)

emotion_entry.grid(row=1, column=4)

im_path = mypath + filenames[item_num]

# opens the image
img = Image.open(im_path)

width, height = img.size
print(width)
print(height)

# resize the image and apply a high-quality down sampling filter
img = img.resize((width, height), Image.ANTIALIAS)

# PhotoImage class is used to add image to widgets, icons etc
img = ImageTk.PhotoImage(img)

# create a label
panel = Label(root, image=img)

# set the image as img
panel.image = img
panel.grid(row=14, column=3)

# Create a button and place it into the window using grid layout

btn = Button(root, text="Next", command=open_img).grid(row=1, column=5)

# set that the function can be executed by pressing the "Enter" button on the keyboard

root.bind("<Return>", lambda _: open_img())


# set the program looping
root.mainloop()


# End of program
