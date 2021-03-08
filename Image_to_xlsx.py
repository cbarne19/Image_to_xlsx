#Filer_Search_The_second 
import tkinter as tk
import tkinter.ttk as tkk
from tkinter import filedialog
from imageio import imread         
import xlsxwriter
from PIL import Image
import numpy as np 
import os

def imagesize_ect(imagename):    
    im = Image.open(imagename)       
    width, height = im.size       
    if width/height != 250/250: 
        if width > height:
            width = height 
        else:
            height = width    
    left = 0
    top = 0
    right = width 
    bottom = height  
    
    im1 = im.crop((left, top, right, bottom))   
    im1.save("my_image_square.png", "PNG")
    im1 = Image.open("my_image_square.png")

    size = 250, 250
    im_resized = im1.resize(size, Image.ANTIALIAS)
    im_resized.save("my_image_resized.png", "PNG")
    img = imread('my_image_resized.png')
    os.remove('my_image_resized.png')
    os.remove("my_image_square.png")
    return img 


def fillfunc(colour,img,n,n2,worksheet,workbook):
    if colour == 'R':
        x = 0 
        col = "'#'+let+'0000'"
    if colour == 'G':
        x = 1
        col = "'#'+'00'+let+'00'" 
    if colour == 'B':
        x = 2 
        col = "'#'+'0000'+let"
    for i in range(n):
        perc = (i/(3*n))*100+(x/3)*100
        progress['value'] = perc
        window.update_idletasks()  
        for j in range(n2):
            cell_format = workbook.add_format()
            cell_format.set_pattern(1)  
            let = hex(img[i][j][x])[2:4] 
            if len(let) == 1:
                let = '0'+let
            cell_format.set_bg_color(eval(col))
            worksheet.write((i*3)+x,j, int(img[i][j][x]), cell_format)
            worksheet.set_row((i*3)+x, 12)
            worksheet.set_column(j, 1)


def browseFiles():
    label_perc.configure(text='')
    filename = filedialog.askopenfilename(initialdir = "/",
										filetypes = (("png files",
														"*.png*"),
													("jpg files",
                                                          "*.jpg*")))   
    select_and_do(filename)
    label_perc.configure(text='The conversion is done')

def select_and_do(filename):														        
    img = imagesize_ect(filename)
    workbook = xlsxwriter.Workbook(filename[0:len(filename)-4]+'.xlsx')
    worksheet = workbook.add_worksheet()    
    colours = ['R','G','B']     
    n = int(len(img))
    n2 = int(len(img[1]))    
    for colour in colours:
        fillfunc(colour,img,n,n2,worksheet,workbook)
    workbook.close()
    progress['value'] = 100
    window.update_idletasks()  
window =tk.Tk()

progress = tkk.Progressbar(window, orient = tk.HORIZONTAL, 
            length = 200, mode = 'indeterminate') 
     
window.title('Image to .xlsx')
window.geometry("500x500")
window.config(background = "white")

label_file_explorer =tk.Label(window, 
							text = "Image to .xlsx",
							width = 100, height = 4, 
							fg = "blue") 
	
label_perc = tk.Label(window,text = "",
							width = 100, height = 4)

button_explore = tk.Button(window, 
						text = "Browse Files",
						command = browseFiles) 

progress.grid(column = 1, row=3)

label_file_explorer.grid(column = 1, row = 1)
button_explore.grid(column = 1, row = 2)
label_perc.grid(column=1,row=4)

window.mainloop()
