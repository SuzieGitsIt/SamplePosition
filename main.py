import configparser                                     # parsing multiple GUI's
import datetime as dt                                   # Date library
import keyboard                                         # windows right key
import os                                               # closing an executable
import pyautogui                                        # automating screen clicks
import pymem                                            # checking if .exe is open
import pymem.process                                    # checking if .exe is open
import pywinauto                                        # bringing an .exe to the foreground
import subprocess                                       # open an executable
import time                                             # call time to count/pause
import tkinter as tk                                    # Tkinter's Tk class
import tkinter.ttk as ttk                               # Tkinter's Tkk class
import win32con                                         # justify right or left the GUI.
import win32gui                                         # bring apps to front foreground

from functools import partial                           # freezing one function while executing another
from openpyxl import *                                  # Write to excel
from pathlib import PureWindowsPath                     # library that cleans up windows path extensions
from PIL import ImageTk, Image                          # Displaying LAL background photo
import imagesearch                                      # opening images, pip package
from tkinter import messagebox                          # Exit standard message box
from win32gui import GetWindowText, GetForegroundWindow # check position of a window

config = configparser.ConfigParser()
btn_pres_cnt = 1                                        # pink button, global variable to keep count throughout all GUI's
samp_arr_raw =[]

##########################################################################################################################################
########################################################     Run Test       ##############################################################
##########################################################################################################################################
class Test(tk.Toplevel):
    def __init__(self):                                                   
        super().__init__()

##########################################################################################################################################
########################################################    GUI STYLES      ##############################################################
##########################################################################################################################################
class GuiStyles():                  # create class to store all GUI styles
    def __init__(self, win_style):  # In __init__ method add an argument to set which window to style
        win_style.configure(background='white')              # Set background color
        win_style.option_add('*foreground', 'black')         # set the text color, hex works too '#FFFFFF'
        win_style.option_add('*background', 'white')         # set the background to white

        #################################################     TTK BUTTON & LABEL STYLE         ################################################
        win_style.style = ttk.Style()
        win_style.style.theme_use('default')  # alt, default, clam and classic

        win_style.style.map('T.TButton', background=[('active', 'pressed', 'white'), ('!active', 'white'),('active', '!pressed', 'grey')])  # active, not active, not pressed
        win_style.style.map('T.TButton', relief=[('pressed', 'sunken'), ('!pressed', 'raised')])  # pressed, not pressed
        win_style.style.configure('T.TButton', font=('Helvetica', '10'))

        win_style.style.map('B.TButton', background=[('active', 'pressed', 'white'), ('!active', 'white'), ('active', '!pressed', 'grey')])  # Press me Button always hot pink when pressed
        win_style.style.map('B.TButton', relief=[('pressed', 'sunken'), ('!pressed', 'raised')])  # pressed, not pressed
        win_style.style.configure('B.TButton', font=('helvetica', '12', 'bold'))

        win_style.style.configure('R.TLabel', font=('helvetica', '12'), foreground='black', background='white')
        win_style.style.configure('B.TLabel', font=('helvetica', '10', 'bold'), foreground='black', background='white')

        win_style.style.map('TCombobox', background= [('readonly', 'white')])
        win_style.style.map('TCombobox', fieldbackground= [('readonly', 'white')], foreground= 'black')
        win_style.style.map('TCombobox', selectbackground= [('readonly', 'white')], selectforeground= 'black')
        win_style.style.configure('TCombobox', font=('helvetica', '12', 'bold'))

##########################################################################################################################################
###############################################             Sample Selection         #####################################################
##########################################################################################################################################
class Samples(tk.Toplevel):
    def __init__(self):                                                         # Special Method, first argument is self.
        super().__init__()

        self.geometry('955x250')
        self.title("Samples")                                                   # title might have to be main window, from the last statement; if __name__ == "__main__":
        GuiStyles(win_style=self)                                               # Create an instance of 'GuiStyles' class and add 'self' as an argument
        self.style = ttk.Style()                                                # Still have to create 'self.style' or else we get an error

        samp_arr_raw.clear()                                                    # default write 1-25 in main GUI. Clear here before starting 
        samp_temp_del  = []                                                     # initialize an empty array to hold a list of array values to be removed from the samp_arr_raw
        samp_temp      = []                                                     # initialize an empty array to hold a list of values in Toggle
        samp_arr_mine  = []                                                     # initialize an empty array and writing all the sample numbers to it ... idk if we need this...
        btn_text       = []                                                     # initialize an empty array to hold a list of button texts
        btn_ids = []                                                            # initialize an empty array to hold a list of button id's   
        btn_row1 = []                                                           # initialize an empty array to hold a list of button variables that will be assigned a button
        btn_row2 = []
        btn_row3 = []

        count_btn_txt = 0                                                       # button text starts at 1, index starts at 0

        set1_var = tk.IntVar()
        set2_var = tk.IntVar()
        set3_var = tk.IntVar()

        samp_sz_var01 = tk.StringVar()
        samp_sz_var26 = tk.StringVar()
        samp_sz_var51 = tk.StringVar()
        samp_sz_var25 = tk.StringVar()
        samp_sz_var50 = tk.StringVar()
        samp_sz_var75 = tk.StringVar()

        def toggle(txt_i):                                                      # function to toggle what button click does
            btn_name = (btn_ids[txt_i])                                         # create 50 unique button names from btn_ids[n]
            btn_name.config(text=f"{btn_text[txt_i]}") 
            print("button name: ", btn_name)                                    # btn_name = .!window.!button##
            if btn_name.config('relief')[-1] == 'sunken':                       # if btn_name relief configuration is equal to sunken
                btn_name.config(relief="raised", bg='white')                    # second button click raises the button back up
                samp_temp_del.remove(txt_i+1)                                   # second button click removes the button # from the temp array of samples to be deleted from samp_arr_raw
                print("Sample removed from temporary delete list: ", samp_temp_del)
            else:
                btn_name.config(relief="sunken", bg='grey')                     # first button click sinks the button
                samp_temp_del.append(txt_i+1)                                   # first button click adds the button # to the temp array of samples to be deleted from samp_arr_raw
                print("Sunken: samples to be deleted array: ", samp_temp_del)

            samp_arr_raw = [*set(samp_arr_mine)]                                # remove duplicates from global array samp_arr_raw
            samp_temp = [*set(samp_temp_del)]                                   # remove duplicates from temporay array of samples to delete from samp_arr_raw
            samp_temp.sort(reverse=False)                                       # sort in ascending order.
            for a in samp_temp:
                try:
                    samp_arr_raw.remove(a)                                      # remove duplicates from global array
                except:
                    pass
            print("List after samples removed from array: ", ', '.join(map(str, samp_arr_raw)))
            print("Display: samples to delete array: ", samp_temp)

        def reset():
            samp_arr_mine.clear()                                               # remove all items from local list/array
            samp_arr_raw.clear()                                                # remove all items from global list/array
            samp_temp_del.clear()                                               # remove all items from local temp delete list/array
            print("Mine Sample Array after .clear() ", samp_arr_mine)
            print("Raw Sample Array after .clear() ", samp_arr_raw)
            print("Remove Sample Array after .clear() ", samp_temp_del)

        def save():
            print("Save: Samples to be deleted from the array: ", samp_temp_del)
            for b in samp_temp_del:
                try:
                    samp_arr_raw.remove(b)
                except:
                    pass
            print("List after samples removed from array before passing to Sample Sizes and Dioptics GUI: ", ', '.join(map(str, samp_arr_raw)))
            self.destroy()

        def exit():
            print("Clear and close.")
            reset()
            self.destroy() 

        def chk_btn1(btn_list):
            row = set1_var.get()                                                # get 0 or 1, (if checkbox is checked or not)
            start_pos = (int(Samples.samp_sz01)-1)                              # starting button position
            final_pos = int(Samples.samp_sz25)                                  # final button position
            col  = (final_pos - start_pos) 
            count = start_pos

            for c in range(0,25):                                               # create list btn_row1, and loop 25 times assigning button index location 
                btn_row1.append(btn_list[c])

            if row:
                print("Checkbox 1 selected", row)
                for d in range(col):                                            # c = column, max 25 columns
                    print(count, start_pos, d)
                    btn_row1[count].grid(row=1, column=d)
                    btn_row1[count].config(text=btn_text[start_pos])
                    samp_arr_mine.append(start_pos+1)                           # add all sample numbers to "Mine" list
                    samp_arr_raw.append(start_pos+1)                            # add all sample numbers to samp_arr_arr
                    start_pos += 1
                    count +=1
                print("Mine Sample Array: ", ', '.join(map(str, samp_arr_mine)))
                print("Raw Sample Array: ", ', '.join(map(str, samp_arr_raw)))

            else:                                                               # if row = 0, remove buttons from the GUI
                print("Checkbox 1 unselected", row)
                print("start_pos after writing to the list: ", start_pos)
                for f in range(0,25):                                           # ok to try and remove 25 buttons, won't cause an error.
                    btn_row1[f].grid_remove()

                for g in range(col):    
                    samp_arr_mine.remove(final_pos)                             # add all sample numbers to "Mine" list
                    samp_arr_raw.remove(final_pos)                              # re-add all sample numbers to samp_arr_arr
                    final_pos -=1
                    samp_temp_del.clear()                                       # need to clear only sample numbers 1-25 if in the temporary array
                print("Mine Sample Array: ", samp_arr_mine)
                print("Raw Sample Array:  ", samp_arr_raw)
                print("Temp Delete Array: ", samp_temp_del)

        def chk_btn2(btn_list):
            row = set2_var.get()
            start_pos  = (int(Samples.samp_sz26)-1)       
            final_pos = int(Samples.samp_sz50)
            col  = (final_pos - start_pos) 
            count = (start_pos -25)

            for h in range(25,50):
                btn_row2.append(btn_list[h])

            if row:
                print("Checkbox 2 selected", row)
                for i in range(col):                                                # c = column, max 25 columns
                    print(count, start_pos, i)
                    btn_row2[count].grid(row=2, column=i)
                    btn_row2[count].config(text=btn_text[start_pos])
                    samp_arr_mine.append(start_pos+1)                               # add all sample numbers to "Mine" list
                    samp_arr_raw.append(start_pos+1)                                # add all sample numbers to samp_arr_arr
                    start_pos += 1
                    count +=1
                print("Mine Sample Array: ", ', '.join(map(str, samp_arr_mine)))
                print("Raw Sample Array: ", ', '.join(map(str, samp_arr_raw)))

            else:
                print("Checkbox 2 unselected", row)
                for k in range(0,25): 
                    btn_row2[k].grid_remove()

                for m in range(col):    
                    samp_arr_mine.remove(final_pos)                                 # add all sample numbers to "Mine" list
                    samp_arr_raw.remove(final_pos)                                  # add all sample numbers to samp_arr_arr
                    final_pos -=1
                    samp_temp_del.clear()                                           # need to clear only sample numbers 1-25 if in the temporary array
                print("Mine Sample Array: ", samp_arr_mine)
                print("Raw Sample Array:  ", samp_arr_raw)
                print("Temp Delete Array: ", samp_temp_del)

        def chk_btn3(btn_list):
            row = set3_var.get()
            start_pos  = (int(Samples.samp_sz51) - 1)
            final_pos = int(Samples.samp_sz75)
            col  = (final_pos - start_pos) 
            count = (start_pos - 50)

            for n in range(50,75):
                btn_row3.append(btn_list[n])

            if row:
                print("Checkbox 3 selected", row)
                for p in range(col):                                                # c = column, max 25 columns
                    print(count, start_pos, p)
                    btn_row3[count].grid(row=3, column=p)
                    btn_row3[count].config(text=btn_text[start_pos])
                    samp_arr_mine.append(start_pos+1)                               # add all sample numbers to "Mine" list
                    samp_arr_raw.append(start_pos+1)                                # add all sample numbers to samp_arr_arr
                    start_pos += 1
                    count +=1
                print("Mine Sample Array: ", ', '.join(map(str, samp_arr_mine)))
                print("Raw Sample Array: ", ', '.join(map(str, samp_arr_raw)))      # update count_arr_nums for next button text
            
            else:
                print("Checkbox 3 unselected", row)
                for r in range(0,25): 
                    btn_row3[r].grid_remove()

                for s in range(col):    
                    samp_arr_mine.remove(final_pos)                                 # add all sample numbers to "Mine" list
                    samp_arr_raw.remove(final_pos)                                  # re-add all sample numbers to samp_arr_arr
                    final_pos -=1
                    samp_temp_del.clear()                                           # need to clear only sample numbers 1-25 if in the temporary array
                print("Mine Sample Array: ", samp_arr_mine)
                print("Raw Sample Array:  ", samp_arr_raw)
                print("Temp Delete Array: ", samp_temp_del)  

        def get_samp01(self, *args):                                            
            Samples.samp_sz01 = samp_sz_var01.get()                                 # If the value in the entry box changes, this will create a new array when display, save or remove samples is clicked.
            print(f"get_samp(): The sample size is: ", Samples.samp_sz01)

        def get_samp25(self, *args):                                            
            Samples.samp_sz25 = samp_sz_var25.get()                                 # If the value in the entry box changes, this will create a new array when display, save or remove samples is clicked.
            print(f"get_samp(): The sample size is: ", Samples.samp_sz25)

        def get_samp26(self, *args):                                            
            Samples.samp_sz26 = samp_sz_var26.get()                                  
            print(f"get_samp(): The sample size is: ", Samples.samp_sz26)

        def get_samp50(self, *args):                                            
            Samples.samp_sz50 = samp_sz_var50.get()                                  
            print(f"get_samp(): The sample size is: ", Samples.samp_sz50)

        def get_samp51(self, *args):                                            
            Samples.samp_sz51 = samp_sz_var51.get()                                  
            print(f"get_samp(): The sample size is: ", Samples.samp_sz51)

        def get_samp75(self, *args):                                            
            Samples.samp_sz75 = samp_sz_var75.get()                                  
            print(f"get_samp(): The sample size is: ", Samples.samp_sz75)

        def pink(event):
            global btn_pres_cnt                                                         # initializing btn_pres_cnt as a global varaible so that it adds through every iteration
            if (btn_pres_cnt == 10 or btn_pres_cnt == 20 or btn_pres_cnt == 30 or btn_pres_cnt == 40 or btn_pres_cnt == 50):  # button turns pink when btn_pres_cnt=100, and =200 and = 300.
                self.style.map('T.TButton', background=[('active', 'pressed', '#FF69B4'), ('!active', 'white'), (
                'active', '!pressed', 'grey')])                                         # only the button being pressed turns hot pink
                self.style.configure('T.Button', font=('Helvetica', '12', 'bold'))
            else:                                                                       # else is the normal style
                self.style.map('T.TButton',
                            background=[('active', 'pressed', 'white'), ('!active', 'white'), ('active', '!pressed', 'grey')])
                self.style.configure('T.Button', font=('Helvetica', '12', 'bold'))
            print('btn_pres_cnt = ', btn_pres_cnt)
            btn_pres_cnt += 1                                                           # This is always executed at the end of the if else

        lbl_cmd_dir1 = ttk.Label(self, text="Update initial and final sample position of the set if it does not", style='B.TLabel')
        lbl_cmd_dir1.place(x=740, y=105, anchor='center')
        lbl_cmd_dir2 = ttk.Label(self, text="start & end at the value listed. Select the sample set checkbox", style='B.TLabel')
        lbl_cmd_dir2.place(x=740, y=125, anchor='center')
        lbl_cmd_dir3 = ttk.Label(self, text="that is going to be tested for the work order. Left click a sample", style='B.TLabel')
        lbl_cmd_dir3.place(x=740, y=145, anchor='center')
        lbl_cmd_dir4 = ttk.Label(self, text="number to remove the individual sample number from the set.", style='B.TLabel')
        lbl_cmd_dir4.place(x=740, y=165, anchor='center')

        lbl_cmd_s1 = ttk.Label(self, text="Initial Sample Position: ", style='B.TLabel')    # middle left, asking for user entry
        lbl_cmd_s1.place(x=85, y=110, anchor='w')
        lbl_cmd_s26 = ttk.Label(self, text="Initial Sample Position: ", style='B.TLabel')
        lbl_cmd_s26.place(x=85, y=135, anchor='w')
        lbl_cmd_s51 = ttk.Label(self, text="Initial Sample Position: ", style='B.TLabel')
        lbl_cmd_s51.place(x=85, y=160, anchor='w')

        lbl_cmd_s25 = ttk.Label(self, text="Final Sample Position: ", style='B.TLabel')     # middle left, asking for user entry
        lbl_cmd_s25.place(x=305, y=110, anchor='w')
        lbl_cmd_s50 = ttk.Label(self, text="Final Sample Position: ", style='B.TLabel')
        lbl_cmd_s50.place(x=305, y=135, anchor='w')
        lbl_cmd_s75 = ttk.Label(self, text="Final Sample Position: ", style='B.TLabel')
        lbl_cmd_s75.place(x=305, y=160, anchor='w')

        samp_sz_var01.trace('w', get_samp01)
        entry_samp01 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var01, width=6)     
        entry_samp01.focus_set()                                                            # Places cursor in the first entry box.         
        entry_samp01.insert(0,1)                                                            # auto set value to 1  
        entry_samp01.place(x=240, y=102)

        samp_sz_var25.trace('w', get_samp25)
        entry_samp25 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var25, width=6)       
        entry_samp25.insert(0,25)                                                           # auto set value to 25
        entry_samp25.place(x=455, y=102)
   
        samp_sz_var26.trace('w', get_samp26)
        entry_samp26 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var26, width=6)           
        entry_samp26.insert(0,26)                                                           # auto set value to 26  
        entry_samp26.place(x=240, y=127)

        samp_sz_var50.trace('w', get_samp50)
        entry_samp50 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var50, width=6)             
        entry_samp50.insert(0,50)                                                           # auto set value to 50  
        entry_samp50.place(x=455, y=127)
    
        samp_sz_var51.trace('w', get_samp51)
        entry_samp51 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var51, width=6)          
        entry_samp51.insert(0,51)                                                           # auto set value to 51 
        entry_samp51.place(x=240, y=152)
   
        samp_sz_var75.trace('w', get_samp75)
        entry_samp75 = tk.Entry(self, justify=tk.LEFT, textvariable=samp_sz_var75, width=6)        
        entry_samp75.insert(0,75)                                                           # auto set value to 75  
        entry_samp75.place(x=455, y=152)

        for t in range(75):                                                                 # set button text 1 to 75, index = 0:74
            btn_text.append(count_btn_txt+1)                                                # append the empty list btn_text with text 1-75
            print("btn_text", btn_text[count_btn_txt])                                      # check the text by calling the index
            btn_samp = tk.Button(self, width=4, text=btn_text[count_btn_txt], relief="raised", command=partial(toggle, count_btn_txt))   # create buttons & assign unique arg (i) to run function (change)
            btn_samp.bind('<Button 1>', pink)
            btn_ids.append(btn_samp)                                                        # append the empty list btn_ids1 with 25 button widgets, index location 0:74
            count_btn_txt +=1                                                               # text and index

        btn_rad1 = tk.Checkbutton(self, text='Set 1', onvalue = 1, offvalue = 0, variable=set1_var, command=lambda: chk_btn1(btn_ids))
        btn_rad1.place(x=10, y=110, anchor='w')

        btn_rad2 = tk.Checkbutton(self, text='Set 2', onvalue = 1, offvalue = 0, variable=set2_var, command=lambda: chk_btn2(btn_ids))
        btn_rad2.place(x=10, y=135, anchor='w')

        btn_rad3 = tk.Checkbutton(self, text='Set 3', onvalue = 1, offvalue = 0, variable=set3_var, command=lambda: chk_btn3(btn_ids))
        btn_rad3.place(x=10, y=160, anchor='w')

        btn_rst = ttk.Button(self, text='Reset All', width = 10, style='T.TButton', command=partial(reset))
        btn_rst.bind('<Button-1>', pink)                                                    # class.function(instance)
        btn_rst.place(x=560, y=200)

        btn_save = ttk.Button(self, text='Save & Close', width = 14, style='T.TButton' ,command=partial(save))
        btn_save.bind('<Button-1>', pink)
        btn_save.place(x=670, y=200)

        btn_exit = ttk.Button(self, text='Exit Without Saving', width = 17, style='T.TButton', command=partial(exit))
        btn_exit.bind('<Button-1>', pink)
        btn_exit.place(x=800, y=200)

##########################################################################################################################################
#################################################              MAIN  SCREEN              #################################################
##########################################################################################################################################
#################################################      INITIALIZING STANDARD DISPLAY     #################################################
##########################################################################################################################################

class Main(tk.Tk):
    def __init__(self):                                                                     # Special Method, first argument is self.
        super().__init__()

        self.geometry('1050x620')                                                           # Set the geometry of the GUI.
        self.title("Main Window")
        GuiStyles(win_style=self)                                                           # Create an instance of 'GuiStyles' class and add 'self' as an argument
        self.style = ttk.Style()                                                            # Still have to create 'self.style' or else we get an error

        #################################################           LAL BACKGROUND IMAGE          ################################################
        # r stands for read, if we wanted to write to the file, we would put 'w'. If we wanted to append, we would put an 'a'     
        def resize_image(event):
            new_width = event.width
            new_height = event.height
            bg_img = copy_img.resize((new_width, new_height))
            new_img = ImageTk.PhotoImage(bg_img)
            lal_img.config(image=new_img)
            lal_img.bg_img = new_img                                                        # avoid garbage collection

        # r stands for read, if we wanted to write to the file, we would put 'w'. If we wanted to append, we would put an 'a'
        bg_img = Image.open(r'LAL.png')
        copy_img = bg_img.copy()
        new_img = ImageTk.PhotoImage(bg_img)
        lal_img = ttk.Label(self, image=new_img, background='white')
        lal_img.bind('<Configure>', resize_image)
        lal_img.pack(fill='both', expand=True)

        ################################################# Variables
        count_arr = 0                                                                       # count for array to 0
        dio_sz_var = tk.StringVar()                                                         # datatype of menu text
        sub_var = tk.StringVar()                                                            # datatype of menu text

        ################################################# Functions
        def get_dio(*args):                                              
            Main.dio_sz = dio_sz_var.get()
            print(f"get_dio(): The dioptic size is: ", Main.dio_sz)

        def get_sub(*args):                                              
            Main.sub_WO = sub_var.get()
            print(f"get_sub(): The sub work order is: ", Main.sub_WO)

        # Display user inputs as outputs
        def cred():
            op_cred = entry_cred.get()                                                      # entry_cred is the variable we are passing
            print("Length of op_cred: ", len(op_cred))
            if len(op_cred) > 3:                                                            # if operator credentials is first upper case, second lower case
                lbl_out_cred.configure(text=op_cred.title())                                # Display cred entry from user on main    
                print("op_cred: ", op_cred)
                Main.opcred = op_cred.title()                                               # entry_cred is the variable we are passing
            else:                                                                           # if not in caps, make it all caps
                if op_cred.isupper() is True:                                               # if work order is all in caps locks
                    lbl_out_cred.configure(text=op_cred)                                    # Display cred entry from user on self     
                    print("op_cred: ", op_cred)
                    Main.opcred = op_cred
                elif op_cred.isupper() is False:                                            # if not in caps, make it all caps
                    lbl_out_cred.configure(text=op_cred.upper())                            # Display cred entry from user on self      
                    print("opcred: ", op_cred.upper())
                    Main.opcred = op_cred.upper()
                           
        def work_ord():
            work_ord = entry_WO.get()[0:10]
            if work_ord.isupper() is True:                                                  # if work order is all in caps locks
                lbl_out_WO.configure(text=work_ord)                                         # Display WO entry from user on
                print("entry_wo: ", work_ord)
                Main.WO = work_ord                                                          # entry_WO is the variable we are passing. Limit 10 characters
            elif work_ord.isupper() is False:                                               # if not in caps, make it all caps
                lbl_out_WO.configure(text=work_ord.upper())                                 # Display WO entry from user on self
                print("entry_wo: ", work_ord.upper())
                Main.WO = work_ord.upper()                                                  # entry_WO is the variable we are passing. Limit 10 characters

        def meas(entry_meas):
            if entry_meas == '-B':
                btn_015 = ttk.Entry(self, width=15, font =('12'))
                btn_015.insert(0, 'Posterior -B')
                btn_015.place(x=220, y=420)
            elif entry_meas == '-A':
                btn_040 = ttk.Entry(self, width=15, font =('12'))
                btn_040.insert(0, 'Anterior -A')
                btn_040.place(x=220, y=420)
            elif entry_meas == '':
                btn_100 = ttk.Entry(self, width=15, font =('12'))
                btn_100.insert(0, 'Full Lens')
                btn_100.place(x=220, y=420)
            print("entry_meas is: ", entry_meas)
            cred()
            work_ord()
            Main.full_wo = Main.WO + Main.sub_WO + entry_meas + ' ' + Main.dio_sz + 'D'
            lbl_out_WO.configure(text=Main.full_wo)

        def display():                                                                      # display all inputs to user
            eq_num_oct = 'EQ# 1364 (Lumedica OCT 1)'
            #eq_num_oct = 'EQ# 2104 (Lumedica OCT 2)'
            print("OCT EQ is: ", eq_num_oct)
            cred()
            work_ord()
            lbl_out_WO.configure(text=Main.full_wo)
            lbl_out_samp.configure(text=', '.join(map(str, samp_arr_raw)))
            try:
                filepath = r"\\RXS-FS-02\userdocs\shaynes\My Documents\R&D - Software\Python/" + eq_num_oct
            except:     
                filepath = r"O:\\Operations\Lumedica OCT Data Backup\Notes to File/" + eq_num_oct
                print("Filepath  O:\\Operations\\Lumedica OCT Data Backup\\Notes to File  exists.")
            else:
                print("Filepath  shaynes\\My Documents\\R&D - Software\\Python  exists.")
            filename = Main.full_wo + '.xlsx'

            win_filepath = PureWindowsPath(filepath)
            if not os.path.exists(win_filepath):
                os.makedirs(win_filepath)

            loc = filepath + '/' + filename
            win_loc = PureWindowsPath(loc)

            lbl_out_fil_pat = ttk.Label(self, text=win_filepath, style='B.TLabel')
            lbl_out_fil_pat.place(x=220, y=60)
            lbl_out_fil_nam = ttk.Label(self, text=filename, style='B.TLabel')
            lbl_out_fil_nam.place(x=220, y=100)

            print("Filepath - WIN: \n", win_filepath)
            print("Filename:       \n", filename)
            print("Location - WIN: \n", win_loc)

        def exit():
            msg_box = tk.messagebox.askquestion('Exit', 'Are you sure you want to exit the application?', icon='warning')
            if msg_box == 'yes':
                self.destroy()
            else:
                tk.messagebox.showinfo('Exit', "Thanks for staying, please continue.")

        def pink(event):
            global btn_pres_cnt                                                             # initializing btn_pres_cnt as a global varaible so that it adds through every iteration
            if (btn_pres_cnt == 10 or btn_pres_cnt == 20 or btn_pres_cnt == 30 or btn_pres_cnt == 40 or btn_pres_cnt == 50):  # button turns pink when btn_pres_cnt=100, and =200 and = 300.
                self.style.map('T.TButton', background=[('active', 'pressed', '#FF69B4'), ('!active', 'white'), ('active', '!pressed', 'grey')])  # only the button being pressed turns hot pink
                self.style.configure('T.Button', font=('Helvetica', '12', 'bold'))
            else:                                                                           # else is the normal style
                self.style.map('T.TButton',background=[('active', 'pressed', 'white'), ('!active', 'white'), ('active', '!pressed', 'grey')])
                self.style.configure('T.Button', font=('Helvetica', '12', 'bold'))
            print('btn_pres_cnt = ', btn_pres_cnt)
            btn_pres_cnt += 1                                                               # This is always executed at the end of the if else

        ################################################                 MAIN BODY                ################################################
        for z in range(25):                                                                 # set button text based on user input of sample size
            samp_arr_raw.append(count_arr + 1)                                              # re-add all sample numbers to samp_arr_arr
            count_arr +=1                                      
        print("Sample Array initial ", ', '.join(map(str, samp_arr_raw)))

        # Display the command label before the entry box to indicate what information the Opterator is to type
        lbl_cmd_date = ttk.Label(self, text="Todays Date is:", style='B.TLabel').place(x=20, y=20)
        lbl_cmd_fold = ttk.Label(self, text="Folder Name:", style='B.TLabel').place(x=20, y=60)
        lbl_cmd_file = ttk.Label(self, text="File Name:", style='B.TLabel').place(x=20, y=100)
        lbl_cmd_cred = ttk.Label(self, text="Enter Operator Credentials:", style='B.TLabel').place(x=20, y=140)
        lbl_cmd_WO   = ttk.Label(self, text="Enter Work Order Number:", style='B.TLabel').place(x=20, y=180)
        lbl_cmd_meas = ttk.Label(self, text="Select Measurement Size:", style='B.TLabel').place(x=20, y=220)
        lbl_cmd_samp = ttk.Label(self, text="Enter Sample Sizes:", style='B.TLabel').place(x=20, y=260)

        # Display the label of what user input as an output
        lbl_disp_cred = ttk.Label(self, text="Credentials:", style='B.TLabel').place(x=20, y=340)
        lbl_disp_WO   = ttk.Label(self, text="Work Order Number:", style='B.TLabel').place(x=20, y=380)
        lbl_disp_meas = ttk.Label(self, text="Measurement Size:", style='B.TLabel').place(x=20, y=420)
        lbl_disp_samp = ttk.Label(self, text="Samples to be Tested:", style='B.TLabel').place(x=20, y=460)

        # Entry boxes to take information from operator
        entry_cred = ttk.Entry(self, width=13, font =('10'))
        entry_cred.focus_set()                                                              # Places cursor in the first entry box.
        entry_cred.place(x=220, y=138)
        entry_WO = ttk.Entry(self, width=13, font =('10'))
        entry_WO.place(x=220, y=178)

        drp_dn_sub_opt = ["-01", "-02", "-03", "-04", "-05", "-06", "-07", "-08", "-09", "-10"] # Dropdown menu options

        sub_var.trace('w', get_sub)
        sub_var.set("Sub WO")                                                               # initial menu text
        drp_dn_sub = ttk.Combobox(self, justify=tk.CENTER, style='C.TCombobox', height='30', font='10', textvariable=sub_var, width=8)  # Create Dropdown menu
        drp_dn_sub['values']= drp_dn_sub_opt 
        drp_dn_sub['state']= 'readonly'
        drp_dn_sub.place(x=365, y=178)

        drp_dn_dio_opt = [                                                                  # Dropdown menu options
            "4", "4.5", "5", "5.5", "6", "6.5", "7", "7.5", "8", "8.5", "9", "9.5", "10", "10.5", "11", "11.5", 
            "12", "12.5", "13", "13.5", "14", "14.5", "15", "15.5", "16", "16.5", "17", "17.5", "18", "18.5",
            "19", "19.5", "20", "20.5", "21", "21.5", "22", "22.5", "23", "23.5", "24", "24.5", "25", "25.5", 
            "26", "26.5", "27", "27.5", "28", "28.5", "29", "29.5", "30"]

        dio_sz_var.trace('w', get_dio)
        dio_sz_var.set("Dioptic Size")                                                      # initial menu text
        drp_dn_dio = ttk.Combobox(self, justify=tk.CENTER, style='C.TCombobox', height='30', font='10', textvariable=dio_sz_var, width=11)  # Create Dropdown menu
        drp_dn_dio['values']= drp_dn_dio_opt 
        drp_dn_dio['state']= 'readonly'
        drp_dn_dio.place(x=475, y=178)

        # Display the user inputs as outputs
        lbl_out_date = ttk.Label(self, text=f'{dt.datetime.now():%b %d, %Y}', style='B.TLabel').place(x=220, y=20)
        lbl_out_cred = ttk.Label(self, text='', style='R.TLabel', width=20)
        lbl_out_cred.place(x=220, y=340)
        lbl_out_WO = ttk.Label(self, text='', style='R.TLabel', width=20)
        lbl_out_WO.place(x=220, y=380)
        lbl_out_samp = ttk.Label(self, text='', style='R.TLabel', wraplength=640)
        lbl_out_samp.place(x=220, y=460)
        
        #################################################        BUTTONS TO BE CLICKED         ################################################  
        btn_015 = ttk.Button(self, text='Posterior', style='T.TButton', command=partial(meas, '-B'))  # Post - 015 
        btn_015.bind('<Button-1>', pink)
        btn_015.place(x=220, y=212)

        btn_040 = ttk.Button(self, text='Anterior', style='T.TButton', command=partial(meas, '-A'))   # Ant - 040 
        btn_040.bind('<Button-1>', pink)
        btn_040.place(x=320, y=212)

        btn_100 = ttk.Button(self, text='Full Lens', style='T.TButton', command=partial(meas, ''))   # Full - 100 
        btn_100.bind('<Button-1>', pink)
        btn_100.place(x=420, y=212)

        btn_samp = ttk.Button(self, text='Select Samples',style='T.TButton', command=partial(self.get_samp))  # Open new GUI 
        btn_samp.bind('<Button-1>', pink)
        btn_samp.place(x=220, y=252)

        btn_dis = ttk.Button(self, text='Display', width = 12, style='B.TButton', command=partial(display))
        btn_dis.bind('<Button-1>', pink)
        btn_dis.place(x=580, y=550)

        btn_run = ttk.Button(self, text='Run Test', width = 12, style='B.TButton', command=partial(self.run_test)) # Open auto test
        btn_run.bind('<Button-1>', pink)
        btn_run.place(x=730, y=550)

        btn_exit = ttk.Button(self, text='Close', width = 12, style='B.TButton', command=partial(exit))
        btn_exit.bind('<Button-1>', pink)
        btn_exit.place(x=880, y=550)

    def get_samp(self):
        samp = Samples()
        samp.grab_set()

    def run_test(self):
        test = Test()
        test.grab_set()
    
# Must be at the end of the program in order for the application to run b/c windows is constantly updating
if __name__ == "__main__":
    main = Main()
    main.mainloop()
