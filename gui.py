import tkinter as tk
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

root = tk.Tk()
root.geometry("400x300")  # Size of the window 
root.title('Powerpoint automator')
my_font1=('times', 18, 'bold')



            

class Data:
    
    def __init__(self, df):
        self.df = df
        
    def upload():
        file = filedialog.askopenfilename(title="Select a file",
                                    filetypes=((".xlsx files", "*.xlsx"), ("all files", "*.*")))
        if file:
            Data.data(path=path)
            print(df.shape)
    
    
    def data(self, file):
        global df
        self.df = pd.read_excel(file)
        return df
        


path = tk.StringVar()
path.set("")


l1 = tk.Label(root, text='Upload File & read', width=30, font=my_font1)  
l1.grid(row=1,column=1)
b1 = tk.Button(root, text='Upload File', width=20,command = Data.upload)
b1.grid(row=2 ,column=1)
l2 = tk.Label(root, textvariable=path, fg='red' )
l2.grid(row=3,column=1) 

# banyak_region = df['responses.region'].value_counts().rename_axis('region').reset_index(name='count').sort_values(by=['region'])
root.mainloop()  # Keep the window open