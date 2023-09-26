from tkinter import filedialog
import tkinter as tk
import tkinter.font as tkFont 
import pandas as pd
import os
import fnmatch
from tkinter import messagebox 


class App:
    def __init__(self, root):
        #setting title
        root.title("Mekong: Avalara to Woocommerce Tax Formatter")
        #setting window size
        width=615
        height=239
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GButton_875=tk.Button(root)
        GButton_875["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        GButton_875["font"] = ft
        GButton_875["fg"] = "#000000"
        GButton_875["justify"] = "center"
        GButton_875["text"] = "Open"
        GButton_875.place(x=70,y=60,width=100,height=30)
        GButton_875["command"] = self.browse_button

        self.GLabel_870=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        self.GLabel_870["font"] = ft
        self.GLabel_870["fg"] = "#333333"
        self.GLabel_870["justify"] = "center"
        self.GLabel_870["text"] = "From folder"
        self.GLabel_870.place(x=190,y=60,width=349,height=33)

        GButton_918=tk.Button(root)
        GButton_918["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        GButton_918["font"] = ft
        GButton_918["fg"] = "#000000"
        GButton_918["justify"] = "center"
        GButton_918["text"] = "Save to"
        GButton_918.place(x=70,y=120,width=100,height=31)
        GButton_918["command"] = self.save_button

        self.GLabel_637=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        self.GLabel_637["font"] = ft
        self.GLabel_637["fg"] = "#333333"
        self.GLabel_637["justify"] = "center"
        self.GLabel_637["text"] = "Save to"
        self.GLabel_637.place(x=190,y=120,width=351,height=32)
        
        #image1 = Image.open("MK5-100x100.png")
        #image1 = image1.resize((50, 50), Image.ANTIALIAS)
        #logo = ImageTk.PhotoImage(image1)
        #GLabel_logo=tk.Label(root,image=logo)
        #GLabel_logo.image = logo
        #GLabel_logo.place(x=15,y=10,width=498,height=30)

        GLabel_223=tk.Label(root)
        ft = tkFont.Font(family='Times',size=18)
        GLabel_223["font"] = ft
        GLabel_223["fg"] = "#333333"
        GLabel_223["justify"] = "center"
        GLabel_223["text"] = "Mekong: Avalara to Woocommerce Tax Converter"
        GLabel_223.place(x=20,y=10,width=498,height=30)

        GButton_686=tk.Button(root)
        GButton_686["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        GButton_686["font"] = ft
        GButton_686["fg"] = "#000000"
        GButton_686["justify"] = "center"
        GButton_686["text"] = "Convert"
        GButton_686.place(x=270,y=180,width=81,height=30)
        GButton_686["command"] = self.covertcsv
        
        self.GLabel_260=tk.Label(root)
        ft = tkFont.Font(family='Times',size=14)
        self.GLabel_260["font"] = ft
        self.GLabel_260["fg"] = "#333333"
        self.GLabel_260["justify"] = "center"
        self.GLabel_260["text"] = ""
        self.GLabel_260.place(x=350,y=180,width=100,height=30)
        
        source_path = ""
        save_path = ""
        
    def show_error(self, *args):
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception',err)

        
    def browse_button(self):
        # Allow user to select a directory and store it in global var
        # called folder_path
        self.GLabel_260.config(text = "")
        filename = filedialog.askdirectory()
        #folder_path.set(filename)
        self.GLabel_870.config(text = filename)
        self.source_path = filename
        print(filename)
        
        
    def save_button(self):
        # Allow user to select a directory and store it in global var
        # called folder_path
        filename = filedialog.askdirectory()
        self.GLabel_260.config(text = "")
        #folder_path.set(filename)
        self.GLabel_637.config(text = filename)
        self.save_path = filename
        print(filename)


    def covertcsv(self):
        self.GLabel_260.config(text = "")
        try:
            for f in os.listdir(self.source_path):
                if fnmatch.fnmatch(f, '*.csv'):
                    print (os.path.splitext(f)[0])
                    taxfile = pd.read_csv(self.source_path+"/"+f)
                    print ('Reading file by Pandas')
                    taxfile = taxfile.drop(['StateRate','EstimatedCountyRate','EstimatedCityRate','EstimatedSpecialRate','RiskLevel'], axis=1)
                    taxfile=taxfile.rename(columns = {'State':'State Code'})
                    taxfile=taxfile.rename(columns = {'ZipCode':'Zip/Postcode'})
                    taxfile=taxfile.rename(columns = {'TaxRegionName':'City'})
                    taxfile=taxfile.rename(columns = {'EstimatedCombinedRate':'Rate %'})
                    print ('Dropped and renamed columns')
                    taxfile.insert(0,'Country Code','US')
                    taxfile['Tax Name'] = taxfile['Country Code']+'-'+taxfile['State Code']+'-'+taxfile['City']
                    taxfile['Priority'] = 1
                    taxfile['Compound'] = 0
                    taxfile['Shipping'] = 1
                    taxfile['Tax Class'] = ''
                    print ('Inserted columns')
                    taxfile['Rate %'] = taxfile['Rate %']*100
                    print ('Calculated tax')
                    taxfile.to_csv(self.save_path+"/"+f,index=False)
                    print ('Saved! Done!')
            self.GLabel_260.config(text = "Saved")
        except BaseException as e:
            messagebox.showerror('Exception',e)
        


    def GButton_686_command(self):
        print("command")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.iconbitmap("MK5-ICO.ico")
    root.mainloop()
    

    
    
