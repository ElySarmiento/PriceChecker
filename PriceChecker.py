import customtkinter
from tkinter import Tk, filedialog
import PriceChecker_Shopee
import PriceChecker_Lazada
import PriceChecker_Tiktok
import os


class MyTabView(customtkinter.CTkTabview):


    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # create tabs
        self.add("Shopee")  
        self.add("Lazada")
        self.add("Tiktok")


        ## SHOPEE PRICE CHECKER TAB AND WIDGET
            
        ## Shopee TB BUTTON
        Shopee_TB_Button = customtkinter.CTkButton(master=self.tab("Shopee"),text="Select the TB file",command= self.S_TB_selectFile)
        Shopee_TB_Button.grid(row=0, column=0, padx=20, pady=10)  
       

        ## Shopee TB Label for the TB file Name
        self.Shopee_TB_Label = customtkinter.CTkLabel(master=self.tab("Shopee"),text='No selected file.')
        self.Shopee_TB_Label.grid(row=0, column=1, padx=20, pady=10,sticky="w")

    
        ## Shopee PWP BUTTON  
        Shopee_PWP_Button = customtkinter.CTkButton(master=self.tab("Shopee"),text="Select the PWP file",command= self.S_PWP_selectFile)
        Shopee_PWP_Button.grid(row=5, column=0, padx=20, pady=10,)
     
        #3 Shopee PWP Label for the PWP file name
        self.Shopee_PWP_Label = customtkinter.CTkLabel(master=self.tab("Shopee"),text='No selected file.')
        self.Shopee_PWP_Label.grid(row=5, column=1, padx=20, pady=10,sticky="w")  


        ## Shopee Destination button 
        Shopee_Savedir_Button = customtkinter.CTkButton(master=self.tab("Shopee"),text="Select the destination",command= self.S_selectFolder)
        Shopee_Savedir_Button.grid(row=10, column=0, padx=20, pady=10)

        ## Shopee Destination Label, to check if the destination directory is right.
        self.Shopee_Savedir_Label = customtkinter.CTkLabel(master=self.tab("Shopee"),text='No selected folder.')
        self.Shopee_Savedir_Label.grid(row=10, column=1, padx=20, pady=10, sticky="w") 

        ## Shopee Promo Label
        self.S_Promo_Label = customtkinter.CTkLabel(master=self.tab("Shopee"),text='Select Promo')
        self.S_Promo_Label.grid(row=15, column=0, padx=25, pady=10, sticky="w")
        self.S_Promo_Label.configure(font=("Helvetica", 15, "normal"))

        ## Shopee Promo ComboBox
        self.S_Promo_combobox = customtkinter.CTkComboBox(master=self.tab("Shopee"), values=['Select Promo'],command= self.S_combobox_callback)
        self.S_Promo_combobox.grid(row=15, column=1, padx=20, pady=10, sticky="w")

        ## Shopee Check the Price Button 
        self.Shopee_CheckPrice_Button = customtkinter.CTkButton(master=self.tab("Shopee"),text="Check the Price",command= PriceChecker_Shopee.main)
        self.Shopee_CheckPrice_Button.grid(row=20, column=0, padx=20, pady=10)
        self.Shopee_CheckPrice_Button.configure(font=("Helvetica", 15, "bold"),state="disabled")
        
        ## END OF SHOPEE PRICE CHECKER TAB  ###########################



        ## LAZADA PRICE CHECKER TAB AND WIDGET
            
        ## Lazada TB BUTTON
        Lazada_TB_Button = customtkinter.CTkButton(master=self.tab("Lazada"),text="Select the TB file",command= self.L_TB_selectFile)
        Lazada_TB_Button.grid(row=0, column=0, padx=20, pady=10)  

        ## Lazada TB Label for the TB file Name
        self.Lazada_TB_Label = customtkinter.CTkLabel(master=self.tab("Lazada"),text='No selected file.')
        self.Lazada_TB_Label.grid(row=0, column=1, padx=20, pady=10,sticky="w")

    
        ## Lazada PWP BUTTON  
        Lazada_PWP_Button = customtkinter.CTkButton(master=self.tab("Lazada"),text="Select the PWP file",command= self.L_PWP_selectFile)
        Lazada_PWP_Button.grid(row=5, column=0, padx=20, pady=10,)
     
        ## Lazada PWP Label for the PWP file name
        self.Lazada_PWP_Label = customtkinter.CTkLabel(master=self.tab("Lazada"),text='No selected file.')
        self.Lazada_PWP_Label.grid(row=5, column=1, padx=20, pady=10,sticky="w")  


        ## Lazada Destination button 
        Lazada_Savedir_Button = customtkinter.CTkButton(master=self.tab("Lazada"),text="Select the destination",command= self.L_selectFolder)
        Lazada_Savedir_Button.grid(row=10, column=0, padx=20, pady=10)

        ## Lazada Destination Label, to check if the destination directory is right.
        self.Lazada_Savedir_Label = customtkinter.CTkLabel(master=self.tab("Lazada"),text='No selected folder.')
        self.Lazada_Savedir_Label.grid(row=10, column=1, padx=20, pady=10, sticky="w") 

        ## Lazada Promo Label
        self.L_Promo_Label = customtkinter.CTkLabel(master=self.tab("Lazada"),text='Select Promo')
        self.L_Promo_Label.grid(row=15, column=0, padx=25, pady=10, sticky="w")
        self.L_Promo_Label.configure(font=("Helvetica", 15, "normal"))

        ## Lazada Promo ComboBox
        self.L_Promo_combobox = customtkinter.CTkComboBox(master=self.tab("Lazada"), values=['Select Promo'],command= self.L_combobox_callback)
        self.L_Promo_combobox.grid(row=15, column=1, padx=20, pady=10, sticky="w")

        ## Lazada Check the Price Button 
        self.Lazada_CheckPrice_Button = customtkinter.CTkButton(master=self.tab("Lazada"),text="Check the Price",command= PriceChecker_Lazada.main)
        self.Lazada_CheckPrice_Button.grid(row=20, column=0, padx=20, pady=10)
        self.Lazada_CheckPrice_Button.configure(font=("Helvetica", 15, "bold"),state="disabled")
        
        ## END OF LAZADA PRICE CHECKER TAB  ###########################


        ## TIKTOK PRICE CHECKER TAB AND WIDGET
            
        ## Tiktok TB BUTTON
        Tiktok_TB_Button = customtkinter.CTkButton(master=self.tab("Tiktok"),text="Select the TB file",command= self.T_TB_selectFile)
        Tiktok_TB_Button.grid(row=0, column=0, padx=20, pady=10)  

        ## Tiktok TB Label for the TB file Name
        self.Tiktok_TB_Label = customtkinter.CTkLabel(master=self.tab("Tiktok"),text='No selected file.')
        self.Tiktok_TB_Label.grid(row=0, column=1, padx=20, pady=10,sticky="w")

    
        ## Tiktok PWP BUTTON  
        Tiktok_PWP_Button = customtkinter.CTkButton(master=self.tab("Tiktok"),text="Select the PWP file",command= self.T_PWP_selectFile)
        Tiktok_PWP_Button.grid(row=5, column=0, padx=20, pady=10,)
     
        ## Tiktok PWP Label for the PWP file name
        self.Tiktok_PWP_Label = customtkinter.CTkLabel(master=self.tab("Tiktok"),text='No selected file.')
        self.Tiktok_PWP_Label.grid(row=5, column=1, padx=20, pady=10,sticky="w")  


        ## Tiktok Destination button 
        Tiktok_Savedir_Button = customtkinter.CTkButton(master=self.tab("Tiktok"),text="Select the destination",command= self.T_selectFolder)
        Tiktok_Savedir_Button.grid(row=10, column=0, padx=20, pady=10)

        ## Tiktok Destination Label, to check if the destination directory is right.
        self.Tiktok_Savedir_Label = customtkinter.CTkLabel(master=self.tab("Tiktok"),text='No selected folder.')
        self.Tiktok_Savedir_Label.grid(row=10, column=1, padx=20, pady=10, sticky="w") 

        ## Tiktok Promo Label
        self.T_Promo_Label = customtkinter.CTkLabel(master=self.tab("Tiktok"),text='Select Promo')
        self.T_Promo_Label.grid(row=15, column=0, padx=25, pady=10, sticky="w")
        self.T_Promo_Label.configure(font=("Helvetica", 15, "normal"))

        ## Tiktok Promo ComboBox
        self.T_Promo_combobox = customtkinter.CTkComboBox(master=self.tab("Tiktok"), values=['Select Promo'],command= self.T_combobox_callback)
        self.T_Promo_combobox.grid(row=15, column=1, padx=20, pady=10, sticky="w")

        ## Tiktok Check the Price Button 
        self.Tiktok_CheckPrice_Button = customtkinter.CTkButton(master=self.tab("Tiktok"),text="Check the Price",command= PriceChecker_Tiktok.main)
        self.Tiktok_CheckPrice_Button.grid(row=20, column=0, padx=20, pady=10)
        self.Tiktok_CheckPrice_Button.configure(font=("Helvetica", 15, "bold"),state="disabled")
        
        ## END OF TIKTOK PRICE CHECKER TAB  ###########################


    
    ## Shopee methods ###

    def S_TB_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the TB file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Shopee_TB_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Shopee.S_tb_file_path(filename)
        else:
            self.Shopee_TB_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Shopee_CheckPrice_Button.configure(state="disabled")
  

    def S_PWP_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the PWP file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Shopee_PWP_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Shopee.S_pwp_file_path(filename)
        else:
            self.Shopee_PWP_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Shopee_CheckPrice_Button.configure(state="disabled")

    
    def S_selectFolder(self, event=None):
        filename = filedialog.askdirectory(title="Select the PWP file")

        if filename:
            self.Shopee_Savedir_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Shopee.S_save_dir(filename)
            PriceChecker_Shopee.S_get_promo()
            val = PriceChecker_Shopee.promo_names
            self.S_Promo_combobox.configure(values=val)

        else:
            self.Shopee_Savedir_Label.configure(text= 'No folder Selected.', text_color='#ED8A80')
            self.Shopee_CheckPrice_Button.configure(state="disabled") 

    def S_combobox_callback(self,choice):
        if(choice == 'Select Promo'):
          self.Shopee_CheckPrice_Button.configure(state="disabled")
        else:    
            PriceChecker_Shopee.S_promo(choice)
            self.Shopee_CheckPrice_Button.configure(state="active")


    

    ## Lazada methods ###
    
    def L_TB_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the TB file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Lazada_TB_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Lazada.L_tb_file_path(filename)
        else:
            self.Lazada_TB_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Lazada_CheckPrice_Button.configure(state="disabled")
  

    def L_PWP_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the PWP file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Lazada_PWP_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Lazada.L_pwp_file_path(filename)
        else:
            self.Lazada_PWP_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Lazada_CheckPrice_Button.configure(state="disabled")

    
    def L_selectFolder(self, event=None):
        filename = filedialog.askdirectory(title="Select the PWP file")

        if filename:
            self.Lazada_Savedir_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Lazada.L_save_dir(filename)
            PriceChecker_Lazada.L_get_promo()
            val = PriceChecker_Lazada.promo_names
            self.L_Promo_combobox.configure(values=val)

        else:
            self.Lazada_Savedir_Label.configure(text= 'No folder Selected.', text_color='#ED8A80')
            self.Lazada_CheckPrice_Button.configure(state="disabled") 

    def L_combobox_callback(self,choice):
        if(choice == 'Select Promo'):
          self.Lazada_CheckPrice_Button.configure(state="disabled")
        else:    
            PriceChecker_Lazada.L_promo(choice)
            self.Lazada_CheckPrice_Button.configure(state="active")




    ## Tiktok methods ###
    
    def T_TB_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the TB file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Tiktok_TB_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Tiktok.T_tb_file_path(filename)
        else:
            self.Tiktok_TB_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Tiktok_CheckPrice_Button.configure(state="disabled")
  

    def T_PWP_selectFile(self, event=None):
        filename = filedialog.askopenfilename(title="Select the PWP file", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.Tiktok_PWP_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Tiktok.T_pwp_file_path(filename)
        else:
            self.Tiktok_PWP_Label.configure(text= 'No file Selected.', text_color='#ED8A80')
            self.Tiktok_CheckPrice_Button.configure(state="disabled")

    
    def T_selectFolder(self, event=None):
        filename = filedialog.askdirectory(title="Select the PWP file")

        if filename:
            self.Tiktok_Savedir_Label.configure(text= os.path.basename(filename), text_color='white')
            PriceChecker_Tiktok.T_save_dir(filename)
            PriceChecker_Tiktok.T_get_promo()
            val = PriceChecker_Tiktok.promo_names
            self.T_Promo_combobox.configure(values=val)

        else:
            self.Tiktok_Savedir_Label.configure(text= 'No folder Selected.', text_color='#ED8A80')
            self.Tiktok_CheckPrice_Button.configure(state="disabled") 

    def T_combobox_callback(self,choice):
        if(choice == 'Select Promo'):
          self.Tiktok_CheckPrice_Button.configure(state="disabled")
        else:    
            PriceChecker_Tiktok.T_promo(choice)
            self.Tiktok_CheckPrice_Button.configure(state="active")





class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("900x400")
        self.title("Price Checker")

        self.resizable(False, False)

        # Create a big label for "PRICE CHECKER" in the upper left
        big_label = customtkinter.CTkLabel(self)
        big_label.configure(text="PRICE CHECKER", font=("Helvetica", 25, "bold"))
        big_label.grid(row=0, column=0, padx=20, pady=20, sticky="nw")


        self.tab_view = MyTabView(master=self)
        self.tab_view.grid(row=0, column=0, padx=20, pady=50)
        self.tab_view.configure(width=850,height=330)



app = App()
app.mainloop()