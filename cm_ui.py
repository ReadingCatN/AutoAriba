from tkinter import PhotoImage
import tkinter as tk
import os

class CM_UI:
    def __init__(self, my_instance):
        self.my_instance = my_instance
        self.root = tk.Tk()
        self.create_window()
        self.create_widgets()
    
    def create_window(self):
        self.root.title("POA CM AUTO TOOL")
        self.root.geometry("400x300")
        self.root.configure(bg="lightblue")
        base_path = self.my_instance.base_path_get()
        asset_path = os.path.join(base_path, "asset")
        cov_path = os.path.join(asset_path, "cov.png")
        icon = PhotoImage(file=cov_path)
        self.root.iconphoto(False, icon)
        self.button_width = 20

    def create_widgets(self):
        # Quality Check button with checkboxes
        quality_check_button = tk.Button(self.root, text="Quality Check", command=self.on_quality_check, width=self.button_width)
        quality_check_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.quality_check_var = tk.IntVar(value=2)
        quality_check_manual = tk.Radiobutton(self.root, text="Manual", variable=self.quality_check_var, value=1, bg="lightblue")
        quality_check_manual.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        quality_check_auto = tk.Radiobutton(self.root, text="Auto", variable=self.quality_check_var, value=2, bg="lightblue")
        quality_check_auto.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # File Upload button with checkboxes
        file_upload_button = tk.Button(self.root, text="File Upload", command=self.on_upload_check, width=self.button_width)
        file_upload_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.file_upload_var = tk.IntVar(value=2)
        file_upload_option1 = tk.Radiobutton(self.root, text="Manual", variable=self.file_upload_var, value=1, bg="lightblue")
        file_upload_option1.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        file_upload_option2 = tk.Radiobutton(self.root, text="Auto", variable=self.file_upload_var, value=2, bg="lightblue")
        file_upload_option2.grid(row=2, column=2, padx=10, pady=10, sticky="w")

        # File Download button with checkboxes
        file_download_button = tk.Button(self.root, text="Change Report", command=self.on_download_check, width=self.button_width)
        file_download_button.grid(row=3, column=0, padx=10, pady=10, sticky="w")

        self.file_download_var = tk.IntVar(value=2)
        file_download_option1 = tk.Radiobutton(self.root, text="Manual", variable=self.file_download_var, value=1, bg="lightblue")
        file_download_option1.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        file_download_option2 = tk.Radiobutton(self.root, text="Auto", variable=self.file_download_var, value=2, bg="lightblue")
        file_download_option2.grid(row=3, column=2, padx=10, pady=10, sticky="w")

        total_cat_get_button = tk.Button(self.root, text="Total Cat Get", command=self.on_total_cat_get, width=self.button_width)
        total_cat_get_button.grid(row=4, column=0, padx=10, pady=10, sticky="w")

    def on_quality_check(self):
        option = self.quality_check_var.get()
        if option == 1:
            self.my_instance.static_cat_quality_check(mode=0)
        elif option == 2:
            self.my_instance.static_cat_quality_check(mode=1)

    def on_upload_check(self):
        option = self.file_upload_var.get()
        if option == 1:
            self.my_instance.ariba_admin_login()
            self.my_instance.ariba_catpage_get(0)
            self.my_instance.ariba_cat_upload(mode=0)
            self.my_instance.ariba_admin_logout()
        elif option == 2:
            self.my_instance.ariba_admin_login()
            self.my_instance.ariba_catpage_get(0)
            self.my_instance.ariba_cat_upload(mode=1)
            self.my_instance.ariba_admin_logout()

    def on_download_check(self):
        option = self.file_download_var.get()
        if option == 1:
            self.my_instance.ariba_admin_login()
            self.my_instance.ariba_catpage_get(1)
            self.my_instance.ariba_cat_download(mode=0, input_type=0)
            self.my_instance.ariba_admin_logout()
        elif option == 2:
            self.my_instance.ariba_admin_login()
            self.my_instance.ariba_catpage_get(1)
            self.my_instance.ariba_cat_download(mode=0, input_type=1)
            self.my_instance.ariba_admin_logout()

    def on_total_cat_get(self):
        self.my_instance.ariba_admin_login(1)
        self.my_instance.ariba_catpage_get(1)
        self.my_instance.ariba_cat_download(mode=1, input_type=1)
        self.my_instance.ariba_admin_logout()

    def run(self):
        self.root.mainloop()
