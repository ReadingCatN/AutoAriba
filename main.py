from ariba_ope import Ariba_Auto
from tkinter import PhotoImage
import tkinter as tk
import os

def create_ui(my_instance):
    def on_quality_check():
        option=quality_check_var.get()
        if option==1:
            my_instance.static_cat_quality_check(mode=0)
        elif option==2:
            my_instance.static_cat_quality_check(mode=1)
    
    def on_upload_check():
        option=file_download_var.get()
        if option==1:
            my_instance.ariba_admin_login()
            my_instance.ariba_catpage_get(0)
            my_instance.ariba_cat_upload(mode=0)
            ariba_case1.ariba_admin_logout()
        elif option==2:
            my_instance.ariba_admin_login()
            my_instance.ariba_catpage_get(0)
            my_instance.ariba_cat_upload(mode=1)
            ariba_case1.ariba_admin_logout()
    
    def on_download_check(): 
        option=file_download_var.get()
        if option==1:
            my_instance.ariba_admin_login()
            my_instance.ariba_catpage_get(1)
            my_instance.ariba_cat_download(mode=0,input_type=0)
            ariba_case1.ariba_admin_logout()
        elif option==2:
            my_instance.ariba_admin_login()
            my_instance.ariba_catpage_get(1)
            my_instance.ariba_cat_download(mode=0,input_type=1)
            ariba_case1.ariba_admin_logout()

    def on_total_cat_get():
        my_instance.ariba_admin_login(1)
        my_instance.ariba_catpage_get(1)
        my_instance.ariba_cat_download(mode=1,input_type=1)
        ariba_case1.ariba_admin_logout()


    # Create the main window
    root = tk.Tk()
    root.title("POA CM AUTO TOOL")

        # Set the size of the window
    root.geometry("400x300")

    # Set the background color of the window
    root.configure(bg="lightblue")

    base_path=my_instance.base_path_get()
    asset_path = os.path.join(base_path, "asset")
    cov_path=os.path.join(asset_path,"cov.png")
    icon = PhotoImage(file=cov_path)
    root.iconphoto(False, icon)

    button_width=20

    # Create Quality Check button with checkboxes
    quality_check_button = tk.Button(root, text="Quality Check", command=on_quality_check,width=button_width)
    quality_check_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    quality_check_var = tk.IntVar(value=2)
    quality_check_manual = tk.Radiobutton(root, text="Manual", variable=quality_check_var, value=1, bg="lightblue")
    quality_check_manual.grid(row=1, column=1, padx=10, pady=10, sticky="w")
    quality_check_auto = tk.Radiobutton(root, text="Auto", variable=quality_check_var, value=2, bg="lightblue")
    quality_check_auto.grid(row=1, column=2, padx=10, pady=10, sticky="w")

    # Create File Upload button with checkboxes
    file_upload_button = tk.Button(root, text="File Upload", command=on_upload_check,width=button_width)
    file_upload_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")

    file_upload_var = tk.IntVar(value=2)
    file_upload_option1 = tk.Radiobutton(root, text="Manual", variable=file_upload_var, value=1, bg="lightblue")
    file_upload_option1.grid(row=2, column=1, padx=10, pady=10, sticky="w")
    file_upload_option2 = tk.Radiobutton(root, text="Auto", variable=file_upload_var, value=2, bg="lightblue")
    file_upload_option2.grid(row=2, column=2, padx=10, pady=10, sticky="w")

    # Create File Download button with checkboxes
    file_download_button = tk.Button(root, text="File Download", command=on_download_check,width=button_width)
    file_download_button.grid(row=3, column=0, padx=10, pady=10, sticky="w")

    file_download_var = tk.IntVar(value=2)
    file_download_option1 = tk.Radiobutton(root, text="Manual", variable=file_download_var, value=1, bg="lightblue")
    file_download_option1.grid(row=3, column=1, padx=10, pady=10, sticky="w")
    file_download_option2 = tk.Radiobutton(root, text="Auto", variable=file_download_var, value=2, bg="lightblue")
    file_download_option2.grid(row=3, column=2, padx=10, pady=10, sticky="w")

    total_cat_get_button = tk.Button(root, text="Total Cat Get", command=on_total_cat_get,width=button_width)
    total_cat_get_button.grid(row=4, column=0, padx=10, pady=10, sticky="w")

    # Run the application
    root.mainloop()

if __name__ == "__main__":
    ariba_case1=Ariba_Auto()
    create_ui(ariba_case1)

