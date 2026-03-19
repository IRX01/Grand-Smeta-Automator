import os
import tkinter as tk
from tkinter import filedialog
from Test_script import run_bot
import time


def choose_folder():
    
    folder_path = filedialog.askdirectory()

    if folder_path:
        print(f"Выбранная папка: {folder_path}")
    entry_field.delete(0, tk.END)
    entry_field.insert(0, folder_path)

def run_script():
    folder_path = entry_field.get()
    os.system("taskkill /f /im excel.exe")
    time.sleep(2)
    run_bot(folder_path)
    


main = tk.Tk()
main.title("Робот")
main.geometry("600x400")

btn = tk.Button(main, text="Выбрать папку для сохранения (папка месяца)",font=("Arial", 14), command= choose_folder)
btn.pack(pady=20)

label = tk.Label(main, text="Выбранная папка:", font=("Arial", 14))
label.pack(anchor="w", padx=5, pady=2)
entry_field = tk.Entry(main, font=("Arial", 14), width=50)
entry_field.pack(fill="x", padx=5, pady=2)

btn_start = tk.Button(main, text="Запустить программу", font=("Arial", 14), command= run_script)
btn_start.pack(pady=50)


main.mainloop()