import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


root = tk.Tk()
root.withdraw()

def open_files():
   path= filedialog.askopenfilename(title="Select a file", filetypes=(("Excel files","*.xlsx"),("all files","*.*")))
   return path

def showMessage(type_of_message):
    root.geometry("150x150")
    if type_of_message == "info":
        messagebox.showinfo("Importing Active Testing", "Active Testing Imported")
    elif type_of_message == "error":
        messagebox.showerror("Importing errors","Could not import Active Testing")
    elif type_of_message == "internet":
        messagebox.showerror("Connection errors","You are not connected to Internet")
    else:
        messagebox.showerror("Errors","Something went wrong")






if __name__ == "__main__":
    showMessage()