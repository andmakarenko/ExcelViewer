from tkinter import *
from tkinter import messagebox
from tkinter import ttk, filedialog
import numpy
import pandas as pd

root = Tk()
root.title("Excel Datasheet Viewer")
root.geometry('1100x400+200+200')

def Open():
    filename = filedialog.askopenfilename(title = "Open a file",
                                          filetypes = (("xlsx Files", ".xlsx"), ("All Files", ".*")))
    
    if filename:
        try:
            filename = r"{}".format(filename)
            df = pd.read_excel(filename)
        except:
            messagebox.showerror("Error", "File unaccessable!")

    tree.delete(*tree.get_children())
    tree['column'] = list(df.columns)
    tree['show'] = 'headings'

    for col in tree["column"]:
        tree.heading(col, text = col)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values = row)

root.iconphoto(False, PhotoImage(file = "logo.png"))

frame = Frame(root)
frame.pack(pady = 25)

tree = ttk.Treeview(frame)
tree.pack()

button = Button(root, text = 'Open', width = 60, height = 2, font = 30, fg = 'white', bg = '#0078d7', command = Open)
button.pack(padx = 10, pady = 20)


root.mainloop()