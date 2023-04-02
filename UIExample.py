#!/usr/local/bin/python3

import tkinter as tk

def okPressed():
    print("Name: " + name.get())
    print("Name2: " + nameTwo.get())
    
window = tk.Tk()
window.title("Downsman Results Generator")

# Create a frame for the text entry box
content = tk.Frame(window)

namelbl = tk.Label(content, text="Name")
name = tk.Entry(content)

nameTwolbl = tk.Label(content, text="Name2")
nameTwo= tk.Entry(content)


ok = tk.Button(content, text="Okay", command=okPressed)

cancel = tk.Button(content, text="Cancel", command=window.destroy)

content.grid(column=0, row=0)
namelbl.grid(column=0, row=0, columnspan=1)
name.grid(column=1, row=0, columnspan=2)
nameTwolbl.grid(column=0, row=1, columnspan=1)
nameTwo.grid(column=1, row=1, columnspan=2)


ok.grid(column=1, row=3)
cancel.grid(column=2, row=3)

window.mainloop()



