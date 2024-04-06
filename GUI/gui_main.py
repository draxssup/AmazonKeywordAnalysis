import tkinter as tk
root = tk.Tk()
root.geometry('500x250')
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
root.columnconfigure(1,weight=1)
tk.Label(root, text='Enter the name of the file without extension', font=('Arial',16)).grid(row=0,column=0)
title = tk.Entry(width=45)
title.grid(row=1,column=0,sticky='nw')
submitbtn = tk.Button

root.mainloop()
