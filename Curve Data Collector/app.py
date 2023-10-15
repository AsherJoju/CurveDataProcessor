from tkinter import Tk, StringVar, Label, Entry, Button
from collector import collect
from threading import Thread
from tkinter.messagebox import showinfo, showerror
from traceback import format_exc
from tkinter.filedialog import askdirectory


root = Tk()
root.title("Curve Data Collector")
root.geometry("500x500")

directory = StringVar()
sheet = StringVar()

def submit():
    try:
        collect(directory.get(), sheet.get())
        showinfo(
            title="Collection Complete!",
            message="Collection is complete without any errors"
        )
    except Exception as e:
        showerror(
            title="An Error Occurred!",
            message=str(e) + "\n\nTraceback:\n" + format_exc()
        )

def start():
    Thread(target=submit).start()
    
    showinfo("Collection Has Started!", "Began to collect data")

Label(text="Enter Directory").pack(fill="both", expand=True)
Entry(textvariable=directory).pack(fill="both", expand=True)
Button(text="Open a Directory", command=lambda: directory.set(askdirectory(
    title="Open Directory",
    mustexist=True
))).pack(fill="both", expand=True)

Label(text="Enter Sheet Name").pack(fill="both", expand=True)
Entry(textvariable=sheet).pack(fill="both", expand=True)

Button(text="Collect Data", command=start).pack(fill="both", expand=True)

root.mainloop()
