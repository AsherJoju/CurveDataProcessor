from tkinter import Tk, Label, Entry, Button, StringVar
from tkinter.messagebox import showerror, showinfo
from tkinter.filedialog import askopenfilenames
from threading import Thread
from traceback import format_exc
from consolidator import consolidate
from extracter import Colors, extract


root = Tk()
root.title("Curve Data Consolidator")
root.geometry("500x750")

files = []

filenames = StringVar()
sheet_no = StringVar()
color_cells = StringVar()


def set_files_list(list):
    for element in list:
        if element not in files:
            files.append(element)
    
    string = ""
    
    for file in files:
        string += " , " + file
    
    filenames.set(string[3:])


def submit():
    errors = ""
    
    try:
        for file in filenames.get().split(","):
            print(file)
            
            try:
                extracted = extract(
                    file,
                    sheet_no.get(),
                    Colors(
                        Colors.set_color(color_cells.get().split()[0]),
                        Colors.set_color(color_cells.get().split()[1]),
                        Colors.set_color(color_cells.get().split()[2])
                    )
                )
                
                consolidate(
                    file,
                    extracted
                )
            except:
                errors += file + " : " + format_exc() + "\n"
    except:
        errors = "Error In Input File Names"
    
    if errors == "":
        showinfo(
            "Consolidation Completed Successfully!",
            "The Consoldation has completed successfully without any internal errors !"
        )
    else:
        showerror(
            "Consolidation Complete With Errors!",
            errors
        )


def start():
    Thread(target=submit).start()
    
    showinfo("Consolidation Has Started!", "Began to consolidate data")


Label(text="Enter Input File Names").pack(fill="both", expand=True)
Entry(textvariable=filenames).pack(fill="both", expand=True)
Button(
    text="Open File Names", command=lambda: set_files_list(
        askopenfilenames(
            filetypes=(
                ("Excel Worksheets", "*.xlsx"),
                ("Excel Worksheets", "*.xltx"),
                ("Excel Worksheets", "*.xlsm"),
                ("Excel Worksheets", "*.xltm"),
                ("All Files", "*.*")
            )
        )
    )
).pack(fill="both", expand=True)

Label(text="Enter Input Sheet Index (1 based)").pack(fill="both", expand=True)
Entry(textvariable=sheet_no).pack(fill="both", expand=True)

Label(text="Enter Color Cells (Tangent CV Curve)").pack(fill="both", expand=True)
Entry(textvariable=color_cells).pack(fill="both", expand=True)


Button(text="Submit", command=start).pack(fill="both", expand=True)


root.mainloop()
