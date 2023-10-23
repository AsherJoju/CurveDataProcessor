from extracter import extract
from marker import mark
from threading import Thread
from tkinter import Button, Entry, Label, StringVar, Tk
from tkinter.filedialog import askopenfilename, askopenfilenames
from tkinter.messagebox import showerror, showinfo
from traceback import format_exc


root = Tk()
root.title("Curve Data Marker")
root.geometry("500x500")


files = []

input_filename = StringVar()
output_filenames = StringVar()
input_sheetname = StringVar()
output_sheetname = StringVar()


def set_files_list(list):
    for element in list:
        if element not in files:
            files.append(element)
    
    string = ""
    
    for file in files:
        string += " , " + file
    
    output_filenames.set(string[3:])


def submit():
    errors = ""
    
    try:
        for file in output_filenames.get().split(","):
            file = file.strip()
            print(file)
            try:
                extracted_data = extract(input_filename.get(), input_sheetname.get())
                
                mark(file, output_sheetname.get(), extracted_data)
            except:
                errors += file + " : " + format_exc() + "\n"
    except:
        errors = "Error In Input File Names"
    
    if errors == "":
        showinfo(
            "Marking Completed Successfully!",
            "The Marking has completed successfully without any internal errors !"
        )
    else:
        showerror(
            "Marking Failed With Errors!",
            errors
        )


def start():
    Thread(target=submit).start()
    
    showinfo(
        "Marking Has Started!",
        "The Program have began to mark data, " +
        "Please wait until it completes."
    )


Label(text="Enter Input File Name").pack(fill="both", expand=True)
Entry(textvariable=input_filename).pack(fill="both", expand=True)
Button(
    text="Open File Name", command=lambda: input_filename.set(
        askopenfilename(
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

Label(text="Enter Input Sheet Name").pack(fill="both", expand=True)
Entry(textvariable=input_sheetname).pack(fill="both", expand=True)


Label(text="Enter Output File Names").pack(fill="both", expand=True)
Entry(textvariable=output_filenames).pack(fill="both", expand=True)
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

Label(text="Enter Output Sheet Name").pack(fill="both", expand=True)
Entry(textvariable=output_sheetname).pack(fill="both", expand=True)

Button(text="Submit", command=start).pack(fill="both", expand=True)


root.mainloop()
