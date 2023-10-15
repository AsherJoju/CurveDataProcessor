from consolidator import consolidate
from extracter import Colors, extract


filenames = [
    "/home/asherjoju/Downloads/attachments/06-Jan-18 110922 AM(180106 AM).xlsx",
    "/home/asherjoju/Downloads/attachments/23-Dec-17 22458 PM (171223 PM).xlsx",
    "/home/asherjoju/Downloads/attachments/12162017 103539 AM ANLS.xlsx"
]
sheet_no = 1

tan_cell = "O1"
cv_cell = "O2"
curve_cell = "O3"

for file in filenames:
    print(file)
    
    extracted = extract(
        file,
        sheet_no,
        Colors(
            Colors.set_color(file, 1, tan_cell),
            Colors.set_color(file, 1, cv_cell),
            Colors.set_color(file, 1, curve_cell)
        )
    )
    
    consolidate(
        file,
        extracted
    )
