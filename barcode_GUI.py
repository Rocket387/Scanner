import openpyxl
import xlsxwriter

#creates excel sheey
workbook = xlsxwriter.Workbook('BarcodeScans.xlsx')
worksheet = workbook.add_worksheet()

#header
worksheet.write('A1', 'Restore_file_barcode')

#GUI window
from tkinter import *
root = Tk()
root.geometry("700x700")
root.title("Barcode Scanner")


my_entries = []

#Row loop
for y in range(10):
    #column loop
    for x in range(5):
        my_entry = Entry(root)
        my_entry.grid(row=y, column=x, pady= 20, padx = 5)
        my_entries.append(my_entry)

#writes to excel file in python program folder
def writetofile():
    wb = openpyxl.load_workbook(filename='BarcodeScans.xlsx')
    row = 1
    col = 0
    entry_list = ''

    for entries in my_entries:
        entry_list = entry_list + str(entries.get())
        worksheet.write(row, col, entries.get())
        row += 1

        #label says when items have been added
        my_label.config(text='Items added')

#button adds data to excel file
applyButton = Button(root, text="Add Data", command=writetofile)
applyButton.grid(row=10, column=1)

my_label = Label(root, text='')
my_label.grid(row=10, column=0, pady=20)

root.mainloop()
workbook.close()

