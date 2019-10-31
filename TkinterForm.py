from openpyxl import *
from tkinter import *

wb = load_workbook('C:\\Users\\Nim_Ish\\Desktop\\Book1.xlsx')

sheet = wb.active


def excel():
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 40
    sheet.column_dimensions['D'].width = 50

    sheet.cell(row=1, column=1).value = "Name"
    sheet.cell(row=1, column=2).value = "Contact Nmber"
    sheet.cell(row=1, column=3).value = "Email id"
    sheet.cell(row=1, column=4).value = "Address"


def clear():
    name_field.delete(0, END)
    contact_no_field.delete(0, END)
    email_id_field.delete(0, END)
    address_field.delete(0, END)


def insert():
    if (name_field.get() == "" and
            contact_no_field.get() == "" and
            email_id_field.get() == "" and
            address_field.get() == ""):

        print("empty input")

    else:

        current_row = sheet.max_row
        current_column = sheet.max_column

        sheet.cell(row=current_row + 1, column=1).value = name_field.get()
        sheet.cell(row=current_row + 1, column=2).value = contact_no_field.get()
        sheet.cell(row=current_row + 1, column=3).value = email_id_field.get()
        sheet.cell(row=current_row + 1, column=4).value = address_field.get()

        wb.save('C:\\Users\\Nim_Ish\\Desktop\\Book1.xlsx')
        name_field.focus_set()
        clear()


if __name__ == "__main__":
    root = Tk()
    root.configure(background='grey')
    root.title("registration form")
    root.geometry("500x300")

    excel()
    heading = Label(root, text="Form", bg="grey")
    name = Label(root, text="Name", bg="grey")
    contact_no = Label(root, text="Contact No.", bg="grey")
    email_id = Label(root, text="Email id", bg="grey")
    address = Label(root, text="Address", bg="grey")

    heading.grid(row=0, column=1)
    name.grid(row=1, column=0)
    contact_no.grid(row=2, column=0)
    email_id.grid(row=3, column=0)
    address.grid(row=4, column=0)

    name_field = Entry(root)
    contact_no_field = Entry(root)
    email_id_field = Entry(root)
    address_field = Entry(root)

    name_field.grid(row=1, column=1, ipadx="100")
    contact_no_field.grid(row=2, column=1, ipadx="100")
    email_id_field.grid(row=3, column=1, ipadx="100")
    address_field.grid(row=4, column=1, ipadx="100")

    excel()

    submit = Button(root, text="Submit", fg="Black",
                    bg="Blue", command=insert)
    submit.grid(row=8, column=1)

    # start the GUI
    root.mainloop()
