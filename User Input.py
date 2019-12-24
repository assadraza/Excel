import tkinter
from tkinter import *
import openpyxl
from openpyxl.styles import Font
import xlsxwriter

# print ('%s%s Hello World !!! %s' % (fg(11), bg(7), attr(0)))

# Creating Windows
window = Tk()
window.title("User Input")
window.geometry("1080x720")
window.configure(background="silver")

# canvas = Canvas(window, width=100, height=100, bg="blue")
# canvas.pack()

# topFrame = Frame(window)
# topFrame.pack ()
# bottomFrame = Frame(window)
# bottomFrame.pack(side=BOTTOM)
#
#
# button1 = Button(topFrame, text="Button1", fg=("red"))
#
# button1.pack(side=LEFT)


# frame1 = LabelFrame(window, text="Please Enter VLAN ID", padx=5,pady=5)
# entry = Entry(frame1)
# frame1.pack()
#

# # # Creating Labels
#
# entry1 = Label(window, width=20, height=1,text="  VLAN ID: ", background='grey', foreground='Yellow')  # Creating Lable
# # entry1.grid(row=0, column=0)  # Aligning Lableone
# entry1.place(x=1, y=300)
#
#
# entry2 = Label(window, width=30, text="*Please enter VLAN ID only ", foreground='red', bg='silver')  # Creating Lable
# # entry2.grid(row=1, column=1)  # Aligning Labletwo
# entry2.place(x=170, y=323)
#
#
# # Creating Variable to Store User Input in Box1
# box1Input = tkinter.StringVar()
#
# # Creating User Input Box
# box1 = Entry(window, width=35, textvariable=box1Input)  # Creating Data Entry Box
# # box1.grid(row=0, column=1)  # Aligning Entry Box
# box1.place(x=160, y=301)
#
# # Defining Action
# def action():
#     entry2.configure(text="You have entered " + box1.get())
#
#
# # Creating Button
# btn1 = Button(window, text="Submit", command=action)
# btn1.grid(row=0, column=2)
# btn1.place(x=400, y=301)


##################################################################################


# ws2['A1'].value = 'Sata'
# ws2['A1'].font = Font(color = "FF0005", bold=4, size=13)
#
# ws2['B1'].value = 'Data2'
# ws2['B1'].font = Font(color = "FF0005", bold=4, size=13)


# Creating Variable to Store User Input in Box2
box2Input = tkinter.StringVar()

# Creating User Input Box
box2 = Entry(window, width=35, textvariable=box2Input)  # Creating Data Entry Box
# box2.grid(row=0, column=1)  # Aligning Entry Box
box2.place(x=100, y=200)


entry3 = Label(window, width=30, text="*Please enter data only ", foreground='red', bg='silver')  # Creating Lable
# entry2.grid(row=1, column=1)  # Aligning Labletwo
entry3.place(x=10, y=22)


# Defining Action

def write():
    # entry3.configure(text="You have entered " + box2.get())
    wb = openpyxl.load_workbook('data1.xlsx')
    print(wb.sheetnames)

    ws2 = wb['mysheet']
    ws2.sheet_properties.tabColor = "1072BA"

    ws2['A1'].value = "I Are"
    wb.save('data1.xlsx')

# Creating Button

btn2 = Button(window, text="Submit", command=write())
btn2.grid(row=0, column=2)
btn2.place(x=44, y=50)


window.mainloop()
