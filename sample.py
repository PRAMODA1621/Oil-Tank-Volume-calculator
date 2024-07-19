import tkinter
import subprocess
import os
from tkinter import *
from tkinter.messagebox import showinfo
from tkinter.simpledialog import askstring
from tkinter import ttk, messagebox, filedialog
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import openpyxl
import tkinter as tk

root = Tk()
root.geometry('1100x950')
root.title('Oil Tank')
root.resizable(0, 0)
root.configure(bg='yellow')
frame = tkinter.Frame(bg='Yellow')

# Excel
filename = "data.xlsx"
wb = openpyxl.Workbook()
sheet = wb.active
t = "API2"

global k, k1, i, i1, int_value, r, rdv,z1,a,b,c,o,pro,finale,l1,b1,c1,u,v,y,q,s,sv
k = ""
k1 = ""
i = ""
i1 = ""
int_value = 0
r = 0
rdv = 0
rdv1=0
z1=None
l=0.0
b=0.0
h=0.0
o=0
pro=0.0
finale=0
l1=0
b1=0
c1=0
u=""
v=""
y=""
q=""
s=""
sv=""
def button(t):
    global k
    k = t
    return k


t1 = "API3"


def button1(t1):
    global k1
    k1 = t1
    return k1


# OilFlow
def convert():
    try:
        sv = float(OF1.get())
        sv1 = sv * 3.785
        h1 = sv1
        OF.delete(0, END)  # Clear the current value in the entry widget
        OF.insert(0, str(h1))  # Insert the converted value into the entry widget
        print(sv1)
    except:
        print("Invalid")
print(sv)
OF_label = tkinter.Label(frame, text='Oil Flow', bg='white', fg='black')
OF = tkinter.Entry(frame, font=('Ariel', 10))
OF1 = tkinter.Entry(frame, font=('Ariel', 10))
OF2=Button(frame,text="Convert",command=convert).grid(row=0,column=3,padx=20,pady=10)

# Design
v1 = BooleanVar()
v2 = BooleanVar()


def calc():
    global i, r, int_value
    conv_value=float(OF.get())
    str_value = conv_value
    try:
        int_value = int(str_value)
        i = int_value * 8
        r = int_value
        print(i)
    except ValueError:
        print("Invalid Input")


def calc1():
    global i1
    str_val = OF.get()
    try:
        int_val = int(str_val)
        i1 = int_val * 5
        print(i1)
    except ValueError:
        print("Invalid Input")

def button_click2():
    D.configure(bg='red',font="bold")
def button_click3():
    D1.configure(bg='red',font="bold")
D_label = tkinter.Label(frame, text='Design Std', bg='white', fg='black')
D = Button(frame, text='API2', command=lambda: (calc(), button(t),button_click2()))
D1 = Button(frame, text='API3', command=lambda: (calc1(), button1(t1),button_click3()))



def calc2():
    global z1
    z = tkinter.Label(frame, text="Volume:", bg='white', fg='black')
    z1 = Entry(frame, font=('Ariel', 10))
    z.grid(row=2, column=3)
    z1.grid(row=2, column=4)

rdv_text = tk.StringVar()
rdv1_text = tk.StringVar()
def volcal():
    global rdv
    s_value = z1.get()
    try:
        print(r)
        in_val = int(s_value)
        f = 2 * r
        f1 = (10 / 100) * r
        rdv = f + f1 + in_val
        print(rdv)
        print(f)
        print(f1)
        rdv_text.set(str(f + f1 + in_val))
    except ValueError:
        print("Invalid operation")


def volcal1():
    global rdv1
    try:
        print(r)
        f3 = 2 * r
        f2 = (10 / 100) * r
        rdv1 = f3 + f2 + 0
        rdv1_text.set(str(f3 + f2 + 0))
    except ValueError:
        print("Invalid operation")

    print(rdv1)
def button_click():
    RT.configure(bg='red')
def button_click1():
    RT1.configure(bg='red')
RT_label = tkinter.Label(frame, text='Rundown tank', bg='white', fg='black')
RT = tk.Button(frame, text='YES',activeforeground="Red", command=lambda: (calc2(),button_click()))
RT1 = tk.Button(frame, text='NO',activeforeground="Red", command=lambda:(volcal1(),button_click1()))
RT2 = tk.Button(frame, text="Calculate", command=lambda: (volcal()))
RT2.grid(row=2, column=5, padx=20, pady=10)

var1 = IntVar()
var2 = IntVar()
var3=  IntVar()

def checkbuttonclick():
    global u, v, y
    if var1.get() == 1:
        u = "SS304"
        sheet["H2"] = "SS304"
    else:
        u = ""
    if var2.get() == 1:
        v = "SS316"
        sheet["H2"] = "SS316"
    else:
        v = ""
    if var3.get() == 1:
        y = "SS316L"
        sheet["H2"] = "SS316L"
    else:
        y = ""


# material
m_label = tkinter.Label(frame, text="Material", bg='white', fg='black')
m1 = Checkbutton(frame, text='SS304', onvalue=1, offvalue=0,variable=var1,command=checkbuttonclick)
m2 = Checkbutton(frame, text='SS316', onvalue=1, offvalue=0,variable=var2,command=checkbuttonclick)
m3 = Checkbutton(frame, text='SS316L', onvalue=1, offvalue=0,variable=var3,command=checkbuttonclick)

va1=IntVar()
va2=IntVar()
def check():
    global o,q,s
    if va1.get()==1:
        q="YES"
        sheet["J2"]="YES"
        o=i+rdv
    else:
        q=""
    if va2.get()==1:
        s="NO"
        sheet["K2"]="NO"
    else:
        s=""
# components
c_label = tkinter.Label(frame, text="Are components mounted on tank", bg='white', fg='black')
c1 = Checkbutton(frame, text='YES', onvalue=1, offvalue=0,variable=va1,command=check)
c2 = Checkbutton(frame, text='NO', onvalue=1, offvalue=0,variable=va2,command=check)


# Size
size_label = tkinter.Label(frame, text="Size of tank:(l,b,h)/Standard", bg="white", fg="black")
size = Entry(frame, font=('Ariel', 10), width=2)
s = tkinter.Label(frame, text="Length", bg='white', fg='black')
size1 = Entry(frame, font=('Ariel', 10), width=2)
s1 = tkinter.Label(frame, text="Breadth", bg='white', fg='black')
size2 = Entry(frame, font=('Ariel', 10), width=2)
s2 = tkinter.Label(frame, text="Height", bg="white", fg="black")
def ch():
    global pro
    if int_value<=100:
        l=2.5
        b=2
        h=0.5
        pro=l*b*h
        print(pro)
    elif int_value > 101 :
        l = 3
        b = 2.5
        h = 0.75
        pro=l*b*h
        print(pro)


def user():
    global finale
    l1 = size.get()
    b1 = size1.get()
    c1 = size2.get()

    if l1 is not None and b1 is not None and c1 is not None:
        l1 = int(l1)
        b1 = int(b1)
        c1 = int(c1)

        finale = l1 * b1 * c1
        print(finale)
def button_click5():
    size3.configure(bg='red',font="bold")
size3 = Button(frame, text="Standard",command=lambda:(ch(),button_click5()))
def button_click6():
    size4.configure(bg='red',font="bold")
size4=Button(frame,text="Calculate",command=lambda:(user(),button_click6()))

def checking():
    if pro >= o or finale >= o:
        messagebox.showerror("ERROR", "PLEASE RECHECK THE GIVEN MEASUREMENTS ARE NOT COMPATIBLE")
        sheet["M2"]="Not Compatible"
        root.destroy()
        frame.destroy()
def finish():
    if pro<o :
        sheet["M2"]="Compatible"
# button

# positions
OF_label.grid(row=0, column=0)
OF.grid(row=0, column=1, padx=20, pady=10)
OF1.grid(row=0, column=2, padx=20, pady=10)

D_label.grid(row=1, column=0)
D.grid(row=1, column=1, padx=20, pady=10)
D1.grid(row=1, column=2, padx=20, pady=10)

RT_label.grid(row=2, column=0)
RT.grid(row=2, column=1, padx=20, pady=10)
RT1.grid(row=2, column=2, padx=20, pady=10)

m_label.grid(row=3, column=0)
m1.grid(row=3, column=1, padx=20, pady=10)
m2.grid(row=3, column=2, padx=20, pady=10)
m3.grid(row=3, column=3, padx=20, pady=10)

c_label.grid(row=4, column=0)
c1.grid(row=4, column=1, padx=20, pady=10)
c2.grid(row=4, column=2, padx=20, pady=10)

size_label.grid(row=5, column=0)
size.grid(row=5, column=2, padx=20, pady=10)
s.grid(row=5, column=1)
size1.grid(row=6, column=2, padx=20, pady=10)
s1.grid(row=6, column=1)
size2.grid(row=7, column=2, padx=20, pady=10)
s2.grid(row=7, column=1)
size3.grid(row=5, column=3, padx=20, pady=10)
size4.grid(row=9,column=1,padx=20,pady=10)

# Calculating oil reservoir volume

filename = "data.xlsx"
sheet_name = "Sheet 1"
border = Border(
    left=Side(border_style="thick"),
    right=Side(border_style="thick"),
    top=Side(border_style="thick"),
    bottom=Side(border_style="thick"),
)
for row in sheet["A1:N3"]:
    for cell in row:
        cell.border = border


def save():
    sheet["A1"] = "Ret Capacity"
    sheet["A2"] = i
    sheet["A3"] = i1
    sheet["C1"] = "Design Std"
    sheet["C2"] = k
    sheet["C3"] = k1
    sheet["E1"] = "Rundown Volume"
    sheet["E2"]="YES"
    sheet["F2"]="NO"
    sheet["E3"] = rdv
    sheet["F3"]=rdv1
    sheet["H1"]="Material"
    checkbuttonclick()
    sheet["J1"]="Component mounted"
    check()
    sheet["M1"]="Compatibility"
    checking()
    finish()
    wb.save(filename)


def output():
    output_label = tk.Label(frame, text="", font=('Arial', 15))
    output_label.grid(row=12, column=0, padx=20, pady=10)

    text = "Oil tank std: " + str(k) + " " + str(k1) + "\nretention capacity: " + str(i) + " " + str(i1) + "\nApprox size = " + str(pro) + " " + str(finale) + "\nMaterial: " + str(u) + str(v) + str(y) + "\nComponents top mounted: " + str(q) + str(s)+"\nRundown Tank Volume: "+str(rdv)+""+str(rdv1)
    output_label.config(text=text)


def open_excel():
    # Open file dialog to select Excel file
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx")])

    if filepath:
        try:
            # Open the selected Excel file with the default program
            subprocess.Popen(['start', '', filepath], shell=True)
        except Exception as e:
            print("Error:", str(e))

def button_click4():
    b.configure(bg='red',font="bold")


b = Button(frame, text="Submit", command=lambda:(save(),checking(),output(),button_click4()))
b.grid(row=9, column=0)
def button_click7():
    oie.configure(bg='red',font="bold")
oie=Button(frame,text="Open in excel",command=lambda:(open_excel(),button_click7()))
oie.grid(row=12,column=3,padx=20,pady=10)
def exit1():
    frame.destroy()
    root.destroy()
exit=Button(frame,text="Exit",command=exit1)
exit.grid(row=13,column=0,padx=20,pady=10)
frame.pack(pady=20, side=TOP, anchor="nw", expand=True)
root.mainloop()