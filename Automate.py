import os
import tkinter as tk
from tkinter import font as tkfont, filedialog, messagebox
from tkinter import *

import pandas as pd


class SampleApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        x2 = 400
        y2 = 400
        screenw = self.winfo_screenwidth()
        screenh = self.winfo_screenheight()
        x = (screenw / 2) - (x2 / 2)
        y = (screenh / 2) - (y2 / 2)

        self.geometry(f'{x2}x{y2}+{int(x)}+{int(y)}')
        self.title('Script Automation')

        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.title_font = tkfont.Font(family='Helvetica', size=40, weight="bold", slant="italic")
        self.resizable(False, False)

        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        self.container = tk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (StartPage, Page2, Page3, Page4, Page5, ThankPage):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def show_frame(self, page_name):
        """Show a frame for the given page name"""
        frame = self.frames[page_name]
        frame.tkraise()

    def on_exit(self):
        """When you click to exit, this function is called"""
        if messagebox.askyesno("Exit", "Do you want to quit the application?"):
            app.destroy()

    def get_page(self, page_class):

        return self.frames[page_class]


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        global full
        self.full = StringVar(value="off")

        self.main = Label(self, text="Welcome to Script Automation", font=("Arial", 15))
        self.main.place(x=75, y=50)

        self.main1 = Label(self, text="Select the method of creation", font=("Arial", 14))
        self.main1.place(x=90, y=100)

        self.a = Radiobutton(self, text='Enter Manually', value=1, variable=self.full, background='#E7EBEC')
        self.a.place(x=100, y=150)

        self.b = Radiobutton(self, text='Select an excel containing Keywords', value=2, variable=self.full,
                             background='#E7EBEC')
        self.b.place(x=100, y=180)

        self.btn1 = Button(self, text='Next', command=lambda: [self.select()])
        self.btn1.place(x=190, y=300)

    def select(self):
        full = self.full.get()
        print(full)
        if full == "1":
            self.controller.show_frame("Page2")
        if full == "2":
            self.controller.show_frame("Page4")


class Page2(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.c2 = Label(self, text="Enter the Keyword", font=("Arial", 15))
        self.c2.place(x=125, y=50)

        self.t1 = Entry(self, width=28, font=("Arial", 15))
        self.t1.place(x=50, y=100, height=30)

        self.btnr = Button(self, text='Next', command=self.two)
        self.btnr.place(x=190, y=300)

    def two(self):
        global inp
        inp = self.t1.get()
        if len(self.t1.get()) == 0:
            messagebox.showerror("Error", "Input the Keyword")
        else:
            a = self.controller.get_page("Page3")
            a.chexkbox()
            self.controller.show_frame("Page3")


class Page3(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        self.cb8 = None
        self.cb7 = None
        self.cb6 = None
        self.cb5 = None
        self.cb4 = None
        self.cb3 = None
        self.cb2 = None
        self.cb = None
        self.cb1 = None
        self.controller = controller

        self.c1 = Label(self, text="Select the Jobsites that you want to create")
        self.c1.place(x=75, y=50)

        self.btn3 = Button(self, text='Create', command=self.angle)
        self.btn3.place(x=190, y=300)
        # self.btn2 = Button(self, text='Project', command=self.projects)
        # self.btn2.place(x=700, y=300)

    def exit(self):
        app.destroy()

    def on_exit(self):
        """When you click to exit, this function is called"""
        if messagebox.askyesno("Exit", "Do you want to quit the application?"):
            app.destroy()

    def chexkbox(self):
        global Allproj, Dice, Monster, Google, SimplyHired, Juju, Googlesearch,Indeed,Career

        self.cb = Checkbutton(self, text='All', variable=Allproj, onvalue='on', offvalue='off')
        self.cb.place(x=100, y=80)  # create new item in vars array
        self.cb1 = Checkbutton(self, text='Dice', variable=Dice, onvalue='on', offvalue='off')
        self.cb1.place(x=100, y=110)
        self.cb2 = Checkbutton(self, text='Monster', variable=Monster, onvalue='on', offvalue='off')
        self.cb2.place(x=100, y=140)
        self.cb3 = Checkbutton(self, text='Google', variable=Google, onvalue='on', offvalue='off')
        self.cb3.place(x=100, y=170)
        self.cb4 = Checkbutton(self, text='SimplyHired', variable=SimplyHired, onvalue='on', offvalue='off')
        self.cb4.place(x=100, y=200)
        self.cb5 = Checkbutton(self, text='Juju', variable=Juju, onvalue='on', offvalue='off')
        self.cb5.place(x=100, y=230)
        self.cb6 = Checkbutton(self, text='GoogleSearch', variable=Googlesearch, onvalue='on', offvalue='off')
        self.cb6.place(x=100, y=260)
        self.cb7 = Checkbutton(self, text='Indeed', variable=Indeed, onvalue='on', offvalue='off')
        self.cb7.place(x=100, y=290)
        self.cb8 = Checkbutton(self, text='Career', variable=Career, onvalue='on', offvalue='off')
        self.cb8.place(x=100, y=320)

    def angle(self):
        global selection, Allproj, inp
        selection = []
        if Allproj.get() == "on":
            selection.append(self.cb1.cget('text'))
            selection.append(self.cb2.cget("text"))
            selection.append(self.cb3.cget("text"))
            selection.append(self.cb4.cget("text"))
            selection.append(self.cb5.cget("text"))
            selection.append(self.cb6.cget("text"))
            selection.append(self.cb7.cget("text"))
            selection.append(self.cb8.cget("text"))
        else:
            if Dice.get() == "on":
                selection.append(self.cb1.cget("text"))
            if Monster.get() == "on":
                selection.append(self.cb2.cget("text"))
            if Google.get() == "on":
                selection.append(self.cb3.cget("text"))
            if SimplyHired.get() == "on":
                selection.append(self.cb4.cget("text"))
            if Juju.get() == "on":
                selection.append(self.cb5.cget("text"))
            if Googlesearch.get() == "on":
                selection.append(self.cb6.cget("text"))
            if Indeed.get() == "on":
                selection.append(self.cb7.cget("text"))
            if Career.get() == "on":
                selection.append(self.cb8.cget("text"))
        lib = os.getcwd()
        if not selection:
            messagebox.showerror("Error", "Select the Website")
        else:
            if "Dice" in selection:
                with open(os.path.join(lib, "lib\Dice.py")) as f1:
                    text = f1.read()
                n1 = str(inp)
                n1 = n1.replace(" ", "%20")
                text = text.replace("Boomi", n1)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n11 = str(inp).replace(" ", "_")
                n11 = "Result\Simply\DD\Dice_" + n11
                d1 = os.path.join(lib, n11)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "Monster" in selection:
                with open(os.path.join(lib, 'lib\Monster.py')) as f1:
                    text = f1.read()
                n2 = str(inp)
                n2 = n2.replace(" ", "_")
                text = text.replace("Boomi", n2)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n22 = str(inp).replace(" ", "_")
                n22 = "Result\Simply\DD\Monster_" + n22
                d1 = os.path.join(lib, n22)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "Google" in selection:
                with open(os.path.join(lib, 'lib\Google.py')) as f1:
                    text = f1.read()
                n3 = str(inp)
                n3 = n3.replace(" ", "_")
                text = text.replace("Boomi", n3)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n33 = str(inp).replace(" ", "_")
                n33 = "Result\Simply\DD\Google_Jobs_" + n33
                d1 = os.path.join(lib, n33)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "SimplyHired" in selection:
                with open(os.path.join(lib, 'lib\SimplyHired.py')) as f1:
                    text = f1.read()
                n4 = str(inp)
                n4 = n4.replace(" ", "_")
                text = text.replace("Boomi", n4)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n44 = str(inp).replace(" ", "_")
                n44 = "Result\Simply\DD\SimplyHired_" + n44
                d1 = os.path.join(lib, n44)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "Juju" in selection:
                with open(os.path.join(lib, 'lib\Juju.py')) as f1:
                    text = f1.read()
                n5 = str(inp)
                n5 = n5.replace(" ", " ")
                text = text.replace("Boomi", n5)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n55 = str(inp).replace(" ", "_")
                n55 = "Result\Simply\DD\Juju_" + n55
                d1 = os.path.join(lib, n55)
                f1 = open(d1 + ".py", 'x')  
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "GoogleSearch" in selection:
                with open(os.path.join(lib, 'lib\GoogleSearch.py')) as f1:
                    text = f1.read()
                n5 = str(inp)
                n5 = n5.replace(" ", " ")
                text = text.replace("Boomi", n5)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n55 = str(inp).replace(" ", "_")
                n55 = "Result\Simply\DD\Googlesearch_" + n55
                d1 = os.path.join(lib, n55)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "Indeed" in selection:
                with open(os.path.join(lib, 'lib\Indeed.py')) as f1:
                    text = f1.read()
                n5 = str(inp)
                n5 = n5.replace(" ", " ")
                text = text.replace("Boomi", n5)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n55 = str(inp).replace(" ", "_")
                n55 = "Result\Simply\DD\Indeed_" + n55
                d1 = os.path.join(lib, n55)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            if "Career" in selection:
                with open(os.path.join(lib, 'lib\Career.py')) as f1:
                    text = f1.read()
                n5 = str(inp)
                n5 = n5.replace(" ", " ")
                text = text.replace("Boomi", n5)
                text = text.replace("boomi.xlsx", str(inp) + ".xlsx")
                n55 = str(inp).replace(" ", "_")
                n55 = "Result\Simply\DD\Careerbuilder_" + n55
                d1 = os.path.join(lib, n55)
                f1 = open(d1 + ".py", 'x')
                with open(d1 + '.py', 'w') as f:
                    f.write(text)
            self.controller.show_frame("ThankPage")


class Page4(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.c2 = Label(self, text="Select the Excel File", font=("Arial", 15))
        self.c2.place(x=125, y=50)
        self.btnw = Button(self, text='Browse', command=lambda: [self.browse_button2()])
        self.btnw.place(x=190, y=150)

        self.t1 = Entry(self, width=50)
        self.t1.place(x=50, y=100, height=30)

        self.btnr = Button(self, text='Next', command=self.two)
        self.btnr.place(x=190, y=360)

    def browse_button2(self):
        global outdir
        self.t1.delete(0, 'end')
        outdir = filedialog.askopenfile(filetypes=[('Excel File', ".xlsx")])
        print(outdir.name)
        self.t1.insert(END, str(outdir.name))

    def two(self):
        global inp1
        inp1 = self.t1.get()
        if "xlsx" not in inp1:
            messagebox.showerror("Error", "Select the Excel File")
        else:
            a = self.controller.get_page("Page5")
            a.chexkbox()
            self.controller.show_frame("Page5")


class Page5(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        self.cb8 = None
        self.cb7 = None
        self.cb6 = None
        self.cb5 = None
        self.cb4 = None
        self.cb3 = None
        self.cb2 = None
        self.cb = None
        self.cb1 = None
        self.controller = controller

        self.c1 = Label(self, text="Select the Jobsites that you want to create")
        self.c1.place(x=75, y=50)

        self.btn3 = Button(self, text='Create', command=self.angle)
        self.btn3.place(x=190, y=340)
        # self.btn2 = Button(self, text='Project', command=self.projects)
        # self.btn2.place(x=700, y=300)

    def exit(self):
        app.destroy()

    def on_exit(self):
        """When you click to exit, this function is called"""
        if messagebox.askyesno("Exit", "Do you want to quit the application?"):
            app.destroy()

    def chexkbox(self):
        global Allproj, Dice, Monster, Google, SimplyHired, Juju, Googlesearch,Indeed,Career

        self.cb = Checkbutton(self, text='All', variable=Allproj, onvalue='on', offvalue='off')
        self.cb.place(x=100, y=80)  # create new item in vars array
        self.cb1 = Checkbutton(self, text='Dice', variable=Dice, onvalue='on', offvalue='off')
        self.cb1.place(x=100, y=110)
        self.cb2 = Checkbutton(self, text='Monster', variable=Monster, onvalue='on', offvalue='off')
        self.cb2.place(x=100, y=140)
        self.cb3 = Checkbutton(self, text='Google', variable=Google, onvalue='on', offvalue='off')
        self.cb3.place(x=100, y=170)
        self.cb4 = Checkbutton(self, text='SimplyHired', variable=SimplyHired, onvalue='on', offvalue='off')
        self.cb4.place(x=100, y=200)
        self.cb5 = Checkbutton(self, text='Juju', variable=Juju, onvalue='on', offvalue='off')
        self.cb5.place(x=100, y=230)
        self.cb6 = Checkbutton(self, text='GoogleSearch', variable=Googlesearch, onvalue='on', offvalue='off')
        self.cb6.place(x=100, y=260)
        self.cb7 = Checkbutton(self, text='Indeed', variable=Indeed, onvalue='on', offvalue='off')
        self.cb7.place(x=100, y=290)
        self.cb8 = Checkbutton(self, text='Career', variable=Career, onvalue='on', offvalue='off')
        self.cb8.place(x=100, y=320)

    def angle(self):
        global selection, Allproj, inp1
        selection = []
        if Allproj.get() == "on":
            selection.append(self.cb1.cget('text'))
            selection.append(self.cb2.cget("text"))
            selection.append(self.cb3.cget("text"))
            selection.append(self.cb4.cget("text"))
            selection.append(self.cb5.cget("text"))
            selection.append(self.cb6.cget("text"))
            selection.append(self.cb7.cget("text"))
            selection.append(self.cb8.cget("text"))
        else:
            if Dice.get() == "on":
                selection.append(self.cb1.cget("text"))
            if Monster.get() == "on":
                selection.append(self.cb2.cget("text"))
            if Google.get() == "on":
                selection.append(self.cb3.cget("text"))
            if SimplyHired.get() == "on":
                selection.append(self.cb4.cget("text"))
            if Juju.get() == "on":
                selection.append(self.cb5.cget("text"))
            if Googlesearch.get() == "on":
                selection.append(self.cb6.cget("text"))
            if Indeed.get() == "on":
                selection.append(self.cb7.cget("text"))
            if Career.get() == "on":
                selection.append(self.cb8.cget("text"))
        lib = os.getcwd()
        if not selection:
            messagebox.showerror("Error", "Select the Website")
        else:
            df = pd.read_excel(inp1)
            print(len(df))
            print(df)
            title = (df.columns[0])
            for i in range(len(df)):
                if "Dice" in selection:
                    with open(os.path.join(lib, "lib\Dice.py")) as f1:
                        text = f1.read()
                    n1 = str(df[title][i])
                    n1 = n1.replace(" ", "%20")
                    text = text.replace("Boomi", n1)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n11 = str(df[title][i]).replace(" ", "_")
                    n11 = "Result\Simply\DD\Dice_" + n11
                    d1 = os.path.join(lib, n11)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "Monster" in selection:
                    with open(os.path.join(lib, 'lib\Monster.py')) as f1:
                        text = f1.read()
                    n2 = str(df[title][i])
                    n2 = n2.replace(" ", "_")
                    text = text.replace("Boomi", n2)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n22 = str(df[title][i]).replace(" ", "_")
                    n22 = "Result\Simply\DD\Monster_" + n22
                    d1 = os.path.join(lib, n22)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "Google" in selection:
                    with open(os.path.join(lib, 'lib\Google.py')) as f1:
                        text = f1.read()
                    n3 = str(df[title][i])
                    n3 = n3.replace(" ", "_")
                    text = text.replace("Boomi", n3)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n33 = str(df[title][i]).replace(" ", "_")
                    n33 = "Result\Simply\DD\Google_Jobs_" + n33
                    d1 = os.path.join(lib, n33)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "SimplyHired" in selection:
                    with open(os.path.join(lib, 'lib\SimplyHired.py')) as f1:
                        text = f1.read()
                    n4 = str(df[title][i])
                    n4 = n4.replace(" ", "_")
                    text = text.replace("Boomi", n4)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n44 = str(df[title][i]).replace(" ", "_")
                    n44 = "Result\Simply\DD\SimplyHired_" + n44
                    d1 = os.path.join(lib, n44)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "Juju" in selection:
                    with open(os.path.join(lib, 'lib\Juju.py')) as f1:
                        text = f1.read()
                    n5 = str(df[title][i])
                    n5 = n5.replace(" ", " ")
                    text = text.replace("Boomi", n5)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n55 = str(df[title][i]).replace(" ", "_")
                    n55 = "Result\Simply\DD\Juju_" + n55
                    d1 = os.path.join(lib, n55)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "GoogleSearch" in selection:
                    with open(os.path.join(lib, 'lib\GoogleSearch.py')) as f1:
                        text = f1.read()
                    n5 = str(df[title][i])
                    n5 = n5.replace(" ", " ")
                    text = text.replace("Boomi", n5)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n55 = str(df[title][i]).replace(" ", "_")
                    n55 = "Result\Simply\DD\Googlesearch_" + n55
                    d1 = os.path.join(lib, n55)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "Indeed" in selection:
                    with open(os.path.join(lib, 'lib\Indeed.py')) as f1:
                        text = f1.read()
                    n5 = str(df[title][i])
                    n5 = n5.replace(" ", " ")
                    text = text.replace("Boomi", n5)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n55 = str(df[title][i]).replace(" ", "_")
                    n55 = "Result\Simply\DD\Indeed_" + n55
                    d1 = os.path.join(lib, n55)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
                if "Career" in selection:
                    with open(os.path.join(lib, 'lib\Career.py')) as f1:
                        text = f1.read()
                    n5 = str(df[title][i])
                    n5 = n5.replace(" ", " ")
                    text = text.replace("Boomi", n5)
                    text = text.replace("boomi.xlsx", str(df[title][i]) + ".xlsx")
                    n55 = str(df[title][i]).replace(" ", "_")
                    n55 = "Result\Simply\DD\Careerbuilder_" + n55
                    d1 = os.path.join(lib, n55)
                    f1 = open(d1 + ".py", 'x')
                    with open(d1 + '.py', 'w') as f:
                        f.write(text)
            self.controller.show_frame("ThankPage")


class ThankPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.btn3 = Label(self, text='Thanks for using the app', font=("Arial", 15))
        self.btn3.place(x=75, y=50)

        self.btn4 = Button(self, text='Finish', command=self.exit)
        self.btn4.place(x=190, y=300)

    def exit(self):
        app.destroy()


if __name__ == "__main__":
    app = SampleApp()
    inp = StringVar()
    inp1 = StringVar()
    selection = StringVar()
    Allproj = StringVar(value='off')
    Dice = StringVar(value='off')
    Monster = StringVar(value='off')
    Google = StringVar(value='off')
    SimplyHired = StringVar(value='off')
    Juju = StringVar(value='off')
    Googlesearch = StringVar(value='off')
    Indeed = StringVar(value='off')
    Career = StringVar(value='off')
    full = StringVar(value="off")
    app.mainloop()
