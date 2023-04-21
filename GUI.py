#GUI 
from tkinter import *
from tkinter.ttk import*

#from tkinter.ttk import Style
    

def create_new_canvas():
    root = Tk() #create object
    root.title('Team 4 Voice Assistant Demo')
   # Set window size
    root.geometry('400x400')
    root.configure(background='white')

    style = Style()
    style.configure('H.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'green')

    button1 = Button(root, text = 'How can I help you?', command = None , style='H.TButton')
    button1.grid(row = 5, column = 3, pady = 10, padx = 100)
    button1.place(relx=0.5, rely=0.3, anchor=CENTER)

    style.configure('st.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'orange')

    button2 = Button(root, text = 'STOP', command = None, style='st.TButton')
    button2.grid(row = 6, column = 1, pady = 10, padx = 100)
    button2.place(relx=0.5, rely=0.5, anchor=CENTER)

    style.configure('cl.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'red')

    button3 = Button(root, text = 'CLOSE', command = create_out_canvas,  style='cl.TButton')
    button3.grid(row = 6, column = 6, pady = 10, padx = 100)
    button3.place(relx=0.5, rely=0.7, anchor=CENTER)

    label_font = ('Arial Black', 12)
    label1 = Label(root, text='Virtual Assistant : Jarvis', font = label_font)
    label2 = Label(root, text='TESTING', font = label_font)
    label1.pack()
    label2.pack()

   # root.mainloop()

def create_out_canvas():
    root = Tk() #create object
    root.title('Team 4 Voice Assistant Demo')
   # Set window size
    root.geometry('400x400')
    root.configure(background='white')

    style = Style()
    style.configure('H.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'green')

    # button1 = Button(root, text = 'How can I help you?', command = None , style='H.TButton')
    # button1.grid(row = 5, column = 3, pady = 10, padx = 100)
    # button1.place(relx=0.5, rely=0.3, anchor=CENTER)

    # style.configure('st.TButton', font =
    #             ('calibri', 10, 'bold', 'underline'),
    #                 foreground = 'orange')

    # button2 = Button(root, text = 'STOP', command = None, style='st.TButton')
    # button2.grid(row = 6, column = 1, pady = 10, padx = 100)
    # button2.place(relx=0.5, rely=0.5, anchor=CENTER)

    style.configure('cl.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'red')

    button3 = Button(root, text = 'CLOSE', command = None,  style='cl.TButton')
    button3.grid(row = 6, column = 6, pady = 10, padx = 100)
    button3.place(relx=0.5, rely=0.7, anchor=CENTER)

    label_font = ('Arial Black', 12)
    label1 = Label(root, text='Virtual Assistant : Jarvis', font = label_font)
    label2 = Label(root, text='OUT', font = label_font)
    label1.pack()
    label2.pack()




def create_gui():
    root = Tk() #create object
    root.title('Team 4 Voice Assistant Demo')
   # Set window size
    root.geometry('400x400')
    root.configure(background='white')

   

    style = Style()
    style.configure('W.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'green')

    button1 = Button(root, text = 'START', command = create_new_canvas , style='W.TButton')
    button1.grid(row = 5, column = 3, pady = 10, padx = 100)
    button1.place(relx=0.5, rely=0.3, anchor=CENTER)

    style.configure('S.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'orange')

    button2 = Button(root, text = 'STOP', command = None, style='S.TButton')
    button2.grid(row = 6, column = 3, pady = 10, padx = 100)
    button2.place(relx=0.5, rely=0.5, anchor=CENTER)

    style.configure('C.TButton', font =
                ('calibri', 10, 'bold', 'underline'),
                    foreground = 'red')

    button3 = Button(root, text = 'CLOSE', command = create_out_canvas,  style='C.TButton')
    button3.grid(row = 7, column = 3, pady = 10, padx = 100)
    button3.place(relx=0.5, rely=0.7, anchor=CENTER)

    label_font = ('Arial Black', 12)
    label1 = Label(root, text='Virtual Assistant : Jarvis', font = label_font)
    label2 = Label(root, text='Welcome, How can I help you?', font = label_font)
    label1.pack()
    label2.pack()

    root.mainloop()

if __name__ == '__main__':
    create_gui()

