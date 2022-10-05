from tkinter import *
from PIL import ImageTk, Image
from tkinter import ttk, messagebox
from tkcalendar import Calendar, DateEntry
from datetime import datetime
import smtplib

window = Tk()
window.resizable(False, False)
window.geometry("810x720+350+60")
font1 = ('Consolas  Bold', 18)
font2 = ('Consolas  Bold', 12)
font3 = ('Consolas', 10)
window.config(bg ='white')


blank = Label(window, text = "", borderwidth = 10)
blank.config(font = font2)

back = 'image_fav.png'
file = Image.open(back)
image = ImageTk.PhotoImage(file)
img = Label(window, image = image)
img.image = image
img.place(x = 10, y = 10)
img.config(bg = 'white')


def modif(replace_this, by_this):
    with open('modele_changeant_fav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
        contents = contents.replace(replace_this, by_this)
    with open(r'modele_changeant_fav.txt', 'w', encoding = 'utf-8') as file:
        file.write(contents)

def cb1():
    win1 = Toplevel(window)
    # win1.resizable(False, False)
    win1.geometry("155x33+700+80")
    gender = ['Madame', 'Monsieur']
    def madame():
        b1.config(bg = '#47B12F')
        b10.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        appellation = "Madame" 
        win1.destroy()
        modif("GENDER", appellation)
    choix = Button(win1, text = "Madame", command = madame)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2)

    def monsieur():
        b1.config(bg = '#47B12F')
        b10.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        appellation = "Monsieur" 
        win1.destroy()
        modif("GENDER", appellation)
    choix2 = Button(win1, text = "Monsieur", command = monsieur)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)


b1 = Button(window, text = "", width = 3, height = 1, command = cb1)
b1.place(x = 167, y = 20)
b1.config(bg = '#B50000')

def cb2():
    win2 = Toplevel(window)
    # win1.resizable(False, False)
    win2.geometry("265x53+700+80")
    def ok():
        b2.config(bg = '#47B12F')
        poste = choix.get() 
        win2.destroy()
        modif('POSTE', poste)
    choix = Entry(win2, width = 37)
    choix.focus_force()
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font3)
    save = Button(win2, text = 'OK', command = ok)
    save.grid(row = 2, column = 1 , sticky = W+E)
    save.config(font = font2)
b2 = Button(window, text = "", width = 4, height = 1, command = cb2)
b2.place(x = 303, y = 75)
b2.config(bg = '#B50000')
global contrat
def cb3():
    win3 = Toplevel(window)
    # win1.resizable(False, False)
    win3.geometry("145x99+700+80")
    def CDD():
        b3.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        contrat = "CDD" 
        win3.destroy()
        modif('CONTRAT+', contrat)
    choix = Button(win3, text = "CDD", command = CDD)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2) 

    def CDI():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "CDI"
        win3.destroy()
        modif('CONTRAT+', contrat)
    choix2 = Button(win3, text = "CDI", command = CDI)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)

    def Stage():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Stage"
        win3.destroy()
        modif('CONTRAT+', contrat)
    choix3 = Button(win3, text = "Stage", command = Stage)
    choix3.grid(row = 2, column = 1, sticky = W+E)
    choix3.config(font = font2)

    def Alternance():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Alternance"
        win3.destroy()
        modif('CONTRAT+', contrat)
    choix4 = Button(win3, text = "Alternance", command = Alternance)
    choix4.grid(row = 2, column = 2, sticky = W+E)
    choix4.config(font = font2)

    def Cdp():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Contrat de professionalisation"
        win3.destroy()
        modif('CONTRAT+', contrat)
    choix5 = Button(win3, text = "Contrat de prof.", command = Cdp)
    choix5.grid(row = 3, column = 1, columnspan = 2, sticky = W+E)
    choix5.config(font = font2)
b3 = Button(window, text = "", width = 5, height = 1, command = cb3)
b3.place(x = 485, y = 95)
b3.config(bg = '#B50000')

def cb4():
    win4 = Toplevel(window)
    # win1.resizable(False, False)
    win4.geometry("265x53+700+80")
    def ok():
        b4.config(bg = '#47B12F')
        durée = " pour une durée de "+choix.get()
        win4.destroy()
        modif("DUREE", durée)
    choix = Entry(win4, width = 37)
    choix.focus_force()
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font3)
    save = Button(win4, text = 'OK', command = ok)
    save.grid(row = 2, column = 1 , sticky = W+E)
    save.config(font = font2)
b4 = Button(window, text = "", width = 8, height = 1, command = cb4)
b4.place(x = 132, y = 115)
b4.config(bg = '#B50000')

def cb5():
    win5 = Toplevel(window) 
    win5.geometry("80x90+700+380") 
    win5.config(bg = '#47B12F')
    now = datetime.now()
    date = now.strftime("%d/%m/%Y")
    # cal = Calendar(win5, selectmode = 'day', 
    #             year = int(date[6:11]), month = int(date[3:5]), 
    #             day = int(date[0:2])) 
    cal  = DateEntry(win5, selectmode = 'day', 
                year = int(date[6:11]), month = int(date[3:5]), 
                day = int(date[0:2])) 
    Label(win5, text = "", bg = '#47B12F').grid(row = 1)
    Label(win5, text = "", bg = '#47B12F').grid(row = 4)
    Label(win5, text = "", bg = '#47B12F').grid(column = 1)
    Label(win5, text = "", bg = '#47B12F').grid(column = 3)
    cal.grid(row = 2, column = 2, sticky = W+E)
    def grad_date(): 
        b5.config(bg = '#47B12F') 
        debut = (cal.get_date()).strftime("%d/%m/%Y")
        # print(debut)
        win5.destroy()
        modif("DEBUT", debut)
    Button(win5, text = "OK", command = grad_date).grid(row = 3, column = 2, sticky = W+E) 
b5 = Button(window, text = "", width = 7, height = 1, command = cb5)
b5.place(x = 570, y = 115)
b5.config(bg = '#B50000')

def cb6():
    win6 = Toplevel(window)
    # win1.resizable(False, False)
    win6.geometry("265x188+700+80")
    def ok():
        b6.config(bg = '#47B12F')
        détails = choix.get("1.0",'end-1c')
        win6.destroy()
        modif("DETAILS", détails)
    choix = Text(win6, width = 37, height = 10)
    choix.focus_force()
    choix.grid(row = 1, rowspan = 2, column = 1)
    choix.config(font = font3)
    save = Button(win6, text = 'OK', command = ok)
    save.grid(row = 4, column = 1 , sticky = W+E)
    save.config(font = font2)
b6 = Button(window, text = "", width = 4, height = 1, command = cb6)
b6.place(x = 446, y = 150)
b6.config(bg = '#B50000')

def cb7():
    win7 = Toplevel(window)
    # win1.resizable(False, False)
    win7.geometry("265x53+700+80")
    def ok():
        b7.config(bg = '#47B12F')
        durée_essai = choix.get() 
        win7.destroy()
        modif("ESSAI", durée_essai)
    choix = Entry(win7, width = 37)
    choix.focus_force()
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font3)
    save = Button(win7, text = 'OK', command = ok)
    save.grid(row = 2, column = 1 , sticky = W+E)
    save.config(font = font2)
b7 = Button(window, text = "", width = 3, height = 1, command = cb7)
b7.place(x = 455, y = 243)
b7.config(bg = '#B50000')

def cb8():
    win8 = Toplevel(window) 
    win8.geometry("80x90+700+380") 
    win8.config(bg = '#47B12F')
    now = datetime.now()
    date = now.strftime("%d/%m/%Y")
    cal  = DateEntry(win8, selectmode = 'day', 
                year = int(date[6:11]), month = int(date[3:5]), 
                day = int(date[0:2])) 
    Label(win8, text = "", bg = '#47B12F').grid(row = 1)
    Label(win8, text = "", bg = '#47B12F').grid(row = 4)
    Label(win8, text = "", bg = '#47B12F').grid(column = 1)
    Label(win8, text = "", bg = '#47B12F').grid(column = 3)
    cal.grid(row = 2, column = 2, sticky = W+E)
    def grad_date(): 
        b8.config(bg = '#47B12F') 
        delai = (cal.get_date()).strftime("%d/%m/%Y")
        win8.destroy()
        modif("DELAI", delai)
    Button(win8, text = "OK", command = grad_date).grid(row = 3, column = 2, sticky = W+E) 
b8 = Button(window, text = "", width = 6, height = 1, command = cb8)
b8.place(x = 625, y = 296)
b8.config(bg = '#B50000')

def cb9():
    win9 = Toplevel(window) 
    win9.geometry("80x90+700+380") 
    win9.config(bg = '#47B12F')
    now = datetime.now()
    date = now.strftime("%d/%m/%Y")
    cal  = DateEntry(win9, selectmode = 'day', 
                year = int(date[6:11]), month = int(date[3:5]), 
                day = int(date[0:2])) 
    Label(win9, text = "", bg = '#47B12F').grid(row = 1)
    Label(win9, text = "", bg = '#47B12F').grid(row = 4)
    Label(win9, text = "", bg = '#47B12F').grid(column = 1)
    Label(win9, text = "", bg = '#47B12F').grid(column = 3)
    cal.grid(row = 2, column = 2, sticky = W+E)
    def grad_date(): 
        b9.config(bg = '#47B12F') 
        start = (cal.get_date()).strftime("%d/%m/%Y")
        win9.destroy()
        modif("START", start)
    Button(win9, text = "OK", command = grad_date).grid(row = 3, column = 2, sticky = W+E) 
b9 = Button(window, text = "", width = 6, height = 1, command = cb9)
b9.place(x = 626, y = 337)
b9.config(bg = '#B50000')

def cb10():
    win10 = Toplevel(window)
    # win1.resizable(False, False)
    win10.geometry("155x33+700+80")
    gender = ['Madame', 'Monsieur']
    def madame():
        b10.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        appellation = "Madame" 
        win10.destroy()
    choix = Button(win10, text = "Madame", command = madame)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2)

    def monsieur():
        b10.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        appellation = "Monsieur" 
        win10.destroy()
    choix2 = Button(win10, text = "Monsieur", command = monsieur)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)
b10 = Button(window, text = "", width = 6, height = 1, command = cb10)
b10.place(x = 290, y = 463)
b10.config(bg = '#B50000')

mail = Entry(window, width = 25)
mail.place(x = 430, y = 580)
mail.config(font = font2)

def sendmail():
    SMTP_SERVER = "smtp-mail.outlook.com"
    SMTP_PORT = 587
    SMTP_USERNAME = "recrutement.equo@outlook.com"
    SMTP_PASSWORD = "Equo2022"
    EMAIL_FROM = "recrutement.equo@outlook.com"
    EMAIL_TO = str(mail.get())
    EMAIL_SUBJECT = "[EQUO CONSTRUCTION: Reponse Candidature]"

    with open('modele_changeant_fav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
    EMAIL_MESSAGE = contents

    s = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    s.starttls()
    s.login(SMTP_USERNAME, SMTP_PASSWORD)
    message = 'Subject: {}\n\n{}'.format(EMAIL_SUBJECT, EMAIL_MESSAGE)
    s.sendmail(EMAIL_FROM, EMAIL_TO, message.encode('utf-8'))
    s.quit()

    print("email envoyé à ", EMAIL_TO," avec succès")

    with open('modele_fixe_fav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
    with open(r'modele_changeant_fav.txt', 'w', encoding = 'utf-8') as file:
        file.write(contents)

    window.destroy()
    messagebox.showinfo("Info", "Mail envoyé : succès")




send = Button(window, text = "Envoyer le mail", command = sendmail)
send.place(x = 430, y = 605)
send.config(bg = '#47B12F', fg = 'white')


window.mainloop()