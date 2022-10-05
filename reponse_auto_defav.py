from tkinter import *
from PIL import ImageTk, Image
from tkinter import ttk, messagebox
from tkcalendar import Calendar 
from datetime import datetime
import smtplib

window = Tk()
window.resizable(False, False)
window.geometry("700x540+350+60")
font1 = ('Consolas  Bold', 18)
font2 = ('Consolas  Bold', 12)
font3 = ('Consolas', 10)
window.config(bg ='white')


blank = Label(window, text = "", borderwidth = 10)
blank.config(font = font2)

back = 'image_defav.png'
file = Image.open(back)
image = ImageTk.PhotoImage(file)
img = Label(window, image = image)
img.image = image
img.place(x = 10, y = 10)
img.config(bg = 'white')


def modif(replace_this, by_this):
    with open('modele_changeant_defav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
        contents = contents.replace(replace_this, by_this)
    with open(r'modele_changeant_defav.txt', 'w', encoding = 'utf-8') as file:
        file.write(contents)

def cb1():
    win1 = Toplevel(window)
    # win1.resizable(False, False)
    win1.geometry("155x33+700+80")
    gender = ['Madame', 'Monsieur']
    def madame():
        b1.config(bg = '#47B12F')
        b4.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        appellation = "Madame" 
        win1.destroy()
        modif("GENDER", appellation)
    choix = Button(win1, text = "Madame", command = madame)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2)
    def monsieur():
        b1.config(bg = '#47B12F')
        b4.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        appellation = "Monsieur" 
        win1.destroy()
        modif("GENDER", appellation)
    choix2 = Button(win1, text = "Monsieur", command = monsieur)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)
b1 = Button(window, text = "", width = 4, height = 1, command = cb1)
b1.place(x = 96, y = 30)
b1.config(bg = '#B50000')

def cb3():
    win3 = Toplevel(window)
    # win1.resizable(False, False)
    win3.geometry("145x99+700+80")
    def CDD():
        b3.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        contrat = "CDD" 
        win3.destroy()
        modif('CONTRAT', contrat)
    choix = Button(win3, text = "CDD", command = CDD)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2) 
    def CDI():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "CDI"
        win3.destroy()
        modif('CONTRAT', contrat)
    choix2 = Button(win3, text = "CDI", command = CDI)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)
    def Stage():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Stage"
        win3.destroy()
        modif('CONTRAT', contrat)
    choix3 = Button(win3, text = "Stage", command = Stage)
    choix3.grid(row = 2, column = 1, sticky = W+E)
    choix3.config(font = font2)
    def Alternance():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Alternance"
        win3.destroy()
        modif('CONTRAT', contrat)
    choix4 = Button(win3, text = "Alternance", command = Alternance)
    choix4.grid(row = 2, column = 2, sticky = W+E)
    choix4.config(font = font2)
    def Cdp():
        b3.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        contrat = "Contrat de professionalisation"
        win3.destroy()
        modif('CONTRAT', contrat)
    choix5 = Button(win3, text = "Contrat de prof.", command = Cdp)
    choix5.grid(row = 3, column = 1, columnspan = 2, sticky = W+E)
    choix5.config(font = font2)
b3 = Button(window, text = "", width = 4, height = 1, command = cb3)
b3.place(x = 363, y = 61)
b3.config(bg = '#B50000')


def cb4():
    win10 = Toplevel(window)
    # win1.resizable(False, False)
    win10.geometry("155x33+700+80")
    gender = ['Madame', 'Monsieur']
    def madame():
        b4.config(bg = '#47B12F')
        choix.config(bg = '#47B12F')
        appellation = "Madame" 
        win10.destroy()
    choix = Button(win10, text = "Madame", command = madame)
    choix.grid(row = 1, column = 1, sticky = W+E)
    choix.config(font = font2)

    def monsieur():
        b4.config(bg = '#47B12F')
        choix2.config(bg = '#47B12F')
        appellation = "Monsieur" 
        win10.destroy()
    choix2 = Button(win10, text = "Monsieur", command = monsieur)
    choix2.grid(row = 1, column = 2, sticky = W+E)
    choix2.config(font = font2)
b4 = Button(window, text = "", width = 5, height = 1, command = cb4)
b4.place(x = 146, y = 305)
b4.config(bg = '#B50000')

mail = Entry(window, width = 25)
mail.place(x = 430, y = 435)
mail.config(font = font2)

def sendmail():
    SMTP_SERVER = "smtp-mail.outlook.com"
    SMTP_PORT = 587
    SMTP_USERNAME = "recrutement.equo@outlook.com"
    SMTP_PASSWORD = "Equo2022"
    EMAIL_FROM = "recrutement.equo@outlook.com"
    EMAIL_TO = str(mail.get())
    EMAIL_SUBJECT = "[EQUO CONSTRUCTION: Reponse Candidature]"

    with open('modele_changeant_defav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
    EMAIL_MESSAGE = contents

    s = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    s.starttls()
    s.login(SMTP_USERNAME, SMTP_PASSWORD)
    message = 'Subject: {}\n\n{}'.format(EMAIL_SUBJECT, EMAIL_MESSAGE)
    s.sendmail(EMAIL_FROM, EMAIL_TO, message.encode('utf-8'))
    s.quit()

    print("email envoyé à ", EMAIL_TO," avec succès")

    with open('modele_fixe_defav.txt', 'r', encoding = 'utf-8') as f:
        contents = f.read()
    with open(r'modele_changeant_defav.txt', 'w', encoding = 'utf-8') as file:
        file.write(contents)
    
    window.destroy()
    messagebox.showinfo("Info", "Mail envoyé : succès")


send = Button(window, text = "Envoyer le mail", command = sendmail)
send.place(x = 430, y = 460)
send.config(bg = '#47B12F', fg = 'white')


window.mainloop()


