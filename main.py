import subprocess
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import Converter
from PIL import ImageTk, Image
import os

# Fonctions

def ouvrir_fichier_PDF():
    root = tk.Tk()
    root.withdraw()
    fichier_PDF = filedialog.askopenfilename(filetypes=[("fichier pdf", ".pdf")])
    entry_fichier_PDF.delete(0, tk.END)
    entry_fichier_PDF.insert(tk.END, fichier_PDF)

def ouvrir_fichier_WORD():
    root = tk.Tk()
    root.withdraw()
    fichier_WORD = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("fichier word", ".docx")])
    entry_fichier_WORD.delete(0, tk.END)
    entry_fichier_WORD.insert(tk.END, fichier_WORD)

def convertir_PDF_to_WORD():
    fichier_pdf = entry_fichier_PDF.get()
    fichier_word = entry_fichier_WORD.get()

    if fichier_pdf and fichier_word:
        try:
            converter = Converter(fichier_pdf)
            converter.convert(fichier_word)
            converter.close()

            reponse = messagebox.askquestion("", "Conversion réussie! Voulez-vous ouvrir le fichier converti?",
                                             parent=root)
            if reponse == 'yes':
                subprocess.Popen([fichier_word], shell=True)
        except Exception as e:
            messagebox.showerror("Erreur", "Conversion échouée!", parent=root)
    else:
        messagebox.showerror("Erreur", "Sélectionner un fichier PDF SVP!", parent=root)


# Interface Graphique

root = tk.Tk() # fenetre principale
root.title("Convertir PDF en WORD") # titre de la fenetre
root.geometry("600x250") # dimenssions de la fentere
root.iconbitmap("icon.ico") # icône de la fenetre
root.resizable(width=False, height=False) # empechement le rendimensionnement de la fenetre
#root.configure(bg='red')
image = Image.open("pdftoword.jpg")
bg_image = ImageTk.PhotoImage(image)
bg_label = tk.Label(root, image=bg_image)
bg_label.place(x=0, y=40, relwidth=1, relheight=1)

frame = tk.Frame(root, bg="red")
frame.place(x=0, y=0, width=600, height=90)
title = tk.Label(frame, bg="red", fg="white", font=('algerian', 20)).place(x=5, y=5)


label_fichier_PDF = tk.Label(frame, text="Selectionner le fichier PDF", bg='red')
label_fichier_PDF.grid(row=0, column=0, padx=5, pady=10)

entry_fichier_PDF = tk.Entry(frame, width=45)
entry_fichier_PDF.grid(row=0, column=1, padx=5, pady=10 )

button_fichier_PDF = tk.Button(frame, text= "Parcourir", command=ouvrir_fichier_PDF)
button_fichier_PDF.grid(row=0, column=2, padx=5, pady=10)

label_fichier_WORD = tk.Label(frame, text="Enregistrer sous fichier WORD", bg='red')
label_fichier_WORD.grid(row=1, column=0, padx=5, pady=10)

entry_fichier_WORD = tk.Entry(frame, width=45)
entry_fichier_WORD.grid(row=1, column=1, padx=5, pady=10 )

button_fichier_WORD = tk.Button(frame, text= "Parcourir", command=ouvrir_fichier_WORD)
button_fichier_WORD.grid(row=1, column=2, padx=5, pady=10)

button_convertir = tk.Button(frame, text= "Convertir", command=convertir_PDF_to_WORD)
button_convertir.grid(row=1, column=3, padx=5, pady=10)

#label_resultat = tk.Label(root, text="")
#label_resultat.grid(row=3, column=1)



root.mainloop()