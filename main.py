import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import Converter
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

            messagebox.showinfo("", "Conversion réussie", parent=root)
        except Exception as e:
            messagebox.showierror("Erreur", "Coversion échouée!", parent=root)
    else:
        messagebox.showerror("Erreur", "Selectionner un fichier PDF SVP!", parent=root)


# Interface Graphique

root = tk.Tk()
root.title("Convertir PDF en WORD")
root.geometry("600x250")
root.iconbitmap("C:\\Users\\nassi\\PycharmProjects\\PDF_to_WORD\\icon1.ico")
root.resizable(width=False, height=False)

label_fichier_PDF = tk.Label(root, text="Selectionner le fichier PDF")
label_fichier_PDF.grid(row=0, column=0, padx=5, pady=10)

entry_fichier_PDF = tk.Entry(root, width=45)
entry_fichier_PDF.grid(row=0, column=1, padx=5, pady=10 )

button_fichier_PDF = tk.Button(root, text= "Parcourir", command=ouvrir_fichier_PDF)
button_fichier_PDF.grid(row=0, column=2, padx=5, pady=10)

label_fichier_WORD = tk.Label(root, text="Enregistrer sous fichier WORD")
label_fichier_WORD.grid(row=1, column=0, padx=5, pady=10)

entry_fichier_WORD = tk.Entry(root, width=45)
entry_fichier_WORD.grid(row=1, column=1, padx=5, pady=10 )

button_fichier_WORD = tk.Button(root, text= "Parcourir", command=ouvrir_fichier_WORD)
button_fichier_WORD.grid(row=1, column=2, padx=5, pady=10)

button_convertir = tk.Button(root, text= "Convertir", command=convertir_PDF_to_WORD)
button_convertir.grid(row=2, column=1, padx=5, pady=10)

#label_resultat = tk.Label(root, text="")
#label_resultat.grid(row=3, column=1)



root.mainloop()