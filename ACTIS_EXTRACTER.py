# ------------------------------------------------------------------------
# ACTIS Extracter - Application pour convertir et remplacer des fichiers
# Droits d'auteur (c) 2023 KRAOUCH Yassir
# Ce programme est protégé par les lois de droits d'auteur et de marque.
# Tous droits réservés.
# ------------------------------------------------------------------------



import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import StringIO
from sylk_parser import SylkParser
import tkinter as tk
from tkinter import filedialog, messagebox

def convert_and_replace():
    # Récupérer les chemins des fichiers d'entrée et de sortie
    slk_file = entry_entree.get()
    xlsx_file = entry_sortie.get()

    if not slk_file:
        messagebox.showerror("Erreur", "Veuillez entrer le chemin du fichier SLK.")
        return

    if not xlsx_file:
        messagebox.showerror("Erreur", "Veuillez entrer le nom du fichier de sortie.")
        return

    if not xlsx_file.endswith('.xlsx'):
        xlsx_file += '.xlsx'

    # Convertir le fichier SLK en CSV
    parser = SylkParser(slk_file)
    fbuf = StringIO()
    parser.to_csv(fbuf)
    fbuf.seek(0)

    # Convertir le CSV en DataFrame
    df = pd.read_csv(fbuf)

    # Remplacer les caractères spéciaux
    caracteres_a_remplacer = {  'Ã©': 'é',
                            'Ã¨': 'è',
                            'Ã': 'à', 
                            'Ã´': 'ô', 
                            'Â': '', 
                            'Ã€': 'À',
                            'Ãª': 'ê',
                            'Ã´': 'ô',
                            'Ã¹': 'ù',
                            'Ã¢': 'â',
                            'Ã®': 'î',
                            'Ã¶': 'ö',
                            'Ã¯': 'ï',
                            'Ã«': 'ë',
                            'Ã¼': 'ü',
                            'Ã§': 'ç',
                            'Ã€': 'À',
                            'Ã‚': 'Â',
                            'Ãƒ': 'Ã',
                            'Ã„': 'Ä',
                            'Ã…': 'Å',
                            'Ã†': 'Æ',
                            'Ã‡': 'Ç',
                            'Ãˆ': 'È',
                            'Ã‰': 'É',
                            'ÃŠ': 'Ê',
                            'Ã‹': 'Ë',
                            'ÃŒ': 'Ì',
                            'ÃŽ': 'Î',
                            'Ã‘': 'Ñ',
                            'Ã’': 'Ò',
                            'Ã“': 'Ó',
                            'Ã”': 'Ô',
                            'Ã•': 'Õ',
                            'Ã–': 'Ö',
                            'Ã—': '×',
                            'Ã˜': 'Ø',
                            'Ã™': 'Ù',
                            'Ãš': 'Ú',
                            'Ã›': 'Û',
                            'Ãœ': 'Ü',
                            'ÃŸ': 'ß',
                            'Ã ': 'à',
                            'Ã¡': 'á',
                            'Ã¢': 'â',
                            'Ã£': 'ã',
                            'Ã¤': 'ä',
                            'Ã¥': 'å',
                            'Ã¦': 'æ',
                            'Ã§': 'ç',
                            'Ã¨': 'è',
                            'Ã©': 'é',
                            'Ãª': 'ê',
                            'Ã«': 'ë',
                            'Ã¬': 'ì',
                            'Ã­': 'í',
                            'Ã®': 'î',
                            'Ã¯': 'ï',
                            'Ã°': 'ð',
                            'Ã±': 'ñ',
                            'Ã²': 'ò',
                            'Ã³': 'ó',
                            'Ã´': 'ô',
                            'Ãµ': 'õ',
                            'Ã¶': 'ö',
                            'Ã·': '÷',
                            'Ã¸': 'ø',
                            'Ã¹': 'ù',
                            'Ãº': 'ú',
                            'Ã»': 'û',
                            'Ã¼': 'ü',
                            'Ã½': 'ý',
                            'Ã¾': 'þ',
                            'Ã¿': 'ÿ'
                          }
    
    def modifier_nom_colonne(col_name):
        for caractere, remplacement in caracteres_a_remplacer.items():
            col_name = col_name.replace(caractere, remplacement)
        return col_name

    df = df.rename(columns=modifier_nom_colonne) 
    df = df.replace(caracteres_a_remplacer, regex=True)

    # Créer un classeur Excel en utilisant openpyxl
    workbook = Workbook()
    sheet = workbook.active

    # Ajouter les données du DataFrame au classeur Excel
    for row in dataframe_to_rows(df, index=False, header=True):
        sheet.append(row)

    # Enregistrer le classeur Excel au format .xlsx
    workbook.save(xlsx_file)
    messagebox.showinfo("Conversion terminée", f"Fichier Excel '{xlsx_file}' créé avec succès.")

# Fonction pour parcourir le fichier d'entrée
def parcourir_entree():
    filename = filedialog.askopenfilename(filetypes=[("SLK files", "*.slk")])
    if filename:
        entry_entree.delete(0, tk.END)
        entry_entree.insert(0, filename)

# Fonction pour parcourir le fichier de sortie
def parcourir_sortie():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        entry_sortie.delete(0, tk.END)
        entry_sortie.insert(0, filename)

# Créer une fenêtre Tkinter
app = tk.Tk()
app.title("ACTIS Extracter")

# Icône de la fenêtre
icon_photo = tk.PhotoImage(file="download.png")
app.iconphoto(False, icon_photo)

# Interface graphique
frame_entree = tk.Frame(app)
label_entree = tk.Label(frame_entree, text="Fichier d'Entrée:")
entry_entree = tk.Entry(frame_entree, width=40)
bouton_parcourir_entree = tk.Button(frame_entree, text="Parcourir", command=parcourir_entree)

frame_sortie = tk.Frame(app)
label_sortie = tk.Label(frame_sortie, text="Fichier de Sortie:")
entry_sortie = tk.Entry(frame_sortie, width=40)
bouton_parcourir_sortie = tk.Button(frame_sortie, text="Parcourir", command=parcourir_sortie)

bouton_executer = tk.Button(app, text="Convertir et Remplacer", command=convert_and_replace)
label_resultat = tk.Label(app, text="")

frame_entree.pack(pady=10)
label_entree.pack(side=tk.LEFT)
entry_entree.pack(side=tk.LEFT)
bouton_parcourir_entree.pack(side=tk.LEFT)

frame_sortie.pack(pady=10)
label_sortie.pack(side=tk.LEFT)
entry_sortie.pack(side=tk.LEFT)
bouton_parcourir_sortie.pack(side=tk.LEFT)

bouton_executer.pack(pady=10)
label_resultat.pack()

# Crédit pour l'auteur du code
auteur_label = tk.Label(app, text="Code réalisé par KRAOUCH", font=("Helvetica", 8), fg="gray")
auteur_label.pack(side=tk.RIGHT, padx=10, pady=5, anchor='se')

# Lancer la boucle principale de l'interface graphique
app.mainloop()
