import os
import tkinter as tk
from tkinter import messagebox
from docx import Document

def save_to_word():
    # Obtient le répertoire des téléchargements
    download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = os.path.join(download_folder, "document.docx")

    # Crée un document Word
    doc = Document()
    # Ajoute le texte de la zone de texte au document Word
    doc.add_paragraph(text_area.get("1.0", tk.END))
    # Sauvegarde le document dans le dossier Téléchargements
    doc.save(file_path)

    print(f"Document saved to {file_path}")

def on_closing():
    save_to_word()  # Sauvegarde le document lorsque la fenêtre est fermée
    root.destroy()  # Ferme l'application

# Configuration de la fenêtre principale
root = tk.Tk()
root.title("Application de traitement de texte")

# Création de la zone de texte
text_area = tk.Text(root, wrap=tk.WORD)
text_area.pack(expand=True, fill=tk.BOTH)

# Gestionnaire d'événement pour la fermeture de la fenêtre
root.protocol("WM_DELETE_WINDOW", on_closing)

# Démarrer la boucle principale de l'application Tkinter
root.mainloop()
