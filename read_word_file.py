from docx import Document
from tkinter import Tk, filedialog, messagebox

def read_word_file(file_path):
    # Charge le document Word
    doc = Document(file_path)
    
    # Lit et retourne tout le texte du document
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

def open_file_dialog():
    # Ouvre une boîte de dialogue pour sélectionner un fichier Word
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        # Lis le fichier Word sélectionné
        content = read_word_file(file_path)
        messagebox.showinfo("Contenu du fichier", content)

# Configuration de l'interface Tkinter
root = Tk()
root.withdraw()  # Cache la fenêtre principale

# Ouvre la boîte de dialogue de fichier et lit le fichier Word
open_file_dialog()
