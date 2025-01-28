import tkinter as tk
from tkinter import filedialog, messagebox
from logic import load_sheet_as_dict, generate_cerfa
import os

class CerfaGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("cerfasso")
        self.root.geometry("350x325")
        self.root.resizable(0, 0)
        
        self.setup_ui()
    
    def setup_ui(self):
        tk.Label(self.root, text="Fichier excel:", justify="left", anchor="w").pack(padx=10, anchor="w")
        self.excel_path = tk.StringVar()
        tk.Entry(self.root, textvariable=self.excel_path, width=40).pack(padx=10, anchor="w")
        tk.Button(self.root, text="Choisir un fichier excel", command=self.select_excel_file, justify="left", anchor="w").pack(pady=5, padx=10, anchor="w")
        
        tk.Label(self.root, text="Signature:", justify="left", anchor="w").pack(padx=10, anchor="w")
        self.signature_path = tk.StringVar()
        tk.Entry(self.root, textvariable=self.signature_path, width=40).pack(padx=10, anchor="w")
        tk.Button(self.root, text="Choisir une fichier singature", command=self.select_signature_file, justify="left", anchor="w").pack(pady=5, padx=10, anchor="w")
        
        tk.Label(self.root, text="Dossier de destination:", justify="left").pack(padx=10, anchor="w")
        self.folder_path = tk.StringVar()
        tk.Entry(self.root, textvariable=self.folder_path, width=40).pack(padx=10, anchor="w")
        tk.Button(self.root, text="Choisir un dossier", command=self.select_folder_path).pack(pady=5, padx=10, anchor="w")
        
        tk.Button(self.root, text="Générer les cerfa", command=self.generate).pack(pady=10)
    
    def select_excel_file(self):
        if file_path := filedialog.askopenfilename():
            self.excel_path.set(file_path)
    
    def select_signature_file(self):
        if file_path := filedialog.askopenfilename():
            self.signature_path.set(file_path)
    
    def select_folder_path(self):
        if folder_path := filedialog.askdirectory():
            self.folder_path.set(folder_path)
    
    def generate(self):
        if not all([self.excel_path.get(), self.folder_path.get(), self.signature_path.get()]):
            messagebox.showerror("Erreur", "Vérifiez que tous les éléments suivants soient spécifiés: fichier excel, signature, dossier de destination.")
            return
        
        excel_path = self.excel_path.get()
        signature=self.signature_path.get()
        destination=self.folder_path.get()
        sheet_name = None
        dons, asso = load_sheet_as_dict(excel_path, sheet_name)
        
        dirname = os.path.dirname(__file__)
        cerfa_path = os.path.join(dirname, "assets/2041-rd_4298.pdf")
        
        for donateur in dons:
            generate_cerfa(cerfa_path,asso,donateur,signature,destination)
        
        messagebox.showinfo("Succès", "Les cerfa sont disponibles dans le dossier de destination.")