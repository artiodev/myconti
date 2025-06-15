import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import shutil

year = '2025'

class CSVLoaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Caricamento Report CSV")
        self.root.configure(bg="black")  # Sfondo finestra

        self.categories = {
            "ContoCorrente": [],
            "CartaPrepagata": [],
            "CartaPaypal": []
        }

        self.buttons = {}
        self.status_labels = {}
        self.name_labels = {}

        row = 0
        for category in self.categories:
            # Button di caricamento
            self.buttons[category] = tk.Button(
                root,
                text="Carica file {}".format(category),
                width=50,
                command=lambda c=category: self.load_files(c),
                fg="white",
                highlightbackground="#333333",
                activebackground="#555555"
            )
            self.buttons[category].grid(row=row, column=0, padx=10, pady=5)

            row += 1

        # Bottone per eseguire main.py
        self.run_button = tk.Button(
            root,
            text="Esegui Analisi (main.py)",
            command=self.run_main_script,
            bg="#228B22",
            fg="white",
            width=40
        )
        self.run_button.grid(row=row, column=0, columnspan=3, pady=20)

    def load_files(self, category):
        files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        if files:
            self.categories[category].extend(files)

    def save_files_to_folders(self):
        base_path = os.getcwd()
        for category, file_list in self.categories.items():
            folder_path = os.path.join(base_path, category, year)
            os.makedirs(folder_path, exist_ok=True)
            for filepath in file_list:
                filename = os.path.basename(filepath)
                destination = os.path.join(folder_path, filename)
                shutil.copy(filepath, destination)

    def run_main_script(self):
        if all(self.categories[cat] for cat in self.categories):
            self.save_files_to_folders()
            try:
                subprocess.run(["python", "main.py"], check=True)
                messagebox.showinfo("Eseguito", "main.py è stato eseguito con successo.")
            except subprocess.CalledProcessError:
                messagebox.showerror("Errore", "Errore durante l'esecuzione di main.py.")
        else:
            messagebox.showwarning("File mancanti", "Carica almeno un file per ogni categoria prima di avviare l'analisi.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVLoaderApp(root)
    root.mainloop()
