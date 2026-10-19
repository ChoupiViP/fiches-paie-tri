#!/usr/bin/env python3
"""Petite fenêtre pour trier ses fiches de paie. Tout reste sur l'ordinateur."""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import trier_fiches_paie as core


def ouvrir(chemin):
    """Ouvre un fichier ou un dossier avec le programme par défaut du système."""
    chemin = str(chemin)
    if sys.platform.startswith("win"):
        os.startfile(chemin)  # noqa: S606 (Windows uniquement)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", chemin])
    else:
        subprocess.Popen(["xdg-open", chemin])


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tri des fiches de paie")
        self.geometry("720x520")
        self.minsize(600, 420)
        self.file = queue.Queue()
        self.resultat = None

        self.var_src = tk.StringVar()
        self.var_dst = tk.StringVar()
        self.var_debug = tk.BooleanVar(value=False)

        cadre = ttk.Frame(self, padding=12)
        cadre.pack(fill="both", expand=True)
        cadre.columnconfigure(1, weight=1)
        cadre.rowconfigure(5, weight=1)

        ttk.Label(
            cadre,
            text="Choisis le dossier qui contient tes fiches de paie (PDF). "
                 "Tout reste sur cet ordinateur : rien n'est envoyé sur Internet.",
            wraplength=680,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(cadre, text="Dossier des fiches :").grid(row=1, column=0, sticky="w")
        ttk.Entry(cadre, textvariable=self.var_src).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(cadre, text="Parcourir…", command=self.choisir_src).grid(row=1, column=2)

        ttk.Label(cadre, text="Enregistrer le résultat dans :").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(cadre, textvariable=self.var_dst).grid(row=2, column=1, sticky="ew", padx=6)
        ttk.Button(cadre, text="Parcourir…", command=self.choisir_dst).grid(row=2, column=2)

        ttk.Checkbutton(
            cadre, text="Mode diagnostic (avancé : garde le texte lu dans chaque PDF)",
            variable=self.var_debug,
        ).grid(row=3, column=0, columnspan=3, sticky="w")

        barre = ttk.Frame(cadre)
        barre.grid(row=4, column=0, columnspan=3, sticky="w", pady=10)
        self.bouton_go = ttk.Button(barre, text="Lancer le tri", command=self.lancer)
        self.bouton_go.pack(side="left")
        self.bouton_xlsx = ttk.Button(barre, text="Ouvrir le tableau Excel",
                                      command=lambda: ouvrir(self.resultat["xlsx"]), state="disabled")
        self.bouton_xlsx.pack(side="left", padx=8)
        self.bouton_dossier = ttk.Button(barre, text="Ouvrir le dossier",
                                         command=lambda: ouvrir(self.resultat["sortie"]), state="disabled")
        self.bouton_dossier.pack(side="left")

        self.journal = ScrolledText(cadre, height=12, state="disabled", wrap="word")
        self.journal.grid(row=5, column=0, columnspan=3, sticky="nsew")

        self.ecrire("Prêt. Choisis un dossier puis clique sur « Lancer le tri ».")
        self.after(100, self.vider_file)

    # ---- helpers ----
    def ecrire(self, texte):
        self.journal.configure(state="normal")
        self.journal.insert("end", texte + "\n")
        self.journal.see("end")
        self.journal.configure(state="disabled")

    def vider_file(self):
        try:
            while True:
                type_, contenu = self.file.get_nowait()
                if type_ == "log":
                    self.ecrire(contenu)
                elif type_ == "fini":
                    self.resultat = contenu
                    self.bouton_go.configure(state="normal")
                    self.bouton_xlsx.configure(state="normal")
                    self.bouton_dossier.configure(state="normal")
                elif type_ == "erreur":
                    self.bouton_go.configure(state="normal")
                    self.ecrire(f"Erreur : {contenu}")
                    messagebox.showerror("Tri impossible", str(contenu))
        except queue.Empty:
            pass
        self.after(100, self.vider_file)

    def choisir_src(self):
        d = filedialog.askdirectory(title="Dossier des fiches de paie")
        if d:
            self.var_src.set(d)
            if not self.var_dst.get():
                self.var_dst.set(str(Path(d).parent))

    def choisir_dst(self):
        d = filedialog.askdirectory(title="Où enregistrer le résultat ?")
        if d:
            self.var_dst.set(d)

    # ---- traitement ----
    def lancer(self):
        src = self.var_src.get().strip()
        base = self.var_dst.get().strip()
        if not src:
            messagebox.showinfo("Dossier manquant", "Choisis d'abord le dossier des fiches de paie.")
            return
        if not base:
            base = str(Path(src).parent)
        # un nouveau dossier à chaque lancement : rien n'est écrasé
        sortie = Path(base) / f"Fiches_de_paie_triees_{datetime.now():%Y%m%d_%H%M%S}"
        self.bouton_go.configure(state="disabled")
        self.bouton_xlsx.configure(state="disabled")
        self.bouton_dossier.configure(state="disabled")
        self.ecrire("-" * 50)
        threading.Thread(
            target=self.travail, args=(src, sortie, self.var_debug.get()), daemon=True
        ).start()

    def travail(self, src, sortie, debug):
        try:
            res = core.traiter(src, sortie, debug, log=lambda m: self.file.put(("log", m)))
            self.file.put(("fini", res))
        except Exception as e:  # noqa: BLE001 : on affiche toute erreur à l'utilisateur
            self.file.put(("erreur", e))


if __name__ == "__main__":
    Application().mainloop()
