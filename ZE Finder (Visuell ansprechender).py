import pandas as pd
from tabulate import tabulate
import tkinter as tk
from tkinter import ttk, messagebox
import os


def lade_excel_datei(pfad):
    """
    Lädt die Excel-Datei und gibt den DataFrame zurück.
    """
    try:
        df = pd.read_excel(pfad)
        return df
    except FileNotFoundError:
        messagebox.showerror("Dateifehler", f"Die Datei wurde nicht gefunden:\n{pfad}")
        return None
    except Exception as e:
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}")
        return None


def suche_daten(df, suchbegriff):
    """
    Sucht nach dem Suchbegriff in den Spalten 'OPS-Text' und 'Handelsnamen'.
    """
    suchbegriff = suchbegriff.lower()
    suchspalten = ['OPS-Text', 'Handelsnamen']

    fehlende_spalten = [spalte for spalte in suchspalten if spalte not in df.columns]
    if fehlende_spalten:
        messagebox.showerror("Spaltenfehler",
                             f"Die folgenden erforderlichen Spalten fehlen in der Excel-Datei:\n{', '.join(fehlende_spalten)}")
        return pd.DataFrame()

    mask = df[suchspalten].astype(str).apply(lambda x: x.str.lower().str.contains(suchbegriff, na=False)).any(axis=1)
    ergebnisse = df[mask]
    return ergebnisse


def zeige_ergebnisse(ergebnisse, tree, relevante_spalten):
    """
    Zeigt die Suchergebnisse in der Treeview an.
    """
    # Lösche vorherige Einträge
    for item in tree.get_children():
        tree.delete(item)

    if ergebnisse.empty:
        messagebox.showinfo("Keine Ergebnisse", "Keine passenden Einträge gefunden.")
    else:
        # Überprüfen, welche relevanten Spalten vorhanden sind
        vorhandene_spalten = [spalte for spalte in relevante_spalten if spalte in ergebnisse.columns]
        fehlende_spalten = [spalte for spalte in relevante_spalten if spalte not in ergebnisse.columns]
        if fehlende_spalten:
            messagebox.showwarning("Spaltenwarnung",
                                   f"Die folgenden Spalten fehlen in den Daten und werden nicht angezeigt:\n{', '.join(fehlende_spalten)}")

        # Definiere die Spalten der Treeview
        tree["columns"] = vorhandene_spalten
        for col in vorhandene_spalten:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor='center')

        # Füge die Daten zur Treeview hinzu
        for _, row in ergebnisse[vorhandene_spalten].iterrows():
            tree.insert("", "end", values=list(row))


def start_suche(df, tree, eingabe, relevante_spalten):
    """
    Startet die Suche und zeigt die Ergebnisse an.
    """
    suchbegriff = eingabe.get().strip()
    if suchbegriff.lower() == 'exit':
        root.quit()
        return
    if suchbegriff == "":
        messagebox.showwarning("Eingabefehler", "Bitte geben Sie einen gültigen Suchbegriff ein.")
        return
    ergebnisse = suche_daten(df, suchbegriff)
    zeige_ergebnisse(ergebnisse, tree, relevante_spalten)


# Hauptfenster erstellen
root = tk.Tk()
root.title("OPS-Text und Handelsnamen Suche")
root.geometry("800x600")
root.resizable(False, False)

# Excel-Dateipfad
excel_pfad = r"C:\Users\hamad\OneDrive\Desktop\ZE Liste.xlsx"

# Daten laden
df = lade_excel_datei(excel_pfad)

if df is None:
    root.destroy()  # Beende die Anwendung, wenn die Datei nicht geladen werden kann
else:
    # Suchbereich
    frame_search = ttk.Frame(root, padding="10")
    frame_search.pack(fill='x')

    label = ttk.Label(frame_search, text="Suchbegriff (Teil des OPS-Textes oder Handelsnamens):")
    label.pack(side='left', padx=(0, 10))

    eingabe = ttk.Entry(frame_search, width=50)
    eingabe.pack(side='left', fill='x', expand=True, padx=(0, 10))

    button_suchen = ttk.Button(frame_search, text="Suchen",
                               command=lambda: start_suche(df, tree, eingabe, relevante_spalten))
    button_suchen.pack(side='left')

    # Treeview für Ergebnisse
    frame_results = ttk.Frame(root, padding="10")
    frame_results.pack(fill='both', expand=True)

    # Definiere die relevanten Spalten
    relevante_spalten = ['ZE', 'OPS', 'OPS-Text', 'Handelsnamen', 'Wirkstoffklasse', 'Infos', 'Betrag']

    tree = ttk.Treeview(frame_results, columns=relevante_spalten, show='headings')
    tree.pack(side='left', fill='both', expand=True)

    # Scrollbar hinzufügen
    scrollbar = ttk.Scrollbar(frame_results, orient="vertical", command=tree.yview)
    scrollbar.pack(side='right', fill='y')
    tree.configure(yscrollcommand=scrollbar.set)

    # Beispiel: Setzen der Spaltenbreite und Ausrichtung
    for spalte in relevante_spalten:
        tree.heading(spalte, text=spalte)
        tree.column(spalte, width=100, anchor='center')

    # Statusleiste
    status = ttk.Label(root, text="Geben Sie einen Suchbegriff ein und klicken Sie auf 'Suchen'.", relief='sunken',
                       anchor='w')
    status.pack(side='bottom', fill='x')


    # Funktion zur Aktualisierung der Statusleiste
    def update_status(message):
        status.config(text=message)


    # Aktualisiere die Statusleiste nach der Suche
    def zeige_ergebnisse(ergebnisse, tree, relevante_spalten):
        # Lösche vorherige Einträge
        for item in tree.get_children():
            tree.delete(item)

        if ergebnisse.empty:
            messagebox.showinfo("Keine Ergebnisse", "Keine passenden Einträge gefunden.")
            update_status("Keine Ergebnisse gefunden.")
        else:
            # Überprüfen, welche relevanten Spalten vorhanden sind
            vorhandene_spalten = [spalte for spalte in relevante_spalten if spalte in ergebnisse.columns]
            fehlende_spalten = [spalte for spalte in relevante_spalten if spalte not in ergebnisse.columns]
            if fehlende_spalten:
                messagebox.showwarning("Spaltenwarnung",
                                       f"Die folgenden Spalten fehlen in den Daten und werden nicht angezeigt:\n{', '.join(fehlende_spalten)}")

            # Definiere die Spalten der Treeview
            tree["columns"] = vorhandene_spalten
            for col in vorhandene_spalten:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor='center')

            # Füge die Daten zur Treeview hinzu
            for _, row in ergebnisse[vorhandene_spalten].iterrows():
                tree.insert("", "end", values=list(row))

            update_status(f"{len(ergebnisse)} Einträge gefunden.")


    # Aktualisiere die 'zeige_ergebnisse' Funktion
    def start_suche(df, tree, eingabe, relevante_spalten):
        suchbegriff = eingabe.get().strip()
        if suchbegriff.lower() == 'exit':
            root.quit()
            return
        if suchbegriff == "":
            messagebox.showwarning("Eingabefehler", "Bitte geben Sie einen gültigen Suchbegriff ein.")
            return
        ergebnisse = suche_daten(df, suchbegriff)
        zeige_ergebnisse(ergebnisse, tree, relevante_spalten)


    # Starte die GUI
    root.mainloop()
