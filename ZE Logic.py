import pandas as pd
from tabulate import tabulate


def lade_excel_datei(pfad):
    """
    Lädt die Excel-Datei und gibt den DataFrame zurück.
    """
    try:
        df = pd.read_excel(pfad)
        return df
    except FileNotFoundError:
        print(f"Die Datei wurde nicht gefunden: {pfad}")
        return None
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        return None


def suche_daten(df, suchbegriff):
    """
    Sucht nach dem Suchbegriff in den Spalten 'OPS-Text' und 'Handelsnamen'.
    """
    # Konvertiere Suchbegriff zu Kleinbuchstaben für eine case-insensitive Suche
    suchbegriff = suchbegriff.lower()

    # Überprüfe, welche Spalten durchsucht werden sollen
    suchspalten = ['OPS-Text', 'Handelsnamen']

    # Prüfe, ob die Suchspalten im DataFrame vorhanden sind
    fehlende_spalten = [spalte for spalte in suchspalten if spalte not in df.columns]
    if fehlende_spalten:
        print(f"Die folgenden erforderlichen Spalten fehlen in der Excel-Datei: {', '.join(fehlende_spalten)}")
        return pd.DataFrame()  # Leerer DataFrame

    # Erstelle eine Maske für Zeilen, die den Suchbegriff enthalten
    mask = df[suchspalten].astype(str).apply(lambda x: x.str.lower().str.contains(suchbegriff, na=False)).any(axis=1)

    # Filtere den DataFrame
    ergebnisse = df[mask]

    return ergebnisse


def zeige_ergebnisse(ergebnisse):
    """
    Zeigt die Suchergebnisse als strukturierte Tabelle an.
    """
    if ergebnisse.empty:
        print("Keine passenden Einträge gefunden.")
    else:
        # Definiere die relevanten Spalten
        relevante_spalten = ['ZE', 'OPS', 'OPS-Text', 'Handelsnamen', 'Wirkstoffklasse', 'Infos', 'Betrag']

        # Überprüfen, ob die Spalten existieren
        vorhandene_spalten = [spalte for spalte in relevante_spalten if spalte in ergebnisse.columns]

        fehlende_spalten = [spalte for spalte in relevante_spalten if spalte not in ergebnisse.columns]
        if fehlende_spalten:
            print(
                f"Die folgenden Spalten fehlen in den Daten und werden nicht angezeigt: {', '.join(fehlende_spalten)}")

        # Bereite die Daten für die tabulate-Bibliothek vor
        tabelle = ergebnisse[vorhandene_spalten]

        # Zeige die Tabelle an
        print(tabulate(tabelle, headers='keys', tablefmt='fancy_grid', showindex=False))


def main():
    excel_pfad = r"C:\Users\hamad\OneDrive\Desktop\ZE Liste.xlsx"
    df = lade_excel_datei(excel_pfad)

    if df is not None:
        while True:
            suchbegriff = input(
                "Geben Sie einen Teil des OPS-Textes oder Handelsnamens ein (oder 'exit' zum Beenden): ").strip()
            if suchbegriff.lower() == 'exit':
                print("Programm wird beendet.")
                break
            if suchbegriff == "":
                print("Bitte geben Sie einen gültigen Suchbegriff ein.")
                continue
            ergebnisse = suche_daten(df, suchbegriff)
            zeige_ergebnisse(ergebnisse)
            print("\n")  # Neue Zeile für bessere Lesbarkeit


if __name__ == "__main__":
    main()
