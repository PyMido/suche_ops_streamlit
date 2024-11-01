import pandas as pd
from tabulate import tabulate  # Für eine bessere Tabellenanzeige


def load_data(file_path):
    """
    Lädt die Excel-Tabelle in einen pandas DataFrame.

    :param file_path: Pfad zur Excel-Datei.
    :return: pandas DataFrame.
    """
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except FileNotFoundError:
        print(f"Datei nicht gefunden: {file_path}")
        return None
    except Exception as e:
        print(f"Fehler beim Laden der Datei: {e}")
        return None


def preprocess_data(df):
    """
    Bereitet die Daten vor, indem Mehrfacheinträge in Handelsnamen aufgeteilt und Leerzeichen entfernt werden.

    :param df: pandas DataFrame.
    :return: Vorgebereiteter DataFrame.
    """
    # Entfernen von führenden/trailenden Leerzeichen in den Spaltennamen
    df.columns = df.columns.str.strip()

    # Entfernen von führenden/trailenden Leerzeichen in allen Zellen (nur für Strings)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    return df


def split_trade_names(df):
    """
    Teilt die Handelsnamen in separate Zeilen auf, wenn sie durch '|' getrennt sind.

    :param df: pandas DataFrame.
    :return: DataFrame mit aufgeteilten Handelsnamen.
    """
    spalte = 'Handelsnamen | Alternativbezeichnung, Synonym'
    if spalte in df.columns:
        df[spalte] = df[spalte].str.split('|')
        df = df.explode(spalte)
        # Entfernen von führenden/trailenden Leerzeichen nach dem Split
        df[spalte] = df[spalte].str.strip()
    else:
        print(f"Spalte '{spalte}' nicht gefunden.")
    return df


def get_medication_info(df, ops_text):
    """
    Gibt alle Informationen zu einem gegebenen OPS-Text zurück (teilweise Übereinstimmung).

    :param df: pandas DataFrame.
    :param ops_text: Der eingegebene OPS-Text (teilweise oder vollständige Angabe).
    :return: Gefilterter DataFrame mit den relevanten Informationen.
    """
    spalte = 'OPS-Text'
    if spalte in df.columns:
        # Filtern nach teilweiser Übereinstimmung des OPS-Texts (nicht case-sensitiv)
        filtered_df = df[df[spalte].str.contains(ops_text, case=False, na=False)]
    else:
        print(f"Spalte '{spalte}' nicht gefunden.")
        return None

    if filtered_df.empty:
        print(f"Keine Einträge für den OPS-Text '{ops_text}' gefunden.")
        return None

    return filtered_df


def display_information(filtered_df):
    """
    Zeigt die gefilterten Informationen strukturiert an.

    :param filtered_df: Gefilterter pandas DataFrame.
    """
    if filtered_df is not None and not filtered_df.empty:
        # Auswahl der relevanten Spalten
        relevante_spalten = ['ZE', 'OPS', 'OPS-Text', 'Handelsnamen | Alternativbezeichnung, Synonym',
                             'Wirkstoffklasse', 'Infos', 'Betrag']
        fehlende_spalten = [col for col in relevante_spalten if col not in filtered_df.columns]
        if fehlende_spalten:
            print(f"Die folgenden erwarteten Spalten fehlen im DataFrame: {fehlende_spalten}")
            return

        display_df = filtered_df[relevante_spalten]

        # Anzeige als Tabelle mit tabulate für bessere Lesbarkeit
        print("\nGefundene Einträge:")
        print(tabulate(display_df, headers='keys', tablefmt='fancy_grid', showindex=False))
    else:
        print("Keine Informationen zum Anzeigen.")


def main():
    # === Hier ist Ihr Dateipfad integriert ===
    file_path = r'C:\Users\hamad\OneDrive\Desktop\ZE Liste.xlsx'

    # Laden der Daten
    df = load_data(file_path)
    if df is None:
        return

    # Datenvorbereitung
    df = preprocess_data(df)

    # Aufteilen der Handelsnamen
    df = split_trade_names(df)

    while True:
        # Eingabe des OPS-Texts
        ops_text = input("Geben Sie den OPS-Text ein (oder 'exit' zum Beenden): ").strip()
        if ops_text.lower() == 'exit':
            print("Programm beendet.")
            break
        elif ops_text == '':
            print("Bitte geben Sie einen gültigen OPS-Text ein.")
            continue

        # Abrufen der Informationen
        filtered_df = get_medication_info(df, ops_text)

        # Anzeigen der Informationen
        display_information(filtered_df)
        print("\n" + "-" * 80 + "\n")


if __name__ == "__main__":
    main()
