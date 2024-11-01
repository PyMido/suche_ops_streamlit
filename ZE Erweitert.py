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

    # Entfernen von führenden/trailenden Leerzeichen in allen String-Zellen
    str_cols = df.select_dtypes(include=['object']).columns
    df[str_cols] = df[str_cols].apply(lambda x: x.str.strip())

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


def get_medication_info(df, ops_text=None, handelsname=None):
    """
    Gibt alle Informationen zu einem gegebenen OPS-Text oder Handelsnamen zurück (teilweise Übereinstimmung).

    :param df: pandas DataFrame.
    :param ops_text: Der eingegebene OPS-Text (teilweise oder vollständige Angabe).
    :param handelsname: Der eingegebene Handelsname (teilweise oder vollständige Angabe).
    :return: Gefilterter DataFrame mit den relevanten Informationen.
    """
    if ops_text:
        spalte = 'OPS-Text'
        if spalte in df.columns:
            filtered_df_ops = df[df[spalte].str.contains(ops_text, case=False, na=False)]
        else:
            print(f"Spalte '{spalte}' nicht gefunden.")
            filtered_df_ops = pd.DataFrame()
    else:
        filtered_df_ops = pd.DataFrame()

    if handelsname:
        spalte = 'Handelsnamen | Alternativbezeichnung, Synonym'
        if spalte in df.columns:
            filtered_df_handelsname = df[df[spalte].str.contains(handelsname, case=False, na=False)]
        else:
            print(f"Spalte '{spalte}' nicht gefunden.")
            filtered_df_handelsname = pd.DataFrame()
    else:
        filtered_df_handelsname = pd.DataFrame()

    if ops_text and handelsname:
        # Kombiniere die beiden Filter mit einer logischen OR-Bedingung
        filtered_df = pd.concat([filtered_df_ops, filtered_df_handelsname]).drop_duplicates()
    elif ops_text:
        filtered_df = filtered_df_ops
    elif handelsname:
        filtered_df = filtered_df_handelsname
    else:
        filtered_df = pd.DataFrame()

    if filtered_df.empty:
        search_field = []
        if ops_text:
            search_field.append(f"OPS-Text '{ops_text}'")
        if handelsname:
            search_field.append(f"Handelsname '{handelsname}'")
        search_description = " und ".join(search_field)
        print(f"Keine Einträge für den/die {search_description} gefunden.")
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
        print("\nWählen Sie die Suchoption:")
        print("1. OPS-Text")
        print("2. Handelsname | Alternativbezeichnung, Synonym")
        print("3. Beide")
        print("4. Beenden")

        auswahl = input("Geben Sie die Nummer Ihrer Wahl ein: ").strip()

        if auswahl == '4' or auswahl.lower() == 'exit':
            print("Programm beendet.")
            break
        elif auswahl not in ['1', '2', '3']:
            print("Ungültige Auswahl. Bitte wählen Sie eine der Optionen 1, 2, 3 oder 4.")
            continue

        ops_text = None
        handelsname = None

        if auswahl in ['1', '3']:
            ops_text = input("Geben Sie den OPS-Text ein (teilweise oder vollständig): ").strip()
            if ops_text.lower() == 'exit':
                print("Programm beendet.")
                break
            elif ops_text == '':
                print("Bitte geben Sie einen gültigen OPS-Text ein.")
                ops_text = None

        if auswahl in ['2', '3']:
            handelsname = input(
                "Geben Sie den Handelsnamen oder eine Alternativbezeichnung ein (teilweise oder vollständig): ").strip()
            if handelsname.lower() == 'exit':
                print("Programm beendet.")
                break
            elif handelsname == '':
                print("Bitte geben Sie einen gültigen Handelsnamen ein.")
                handelsname = None

        # Abrufen der Informationen
        filtered_df = get_medication_info(df, ops_text, handelsname)

        # Anzeigen der Informationen
        display_information(filtered_df)
        print("\n" + "-" * 80 + "\n")


if __name__ == "__main__":
    main()
