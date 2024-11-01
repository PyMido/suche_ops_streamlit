import streamlit as st
import pandas as pd

def lade_excel_datei(pfad):
    """
    Lädt die Excel-Datei und gibt den DataFrame zurück.
    """
    try:
        df = pd.read_excel(pfad)
        return df
    except FileNotFoundError:
        st.error(f"Die Datei wurde nicht gefunden: {pfad}")
        return None
    except Exception as e:
        st.error(f"Ein Fehler ist aufgetreten: {e}")
        return None

def suche_daten(df, suchbegriff):
    """
    Sucht nach dem Suchbegriff in den Spalten 'OPS-Text' und 'Handelsnamen'.
    """
    suchbegriff = suchbegriff.lower()
    suchspalten = ['OPS-Text', 'Handelsnamen']

    # Überprüfen, ob die Suchspalten vorhanden sind
    fehlende_spalten = [spalte for spalte in suchspalten if spalte not in df.columns]
    if fehlende_spalten:
        st.error(f"Die folgenden erforderlichen Spalten fehlen in der Excel-Datei: {', '.join(fehlende_spalten)}")
        return pd.DataFrame()

    mask = df[suchspalten].astype(str).apply(lambda x: x.str.lower().str.contains(suchbegriff, na=False)).any(axis=1)
    ergebnisse = df[mask]
    return ergebnisse

def main():
    st.title("OPS-Text und Handelsnamen Suche")

    # Angepasster Pfad zur Excel-Datei im Hauptverzeichnis
    excel_pfad = "ZE Liste.xlsx"

    st.sidebar.header("Daten laden")
    st.sidebar.write(f"**Aktueller Excel-Pfad:** {excel_pfad}")

    # Daten laden beim Start der Anwendung
    if 'df' not in st.session_state:
        df = lade_excel_datei(excel_pfad)
        st.session_state['df'] = df
    else:
        df = st.session_state['df']

    if df is not None:
        st.success("Daten erfolgreich geladen!")
        st.subheader("Suche nach OPS-Text oder Handelsnamen")
        suchbegriff = st.text_input("Suchbegriff (Teil des OPS-Textes oder Handelsnamens)")

        if st.button("Suchen"):
            if suchbegriff.strip() == "":
                st.warning("Bitte geben Sie einen gültigen Suchbegriff ein.")
            else:
                ergebnisse = suche_daten(df, suchbegriff)
                if not ergebnisse.empty:
                    st.success(f"{len(ergebnisse)} Einträge gefunden.")
                    # Auswahl der relevanten Spalten
                    relevante_spalten = ['ZE', 'OPS', 'OPS-Text', 'Handelsnamen', 'Wirkstoffklasse', 'Infos', 'Betrag']
                    vorhandene_spalten = [spalte for spalte in relevante_spalten if spalte in ergebnisse.columns]
                    fehlende_spalten = [spalte for spalte in relevante_spalten if spalte not in ergebnisse.columns]

                    if fehlende_spalten:
                        st.warning(
                            f"Die folgenden Spalten fehlen in den Daten und werden nicht angezeigt: {', '.join(fehlende_spalten)}")

                    # Anzeige der Ergebnisse
                    st.dataframe(ergebnisse[vorhandene_spalten].reset_index(drop=True))

                    # Download-Option
                    csv = ergebnisse[vorhandene_spalten].to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Ergebnisse als CSV herunterladen",
                        data=csv,
                        file_name='ergebnisse.csv',
                        mime='text/csv',
                    )
                else:
                    st.info("Keine passenden Einträge gefunden.")
    else:
        st.error("Die Excel-Datei konnte nicht geladen werden. Bitte überprüfen Sie den Pfad und die Datei.")

if __name__ == "__main__":
    main()
