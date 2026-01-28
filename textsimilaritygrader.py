import pandas as pd
import re
import os
import glob
import time
import json
from pathlib import Path
from difflib import SequenceMatcher

# ==========================================
# 1. EINSTELLUNGEN
# ==========================================

TRANSCRIPT_FOLDER = Path("./transcripts")
EXCEL_FILE = "Lösungen.xlsx"
OUTPUT_FILE = "Auswertung_Ergebnisse.xlsx"

# "fuzzy" = Tippfehler erlaubt (Empfohlen!)
# "strict" = Muss zu 100% stimmen
SCORING_MODE = "fuzzy"
FUZZY_THRESHOLD = 0.85 

# ==========================================
# 2. HILFSFUNKTIONEN
# ==========================================

def print_banner():
    print("="*50)
    print("   TRANSCRIPT EVALUATOR (Automatische Auswertung)")
    print("="*50)
    print(f" Transkript-Ordner: {TRANSCRIPT_FOLDER}")
    print(f" Lösungs-Datei:     {EXCEL_FILE}")
    print(f" Bewertungs-Modus: {SCORING_MODE}")
    print("="*50 + "\n")

def clean_text(text):
    """Bereinigt Text: Alles klein, nur Buchstaben/Zahlen (auch Umlaute)."""
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    # Erlaubt a-z, 0-9 und äöüß. Alles andere (Punkt, Komma) fliegt raus.
    text = re.sub(r'[^\w\säöüß]', '', text, flags=re.IGNORECASE)
    # Doppelte Leerzeichen reduzieren
    text = ' '.join(text.split())
    return text

def extract_from_json(content):
    """Holt den reinen Text aus dem Gladia-JSON (Robust)."""
    try:
        data = json.loads(content)
        
        # Zuerst: Volles Transkript direkt suchen (falls Struktur flach)
        if isinstance(data, dict) and "full_transcript" in data:
            return data["full_transcript"]
        
        # Gladia-Standard: unter "transcription"
        transcription = data.get("transcription", {})
        
        if isinstance(transcription, dict):
            # Fall 1: Volles Transkript vorhanden
            if "full_transcript" in transcription:
                return transcription["full_transcript"]
            
            # Fall 2: Utterances zusammenbauen
            elif "utterances" in transcription and isinstance(transcription["utterances"], list):
                utterances = transcription["utterances"]
                # Sicherstellen, dass es Strings sind
                return " ".join([str(u.get("text", "")) for u in utterances])
        
        # Fall 3: Falls das JSON unerwartet nur ein String ist
        if isinstance(data, str):
            return data
            
        # Fall 4: Notfall - gib den gesamten Inhalt als String zurück
        return str(data)
        
    except json.JSONDecodeError:
        return content  # War wohl doch kein JSON, gib Text zurück
    except Exception:
        return content  # Bei anderen Fehlern: Text zurückgeben

def get_file_content(filepath):
    """Liest Datei (txt/json) und behandelt Encoding-Probleme."""
    content = ""
    success = False

    # 1. Datei roh einlesen (Encoding-Fallback)
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read().strip()
            success = True
    except UnicodeDecodeError:
        try:
            with open(filepath, "r", encoding="latin-1") as f:
                content = f.read().strip()
                success = True
        except:
            return "[DATEI NICHT LESBAR]", False
    except FileNotFoundError:
        return "[DATEI FEHLT]", False

    # 2. Falls JSON -> Parsen
    if success and filepath.suffix.lower() == '.json':
        content = extract_from_json(content)

    return content, success

def calculate_score(target, actual, mode):
    """Vergibt Punkte basierend auf dem Modus."""
    t_clean = clean_text(target)
    a_clean = clean_text(actual)
    
    # Wie ähnlich sind sich die Texte? (0.0 bis 1.0)
    similarity = SequenceMatcher(None, t_clean, a_clean).ratio()
    points = 0
    
    if mode == "strict":
        points = 1 if t_clean == a_clean else 0
        
    elif mode == "contains":
        points = 1 if t_clean in a_clean else 0
        
    elif mode == "fuzzy":
        # SPEZIAL-LOGIK:
        # Wenn das Lösungswort sehr kurz ist (<= 4 Buchstaben, z.B. "Eis"),
        # ist Fuzzy gefährlich (verwechselt "Eis" mit "Ein").
        # Da nutzen wir lieber 'contains'.
        if len(t_clean) <= 4:
            points = 1 if t_clean in a_clean else 0
        else:
            # Sonst: Fuzzy Matching (erlaubt Tippfehler)
            points = 1 if similarity >= FUZZY_THRESHOLD or t_clean in a_clean else 0
            
    return points, similarity

# ==========================================
# 3. HAUPT-PROGRAMM
# ==========================================

def main():
    start_time = time.time()
    print_banner()
    
    # --- SCHRITT 1: Excel laden (mit Fehlerschutz) ---
    if not os.path.exists(EXCEL_FILE):
        print(f" FEHLER: Die Datei '{EXCEL_FILE}' wurde nicht gefunden!")
        print("  Bitte erstelle die Excel-Datei im gleichen Ordner.")
        input("\nDrücke ENTER zum Beenden...")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        # Spalten normalisieren
        if len(df.columns) < 2:
            raise ValueError("Die Excel-Datei braucht mindestens 2 Spalten!")
        
        df.columns.values[0] = "Dateiname"
        df.columns.values[1] = "Soll_Text"
        
    except PermissionError:
        print(f" FEHLER: Zugriff verweigert!")
        print(f"   Ist die Datei '{EXCEL_FILE}' vielleicht noch in Excel geöffnet?")
        print("   -> Bitte Excel schließen und nochmal versuchen.")
        input("\nDrücke ENTER zum Beenden...")
        return
    except Exception as e:
        print(f" Kritisches Problem mit der Excel-Datei: {e}")
        input("\nDrücke ENTER zum Beenden...")
        return

    # --- SCHRITT 2: Auswertung ---
    results = []
    total_rows = len(df)
    print(f" Starte Auswertung für {total_rows} Einträge...\n")

    for index, row in df.iterrows():
        # Kleiner Fortschritts-Indikator
        if index % 50 == 0 and index > 0:
            print(f"   ... verarbeite Zeile {index} von {total_rows}")

        raw_filename = str(row["Dateiname"]).strip()
        target_text = str(row["Soll_Text"]).strip()
        
        # Dateinamen ohne Endung holen (z.B. "Audio1.wav" -> "Audio1")
        base_name = Path(raw_filename).stem
        
        # Wir suchen flexibel nach .txt oder .json
        found = False
        actual_text = "[NICHT GEFUNDEN]"
        
        potential_files = [
            TRANSCRIPT_FOLDER / f"{base_name}.json",  # JSON zuerst probieren
            TRANSCRIPT_FOLDER / f"{base_name}.txt"
        ]
        
        for p in potential_files:
            if p.exists():
                content, success = get_file_content(p)
                if success:
                    actual_text = content
                    found = True
                    break
        
        # Punkte berechnen
        points, sim = calculate_score(target_text, actual_text, SCORING_MODE)
        
        results.append({
            "Dateiname": raw_filename,
            "Soll (Excel)": target_text,
            "Ist (Gladia)": actual_text,
            "Punkte": points,
            "Ähnlichkeit (%)": round(sim * 100, 1),
            "Status": "OK" if found else "FEHLT"
        })

    # --- SCHRITT 3: Speichern ---
    df_result = pd.DataFrame(results)
    
    # Statistik
    correct = df_result["Punkte"].sum()
    quote = (correct / total_rows * 100) if total_rows > 0 else 0
    duration = time.time() - start_time
    
    print("\n" + "="*30)
    print(f" ERGEBNIS")
    print(f"   Dateien gesamt:  {total_rows}")
    print(f"   Punkte vergeben: {correct}")
    print(f"   Trefferquote:    {quote:.1f}%")
    print(f"   Dauer:           {duration:.2f} Sekunden")
    print("="*30)
    
    try:
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"\n Erfolgreich gespeichert in:\n   {OUTPUT_FILE}")
    except PermissionError:
        print(f"\n FEHLER beim Speichern!")
        print(f"   Die Datei '{OUTPUT_FILE}' ist noch geöffnet.")
        print("   Bitte schließen und Skript neu starten.")
    except Exception as e:
        print(f"\n Unbekannter Fehler beim Speichern: {e}")

    # WICHTIG für Windows: Fenster offen lassen
    input("\n Fertig. Drücke ENTER zum Schließen...")

if __name__ == "__main__":
    main()
