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
SCORING_MODE = "fuzzy"
FUZZY_THRESHOLD = 0.75 

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

def extract_keywords_from_text(text, min_length=3):
    """Extrahiert alle Wörter aus dem Text für Keyword-Suche."""
    words = re.findall(r'\b[\wäöüßÄÖÜ-]+\b', text.lower())
    # Filtere sehr kurze Wörter (falls gewünscht)
    return [w for w in words if len(w) >= min_length]

def find_best_match(target, actual, mode):
    """
    Findet das beste Matching zwischen Target-Keyword und Wörtern im Text.
    Gibt zurück: (gefundenes_wort, ähnlichkeit, punkte)
    """
    t_clean = clean_text(target)
    a_clean = clean_text(actual)
    
    # Falls kein Text gefunden wurde
    if not a_clean or a_clean in ["[nicht gefunden]", "[datei nicht lesbar]", "[datei fehlt]"]:
        return None, 0, 0
    
    # Wörter aus dem tatsächlichen Text extrahieren
    actual_words = extract_keywords_from_text(a_clean)
    
    # MODUS: strict
    if mode == "strict":
        if t_clean in actual_words:
            return t_clean, 100, 1
        else:
            return None, 0, 0
    
    # MODUS: contains
    elif mode == "contains":
        # Prüfe ob Keyword in einem der Wörter enthalten ist
        for word in actual_words:
            if t_clean in word:
                # Berechne Ähnlichkeit basierend auf Längenverhältnis
                similarity = len(t_clean) / len(word) * 100
                return word, similarity, 1
        return None, 0, 0
    
    # MODUS: fuzzy (EMPFEHLENSWERT)
    elif mode == "fuzzy":
        # Für den Fall, dass der Text sehr kurz ist (nur ein Wort)
        if len(actual_words) == 1 and len(actual_words[0]) <= 10:
            word = actual_words[0]
            similarity = SequenceMatcher(None, t_clean, word).ratio()
            points = 1 if similarity >= FUZZY_THRESHOLD else 0
            return word, similarity * 100, points
        
        # Levenshtein-Distanz für jedes Wort berechnen
        best_match = None
        best_similarity = 0
        
        for word in actual_words:
            similarity = SequenceMatcher(None, t_clean, word).ratio()
            if similarity > best_similarity:
                best_similarity = similarity
                best_match = word
        
        # Punkte basierend auf Schwellenwert berechnen
        points = 1 if best_similarity >= FUZZY_THRESHOLD else 0
        
        # Spezialfall: Sehr kurze Wörter (<= 3 Buchstaben)
        if len(t_clean) <= 3:
            # Für kurze Wörter strenger
            if t_clean in a_clean:  # Direkt im Text suchen (nicht nur in Wörtern)
                points = 1
                best_similarity = 1.0
                best_match = t_clean
            elif best_similarity >= 0.9:  # Höherer Schwellenwert für kurze Wörter
                points = 1
            else:
                points = 0
        
        return best_match, best_similarity * 100, points
    
    return None, 0, 0

def calculate_score(target, actual, mode):
    """Vergibt Punkte basierend auf dem Modus."""
    best_match, similarity, points = find_best_match(target, actual, mode)
    
    # Formatiere Ausgabe für "Ist (Gladia)" Spalte
    if best_match:
        if similarity >= 100:
            ist_display = best_match
        elif similarity >= 85:
            ist_display = f"{best_match} ({similarity:.0f}%)"
        else:
            ist_display = f"[{best_match}] ({similarity:.0f}%)"
    else:
        ist_display = "[NICHT GEFUNDEN]"
    
    return points, similarity, ist_display

def extract_from_json(content):
    """Holt den reinen Text aus dem Gladia-JSON."""
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
        
        # Punkte berechnen (jetzt mit verbesserter Funktion)
        points, sim, ist_display = calculate_score(target_text, actual_text, SCORING_MODE)
        
        results.append({
            "Dateiname": raw_filename,
            "Soll (Excel)": target_text,
            "Ist (Gladia)": ist_display,
            "Punkte": points,
            "Ähnlichkeit (%)": round(sim, 1),
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
    
    # Detailausgabe
    print("\nDETAILS:")
    print("-" * 80)
    print(f"{'Dateiname':<30} {'Soll':<15} {'Ist':<20} {'Ähnl.':<8} {'Pkt':<4} {'Status':<6}")
    print("-" * 80)
    
    for result in results:
        print(f"{result['Dateiname']:<30} "
              f"{result['Soll (Excel)']:<15} "
              f"{str(result['Ist (Gladia)'])[:20]:<20} "
              f"{result['Ähnlichkeit (%)']:<8.1f} "
              f"{result['Punkte']:<4} "
              f"{result['Status']:<6}")
    
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
