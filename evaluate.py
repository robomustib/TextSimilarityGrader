import pandas as pd
import re
import os
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
# Toleranz: 0.75 erlaubt auch Buchstabendreher (z.B. Schbielplatz)
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
    # WICHTIG: ß zu ss machen, damit Bus/Buß erkannt wird
    text = text.replace("ß", "ss")
    # Erlaubt a-z, 0-9 und äöü. Alles andere (Punkt, Komma) fliegt raus.
    text = re.sub(r'[^\w\säöü]', '', text, flags=re.IGNORECASE)
    # Doppelte Leerzeichen reduzieren
    text = ' '.join(text.split())
    return text

def find_best_match(target, actual, mode):
    """
    Sucht das beste Wort im Satz und gibt das ORIGINAL-Wort zurück.
    """
    t_clean = clean_text(target)
    
    # Wir splitten den Original-Text, damit wir das Original-Wort zurückgeben können
    actual_words_orig = actual.split()
    
    if not actual_words_orig:
        return None, 0, 0

    best_match_word = None
    best_similarity = 0.0
    
    # Jedes Wort im Satz prüfen
    for w_orig in actual_words_orig:
        w_clean = clean_text(w_orig)
        
        current_sim = 0.0
        
        if mode == "strict":
            current_sim = 100.0 if t_clean == w_clean else 0.0
        elif mode == "contains":
            current_sim = 100.0 if t_clean in w_clean else 0.0
        elif mode == "fuzzy":
            if t_clean in w_clean:
                current_sim = 100.0
            else:
                current_sim = SequenceMatcher(None, t_clean, w_clean).ratio() * 100
        
        if current_sim > best_similarity:
            best_similarity = current_sim
            best_match_word = w_orig 

    # Punkte vergeben?
    points = 0
    if best_similarity >= (FUZZY_THRESHOLD * 100):
        points = 1
    
    # Spezialfall kurze Wörter (<= 3 Zeichen)
    if len(t_clean) <= 3:
        if best_similarity < 85: 
             points = 0
        else:
             points = 1

    return best_match_word, best_similarity, points

def extract_from_json(content):
    """Holt den reinen Text aus dem Gladia-JSON (Robust)."""
    try:
        data = json.loads(content)
        def find_text_in_obj(obj):
            if isinstance(obj, dict):
                if "full_transcript" in obj and obj["full_transcript"]: return obj["full_transcript"]
                if "text" in obj and isinstance(obj["text"], str): return obj["text"]
                if "transcription" in obj: return find_text_in_obj(obj["transcription"])
                if "result" in obj: return find_text_in_obj(obj["result"])
                if "utterances" in obj and isinstance(obj["utterances"], list):
                    return " ".join([str(u.get("text", "")) for u in obj["utterances"]])
            return None
        result = find_text_in_obj(data)
        return result if result else ""
    except:
        return content

def get_file_content(filepath):
    """Liest Datei (txt/json) und behandelt Encoding-Probleme."""
    content = ""
    success = False
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

    if success and filepath.suffix.lower() == '.json':
        content = extract_from_json(content)
    return content, success

# ==========================================
# 3. HAUPT-PROGRAMM
# ==========================================

def main():
    start_time = time.time()
    print_banner()
    
    if not os.path.exists(EXCEL_FILE):
        print(f" FEHLER: Datei '{EXCEL_FILE}' nicht gefunden!")
        input("\nDrücke ENTER...")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        if len(df.columns) < 2: raise ValueError("Zu wenige Spalten")
        df.columns.values[0] = "Dateiname"
        df.columns.values[1] = "Soll_Text"
    except Exception as e:
        print(f" Fehler Excel: {e}")
        input("\nDrücke ENTER...")
        return

    results = []
    print(f" Starte Auswertung für {len(df)} Einträge...\n")

    for index, row in df.iterrows():
        raw_filename = str(row["Dateiname"]).strip()
        
        if raw_filename.startswith("_"):
            continue

        raw_target = row["Soll_Text"]
        if pd.isna(raw_target) or str(raw_target).strip().lower() == "nan":
            target = ""
        else:
            target = str(raw_target).strip()

        base_name = Path(raw_filename).stem
        found = False
        actual_raw = "[NICHT GEFUNDEN]"
        
        for ext in [".json", ".txt"]:
            p = TRANSCRIPT_FOLDER / (base_name + ext)
            if p.exists():
                actual_raw, found = get_file_content(p)
                if found: break
        
        # Bewerten
        ist_display = ""
        points = 0
        similarity = 0

        if found and target:
            match_word, similarity, points = find_best_match(target, actual_raw, SCORING_MODE)
            
            if match_word:
                ist_display = match_word
            else:
                ist_display = "-"
        else:
            ist_display = "-"
            points = 0
            similarity = 0

        results.append({
            "Dateiname": raw_filename,
            "Soll": target,
            "Ist (Gefundenes Wort)": ist_display,
            "Transkript (Ganzer Satz)": actual_raw if found else "[FEHLT]",
            "Punkte": points,
            "Ähnlichkeit (%)": round(similarity, 1),
            "Status": "OK" if found else "FEHLT"
        })

    # Speichern
    df_result = pd.DataFrame(results)
    correct = df_result["Punkte"].sum()
    valid_count = len(df_result[df_result["Soll"] != ""])
    quote = (correct / valid_count * 100) if valid_count > 0 else 0
    
    print("\n" + "="*30)
    print(f" ERGEBNIS")
    print(f"   Dateien gesamt:  {len(results)}")
    print(f"   Punkte vergeben: {correct}")
    print(f"   Trefferquote:    {quote:.1f}%")
    print(f"   Dauer:           {time.time() - start_time:.2f} Sek")
    print("="*30)
    
    try:
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"\n Erfolgreich gespeichert in: {OUTPUT_FILE}")
    except Exception as e:
        print(f"\n Fehler beim Speichern: {e}")

    input("\n Fertig. Drücke ENTER zum Schließen...")

if __name__ == "__main__":
    main()
