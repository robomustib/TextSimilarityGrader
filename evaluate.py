"""
TextSimilarityGrader (https://github.com/robomustib/TextSimilarityGrader/)
Copyright (c) 2026 Mustafa Bilgin
Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
"""

import pandas as pd
import re
import os
import time
import json
from pathlib import Path
from difflib import SequenceMatcher

# ==========================================
# 1. SETTINGS
# ==========================================

TRANSCRIPT_FOLDER = Path("./transcripts")
EXCEL_FILE = "Solutions.xlsx"
OUTPUT_FILE = "Grading_Results.xlsx"
SCORING_MODE = "fuzzy"
# Tolerance: 0.75 allows for typos/letter swaps
FUZZY_THRESHOLD = 0.75 

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

def print_banner():
    print("="*50)
    print("   TRANSCRIPT EVALUATOR (Automated Grading)")
    print("="*50)
    print(f" Transcript Folder: {TRANSCRIPT_FOLDER}")
    print(f" Solutions File:    {EXCEL_FILE}")
    print(f" Grading Mode:      {SCORING_MODE}")
    print("="*50 + "\n")

def clean_text(text):
    """Cleans text: Lowercase, alphanumeric only (including umlauts)."""
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    # IMPORTANT: Convert ß to ss so Bus/Buß is recognized e.g. German
    text = text.replace("ß", "ss")
    # Allows a-z, 0-9 and äöü. Everything else (punctuation) is removed.
    text = re.sub(r'[^\w\säöü]', '', text, flags=re.IGNORECASE)
    # Reduce double spaces
    text = ' '.join(text.split())
    return text

def find_best_match(target_input, actual, mode):
    """
    Sucht das beste Wort im Satz.
    NEU: Unterstützt mehrere Synonyme, getrennt durch Komma (z.B. "laufen, läuf").
    """
    # 1. Zelle am Komma aufsplitten -> Liste von Zielen erstellen
    # Z.B. "Schubkarre, Karre" -> ["Schubkarre", "Karre"]
    # str(target_input) sichert ab, falls Excel Zahlen (z.B. 2024) sendet
    targets = [t.strip() for t in str(target_input).split(",")]

    # Hier speichern wir das beste Ergebnis aller Varianten (Global für diesen Aufruf)
    overall_best_word = None
    overall_best_sim = 0.0
    overall_points = 0

    actual_words_orig = actual.split()
    
    # Wenn Transkript leer ist, sofort raus
    if not actual_words_orig:
        return None, 0, 0

    # 2. Jedes Ziel-Wort (Synonym) einzeln prüfen
    for target in targets:
        t_clean = clean_text(target)
        # Falls leeres Synonym (z.B. durch "Wort,,Wort"), überspringen
        if not t_clean: 
            continue

        # Lokale Bestwerte nur für DIESES Synonym
        current_target_best_sim = 0.0
        current_target_best_word = None
        
        # Jedes Wort im Transkript prüfen
        for w_orig in actual_words_orig:
            w_clean = clean_text(w_orig)
            current_sim = 0.0
            
            if mode == "strict":
                current_sim = 100.0 if t_clean == w_clean else 0.0
            elif mode == "contains":
                # Prüft, ob Ziel im Wort steckt (gut für "lauf" in "gelaufen")
                current_sim = 100.0 if t_clean in w_clean else 0.0
            elif mode == "fuzzy":
                if t_clean in w_clean:
                    current_sim = 100.0
                else:
                    current_sim = SequenceMatcher(None, t_clean, w_clean).ratio() * 100
            
            # Ist das aktuelle Wort im Satz ähnlicher als das vorige Wort im Satz?
            if current_sim > current_target_best_sim:
                current_target_best_sim = current_sim
                current_target_best_word = w_orig 

        # Punkte berechnen für dieses spezifische Synonym
        current_points = 0
        if current_target_best_sim >= (FUZZY_THRESHOLD * 100):
            current_points = 1
        
        # Spezialfall kurze Wörter (<= 3 Zeichen)
        if len(t_clean) <= 3:
            if current_target_best_sim < 85: 
                 current_points = 0
            else:
                 current_points = 1

        # 3. Ist dieses Synonym besser als die bisher getesteten Synonyme?
        # Beispiel: "Karre" hatte 50%, "Schubkarre" hat 90%. Wir nehmen die 90%.
        if current_target_best_sim > overall_best_sim:
            overall_best_sim = current_target_best_sim
            overall_best_word = current_target_best_word
            overall_points = current_points

    # Das Beste zurückgeben (gleiches Format wie früher!)
    return overall_best_word, overall_best_sim, overall_points

def extract_from_json(content):
    """Extracts pure text from Gladia JSON."""
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
    """Reads file (txt/json) and handles encoding issues."""
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
            return "[FILE UNREADABLE]", False

    if success and filepath.suffix.lower() == '.json':
        content = extract_from_json(content)
    return content, success

# ==========================================
# 3. MAIN PROGRAM
# ==========================================

def main():
    start_time = time.time()
    print_banner()
    
    if not os.path.exists(EXCEL_FILE):
        print(f" ERROR: File '{EXCEL_FILE}' not found!")
        input("\nPress ENTER to exit...")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        if len(df.columns) < 2: raise ValueError("Too few columns")
        df.columns.values[0] = "Filename"
        df.columns.values[1] = "Target_Text"
    except Exception as e:
        print(f" Excel Error: {e}")
        input("\nPress ENTER...")
        return

    results = []
    print(f" Starting evaluation for {len(df)} entries...\n")

    for index, row in df.iterrows():
        raw_filename = str(row["Filename"]).strip()
        
        # Ignore system files starting with underscore
        if raw_filename.startswith("_"):
            continue

        raw_target = row["Target_Text"]
        if pd.isna(raw_target) or str(raw_target).strip().lower() == "nan":
            target = ""
        else:
            target = str(raw_target).strip()

        base_name = Path(raw_filename).stem
        found = False
        actual_raw = "[NOT FOUND]"
        
        for ext in [".json", ".txt"]:
            p = TRANSCRIPT_FOLDER / (base_name + ext)
            if p.exists():
                actual_raw, found = get_file_content(p)
                if found: break
        
        # Grading
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
            "Filename": raw_filename,
            "Target": target,
            "Actual (Found Word)": ist_display,
            "Transcript (Full Sentence)": actual_raw if found else "[MISSING]",
            "Points": points,
            "Similarity (%)": round(similarity, 1),
            "Status": "OK" if found else "MISSING"
        })

    # Save Results
    df_result = pd.DataFrame(results)
    correct = df_result["Points"].sum()
    valid_count = len(df_result[df_result["Target"] != ""])
    quote = (correct / valid_count * 100) if valid_count > 0 else 0
    
    print("\n" + "="*30)
    print(f" RESULTS")
    print(f"   Total files:     {len(results)}")
    print(f"   Points awarded:  {correct}")
    print(f"   Success rate:    {quote:.1f}%")
    print(f"   Duration:        {time.time() - start_time:.2f} sec")
    print("="*30)
    
    try:
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"\n Successfully saved to: {OUTPUT_FILE}")
    except Exception as e:
        print(f"\n Error saving file: {e}")

    input("\n Done. Press ENTER to close...")

if __name__ == "__main__":
    main()
