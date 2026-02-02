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

def find_best_match(target, actual, mode):
    """
    Searches for the best matching word in the sentence and returns the ORIGINAL word.
    """
    t_clean = clean_text(target)
    
    # Split original text to return the original word
    actual_words_orig = actual.split()
    
    if not actual_words_orig:
        return None, 0, 0

    best_match_word = None
    best_similarity = 0.0
    
    # Check each word in the sentence
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

    # Assign points?
    points = 0
    if best_similarity >= (FUZZY_THRESHOLD * 100):
        points = 1
    
    # Special case for short words (<= 3 chars)
    if len(t_clean) <= 3:
        if best_similarity < 85: 
             points = 0
        else:
             points = 1

    return best_match_word, best_similarity, points

def extract_from_json(content):
    """Extracts pure text from Gladia JSON (Robust)."""
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
