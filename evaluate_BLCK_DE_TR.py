"""
TextSimilarityGrader (https://github.com/robomustib/TextSimilarityGrader/)
Copyright (c) 2026 Mustafa Bilgin
Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
Add-on: Blacklist & Multi-Word (DE/TR)
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
EXCEL_FILE = "Solutions_BLCK.xlsx"
OUTPUT_FILE = "Grading_Results_BLCK.xlsx"
SCORING_MODE = "fuzzy"

# Threshold = 70%
FUZZY_THRESHOLD = 0.70 

# Language toggle: "DE" for German, "TR" for Turkish
EVALUATION_LANGUAGE = "TR" 

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

def print_banner():
    print("="*50)
    print("   TRANSCRIPT EVALUATOR (DE/TR FINAL EDITION)")
    print("="*50)
    print(f" Transcript Folder: {TRANSCRIPT_FOLDER}")
    print(f" Solutions File:    {EXCEL_FILE}")
    print(f" Language:          {EVALUATION_LANGUAGE}")
    print(f" Threshold:         {FUZZY_THRESHOLD}")
    print("="*50 + "\n")

def clean_text(text):
    if not isinstance(text, str): return ""
    if EVALUATION_LANGUAGE == "TR":
        text = text.replace("I", "ı").replace("İ", "i")
    text = text.lower().strip()
    text = text.replace("ß", "ss")
    text = re.sub(r'[^\w\säöüçğış]', '', text, flags=re.IGNORECASE)
    return ' '.join(text.split())

def check_forbidden(forbidden_input, actual_text, matched_word=None, matched_target=None):
    if pd.isna(forbidden_input) or str(forbidden_input).strip() == "":
        return False, None
        
    forbidden_list = [f.strip() for f in str(forbidden_input).split(",")]
    clean_transcript = clean_text(actual_text)
    transcript_words = clean_transcript.split()
    
    for f_word in forbidden_list:
        clean_f = clean_text(f_word)
        if not clean_f: continue
        
        found_in_transcript = False
        count_in_transcript = 0
        
        if " " in clean_f:
            padded_transcript = f" {clean_transcript} "
            padded_f = f" {clean_f} "
            if padded_f in padded_transcript:
                found_in_transcript = True
                count_in_transcript = padded_transcript.count(padded_f)
        else:
            if clean_f in transcript_words:
                found_in_transcript = True
                count_in_transcript = transcript_words.count(clean_f)
                
        if found_in_transcript:
            if matched_word and matched_target:
                clean_match = clean_text(matched_word)
                clean_target = clean_text(matched_target)

                count_in_match = 0
                if " " in clean_f:
                    padded_match = f" {clean_match} "
                    padded_f = f" {clean_f} "
                    count_in_match = padded_match.count(padded_f)
                else:
                    match_words_list = clean_match.split()
                    count_in_match = match_words_list.count(clean_f)
                    
                is_in_target = False
                if " " in clean_f:
                    padded_target = f" {clean_target} "
                    padded_f = f" {clean_f} "
                    if padded_f in padded_target:
                        is_in_target = True
                else:
                    if clean_f in clean_target.split():
                        is_in_target = True
                        
                if count_in_transcript > 0 and (count_in_transcript == count_in_match) and is_in_target:
                    continue 
                    
            return True, f_word
            
    return False, None

def find_best_match(target_input, actual, mode):
    targets = [t.strip() for t in str(target_input).split(",")]

    overall_best_word = None
    overall_best_sim = 0.0
    overall_points = 0
    overall_best_target = None

    actual_words_orig = actual.split()
    if not actual_words_orig:
        return None, 0, 0, None

    for target in targets:
        t_clean = clean_text(target)
        if not t_clean: continue

        target_len = len(t_clean.split())
        current_target_best_sim = 0.0
        current_target_best_word = None
        
        n_grams = []
        if target_len > 0 and len(actual_words_orig) >= target_len:
            for i in range(len(actual_words_orig) - target_len + 1):
                n_grams.append(" ".join(actual_words_orig[i:i+target_len]))
        else:
            n_grams = [" ".join(actual_words_orig)]

        for w_orig in n_grams:
            w_clean = clean_text(w_orig)
            current_sim = 0.0
            
            if mode == "strict":
                current_sim = 100.0 if t_clean == w_clean else 0.0
            elif mode == "fuzzy":
                if t_clean in w_clean:
                    current_sim = 100.0
                else:
                    current_sim = SequenceMatcher(None, t_clean, w_clean).ratio() * 100
            
            if current_sim > current_target_best_sim:
                current_target_best_sim = current_sim
                current_target_best_word = w_orig 

        current_points = 0
        if current_target_best_sim >= (FUZZY_THRESHOLD * 100):
            current_points = 1
        
        # Stricter rule for short words
        if len(t_clean) <= 3:
            if current_target_best_sim < 90: 
                 current_points = 0
            else:
                 current_points = 1

        # Tie-Breaker Logic
        if current_target_best_sim > overall_best_sim:
            overall_best_sim = current_target_best_sim
            overall_best_word = current_target_best_word
            overall_points = current_points
            overall_best_target = target
        elif current_target_best_sim == overall_best_sim and current_target_best_sim > 0:
            if current_target_best_word and overall_best_word:
                if len(current_target_best_word.split()) > len(overall_best_word.split()):
                    overall_best_word = current_target_best_word
                    overall_points = current_points
                    overall_best_target = target

    return overall_best_word, overall_best_sim, overall_points, overall_best_target

def extract_from_json(content):
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
        if len(df.columns) < 2: 
            raise ValueError("Too few columns")
            
        df.columns.values[0] = "Filename"
        df.columns.values[1] = "Target_Text"
        
        if len(df.columns) > 2:
             df.columns.values[2] = "Forbidden"
        else:
             df["Forbidden"] = "" 

    except Exception as e:
        print(f" Excel Error: {e}")
        input("\nPress ENTER to exit...")
        return

    results = []
    print(f" Starting evaluation for {len(df)} entries...\n")

    for index, row in df.iterrows():
        raw_filename = str(row["Filename"]).strip()
        
        if raw_filename.startswith("_"):
            continue

        raw_target = row["Target_Text"]
        target = "" if (pd.isna(raw_target) or str(raw_target).strip().lower() == "nan") else str(raw_target).strip()

        raw_forbidden = row["Forbidden"]
        
        base_name = Path(raw_filename).stem
        found = False
        actual_raw = "[NOT FOUND]"
        
        for ext in [".json", ".txt"]:
            p = TRANSCRIPT_FOLDER / (base_name + ext)
            if p.exists():
                actual_raw, found = get_file_content(p)
                if found: break
        
        ist_display = ""
        points = 0
        similarity = 0
        status_msg = "OK"

        if found and target:
            match_word, similarity, points, matched_target_str = find_best_match(target, actual_raw, SCORING_MODE)
            
            is_forbidden, forbidden_word_found = check_forbidden(
                raw_forbidden, 
                actual_raw, 
                match_word if points > 0 else None, 
                matched_target_str if points > 0 else None
            )
            
            if is_forbidden:
                points = 0
                similarity = 0 
                ist_display = f"FORBIDDEN: {forbidden_word_found}"
                status_msg = "FORBIDDEN"
            else:
                if match_word:
                    ist_display = match_word
                else:
                    ist_display = "-"
        else:
            ist_display = "-"
            points = 0
            similarity = 0
            if not found: status_msg = "MISSING FILE"
            elif not target: status_msg = "NO TARGET"

        results.append({
            "Filename": raw_filename,
            "Target": target,
            "Forbidden": raw_forbidden, 
            "Actual (Found Word)": ist_display,
            "Transcript (Full Sentence)": actual_raw if found else "[MISSING]",
            "Points": points,
            "Similarity (%)": round(similarity, 1),
            "Status": status_msg
        })

    df_result = pd.DataFrame(results)
    correct = df_result["Points"].sum()
    
    valid_entries = df_result[ (df_result["Target"] != "") & (df_result["Transcript (Full Sentence)"] != "[MISSING]") ]
    valid_count = len(valid_entries)
    
    quote = (correct / valid_count * 100) if valid_count > 0 else 0
    
    print("\n" + "="*30)
    print(f" RESULTS")
    print(f"   Total files:     {len(results)}")
    print(f"   Valid entries:   {valid_count}")
    print(f"   Points awarded:  {correct}")
    print(f"   Success rate:    {quote:.1f}%")
    print(f"   Duration:        {time.time() - start_time:.2f} sec")
    print("="*30)
    
    try:
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"\n Successfully saved as: {OUTPUT_FILE}")
    except Exception as e:
        print(f"\n Error saving file: {e}")

    input("\n Done. Press ENTER to close...")

if __name__ == "__main__":
    main()
