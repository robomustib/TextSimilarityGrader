import json
import os
from pathlib import Path

test_cases = [
    # ============================================================
    # TEST 1: Tippfehler "Apfell" statt "Apfel"
    # ============================================================
    {
        "filename": "test_apfel_tippfehler.json",
        "description": "Soll-Text: 'Apfel', Tippfehler: 'Apfell' (doppeltes L), Ähnlichkeit: ~91%",
        "content": {
            "metadata": {"audio_duration": 2.5},
            "transcription": {
                "full_transcript": "Ich habe einen Apfell gekauft.",
                "languages": ["de"],
                "utterances": [{"text": "Ich habe einen Apfell gekauft.", "language": "de", "confidence": 0.88}]
            }
        }
    },
    
    # ============================================================
    # TEST 2: Korrekte Transkription "Katze" - Test für Buchstabendreher-Simulation
    # ============================================================
    {
        "filename": "test_katze_tippfehler.json", 
        "description": "Soll-Text: 'Katze', ASR-Output korrekt. Für Tests: Soll-Text könnte 'Katzte' (Buchstabendreher) sein",
        "content": {
            "transcription": {
                "full_transcript": "Die Katze sitzt auf dem Fensterbrett.",
                "languages": ["de"]
            }
        }
    },
    
    # ============================================================
    # TEST 3: Tippfehler "Buß" statt "Bus" - kurzes Wort, Speziallogik benötigt
    # ============================================================
    {
        "filename": "test_bus_tippfehler.json",
        "description": "Soll-Text: 'Bus', Tippfehler: 'Buß' (ß statt ss), Ähnlichkeit: ~67% (kurzes Wort → Speziallogik)",
        "content": {
            "result": {
                "transcription": {
                    "utterances": [{"text": "Ich fahre mit dem Bus zur Schule", "language": "de"}],
                    "full_transcript": "Ich fahre mit dem Buß zur Schule"
                }
            }
        }
    },
    
    # ============================================================
    # TEST 4: Fehlendes Leerzeichen "istmein" - "Haus" trotzdem enthalten
    # ============================================================
    {
        "filename": "test_haus_tippfehler.json",
        "description": "Soll-Text: 'Haus', Tippfehler: 'istmein' statt 'ist mein' (kein Leerzeichen). Extrahiert: 'Das istmein Haus.' → 'Haus' ist trotzdem enthalten",
        "content": {
            "audio_url": "https://example.com/test.wav",
            "full_transcript": "Das istmein Haus.",
            "language": "de",
            "confidence": 0.85
        }
    },
    
    # ============================================================
    # TEST 5: Tippfehler "Bibliotek" statt "Bibliothek"
    # ============================================================
    {
        "filename": "test_bibliothek_tippfehler.json",
        "description": "Soll-Text: 'Bibliothek', Tippfehler: 'Bibliotek' (k statt kh), Ähnlichkeit: ~90%",
        "content": {
            "metadata": {"audio_duration": 4.2},
            "transcription": {
                "full_transcript": "Die Bibliotek hat viele interessante Büecher.",
                "languages": ["de"],
                "utterances": [{"text": "Die Bibliotek hat viele interessante Büecher.", "language": "de", "confidence": 0.82}]
            }
        }
    },
    
    # ============================================================
    # TEST 6: Phonetischer Tippfehler "Schbielplatz" statt "Spielplatz"
    # ============================================================
    {
        "filename": "test_spielplatz_tippfehler.json",
        "description": "Soll-Text: 'Spielplatz', Tippfehler: 'Schbielplatz' (phonetische Transkription), Ähnlichkeit: ~73% (unter 85% → 0 Punkte bei Fuzzy)",
        "content": {
            "text": "Schbielplatz",
            "status": "success"
        }
    },
    
    # ============================================================
    # TEST 7: Mehrsprachige Transkription mit Sprecherwechsel
    # ============================================================
    {
        "filename": "test_anna.json",
        "description": "Erwartete Lösung: 'Anna'. Test: Verschachtelte Struktur + Code-Switching (de/en) + Sprecher-ID",
        "content": {
            "result": {
                "transcription": {
                    "utterances": [
                        {
                            "text": "Hello, my name is",
                            "language": "en",
                            "speaker": 0
                        },
                        {
                            "text": "Mein Name ist Anna",
                            "language": "de", 
                            "speaker": 1
                        },
                        {
                            "text": "und ich komme aus Berlin",
                            "language": "de",
                            "speaker": 1
                        }
                    ]
                }
            }
        }
    },
    
    # ============================================================
    # TEST 8: Standard Gladia-Struktur - full_transcript vorhanden
    # ============================================================
    {
        "filename": "test_apfel.json",
        "description": "Erwartete Lösung: 'Apfel'. Test: full_transcript vorhanden (Standard Gladia-Struktur)",
        "content": {
            "metadata": {
                "audio_duration": 5.32,
                "transcription_time": 3.45
            },
            "transcription": {
                "full_transcript": "Der Apfel liegt auf dem Tisch.",
                "languages": ["de"],
                "utterances": [
                    {
                        "text": "Der Apfel liegt auf dem Tisch.",
                        "language": "de",
                        "confidence": 0.95
                    }
                ]
            }
        }
    },
    
    # ============================================================
    # TEST 9: Einfachste JSON-Struktur - Minimal-Test
    # ============================================================
    {
        "filename": "test_banane.json",
        "description": "Erwartete Lösung: 'Banane'. Test: Einfachste mögliche Struktur",
        "content": {
            "text": "Banane",
            "status": "success"
        }
    },
    
    # ============================================================
    # TEST 10: Nur utterances - muss kombiniert werden
    # ============================================================
    {
        "filename": "test_hund.json",
        "description": "Erwartete Lösung: 'Hund'. Test: Nur utterances → muss kombiniert werden (kein full_transcript)",
        "content": {
            "metadata": {
                "audio_duration": 3.21
            },
            "transcription": {
                "languages": ["de"],
                "utterances": [
                    {
                        "text": "Hallo",
                        "language": "de",
                        "confidence": 0.98
                    },
                    {
                        "text": "ich suche den Hund",
                        "language": "de",
                        "confidence": 0.92
                    },
                    {
                        "text": "er ist braun",
                        "language": "de",
                        "confidence": 0.87
                    }
                ]
            }
        }
    }
]

# Ordner erstellen
Path("transcripts").mkdir(exist_ok=True)

# Metadaten-Datei mit allen Testbeschreibungen erstellen
metadata = {
    "test_cases": [
        {
            "filename": test["filename"],
            "description": test["description"],
            "expected_keywords": {
                "test_apfel_tippfehler.json": "Apfel",
                "test_katze_tippfehler.json": "Katze", 
                "test_bus_tippfehler.json": "Bus",
                "test_haus_tippfehler.json": "Haus",
                "test_bibliothek_tippfehler.json": "Bibliothek",
                "test_spielplatz_tippfehler.json": "Spielplatz",
                "test_anna.json": "Anna",
                "test_apfel.json": "Apfel",
                "test_banane.json": "Banane",
                "test_hund.json": "Hund"
            }[test["filename"]]
        }
        for test in test_cases
    ],
    "summary": {
        "total_tests": len(test_cases),
        "tippfehler_tests": 6,
        "struktur_tests": 4,
        "challenges": [
            "Verschiedene JSON-Strukturen verarbeiten",
            "Tippfehler in Transkriptionen erkennen", 
            "Mehrsprachige Inhalte extrahieren",
            "Kurze vs. lange Wörter behandeln",
            "Phonetische Abweichungen bewerten"
        ]
    }
}

# JSON-Dateien schreiben
print("Erstelle Test-Dateien...")
for test in test_cases:
    with open(f"transcripts/{test['filename']}", "w", encoding="utf-8") as f:
        json.dump(test["content"], f, indent=2, ensure_ascii=False)
    print(f"  ✓ {test['filename']}: {test['description'][:60]}...")

# TXT-Datei erstellen (kein JSON)
with open("transcripts/test_apfel.txt", "w", encoding="utf-8") as f:
    f.write("Der Apfel liegt auf dem Tisch. Dies ist ein Test.")
print(f"  ✓ test_apfel.txt: Reine Textdatei zum Testen")

# Metadaten-Datei speichern
with open("transcripts/_test_metadata.json", "w", encoding="utf-8") as f:
    json.dump(metadata, f, indent=2, ensure_ascii=False)

print(f"\n{'='*60}")
print(f"ERFOLG: {len(test_cases) + 1} Test-Dateien erstellt! (inkl. TXT)")
print(f"{'='*60}\n")

# Übersicht anzeigen
print("TEST-ÜBERSICHT:")
print("-" * 60)
for i, test in enumerate(test_cases, 1):
    print(f"{i:2d}. {test['filename']:30} → Erwartet: {metadata['test_cases'][i-1]['expected_keywords']}")
print("-" * 60)

print("\nOrdner 'transcripts' enthält:")
for f in sorted(os.listdir("transcripts")):
    if not f.startswith("_"):  # Metadaten-Datei nicht in Hauptliste
        print(f"   {f}")

print(f"\n Metadaten: transcripts/_test_metadata.json")
