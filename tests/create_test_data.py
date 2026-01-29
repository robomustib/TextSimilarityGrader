import json
import os
from pathlib import Path

# Define test cases with English content
test_cases = [
    {
        "filename": "test_apple_typo.json",
        "description": "Expected: 'Apple', Typo: 'Appple' (double p), Similarity: ~91%",
        "content": {
            "metadata": {"audio_duration": 2.5},
            "transcription": {
                "full_transcript": "I bought an Appple yesterday.",
                "languages": ["en"],
                "utterances": [{"text": "I bought an Appple yesterday.", "language": "en", "confidence": 0.88}]
            }
        }
    },
    {
        "filename": "test_cat_clean.json", 
        "description": "Expected: 'Cat', ASR correct. Setup for testing potential transposition errors.",
        "content": {
            "transcription": {
                "full_transcript": "The cat sits on the mat.",
                "languages": ["en"]
            }
        }
    },
    {
        "filename": "test_bus_typo.json",
        "description": "Expected: 'Bus', Typo: 'Buss' (phonetic error), Similarity: high but not exact.",
        "content": {
            "result": {
                "transcription": {
                    "utterances": [{"text": "I take the Buss to school", "language": "en"}],
                    "full_transcript": "I take the Buss to school"
                }
            }
        }
    },
    {
        "filename": "test_house_nospace.json",
        "description": "Expected: 'House', Typo: 'ismy' instead of 'is my'. Keyword 'House' is intact.",
        "content": {
            "audio_url": "https://example.com/test.wav",
            "full_transcript": "This ismy House.",
            "language": "en",
            "confidence": 0.85
        }
    },
    {
        "filename": "test_library_typo.json",
        "description": "Expected: 'Library', Typo: 'Libary' (common misspelling), Similarity: ~90%",
        "content": {
            "metadata": {"audio_duration": 4.2},
            "transcription": {
                "full_transcript": "The libary has many books.",
                "languages": ["en"],
                "utterances": [{"text": "The libary has many books.", "language": "en", "confidence": 0.82}]
            }
        }
    },
    {
        "filename": "test_playground_phonetic.json",
        "description": "Expected: 'Playground', Typo: 'Plaiground' (phonetic), Similarity: ~90%",
        "content": {
            "text": "Plaiground",
            "status": "success"
        }
    },
    {
        "filename": "test_anna_structure.json",
        "description": "Expected: 'Anna'. Test: Nested structure + Multi-speaker",
        "content": {
            "result": {
                "transcription": {
                    "utterances": [
                        {
                            "text": "Hello, who are you?",
                            "language": "en",
                            "speaker": 0
                        },
                        {
                            "text": "My name is Anna",
                            "language": "en", 
                            "speaker": 1
                        },
                        {
                            "text": "nice to meet you",
                            "language": "en",
                            "speaker": 0
                        }
                    ]
                }
            }
        }
    },
    {
        "filename": "test_apple_standard.json",
        "description": "Expected: 'Apple'. Test: Standard structure, correct text.",
        "content": {
            "metadata": {
                "audio_duration": 5.32
            },
            "transcription": {
                "full_transcript": "The apple is on the table.",
                "languages": ["en"],
                "utterances": [
                    {
                        "text": "The apple is on the table.",
                        "language": "en",
                        "confidence": 0.95
                    }
                ]
            }
        }
    },
    {
        "filename": "test_banana_simple.json",
        "description": "Expected: 'Banana'. Test: Simplest possible JSON structure.",
        "content": {
            "text": "Banana",
            "status": "success"
        }
    },
    {
        "filename": "test_dog_utterances.json",
        "description": "Expected: 'Dog'. Test: Only utterances, no full transcript.",
        "content": {
            "metadata": {
                "audio_duration": 3.21
            },
            "transcription": {
                "languages": ["en"],
                "utterances": [
                    {
                        "text": "Hello",
                        "language": "en",
                        "confidence": 0.98
                    },
                    {
                        "text": "I am looking for the dog",
                        "language": "en",
                        "confidence": 0.92
                    },
                    {
                        "text": "it is brown",
                        "language": "en",
                        "confidence": 0.87
                    }
                ]
            }
        }
    }
]

# Ensure directory exists
# Note: Going one level up if run from tests/ folder, or creating local transcripts folder
output_dir = Path("transcripts")
output_dir.mkdir(exist_ok=True)

# Map filenames to expected keywords for metadata
keyword_map = {
    "test_apple_typo.json": "Apple",
    "test_cat_clean.json": "Cat", 
    "test_bus_typo.json": "Bus",
    "test_house_nospace.json": "House",
    "test_library_typo.json": "Library",
    "test_playground_phonetic.json": "Playground",
    "test_anna_structure.json": "Anna",
    "test_apple_standard.json": "Apple",
    "test_banana_simple.json": "Banana",
    "test_dog_utterances.json": "Dog"
}

metadata = {
    "test_cases": [
        {
            "filename": test["filename"],
            "description": test["description"],
            "expected_keywords": keyword_map[test["filename"]]
        }
        for test in test_cases
    ],
    "summary": {
        "total_tests": len(test_cases),
        "notes": "Generated mock data with English content for validation."
    }
}

print("Creating mock transcript files...")
for test in test_cases:
    file_path = output_dir / test['filename']
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(test["content"], f, indent=2, ensure_ascii=False)
    print(f" {test['filename']}")

# Create a dummy TXT file
txt_path = output_dir / "test_apple_plain.txt"
with open(txt_path, "w", encoding="utf-8") as f:
    f.write("The Apple is on the table. This is a plain text test.")
print(f" test_apple_plain.txt")

# Save metadata
meta_path = output_dir / "_test_metadata.json"
with open(meta_path, "w", encoding="utf-8") as f:
    json.dump(metadata, f, indent=2, ensure_ascii=False)

print(f"\n{'='*60}")
print(f"SUCCESS: {len(test_cases) + 1} test files created in '{output_dir}'")
print(f"Metadata saved to: {meta_path}")
print(f"{'='*60}\n")
