# TextSimilarityGrader 
*Automated Fuzzy Evaluation for Speech-to-Text Transcripts*

A robust, automated grading tool designed to compare speech-to-text transcripts against expected answers using fuzzy logic. It is designed to grade large batches of files (e.g., from psychological or linguistic studies) where exact string matching fails due to typos, stuttering, or phonetic spelling errors in the transcript.

## Flowchart
<img src="https://github.com/robomustib/TextSimilarityGrader/blob/main/img/TextSimilarityGrader_Flowchart.svg" alt="Flowchart of TextSimilarityGrader" width="25%"/>

## Related Projects

This tool is the evaluation counterpart to the **[Gladia Batch Transcriber](https://github.com/robomustib/gladia_batch_transcriber)**.
* Use **Gladia Batch Transcriber** to generate transcripts from audio.
* Use **TextSimilarityGrader** (this repo) to evaluate and grade those transcripts.

## Key Features

* **Fuzzy Matching:** Tolerates spelling mistakes and ASR errors (e.g., "apple" vs. "appple" is graded as correct).
* **Smart Parsing:** Automatically finds the relevant text in complex JSON structures (Gladia, generic API responses) or plain `.txt` files.
* **Excel Workflow:** Automatically generates solution templates and exports detailed grading reports.
* **Format Agnostic:** Works with any folder containing `.json` or `.txt` transcripts.

## Project Structure

* `evaluate.py`: The core engine. Scans files, finds the spoken words, and grades them (0 or 1 point).
* `generate_excel.py`: Helper script to create the `Solutions.xlsx` template from your file list.
* `tests/`: Contains scripts to generate dummy data for testing the logic without real transcripts.

## Data Flowchart
<img src="https://github.com/robomustib/TextSimilarityGrader/blob/main/img/TextSimilarityGrader_Dataflowchart.svg" alt="Data Flowchart of TextSimilarityGrader" width="100%"/>

## Installation

**1. Clone the repository:**

```bash
git clone https://github.com/robomustib/TextSimilarityGrader.git
cd TextSimilarityGrader
```

**2. Install dependencies:**

```
   bash
   pip install -r requirements.txt
```

## Usage

**Step 1: Get your Transcripts**

Ensure your transcript files (.json or .txt) are in the transcripts/ folder. (If you need to generate transcripts from audio files first, check out the Gladia Batch Transcriber).

**Step 2: Generate the Solution Template**

Run the helper to scan your folder and create an Excel list:

```
   bash
   python generate_excel.py
```

This creates Solutions.xlsx.

**Step 3: Define Expected Answers**

Open Solutions.xlsx.

- Column A: Filename (auto-filled)
- Column B: Enter the expected keyword here (e.g., "House", "Apple")

Step 4: Run the Grader

```
   bash
   python evaluate.py
```

The tool will compare the actual content of your files against the expected words in Excel. Result: Open Grading_Results.xlsx to see points, similarity percentages, and the exact text found.

## Configuration

You can adjust the strictness of the grading in evaluate.py:

```
   Python
   # Threshold for fuzzy matching (0.0 to 1.0)
   # 0.75 allows for minor typos and phonetic differences
   FUZZY_THRESHOLD = 0.75
```

## Development & Testing

You can verify the grading logic without real data using the included test suite:

1. Generate Mock Data:

```
   bash
   python tests/create_test_data.py
```

Creates dummy files with intentional typos in transcripts/.

2. Generate Test Solutions:

```
   bash
   python tests/setup_test_excel.py
```

Creates a pre-filled Solutions.xlsx.

3. Run Validation:

```
   bash
   python evaluate.py
```

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Citation

If you use this software for your research, please cite it as follows:

**APA Format:**
> Bilgin, M. (2026). *TextSimilarityGrader* (Version 1.0.0) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.18422534

**BibTeX:**
```bibtex

@software{TextSimilarityGrader,
  author       = {Bilgin, Mustafa},
  title        = {TextSimilarityGrader: A Python Tool for Automated Fuzzy Evaluation of Speech-to-Text Transcripts in Research Contexts},
  year         = {2026},
  publisher    = {Zenodo},
  version      = {1.0.0},
  doi          = {10.5281/zenodo.18422534},
  url          = {https://doi.org/10.5281/zenodo.18422534}
}

```
