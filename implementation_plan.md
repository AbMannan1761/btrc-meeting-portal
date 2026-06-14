# Implementation Plan - BTRC Commission Meeting Minutes Parser & Interactive Dashboard

This document outlines the research findings, database construction method, and user interface architecture for the BTRC Commission Meeting Minutes Reader app.

## Research Findings & Technical Strategy

1. **PDF Text Extraction Limitations**:
   - Our analysis of all **31 raw PDF files** showed **0 Bengali Unicode characters** in their text streams. These PDFs are either purely scanned images or have subsetted, randomized font encodings (meaning extracting text results in scrambled characters like `<IqErcTl`).
   - Windows OCR capability (`Language.OCR~~~bn-BD`) is not installed on this machine, and installer commands fail due to Session 0 / Sandboxed terminal execution permissions.
   - Word COM automation fails on PDFs because the PDF Reflow Converter blocks on non-interactive prompts.
   
2. **Alternative Data Goldmine**:
   - In the parent directory, we found several native Word (`.docx`, `.doc`) and Excel (`.xlsx`) files which contain highly structured, complete, and perfectly readable Bengali Unicode text of the meeting agendas, decisions, and E&I directorate reports.
   - We will extract and compile the meeting database from these files:
     - [Summery_Over all.xlsx](file:///c:/Users/mannan/OneDrive%20-%20Bangladesh%20Telecommunication%20Regulatory%20Commission/1.%20E%20&%20I/Meeting/Summery_Over%20all.xlsx): Contains the master index, dates, agenda numbers, and decisions for meetings **278 to 305**.
     - [CD-PS-15042026V1 from sharmin mam.xlsx](file:///c:/Users/mannan/OneDrive%20-%20Bangladesh%20Telecommunication Regulatory Commission/1. E & I/Meeting/CD-PS-15042026V1 from sharmin mam.xlsx): Contains detailed fine amounts and department responsibilities for meetings **287 to 304**.
     - [৩০০তম[2].doc](file:///c:/Users/mannan/OneDrive%20-%20Bangladesh%20Telecommunication Regulatory Commission/1. E & I/Meeting/Commission meeting/৩০০তম[2].doc): Full detailed agenda and background for the **300th meeting**.
     - [Agenda_302.docx](file:///c:/Users/mannan/OneDrive%20-%20Bangladesh%20Telecommunication Regulatory Commission/1. E & I/Meeting/Agenda_302.docx): Full detailed agenda and background for the **302nd meeting**.
     - [list of 307 draft.docx](file:///c:/Users/mannan/OneDrive%20-%20Bangladesh%20Telecommunication Regulatory Commission/1. E & I/Meeting/307/list of 307 draft.docx): Agenda summaries and reports for the **307th meeting**.

---

## Proposed Changes & Components

We will build the system in three steps:

### 1. Database Builder (Python Compiler)
We will write a python script `C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\compile_database.py` that:
- Reads `Summery_Over all.xlsx` Sheet1, cleaning the meeting numbers (e.g. `278.0` -> `২৭৮তম`) and formatting dates in Bengali context.
- Merges fine amounts and departments from `CD-PS-15042026V1 from sharmin mam.xlsx`.
- Integrates the rich text details, background, violation tables, and implementation responsibilities from `৩০০তম[2].doc` (after conversion to `.docx`), `Agenda_302.docx`, and `list of 307 draft.docx`.
- Sorts all meetings in ascending order (২৭৮তম -> ২৮০তম -> ২৮১তম -> ... -> ৩০৭তম).
- Exports the consolidated database as a structured JSON file.

### 2. Premium Single-File HTML Reader App
We will build a responsive, single-file HTML/JS/CSS reader app (`C:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\Commission meeting\BTRC_Meeting_Minutes_Reader.html`) featuring a premium design system:
- **Styling**: Sleek dark mode / glassmorphism theme, smooth gradients, Inter/Outfit typography via Google Fonts, and subtle micro-animations (transitions, hover effects).
- **Layer 1: Dashboard**: A grid of meetings sorted chronologically. Hovering/clicking a meeting card transitions the user to Layer 2.
- **Layer 2: Agenda Directory**: Displays a list of all agendas for the selected meeting (e.g. "(০১) ৩০০তম কমিশন সভার...", "বিবিধ আলোচ্যসূচী- ০১...").
- **Layer 3: Detail Viewer**: Shows:
  - **Subject (বিষয়)**: Clear title block.
  - **Decision (সিদ্ধান্ত)**: A highlighted card (notice block) with a clean card container.
  - **Details Accordion (`<details>` / `<summary>`)**: Slides open to reveal:
    - Presentation Background & Context (কার্যপত্র উপস্থাপন ও প্রেক্ষাপট)
    - Audit/stat/deviation tables (ব্যত্যয়ের ছক) - nicely styled HTML tables.
    - Implementation Responsibility (বাস্তবায়নে).

---

## Verification Plan

### Automated Tests
- Run `compile_database.py` to verify that the JSON database compiles successfully without data loss and contains all unique meetings.
- Verify that dates, fines, laws, and technical telecom terms (like ICX, IIG, TVAS) remain exact and in English text.

### Manual Verification
- Open `BTRC_Meeting_Minutes_Reader.html` in Chrome/Edge, verifying:
  - Responsive layout on desktop and mobile.
  - Correct chronological sorting of meetings.
  - Smooth interaction and transition between the 3 layers.
  - Accordion toggle works and formats the violation tables correctly.
