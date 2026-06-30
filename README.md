# ResumeLibraryBuilder

ResumeLibraryBuilder is a local PowerShell utility for turning a folder of tailored job applications into AI-friendly source material. It scans a folder layout like `MainFolder/Company Name/Job Name/`, finds your resume and cover-letter Word documents, and combines them into formatted Word files plus plain-text exports.

## What it creates

- `Combined_Resumes.docx` — one formatted Word document containing all detected resumes.
- `Combined_Resumes.txt` — one plain-text resume export with detailed separators for AI tools.
- `Combined_CoverLetters.docx` — one formatted Word document containing all detected cover letters.
- `Combined_CoverLetters.txt` — one plain-text cover-letter export with detailed separators.
- `Applications_Index.json` — machine-readable index of detected applications.
- `ScanReport.txt` — human-readable scan report.
- `CombineLog.txt` — run log.

## Project location

The working tool lives in [`ResumeLibraryBuilder/`](ResumeLibraryBuilder/).

Start with the setup guide in [`ResumeLibraryBuilder/README.md`](ResumeLibraryBuilder/README.md).

## Requirements

- Windows
- PowerShell
- Microsoft Word installed locally

Microsoft Word is required because the script uses Word automation to preserve `.docx` formatting when combining files.
