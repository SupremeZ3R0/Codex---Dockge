# ResumeLibraryBuilder

PowerShell utility for combining tailored job-application resumes and cover letters into formatted Word documents and plain text files for AI-assisted resume tailoring.

## Folder layout expected

```text
MainFolder/
  Company Name/
    Job Name/
      Resume.docx
```

The script scans `Company Name/Job Name` folders, finds resumes titled `Bramhadev Emogaje Resume`, detects cover letters by keyword, then writes all generated files to `ResumeLibraryBuilder/Output` by default.

## Outputs

- `Combined_Resumes.docx` — formatted Word document using Word's `InsertFile()` to preserve resume formatting, with only company and job headings.
- `Combined_Resumes.txt` — plain text version separated by full company/job/file/folder metadata.
- `Combined_CoverLetters.docx` — formatted Word document combining detected cover letters.
- `Combined_CoverLetters.txt` — plain text cover-letter export separated by full metadata.
- `Applications_Index.json` — machine-readable index of detected applications.
- `ScanReport.txt` — human-readable scan report.
- `CombineLog.txt` — log file.

## How to run

1. Open PowerShell on Windows with Microsoft Word installed.
2. Edit `Config.ps1` and set `$RootFolder` to your main job-applications folder.
3. Run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Build-ResumeLibrary.ps1
```

Or override the root folder without editing config:

```powershell
.\Build-ResumeLibrary.ps1 -RootFolderOverride "C:\Users\YourName\Documents\Post-College"
```

## Troubleshooting

### Execution policy

Run the execution-policy command before running the build script in the same PowerShell window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Build-ResumeLibrary.ps1 -RootFolderOverride "C:\Users\YourName\Documents\Post-College"
```

## Notes

- Temporary Word files beginning with `~$` are skipped.
- `.docx` and `.doc` files are supported.
- Resume discovery expects files titled `Bramhadev Emogaje Resume` before the extension. Copies like `Bramhadev Emogaje Resume (1).docx` are also accepted.
- PDF extraction is intentionally not included in this first working version because preserving formatting reliably requires a separate PDF pipeline.
