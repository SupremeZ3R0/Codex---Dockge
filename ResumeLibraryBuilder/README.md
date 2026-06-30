# ResumeLibraryBuilder

PowerShell utility for combining tailored job-application resumes into one formatted Word document and one plain text file for AI-assisted resume tailoring.

## Folder layout expected

```text
MainFolder/
  Company Name/
    Job Name/
      Resume.docx
```

The script scans `Company Name/Job Name` folders, finds Word resume files by keywords such as `resume` or `cv`, then writes all generated files to `ResumeLibraryBuilder/Output` by default.

## Outputs

- `Combined_Resumes.docx` — formatted Word document using Word's `InsertFile()` to preserve resume formatting.
- `Combined_Resumes.txt` — plain text version separated by company/job metadata.
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

## Notes

- Temporary Word files beginning with `~$` are skipped.
- `.docx` and `.doc` files are supported.
- PDF extraction is intentionally not included in this first working version because preserving formatting reliably requires a separate PDF pipeline.
