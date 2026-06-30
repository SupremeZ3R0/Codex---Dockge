# ResumeLibraryBuilder settings
# Change this to the top-level folder that contains CompanyName\JobName folders.
$RootFolder = "C:\Users\YourName\Documents\Post-College"

# Output files are written here. The default keeps generated files inside this project.
$OutputFolder = Join-Path $PSScriptRoot "Output"

# File discovery settings.
# Resumes are expected to use this title, for example: Bramhadev Emogaje Resume.docx
$ResumeTitle = "Bramhadev Emogaje Resume"
$CoverLetterKeywords = @("cover", "coverletter", "cover-letter", "letter")
$SupportedExtensions = @(".docx", ".doc")

# Prefer modern Word documents when multiple matching files exist.
$PreferDocx = $true

# Generated output names.
$CombinedResumeDocxName = "Combined_Resumes.docx"
$CombinedResumeTxtName = "Combined_Resumes.txt"
$CombinedCoverLetterDocxName = "Combined_CoverLetters.docx"
$CombinedCoverLetterTxtName = "Combined_CoverLetters.txt"
$ApplicationsIndexName = "Applications_Index.json"
$ScanReportName = "ScanReport.txt"
$LogName = "CombineLog.txt"
