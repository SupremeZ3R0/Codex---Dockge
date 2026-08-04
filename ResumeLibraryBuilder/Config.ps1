# ResumeLibraryBuilder settings
# Change this to the top-level folder that contains CompanyName\JobName folders.
$RootFolder = "C:\Users\Nysup\OneDrive\Documents\Post-College"

# Output files are written here. The default keeps generated files inside this project.
$OutputFolder = Join-Path $PSScriptRoot "Output"

# File discovery settings.
# Resumes are expected to use this title, for example: Bramhadev Emogaje Resume.docx
$ResumeTitle = "Bramhadev Emogaje Resume"
$CoverLetterKeywords = @("cover", "coverletter", "cover-letter", "letter")
$SupportedExtensions = @(".docx", ".doc")

# Prefer modern Word documents when multiple matching files exist.
$PreferDocx = $true

# Combined resume Word document margins, in inches.
$ResumeDocMarginTopInches = 1
$ResumeDocMarginBottomInches = 0.6
$ResumeDocMarginLeftInches = 0.5
$ResumeDocMarginRightInches = 0.5

# Combined cover-letter Word document margins, in inches.
$CoverLetterDocMarginTopInches = 1
$CoverLetterDocMarginBottomInches = 1
$CoverLetterDocMarginLeftInches = 1
$CoverLetterDocMarginRightInches = 1

# Generated output names.
$CombinedResumeDocxName = "Combined_Resumes.docx"
$CombinedResumeTxtName = "Combined_Resumes.txt"
$CombinedCoverLetterDocxName = "Combined_CoverLetters.docx"
$CombinedCoverLetterTxtName = "Combined_CoverLetters.txt"
$ApplicationsIndexName = "Applications_Index.json"
$ScanReportName = "ScanReport.txt"
$LogName = "CombineLog.txt"
