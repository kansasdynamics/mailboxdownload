# Set the working directory to where you want to save files
Set-Location "C:\Users\Public\Downloads"

# Define paths to your Python scripts
$extractPdfScript = "C:\Users\Public\Downloads\scripts\extract_pdf.py"
$cleanupScript = "C:\Users\Public\Downloads\scripts\extract_pdf_cleanup.py"

# Run the extract_pdf.py script
Write-Output "Running extract_pdf.py..."
python $extractPdfScript

# Run the extract_pdf_cleanup.py script
Write-Output "Running extract_pdf_cleanup.py..."
python $cleanupScript

Write-Output "Process completed."

# You can uncomment this Read-Host line if you need to troubleshoot the script output when running it from the shortcut
# Read-Host "Press Enter to close this window..."

# Use this location to make a desktop shortcut
# powershell.exe -ExecutionPolicy Bypass -File "C:\Users\Public\Downloads\scripts\run_extract_pdf.ps1"

# Open PowerShell as Administrator and run this command to setup the script
# Set-ExecutionPolicy RemoteSigned
