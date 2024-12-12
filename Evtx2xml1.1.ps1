# Prompt for the directory containing the .evtx files
$directoryPath = Read-Host -Prompt "Enter the directory containing the .evtx files"

# Prompt for the output directory
$outputDirectory = Read-Host -Prompt "Enter the output directory"

# Prompt for the number to append to the file names
$appendNumber = Read-Host -Prompt "Enter the number to append to the file names"

# Get all .evtx files in the specified directory
$evtxFiles = Get-ChildItem -Path $directoryPath -Filter "*.evtx"

# Loop through each .evtx file
foreach ($evtxFile in $evtxFiles) {
    # Remove "Microsoft-Windows-" from the file name
    $baseName = $evtxFile.BaseName -replace "Microsoft-Windows-", ""

    # Construct the new .clixml file name with the appended number
    $clixmlFileName = $baseName + "_" + $appendNumber + ".clixml"

    # Construct the full path for the output .clixml file
    $clixmlFilePath = Join-Path -Path $outputDirectory -ChildPath $clixmlFileName

    # Convert the .evtx file to .clixml
    Get-WinEvent -Path $evtxFile.FullName | Export-Clixml -Path $clixmlFilePath
}

Write-Host "Conversion complete!"