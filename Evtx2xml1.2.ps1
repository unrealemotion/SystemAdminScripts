# Prompt for the directory containing the .evtx files
$directoryPath = Read-Host -Prompt "Enter the directory containing the .evtx files"

# Prompt for the output directory
$outputDirectory = Read-Host -Prompt "Enter the output directory"

# Prompt for the machine name
$machineName = Read-Host -Prompt "Enter the machine name"

# Get all .evtx files in the specified directory
$evtxFiles = Get-ChildItem -Path $directoryPath -Filter "*.evtx"

# Loop through each .evtx file
foreach ($evtxFile in $evtxFiles) {
    try {
        # Remove "Microsoft-Windows-" from the file name
        $baseName = $evtxFile.BaseName -replace "Microsoft-Windows-", ""

        # Construct the new .clixml file name with the machine name
        $clixmlFileName = $machineName + "__" + $baseName + ".xml"

        # Construct the full path for the output .clixml file
        $clixmlFilePath = Join-Path -Path $outputDirectory -ChildPath $clixmlFileName

        # Construct the new .txt file name with the machine name
        $txtFileName = $machineName + "__" + $baseName + ".txt"

        # Construct the full path for the output .txt file
        $txtFilePath = Join-Path -Path $outputDirectory -ChildPath $txtFileName

        # Convert the .evtx file to .clixml, suppressing error messages
        Get-WinEvent -Path $evtxFile.FullName -ErrorAction SilentlyContinue | Export-Clixml -Path $clixmlFilePath

        # Check if any events were found
        if (-not $?) {
            Write-Host "No events found in: $($evtxFile.FullName). Don't worry, we will proceed with the next file :D"

            # Delete the empty XML file
            Remove-Item -Path $clixmlFilePath -Force
        } else {
            # Export the events as a formatted table to a .txt file with no trimming
            Get-WinEvent -Path $evtxFile.FullName | Format-Table -AutoSize -Wrap | Out-File -FilePath $txtFilePath -Width 512
        }
    } catch {
        Write-Host "Error processing $($evtxFile.FullName): $_"
    }
}

Write-Host "Conversion complete!"