<#
.SYNOPSIS
Prompts for multiple machines, a drive letter, and credentials. Gathers volume
information from all machines, prompts for a uniform shrink amount (MB), validates,
confirms, and then attempts to shrink the specified volume on each machine sequentially,
reporting the outcome for each.

.DESCRIPTION
This script orchestrates shrinking the same volume (by drive letter) across multiple
remote computers by a specified amount in Megabytes (MB).
1. Asks for the number of machines.
2. Collects the names/IP addresses of the target machines.
3. Asks for the drive letter to shrink (must be the same for all).
4. Prompts for administrative credentials valid on ALL target machines.
5. Connects to each machine to retrieve the current and minimum possible size (MB)
   for the specified volume. Displays this aggregated information.
6. Prompts the user for the amount of space (in MB) to shrink EACH volume by.
7. Validates the requested shrink amount against the constraints of ALL machines.
8. Asks for final confirmation before proceeding.
9. Iterates through each machine, attempting the shrink operation using PowerShell
   Remoting (WinRM) and reports the success or failure for that specific machine.
10. After processing all machines, asks the user if they wish to start the entire
    process over.

.NOTES
- Requires PowerShell Remoting (WinRM) enabled and configured on ALL remote machines.
- Requires administrative privileges (provided via credentials) on ALL remote machines.
- Firewalls must allow WinRM traffic (TCP 5985/5986).
- If machines are not domain-joined, TrustedHosts configuration might be needed on the
  machine running the script, or use IP addresses with specific credential types.
- ALWAYS BACK UP DATA before performing disk operations. This script performs potentially
  destructive actions. Use with extreme caution.
- Errors on one machine (e.g., unreachable, drive not found) during info gathering
  will prevent it from being processed further but won't stop info gathering for others.
- Errors during the shrink phase on one machine will be reported, and the script will
  continue to the next machine.
#>

# Loop control variable
$startOver = 'y'

while ($startOver -eq 'y') {
    # Clear screen for a cleaner restart
    Clear-Host
    Write-Host "Multi-Machine Remote Volume Shrink Script (MB Granularity)" -ForegroundColor Yellow
    Write-Host "---------------------------------------------------------"

    # --- Initialize variables for this loop iteration ---
    $numberOfMachines = 0
    $computerNames = @()
    $driveLetterInput = $null
    $driveLetter = $null
    $credential = $null
    $volumeInfoList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $shrinkAmountMB_Input = $null
    $shrinkAmountMB = $null
    $errorOccurredInLoop = $false
    $machinesToProcess = @()

    try {
        # --- Get Number of Machines ---
        while ($numberOfMachines -le 0) {
            $numInput = Read-Host "Enter the number of machines to shrink"
            if ($numInput -match '^\d+$' -and [int]$numInput -gt 0) {
                $numberOfMachines = [int]$numInput
            } else {
                Write-Warning "Please enter a positive integer for the number of machines."
            }
        }

        # --- Get Machine Names ---
        Write-Host "`nEnter the name or IP address for each machine:"
        for ($i = 1; $i -le $numberOfMachines; $i++) {
            $name = ""
            while (-not $name) {
                $name = Read-Host "  Machine #$i"
                if (-not $name) { Write-Warning "Machine name cannot be empty." }
            }
            $computerNames += $name.Trim()
        }
        Write-Host ("Target machines: {0}" -f ($computerNames -join ', ')) -ForegroundColor Cyan

        # --- Get Drive Letter ---
        while (-not $driveLetter) {
            $driveLetterInput = Read-Host "Enter the drive letter to shrink on ALL machines (e.g., C, D, E)"
            $parsedLetter = $driveLetterInput.Trim().Trim(":") | Select-Object -First 1
            if ($parsedLetter.Length -ne 1 -or -not ($parsedLetter -match "[a-zA-Z]")) {
                Write-Warning "Invalid drive letter specified: '$driveLetterInput'. Please enter a single letter."
            } else {
                $driveLetter = $parsedLetter.ToUpper()
            }
        }
        Write-Host ("Target drive letter: {0}" -f $driveLetter) -ForegroundColor Cyan

        # --- Get Credentials ---
        Write-Host "`nEnter credentials with administrative rights on ALL target machines"
        Write-Host "(e.g., Domain\User or Use .\Administrator for local admin on non-domain joined)" -ForegroundColor Gray
        $credential = Get-Credential
        if (-not $credential) { throw "Credentials are required." }

        # --- Step 1: Get Current Size and Minimum Size for ALL Machines ---
        Write-Host "`nConnecting to machines to gather volume information for drive $driveLetter..." -ForegroundColor Yellow

        $maxMinSizeBytesOverall = 0 # Track the largest minimum size found across all valid machines
        $foundValidVolume = $false # Flag if we found at least one valid volume

        foreach ($computerName in $computerNames) {
            Write-Host "  Querying '$computerName'..." -NoNewline
            $machineInfo = $null
            $machineError = $null
            try {
                $machineInfo = Invoke-Command -ComputerName $computerName -Credential $credential -ScriptBlock {
                    param($letter)
                    # Use 'Stop' locally inside ScriptBlock for clean error capture
                    $ErrorActionPreference = 'Stop'
                    try {
                        $partition = Get-Partition -DriveLetter $letter
                        $supportedSize = Get-PartitionSupportedSize -DriveLetter $letter

                        return [PSCustomObject]@{
                            ComputerName     = $env:COMPUTERNAME # Get actual name from remote machine
                            DriveLetter      = $letter
                            CurrentSizeBytes = $partition.Size
                            MinSizeBytes     = $supportedSize.SizeMin
                            CurrentSizeMB    = [Math]::Round($partition.Size / 1MB, 2)
                            MinSizeMB        = [Math]::Round($supportedSize.SizeMin / 1MB, 2)
                            Error            = $null
                        }
                    } catch {
                        # Capture specific error message to return
                        throw "Error on '$($env:COMPUTERNAME)' for drive '$letter': $($_.Exception.Message)"
                    }
                } -ArgumentList $driveLetter

                Write-Host " OK. Current: $($machineInfo.CurrentSizeMB) MB, Min: $($machineInfo.MinSizeMB) MB" -ForegroundColor Green
                $volumeInfoList.Add($machineInfo)
                $foundValidVolume = $true
                 # Update the overall maximum minimum size needed
                if ($machineInfo.MinSizeBytes -gt $maxMinSizeBytesOverall) {
                    $maxMinSizeBytesOverall = $machineInfo.MinSizeBytes
                }

            } catch {
                # Error during Invoke-Command (connection, auth, or scriptblock exception)
                $errMsg = "Failed to get info from '$computerName': $($_.Exception.Message)"
                Write-Warning "`n  $errMsg" # Use Write-Warning for non-fatal errors here
                $machineError = $errMsg
                # Add a placeholder object to indicate failure for this machine
                 $volumeInfoList.Add([PSCustomObject]@{
                     ComputerName = $computerName # Use the name we tried to connect to
                     DriveLetter  = $driveLetter
                     Error        = $machineError
                 })
            }
        } # End foreach computername for info gathering

        # --- Check if any information was gathered ---
        if (-not $foundValidVolume) {
            throw "Could not retrieve valid volume information for drive '$driveLetter' from any of the specified machines. Cannot proceed."
        }

        # Filter list to only those successfully queried
        $machinesToProcess = $volumeInfoList | Where-Object { -not $_.Error }

        # Display Summary
        Write-Host "`n--- Volume Information Summary ---" -ForegroundColor Yellow
        $machinesToProcess | Format-Table ComputerName, DriveLetter, CurrentSizeMB, MinSizeMB
        $maxMinSizeMBOverall = [Math]::Round($maxMinSizeBytesOverall / 1MB, 2)
        Write-Host "Overall minimum required size across all machines: $maxMinSizeMBOverall MB" -ForegroundColor Cyan


        # --- Step 2: Get Shrink Amount from User ---
        $shrinkAmountMB = $null
        while ($null -eq $shrinkAmountMB) {
            $shrinkAmountMB_Input = Read-Host "`nEnter the amount of space to SHRINK EACH volume BY, in MB (e.g., 10240 for 10GB)"
            try {
                $parsedAmount = [double]$shrinkAmountMB_Input
                if ($parsedAmount -le 0) {
                    Write-Warning "Shrink amount must be a positive number."
                } else {
                     $shrinkAmountMB = $parsedAmount
                }
            } catch {
                Write-Warning "Invalid shrink amount specified: '$shrinkAmountMB_Input'. Please enter a positive number."
            }
        }

        # --- Step 3: Calculate and Validate (Across All Machines) ---
        Write-Host "`nValidating shrink amount '$($shrinkAmountMB) MB' against all machines..." -ForegroundColor Yellow
        $shrinkAmountBytes = $shrinkAmountMB * 1MB
        $validationFailed = $false

        foreach ($machineInfo in $machinesToProcess) {
            $desiredNewSizeBytes = $machineInfo.CurrentSizeBytes - $shrinkAmountBytes
            $desiredNewSizeMB = [Math]::Round($desiredNewSizeBytes / 1MB, 2)

            Write-Host ("  Checking '$($machineInfo.ComputerName)': Current ({0:N2} MB) - Shrink ({1:N2} MB) = Target ({2:N2} MB)" -f $machineInfo.CurrentSizeMB, $shrinkAmountMB, $desiredNewSizeMB)

            if ($desiredNewSizeBytes -lt $machineInfo.MinSizeBytes) {
                 Write-Error ("Validation Failed on '$($machineInfo.ComputerName)': Shrinking by {0:N2} MB would result in a size ({1:N2} MB) smaller than its minimum allowed ({2:N2} MB)." -f $shrinkAmountMB, $desiredNewSizeMB, $machineInfo.MinSizeMB)
                 $validationFailed = $true
            }
            # This check implicitly uses the largest minimum size found earlier
            if ($desiredNewSizeBytes -lt $maxMinSizeBytesOverall) {
                Write-Error ("Validation Failed on '$($machineInfo.ComputerName)': Shrinking by {0:N2} MB would result in a size ({1:N2} MB) smaller than the overall required minimum size ({2:N2} MB) dictated by another machine." -f $shrinkAmountMB, $desiredNewSizeMB, $maxMinSizeMBOverall)
                $validationFailed = $true
            }
            if ($desiredNewSizeBytes -ge $machineInfo.CurrentSizeBytes) {
                 Write-Error ("Validation Failed on '$($machineInfo.ComputerName)': Shrink amount ({0:N2} MB) must result in a smaller volume size." -f $shrinkAmountMB)
                 $validationFailed = $true
            }
             # Ensure we are actually shrinking by a noticeable amount
             if ($shrinkAmountBytes -lt 1MB) {
                 Write-Warning "Shrink amount is very small (less than 1 MB)."
             }
        }

        if ($validationFailed) {
            throw "Validation failed for one or more machines. Cannot proceed with shrink operation."
        }

        Write-Host "Validation Successful." -ForegroundColor Green

        # --- Step 4: Confirmation ---
        Write-Host "`n--- Confirmation ---" -ForegroundColor Yellow
        Write-Host "You are about to perform the following shrink operation:"
        Write-Host "  Machines:      $($machinesToProcess.ComputerName -join ', ')"
        Write-Host "  Drive Letter:  $driveLetter"
        Write-Host "  Shrink BY:     $shrinkAmountMB MB (applied to each machine)"
        Write-Host "The resulting size on each machine will vary based on its current size."
        Write-Host "Ensure backups exist for all target machines." -ForegroundColor Red -BackgroundColor Black

        $confirm = Read-Host "Proceed with shrinking ALL listed volumes? (y/n)"
        if ($confirm -ne 'y') {
            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
            # Set flag to prevent 'error' message at the end, but still offer restart
            $errorOccurredInLoop = $true # Treat cancellation like an error for restart prompt logic
            throw "User cancelled operation."
        }

        # --- Step 5: Execute Shrink (Sequentially) ---
        Write-Host "`n--- Starting Shrink Process ---" -ForegroundColor Yellow
        foreach ($machineInfo in $machinesToProcess) {
            $targetComputer = $machineInfo.ComputerName
            $targetDrive = $machineInfo.DriveLetter
            # Recalculate the target size SPECIFIC to this machine
            $targetSizeBytes = $machineInfo.CurrentSizeBytes - $shrinkAmountBytes
            $targetSizeMB = [Math]::Round($targetSizeBytes / 1MB, 2)

            Write-Host "`nProcessing '$targetComputer' (Drive $targetDrive)..."
            Write-Host "  Attempting to resize to $targetSizeMB MB..."

            try {
                Invoke-Command -ComputerName $targetComputer -Credential $credential -ScriptBlock {
                    param($letter, $targetSize)
                    $ErrorActionPreference = 'Stop' # Ensure Resize-Partition errors are caught
                    Write-Host "  Remote Exec: Resize-Partition -DriveLetter $letter -Size $targetSize on $($env:COMPUTERNAME)" -ForegroundColor Gray
                    Resize-Partition -DriveLetter $letter -Size $targetSize
                    # No output needed on success inside scriptblock, handled outside
                } -ArgumentList $targetDrive, $targetSizeBytes

                Write-Host "  SUCCESS: Shrink command completed successfully for '$targetComputer'." -ForegroundColor Green

            } catch {
                # Error during the actual resize operation
                $errMsg = "FAILURE on '$targetComputer': $($_.Exception.Message)"
                Write-Error $errMsg
                # You might want to log this error more permanently here
            }
        } # End foreach machine for shrinking

        Write-Host "`n--- Shrink Process Completed for all targeted machines ---" -ForegroundColor Yellow
        Write-Host "IMPORTANT: Verify the results on each remote machine using Disk Management (diskmgmt.msc) or Get-Volume." -ForegroundColor Yellow

    }
    catch {
        # Catch any error thrown from the main try block (input, validation, cancellation)
        Write-Error "An error occurred in the script: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        $errorOccurredInLoop = $true # Flag that an error happened
    }

    # --- Ask to Start Over ---
    Write-Host "" # Add a blank line
    $startOverInput = Read-Host "Do you want to start the entire process over with new inputs? (y/n)"
    if ($startOverInput -ne 'y') {
        $startOver = 'n' # Set loop control variable to exit
    }
    # If user enters 'y', the loop continues naturally

} # End of While loop

Write-Host "`nScript finished."