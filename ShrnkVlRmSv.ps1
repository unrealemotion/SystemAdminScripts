<#
.SYNOPSIS
Prompts for necessary details, gets current size, asks for shrink amount (MB),
and attempts to shrink a volume on a remote computer, with an option to restart.

.DESCRIPTION
This script interactively shrinks a remote volume. It gathers the remote computer
name, drive letter, and administrative credentials. It then connects to fetch the
current size (MB) and minimum possible size of the specified volume. If successful,
it prompts the user for the amount of space (in MB) to shrink the volume by.
After validation and user confirmation, it attempts the shrink operation using
PowerShell Remoting (WinRM). After any outcome (success, failure, cancellation),
it asks the user if they wish to start the entire process over.

.NOTES
- Requires PowerShell Remoting (WinRM) enabled and configured on the remote machine.
- Requires administrative privileges on the remote machine.
- Firewalls must allow WinRM traffic (TCP 5985/5986).
- If machines are not domain-joined, TrustedHosts configuration might be needed.
- ALWAYS BACK UP DATA before performing disk operations.
#>

# Loop control variable
$startOver = 'y'

while ($startOver -eq 'y') {
    # Clear screen for a cleaner restart
    Clear-Host
    Write-Host "Remote Volume Shrink Script (MB Granularity)" -ForegroundColor Yellow
    Write-Host "------------------------------------------"

    # --- Initialize variables for this loop iteration ---
    $remoteComputer = $null
    $driveLetterInput = $null
    $driveLetter = $null
    $credential = $null
    $volumeInfo = $null
    $shrinkAmountMB_Input = $null
    $shrinkAmountMB = $null
    $currentSizeBytes = $null
    $minSizeBytes = $null
    $desiredNewSizeBytes = $null
    $errorOccurred = $false

    # --- Get User Input ---
    try {
        $remoteComputer = Read-Host "Enter the name or IP address of the remote computer"
        if (-not $remoteComputer) { throw "Remote computer name cannot be empty." }

        $driveLetterInput = Read-Host "Enter the drive letter to shrink (e.g., C, D, E)"
        $driveLetter = $driveLetterInput.Trim().Trim(":") | Select-Object -First 1
        if ($driveLetter.Length -ne 1 -or -not ($driveLetter -match "[a-zA-Z]")) {
            throw "Invalid drive letter specified: '$driveLetterInput'. Please enter a single letter."
        }
        $driveLetter = $driveLetter.ToUpper()

        Write-Host "Enter credentials for '$remoteComputer' (e.g., Domain\User or ComputerName\User)"
        $credential = Get-Credential
        if (-not $credential) { throw "Credentials are required." }

        # --- Step 1: Get Current Size and Minimum Size ---
        Write-Host "`nConnecting to '$remoteComputer' to get volume information for drive $driveLetter..."
        $volumeInfo = Invoke-Command -ComputerName $remoteComputer -Credential $credential -ScriptBlock {
            param($letter)
            $ErrorActionPreference = 'Stop' # Ensure errors inside stop the scriptblock
            try {
                $partition = Get-Partition -DriveLetter $letter
                $supportedSize = Get-PartitionSupportedSize -DriveLetter $letter

                # Return relevant info as a custom object
                return [PSCustomObject]@{
                    CurrentSizeBytes = $partition.Size
                    MinSizeBytes     = $supportedSize.SizeMin
                    CurrentSizeMB    = [Math]::Round($partition.Size / 1MB, 2)
                    MinSizeMB        = [Math]::Round($supportedSize.SizeMin / 1MB, 2)
                    DriveLetterFound = $true
                }
            } catch {
                # Specific handling if drive/partition not found vs other errors
                 if ($_.Exception.Message -like "*No MSFT_Partition objects found with property 'DriveLetter' equal to '$letter'*") {
                     throw "Partition for drive letter '$letter' not found on '$($env:COMPUTERNAME)'."
                 } else {
                     # Rethrow other errors to be caught by the outer catch
                     throw "Error getting volume info on '$($env:COMPUTERNAME)': $($_.Exception.Message)"
                 }
            }
        } -ArgumentList $driveLetter

        # Assign values received from remote host
        $currentSizeBytes = $volumeInfo.CurrentSizeBytes
        $minSizeBytes = $volumeInfo.MinSizeBytes
        $currentSizeMB = $volumeInfo.CurrentSizeMB
        $minSizeMB = $volumeInfo.MinSizeMB

        Write-Host ("Successfully retrieved info for drive $driveLetter on $remoteComputer.") -ForegroundColor Green
        Write-Host ("   Current Size: {0:N2} MB" -f $currentSizeMB)
        Write-Host ("   Minimum Size: {0:N2} MB (Volume cannot be smaller than this)" -f $minSizeMB)

        # --- Step 2: Get Shrink Amount from User ---
        $shrinkAmountMB_Input = Read-Host "`nEnter the amount of space to SHRINK BY, in MB (e.g., 10240 for 10GB)"
        try {
            $shrinkAmountMB = [double]$shrinkAmountMB_Input
            if ($shrinkAmountMB -le 0) { throw "Shrink amount must be a positive number." }
        } catch {
            throw "Invalid shrink amount specified: '$shrinkAmountMB_Input'. Please enter a positive number."
        }

        # --- Step 3: Calculate and Validate ---
        $shrinkAmountBytes = $shrinkAmountMB * 1MB
        $desiredNewSizeBytes = $currentSizeBytes - $shrinkAmountBytes

        Write-Host ("   Calculating: Current ({0:N2} MB) - Shrink ({1:N2} MB) = Target ({2:N2} MB)" -f $currentSizeMB, $shrinkAmountMB, ($desiredNewSizeBytes / 1MB))

        if ($desiredNewSizeBytes -lt $minSizeBytes) {
            throw ("Shrinking by {0:N2} MB would result in a size ({1:N2} MB) smaller than the minimum allowed ({2:N2} MB)." -f $shrinkAmountMB, ($desiredNewSizeBytes / 1MB), $minSizeMB)
        }
        if ($desiredNewSizeBytes -ge $currentSizeBytes) {
             # This should only happen if shrink amount is zero or negative, already checked, but good failsafe
             throw ("Shrink amount ({0:N2} MB) must result in a smaller volume size." -f $shrinkAmountMB)
        }

         # Ensure we are actually shrinking by a noticeable amount (e.g., > 1MB) to avoid tiny/pointless shrinks
         if ($shrinkAmountBytes -lt 1MB) {
             Write-Warning "Shrink amount is very small (less than 1 MB)."
             # You could choose to exit here or let it proceed
         }


        # --- Step 4: Confirmation ---
        $finalTargetSizeMB = $desiredNewSizeBytes / 1MB
        $confirm = Read-Host ("`nWARNING: About to shrink drive $driveLetter on '$remoteComputer'." `
                             + "`n  Current Size: {0:N2} MB" -f $currentSizeMB `
                             + "`n  Shrink By:    {0:N2} MB" -f $shrinkAmountMB `
                             + "`n  New Size:     {0:N2} MB" -f $finalTargetSizeMB `
                             + "`nEnsure backups exist. Proceed? (y/n)")
        if ($confirm -ne 'y') {
            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
            throw "User cancelled operation." # Throw to trigger the 'start over' prompt logic
        }

        # --- Step 5: Execute Shrink ---
        Write-Host "`nSending shrink command..."
        Invoke-Command -ComputerName $remoteComputer -Credential $credential -ScriptBlock {
            param($letter, $targetSize)
            $ErrorActionPreference = 'Stop' # Ensure Resize-Partition errors are caught
            Write-Host "Executing: Resize-Partition -DriveLetter $letter -Size $targetSize"
            Resize-Partition -DriveLetter $letter -Size $targetSize
            Write-Host "Resize command completed successfully on remote host for drive $letter."
        } -ArgumentList $driveLetter, $desiredNewSizeBytes

        Write-Host "`nShrink command successfully sent to '$remoteComputer'." -ForegroundColor Green
        Write-Host "IMPORTANT: Verify the result on the remote machine using Disk Management (diskmgmt.msc) or Get-Volume." -ForegroundColor Yellow

    }
    catch {
        # Catch any error thrown from the try block
        Write-Error "An error occurred: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        $errorOccurred = $true # Flag that an error happened
    }

    # --- Ask to Start Over ---
    # This section runs regardless of success or failure within the try block
    Write-Host "" # Add a blank line
    $startOverInput = Read-Host "Do you want to start the process over? (y/n)"
    if ($startOverInput -ne 'y') {
        $startOver = 'n' # Set loop control variable to exit
    }
    # If user enters 'y', the loop continues naturally

} # End of While loop

Write-Host "`nScript finished."