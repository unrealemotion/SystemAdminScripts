<#
.SYNOPSIS
Shrinks a specified disk volume by a given amount in Megabytes (MB).
Designed to be run via GPO as a computer startup script. Logs errors
to the Windows Application Event Log under the source "myScript".

.DESCRIPTION
This script identifies the target partition based on its drive letter.
It calculates the required target size by subtracting the specified shrink
amount from the current size. It performs checks to ensure the shrink
operation is feasible before attempting the resize. If any step fails,
an error event is logged.

.NOTES
Version:        1.1
Author:         Asher
Requires:       PowerShell 3.0 or later, Administrator privileges (provided when run as SYSTEM via GPO).
WARNING:        Disk operations are potentially destructive. Test thoroughly before deploying widely.

.EXAMPLE
# Run manually (requires Administrator PowerShell) after setting variables:
.\ShrinkVolume.ps1
#>

#Requires -Version 3.0
#Requires -RunAsAdministrator

# --- CONFIGURATION ---

# Specify the drive letter of the volume to shrink (e.g., "C", "D")
$TargetDriveLetter = "O"

# Specify the amount of space to shrink IN MEGABYTES (MB)
# Example: 10240 = 10 GB
$ShrinkAmountMB = 150

# --- END CONFIGURATION ---

# --- Event Log Configuration ---
$EventSourceProviderName = "myScript"
$ErrorEventID = 1001 # Unique ID for shrink errors from this script
# --- End Event Log Configuration ---

# --- Script Body ---

# Set strict error handling
$ErrorActionPreference = 'Stop'

# Function to write to Event Log
function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Error", "Warning", "Information")]
        [string]$EntryType,

        [Parameter(Mandatory = $true)]
        [int]$EventId
    )

    # Check if the Event Source exists, create if not. Needs elevation (SYSTEM context is fine).
    if (-not ([System.Diagnostics.EventLog]::SourceExists($EventSourceProviderName))) {
        try {
            Write-Verbose "Event source '$EventSourceProviderName' not found. Attempting to register..."
            New-EventLog -LogName Application -Source $EventSourceProviderName -ErrorAction Stop
            Write-Verbose "Event source '$EventSourceProviderName' registered successfully."
        } catch {
            # If registration fails even as SYSTEM, something is fundamentally wrong.
            # Log to console/stderr if possible, but Write-EventLog will fail.
            Write-Error "FATAL: Failed to register event source '$EventSourceProviderName'. Error: $($_.Exception.Message)"
            # Cannot reliably write to event log if source registration fails.
            # Consider alternative logging or just exit.
            Exit 1 # Exit script if logging isn't possible.
        }
    }

    # Write the actual event log entry
    try {
        Write-EventLog -LogName Application -Source $EventSourceProviderName -EventId $EventId -EntryType $EntryType -Message $Message -ErrorAction Stop
        Write-Verbose "Logged message to Application event log (Source: $EventSourceProviderName, EventID: $EventId, Type: $EntryType)."
    } catch {
        # Log failure to write the event itself
        Write-Error "Failed to write to event log (Source: $EventSourceProviderName). Error: $($_.Exception.Message)"
        # The original error won't be logged, but this secondary error indicates a logging problem.
    }
}

# --- Main Logic ---
try {
    Write-Verbose "Script started. Target Drive: $TargetDriveLetter, Shrink Amount: ${ShrinkAmountMB}MB"

    # Validate configuration
    if (-not ($TargetDriveLetter -match '^[a-zA-Z]$')) {
        throw "Invalid Drive Letter specified: '$TargetDriveLetter'. Please specify a single letter (e.g., 'C')."
    }
    if ($ShrinkAmountMB -le 0) {
        throw "Invalid Shrink Amount specified: '$ShrinkAmountMB'. Must be a positive number of Megabytes."
    }

    $DrivePath = "$($TargetDriveLetter):\"
    if (-not (Test-Path -Path $DrivePath -PathType Container)) {
        throw "Drive letter '$TargetDriveLetter' does not exist or is not accessible."
    }

    # Get partition information
    Write-Verbose "Getting partition information for drive $TargetDriveLetter..."
    $Partition = Get-Partition -DriveLetter $TargetDriveLetter

    $CurrentSizeBytes = $Partition.Size
    $ShrinkAmountBytes = $ShrinkAmountMB * 1MB # 1MB is a built-in PowerShell multiplier

    Write-Verbose "Current Size: $($CurrentSizeBytes / 1GB) GB ($CurrentSizeBytes bytes)"
    Write-Verbose "Requested Shrink Amount: $($ShrinkAmountBytes / 1MB) MB ($ShrinkAmountBytes bytes)"

    # Calculate the target size
    $TargetSizeBytes = $CurrentSizeBytes - $ShrinkAmountBytes
    Write-Verbose "Calculated Target Size: $($TargetSizeBytes / 1GB) GB ($TargetSizeBytes bytes)"

    # Pre-check: Get minimum and maximum supported sizes for resize
    Write-Verbose "Checking supported resize range..."
    $SupportedSize = Get-PartitionSupportedSize -DriveLetter $TargetDriveLetter
    $MinimumSize = $SupportedSize.SizeMin
    $MaximumSize = $SupportedSize.SizeMax # This is usually the current size

    Write-Verbose "Minimum possible size for partition: $($MinimumSize / 1GB) GB ($MinimumSize bytes)"
    Write-Verbose "Maximum shrinkable space (Current - Min): $(($CurrentSizeBytes - $MinimumSize) / 1MB) MB"

    # Validate if the requested shrink is possible
    if ($TargetSizeBytes -lt $MinimumSize) {
        $MaxPossibleShrinkMB = [Math]::Floor(($CurrentSizeBytes - $MinimumSize) / 1MB)
        throw "Cannot shrink by ${ShrinkAmountMB}MB. The resulting size ($($TargetSizeBytes / 1GB) GB) would be smaller than the minimum allowed size ($($MinimumSize / 1GB) GB). Maximum possible shrink is approx ${MaxPossibleShrinkMB}MB."
    }

    if ($TargetSizeBytes -ge $CurrentSizeBytes) {
         throw "Calculated target size ($($TargetSizeBytes / 1GB) GB) is not smaller than current size ($($CurrentSizeBytes / 1GB) GB). Shrink amount ($ShrinkAmountMB MB) might be too small or zero."
    }

    # Perform the resize operation
    Write-Verbose "Attempting to resize partition $TargetDriveLetter to $($TargetSizeBytes / 1GB) GB..."
    Resize-Partition -DriveLetter $TargetDriveLetter -Size $TargetSizeBytes

    $NewPartitionInfo = Get-Partition -DriveLetter $TargetDriveLetter
    $FinalSizeBytes = $NewPartitionInfo.Size
    Write-Verbose "Resize operation completed successfully. New size: $($FinalSizeBytes / 1GB) GB ($FinalSizeBytes bytes)."

    # Optional: Log success event
    # $SuccessMessage = "Successfully shrank drive $TargetDriveLetter by approx ${ShrinkAmountMB}MB. New size is $($FinalSizeBytes / 1GB) GB."
    # Write-LogEntry -Message $SuccessMessage -EntryType Information -EventId ($ErrorEventID + 1) # Use a different ID for success

    Write-Host "Volume $TargetDriveLetter successfully shrunk." # For interactive testing

} catch {
    # Log failure to the Application Event Log
    $ErrorMessage = @"
Error shrinking volume $TargetDriveLetter by ${ShrinkAmountMB}MB.
Error Details:
$($_.Exception.Message)
Script StackTrace:
$($_.ScriptStackTrace)
"@ -replace "`r`n?", "`n" # Ensure consistent line endings for event log

    Write-LogEntry -Message $ErrorMessage -EntryType Error -EventId $ErrorEventID

    # Optionally write to stderr as well, useful if run interactively for debugging
    Write-Error $ErrorMessage

    # Exit with a non-zero code to indicate failure (GPO might track this)
    Exit 1
}

# Exit with 0 for success
Exit 0