#region Prompts

# Prompt for the lab name
$labName = Read-Host -Prompt "Enter the name of this lab"

# Prompt for the lab storage path
$labStoragePath = Read-Host -Prompt "Enter the root path to store this lab (e.g., C:\Labs)"

# Prompt for the number of network sites
$numSites = [int](Read-Host -Prompt "Enter the number of network sites")

# Prompt for the template VHDX path
$templateVHDXPath = Read-Host -Prompt "Enter the path to the syspreped VHDX file (e.g., C:\Template\template.vhdx)"

# Prompt for the gateway template, use the template vhdx file if type n
$gatewayTemplateVHDXPath = Read-Host -Prompt "Enter the path to the gateway VHDX file, or 'n' to use the main template"
if ($gatewayTemplateVHDXPath -eq 'n') {
    $gatewayTemplateVHDXPath = $templateVHDXPath
}

# Prompt for the amount of RAM (in MB)
$ramMB = [int](Read-Host -Prompt "Enter the amount of RAM for each VM (in MB)")

#endregion

#region Data Structures

# Store site information
$sites = @()
for ($siteNum = 1; $siteNum -le $numSites; $siteNum++) {
    $siteInfo = [PSCustomObject]@{
        SiteNumber = $siteNum
        SwitchName = "${labName}_Site${siteNum}"
        NumDCs     = [int](Read-Host -Prompt "Enter the number of domain controllers for Site $siteNum")
        NumMembers = [int](Read-Host -Prompt "Enter the number of member servers for Site $siteNum")
    }
    $sites += $siteInfo
}

#endregion

#region Path Calculations

# Calculate the main lab directory
$labPath = Join-Path $labStoragePath $labName

# Calculate subdirectories for configuration and VHDX files
$vmPath = Join-Path $labPath "Configuration"
$vhdxPath = Join-Path $labPath "VHDX"

#endregion

#region Create Directories and Virtual Switches

# Create the main lab directory
if (-not (Test-Path $labPath)) {
    Write-Host "Creating lab directory: $labPath"
    New-Item -ItemType Directory -Path $labPath -Force
}

# Create the configuration and VHDX directories.
New-Item -ItemType Directory -Path $vmPath -Force
New-Item -ItemType Directory -Path $vhdxPath -Force

# Create an array to store switch names
$switches = @()

foreach ($site in $sites) {
     # Create site-specific subdirectories
    $siteVmPath = Join-Path $vmPath "Site$($site.SiteNumber)"
    $siteVhdxPath = Join-Path $vhdxPath "Site$($site.SiteNumber)"
    New-Item -ItemType Directory -Path $siteVmPath -Force
    New-Item -ItemType Directory -Path $siteVhdxPath -Force

    # Create virtual switch for the site
    Write-Host "Creating virtual switch: $($site.SwitchName)"
    New-VMSwitch -Name $site.SwitchName -SwitchType Internal
    $switches += $site.SwitchName
}
#endregion

#region VM Creation
foreach ($site in $sites) {
    Write-Host "Configuring Site $($site.SiteNumber)"
    $siteVmPath = Join-Path $vmPath "Site$($site.SiteNumber)"
    $siteVhdxPath = Join-Path $vhdxPath "Site$($site.SiteNumber)"

    # Create Domain Controllers for the site
    for ($i = 1; $i -le $site.NumDCs; $i++) {
        $dcName = "${labName}_Site$($site.SiteNumber)_DC$i"
        $newVHDXPath = Join-Path $siteVhdxPath "$dcName.vhdx"

        # Copy the template VHDX
        Write-Host "Copying template VHDX to $newVHDXPath"
        Copy-Item -Path $templateVHDXPath -Destination $newVHDXPath

        # Create the DC
        Write-Host "Creating domain controller: $dcName"
        New-VM -Name $dcName -MemoryStartupBytes ($ramMB * 1MB) -Generation 2 -VHDPath $newVHDXPath -SwitchName $site.SwitchName -Path $siteVmPath
    }

    # Create Member Servers for the site
    for ($i = 1; $i -le $site.NumMembers; $i++) {
        $vmName = "${labName}_Site$($site.SiteNumber)_VM$i"
        $newVHDXPath = Join-Path $siteVhdxPath "$vmName.vhdx"

        # Copy the template VHDX
        Write-Host "Copying template VHDX to $newVHDXPath"
        Copy-Item -Path $templateVHDXPath -Destination $newVHDXPath

        # Create the member server
        Write-Host "Creating member server: $vmName"
        New-VM -Name $vmName -MemoryStartupBytes ($ramMB * 1MB) -Generation 2 -VHDPath $newVHDXPath -SwitchName $site.SwitchName -Path $siteVmPath
    }
}

# --- Create Gateway VM ---
$gatewayName = "${labName}_GateWay"
$gatewayVhdxPath = Join-Path $vhdxPath "$gatewayName.vhdx" # Gateway VHDX in the main VHDX folder

# Copy the gateway template VHDX
Write-Host "Copying gateway template VHDX to $gatewayVhdxPath"
Copy-Item -Path $gatewayTemplateVHDXPath -Destination $gatewayVhdxPath

Write-Host "Creating gateway VM: $gatewayName"

# Create the VM connected to the *first* switch initially.
New-VM -Name $gatewayName -MemoryStartupBytes ($ramMB * 1MB) -Generation 2 -VHDPath $gatewayVhdxPath -Path $vmPath -SwitchName $switches[0]

# Add network adapters for the remaining switches.
for ($i = 1; $i -lt $switches.Count; $i++) {
    Write-Host "Adding network adapter to $gatewayName for switch: $($switches[$i])"
    Add-VMNetworkAdapter -VMName $gatewayName -SwitchName $switches[$i] -Name "Network Adapter"
}

Write-Host "VM deployment completed!"
