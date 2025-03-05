$input = Read-Host "Please provide us with the parent directory"
$paths = Get-ChildItem $input  # Replace with your target directory
foreach ($item in $paths)
{
$path=$item.fullname

if (!(Test-Path -Path $path -PathType Container)) {
  Write-Host "Error: The directory '$path' does not exist." -ForegroundColor Red
  exit 1  # Exit with an error code
}

# Get the ACL of the directory
try {
    $acl = Get-Acl -Path $path
} catch {
    Write-Host "Error getting ACL for '$path': $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}


# Check if inheritance is enabled
if ($acl.AreAccessRulesProtected) {  # If AreAccessRulesProtected is $true, inheritance is *disabled*.
  Write-Host "$path inheritance is disabled already." -ForegroundColor Yellow
} else {
  Write-Host "Inheritance is enabled for '$path'." -ForegroundColor Green
# --- 1. Get all ACEs (Access Control Entries) ---

$originalAcl = Get-Acl -Path $path
$owner = $originalAcl.Owner
$group = $originalAcl.Group

# Get *all* access rules (before disabling inheritance - we're rebuilding)
$accessRules = $originalAcl.Access

# --- 2. Create a new ACL and set owner/group ---

$acl = New-Object System.Security.AccessControl.DirectorySecurity
$acl.SetOwner([System.Security.Principal.NTAccount]$owner)
$acl.SetGroup([System.Security.Principal.NTAccount]$group)

# --- 3. Sort ACEs into Deny and Allow lists ---

$denyRules = New-Object System.Collections.Generic.List[System.Security.AccessControl.FileSystemAccessRule]
$allowRules = New-Object System.Collections.Generic.List[System.Security.AccessControl.FileSystemAccessRule]

foreach ($rule in $accessRules) {
    if ($rule.AccessControlType -eq "Deny") {
        $denyRules.Add($rule)
    } else {
        $allowRules.Add($rule)
    }
}

# --- 4. Apply ACEs in the correct order (Deny first, then Allow) ---

foreach ($rule in $denyRules) {
    $acl.AddAccessRule($rule)
}
foreach ($rule in $allowRules) {
    $acl.AddAccessRule($rule)
}

# --- 5. Apply the new ACL and *then* disable inheritance ---

Set-Acl -Path $path -AclObject $acl  # Apply the rebuilt ACL

# Disable inheritance and copy existing (now explicit) rules
$acl = Get-Acl -Path $path # Get ACL *after* rebuild
$acl.SetAccessRuleProtection($true, $false)
Set-Acl -Path $path -AclObject $acl # Apply inheritance change

Write-Host "ACL rebuilt, reordered, and inheritance disabled for '$path'"


# -- test

# --- 6. Remove ACEs for "MSI\Users" (or the local equivalent) ---

# Get the computer name (to determine if we need domain or local "Users")
$computerName = [System.Environment]::MachineName

# Try to get the domain "Users" group SID first
try {
    $domainUsers = New-Object System.Security.Principal.NTAccount($computerName, "Users")
    $domainUsersSID = $domainUsers.Translate([System.Security.Principal.SecurityIdentifier])
}
catch {
    # If getting the domain group fails (e.g., not joined to a domain),
    # it's likely a local group.  No need to do anything here, the next step will get it
    $domainUsersSID = $null #Explicit is better than implicit
}

# Get the *local* "Users" group SID (this will *always* work)
$localUsers = New-Object System.Security.Principal.NTAccount("Users") #Builtin group does not require computername
$localUsersSID = $localUsers.Translate([System.Security.Principal.SecurityIdentifier])

# Determine which SID to use (prioritize domain if it exists)
if ($domainUsersSID) {
    $usersSID = $domainUsersSID
    Write-Host "Removing ACEs for domain group: $computerName\Users"
}
else {
    $usersSID = $localUsersSID
    Write-Host "Removing ACEs for local group: Users"
}

# Get the ACL *again* (after disabling inheritance) - crucial for accurate removal.
$acl = Get-Acl -Path $path

# Remove *all* access rules for the identified SID.  This is the most robust approach.
$acl.PurgeAccessRules($usersSID)

# Apply the changes
Set-Acl -Path $path -AclObject $acl
}
}
