# Export-ConditionalAccessPolicies.ps1
# Exports all Conditional Access policies with resolved security principal names

# Only import required Microsoft.Graph modules to avoid function overflow
$modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.DirectoryManagement')
foreach ($mod in $modules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Install-Module -Name $mod -Scope CurrentUser -Force
        Write-Host "Module $mod installed for the current user."
        Write-Host "Note: Modules installed with -Scope CurrentUser are located in:"
        Write-Host "$($env:USERPROFILE)\Documents\WindowsPowerShell\Modules"
        Write-Host "They will NOT appear in C:\Program Files\WindowsPowerShell\Modules unless installed system-wide."
        Write-Host "Please restart PowerShell and re-run this script."
        exit
    }
    Import-Module $mod -Force
}

if (-not (Get-Command Get-MgConditionalAccessPolicy -ErrorAction SilentlyContinue)) {
    Write-Host "Get-MgConditionalAccessPolicy cmdlet not found."
    Write-Host "Please update all Microsoft.Graph modules by running:"
    Write-Host "`n    Update-Module Microsoft.Graph -Force`n"
    Write-Host "Then restart PowerShell and try again."
    exit
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Policy.Read.All, Directory.Read.All"

# Get all Conditional Access policies
$policies = Get-MgConditionalAccessPolicy

# Helper function to resolve principal IDs to names
function Resolve-PrincipalName {
    param(
        [string]$Id,
        [string]$Type
    )
    switch ($Type) {
        'User'   { (Get-MgUser -UserId $Id -ErrorAction SilentlyContinue).DisplayName }
        'Group'  { (Get-MgGroup -GroupId $Id -ErrorAction SilentlyContinue).DisplayName }
        'Role'   { (Get-MgDirectoryRole -DirectoryRoleId $Id -ErrorAction SilentlyContinue).DisplayName }
        default  { $Id }
    }
}

# Prepare output array
$output = @()

foreach ($policy in $policies) {
    $resolvedUsers = @()
    $resolvedGroups = @()
    $resolvedRoles = @()
    if ($policy.Conditions.Users) {
        foreach ($userId in $policy.Conditions.Users.IncludeUsers) {
            $resolvedUsers += Resolve-PrincipalName -Id $userId -Type 'User'
        }
        foreach ($groupId in $policy.Conditions.Users.IncludeGroups) {
            $resolvedGroups += Resolve-PrincipalName -Id $groupId -Type 'Group'
        }
        foreach ($roleId in $policy.Conditions.Users.IncludeRoles) {
            $resolvedRoles += Resolve-PrincipalName -Id $roleId -Type 'Role'
        }
    }
    $output += [PSCustomObject]@{
        PolicyName = $policy.DisplayName
        State = $policy.State
        Users = $resolvedUsers -join ", "
        Groups = $resolvedGroups -join ", "
        Roles = $resolvedRoles -join ", "
        Applications = $policy.Conditions.Applications.IncludeApplications -join ", "
        Locations = $policy.Conditions.Locations.IncludeLocations -join ", "
        Platforms = $policy.Conditions.Platforms.IncludePlatforms -join ", "
        GrantControls = $policy.GrantControls.BuiltInControls -join ", "
        SessionControls = $policy.SessionControls | ConvertTo-Json -Compress
    }
}

# Export to CSV
$output | Export-Csv -Path "ConditionalAccessPolicies.csv" -NoTypeInformation
Write-Host "Export complete: ConditionalAccessPolicies.csv"