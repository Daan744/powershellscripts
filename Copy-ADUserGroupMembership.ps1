[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceUser,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetUser,

    [switch]$SkipDefaultPrimaryGroup
)

Import-Module ActiveDirectory

$source = Get-ADUser -Identity $SourceUser -Properties SamAccountName
$target = Get-ADUser -Identity $TargetUser -Properties SamAccountName

if (-not $source) {
    throw "Brongebruiker '$SourceUser' niet gevonden in Active Directory."
}

if (-not $target) {
    throw "Doelgebruiker '$TargetUser' niet gevonden in Active Directory."
}

$sourceGroups = Get-ADPrincipalGroupMembership -Identity $source
$targetGroups = Get-ADPrincipalGroupMembership -Identity $target

if ($SkipDefaultPrimaryGroup) {
    $sourceGroups = $sourceGroups | Where-Object { $_.Name -ne 'Domain Users' }
}

$targetGroupDns = $targetGroups.DistinguishedName

$groupsToAdd = $sourceGroups | Where-Object {
    $_.DistinguishedName -notin $targetGroupDns
}

if (-not $groupsToAdd) {
    Write-Host "Geen nieuwe groepen om toe te voegen. '$($target.SamAccountName)' heeft alle groepen al."
    return
}

foreach ($group in $groupsToAdd) {
    if ($PSCmdlet.ShouldProcess("$($target.SamAccountName)", "Toevoegen aan groep '$($group.Name)'")) {
        Add-ADGroupMember -Identity $group.DistinguishedName -Members $target.DistinguishedName
        Write-Host "Toegevoegd: $($target.SamAccountName) -> $($group.Name)"
    }
}

Write-Host "Klaar. $($groupsToAdd.Count) groep(en) verwerkt."
