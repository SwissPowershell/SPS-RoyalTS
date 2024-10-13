$ModuleDescription = Get-ChildItem -Path $PSScriptRoot -Filter '*.psd1' | Select-Object -First 1
$ModuleDescription | Select-Object -ExpandProperty FullName | ForEach-Object {Import-Module $_ -Force}
# Set the most constrained mode
Set-StrictMode -Version Latest
# Set the error preference
$ErrorActionPreference = 'Stop'
# Set the verbose preference in order to get some insights
$VerbosePreference = 'Continue'
$DebugStart = Get-Date

############################
# Test your functions here #
############################

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

New-RoyalTSDynamicFolder -Verbose
# git add --all;Git commit -a -am 'Initial Commit';Git push

##################################
# End of the tests show metrics #
##################################

Write-Host '------------------- Ending script -------------------' -ForegroundColor Yellow
$TimeSpentInDebugScript = New-TimeSpan -Start $DebugStart -Verbose:$False -ErrorAction SilentlyContinue
$TimeUnits = [ordered]@{TotalDays = "$($TimeSpentInDebugScript.TotalDays) D.";TotalHours = "$($TimeSpentInDebugScript.TotalHours) h.";TotalMinutes = "$($TimeSpentInDebugScript.TotalMinutes) min.";TotalSeconds = "$($TimeSpentInDebugScript.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentInDebugScript.TotalMilliseconds) ms."}
foreach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentInDebugScript.$($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}}
if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentInDebugScript.Ticks) Ticks"}
Write-Host 'Ending : ' -ForegroundColor Yellow -NoNewLine
Write-Host $($MyInvocation.MyCommand) -ForegroundColor Magenta -NoNewLine
Write-Host ' - TimeSpent : ' -ForegroundColor Yellow -NoNewLine
Write-Host $TimeSpentString -ForegroundColor Magenta


BREAK
$Version = '0.0.1'
$ReleaseTag = 'alpha'
## Update the prerelease tag
$ModuleManifest = Test-ModuleManifest -Path $ModuleDescription
$PrereleaseTag = $ModuleManifest.PrivateData.PSData.Prerelease
$PrereleaseArray = $PrereleaseTag -split '\.|_'
[Int] $NewPrereleaseNumber = $PrereleaseArray[-1]
if ($PrereleaseArray[-1] -ne $ReleaseTag) {
    # Not the same tag so we reset the prerelease number
    $NewPrereleaseNumber = 0 
} else {
    # Same tag so we increment the prerelease number
    $NewPrereleaseNumber = $PrereleaseArray[-1] + 1
}
$NewPrereleaseTag = "$($ReleaseTag)_$($NewPrereleaseNumber + 1)"
## Update the module description
$PSD1Content = Get-Content -Path $ModuleDescription | ForEach-Object {$_ -replace "Prerelease = '$PrereleaseTag'", "Prerelease = '$NewPrereleaseTag'"}
Set-Content -Path $ModuleDescription -Value $PSD1Content
## Commit the changes
$CommitMessage = "Update module to version $($ModuleManifest.Version)-$($NewPrereleaseTag)"
Git add --all
Git commit -a -am $CommitMessage
Git push