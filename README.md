# SPS-RoyalTS
This module is used to interact with RoyalTS using PowerShell.

## Pre-requisites
- PowerShell 5.1 or later
- RoyalTS 5 or later

## Installation
```powershell
Import-Module .\SPS-RoyalTS.psd1
```

## Functions
### New-RoyalTSDynamicFolder
This function is used to create a new dynamic folder in RoyalTS.
it takes the following parameters:
- Name: The name of the config (so you can use different config). (default "Default")
- ConfigurationFile: The path to use (default "%Appdata%\SPS-RoyalTS\RoyalTSConfiguration_*Default*.xml").