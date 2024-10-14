Add-Type -AssemblyName System.Windows.Forms
#region Define the Module Class and Enums
Enum MessageType {
    Information
    Warning
    Error
}
Class RoyalTSRegexp {
    # this class help to define a Regexp pattern and the group order
    [String] ${Property} = 'SamAccountName'
    [String] ${Pattern}
    [String[]] ${GroupOrder} = @()
    RoyalTSRegexp([String] $Pattern) {
        $this.Pattern = $Pattern
    }
    RoyalTSRegexp([String] $Pattern,[String[]] $GroupOrder) {
        $this.Pattern = $Pattern
        $this.GroupOrder = $GroupOrder
    }
    RoyalTSRegexp([System.Xml.XmlElement] $Regexp) {
        $this.Property = $Regexp.Property
        $this.Pattern = $Regexp.Pattern
        $this.GroupOrder = $Regexp.GroupOrder -split ',|;|\s'
    }
    [Microsoft.PowerShell.Commands.MatchInfo] Match([String] $Value) {
        $Match = $Value | Select-String -Pattern $this.Pattern -AllMatches
        Return $Match
    }
    [String] GetPath([String] $Path, [String] $Value) {
        $Match = $this.Match($Value)
        if ($Match) {
            $Groups = $Match.Matches.Groups | Where-Object {$_.Name -ne 0} # exclude the group 0 as it's the whole regex match
            if ($this.GroupOrder.Count -gt 0) {
                ForEach ($GroupName in $this.GroupOrder) {
                    $Value = $Groups | Where-Object Name -eq $GroupName | Select-Object -ExpandProperty 'Value'
                    if ($Value -notlike '') {
                        $Path = "$($Path)\$($Value)"
                    }
                }
            }Else{
                # no order defined use the detected order
                ForEach ($GroupName in $Groups) {
                    $Value = $GroupName.Value
                    if ($Value -notlike '') {
                        $Path = "$($Path)\$($Value)"
                    }
                }
            }
        }
        Return $Path
    }
}
Class RoyalTSADGroupRule {
    [String] ${Name}
    [String] ${Domain}
    [String] ${Path}
    [String] ${UserName} = $Env:UserName
    [String] ${DefaultComputerName}
    [RoyalTSRegexp] ${GroupNameRegexp}
    [RoyalTSRegexp] ${ComputerNameRegexp}
    RoyalTSADGroupRule([String] $Name, [String] $Domain) {
        $this.Name = $Name
        $this.Domain = $Domain
    }
    RoyalTSADGroupRule([System.Xml.XmlElement] $ADGroup,[String] ${Domain}) {
        $this.Name = $ADGroup.Name
        $this.Domain = $Domain
        $this.Path = $ADGroup.Path
        $this.GroupNameRegexp = [RoyalTSRegexp]::new($ADGroup.GroupNameRegexp)
        $this.ComputerNameRegexp = [RoyalTSRegexp]::new($ADGroup.ComputerNameRegexp)
        $this.UserName = $ADGroup.UserName
        $this.DefaultComputerName = $ADGroup.DefaultComputerName
    }
    [System.Collections.Generic.List[RoyalTSObject]] GetComputers() {
        if ($this.Name -notlike '') {
            $List = [System.Collections.Generic.List[RoyalTSObject]]::new()
            # Retrieve the default FQDN
            # $DomainFQDN = Get-ADDomain -Server $this.Domain | Select-Object -ExpandProperty 'DNSRoot'
            # Get the AD Group matching the Name
            $ADGroupFilter = "GroupCategory -eq 'Security' -and ObjectClass -eq 'Group' -and SamAccountName -like '$($This.Name)'"
            $SplatGetGroup = @{
                Filter = $ADGroupFilter
                Server = $this.Domain
            }
            Try {
                $AllADGroups = Get-ADGroup @SplatGetGroup
            }Catch{
                $Message = "An Unexpected error occurs while getting the AD Groups using filter [$($ADGroupFilter)] on Domain [$($This.Domain)]: $($_.Exception.Message)"
                Write-RTSLog -Message $Message -Type [MessageType]::Error
                Throw $Message
            }
            $RootPath = $this.Path
            ForEach ($ADGroup in $AllADGroups) {
                $ThisGroupPath = $RootPath
                if ($This.GroupNameRegexp.Pattern -notlike '') {
                    # there is a pattern defined create the path from it
                    # $RegexResult = $ADGroup.$($this.GroupNameRegexp.Property) | Select-String -Pattern $This.GroupNameRegexp.Pattern -AllMatches
                    # if ($RegexResult) {
                    #     # the regex match, apply the order to define the Path
                    #     $RXgroups = $RegexResult.Matches.Groups | Where-Object {$_.Name -ne 0} # Exclude the group 0 as it's the whole regex match
                    #     if ($This.GroupNameRegexp.GroupOrder.Count -gt 0) {
                    #         # Apply the group Order
                    #         ForEach ($RXGroupName in $This.GroupNameRegexp.GroupOrder) {
                    #             $Value = $RXGroups | Where-Object Name -eq $RXGroupName | Select-Object -ExpandProperty 'Value'
                    #             if ($Value -notlike '') {
                    #                 $ThisGroupPath = "$($ThisGroupPath)\$($Value)"
                    #             }
                    #         }
                    #     }Else{
                    #         # no order defined use the detected order
                    #         ForEach ($RXGroupName in $RXGroups) {
                    #             $Value = $RXGroupName.Value
                    #             if ($Value -notlike ''){
                    #                 $ThisGroupPath = "$($ThisGroupPath)\$($Value)"
                    #             }
                    #         }
                    #     }
                    # }
                    $ThisGroupPath = $This.GroupNameRegexp.GetPath($ThisGroupPath,$ADGroup.$($this.GroupNameRegexp.Property))
                }
                # Retrieve all computers from the group
                Try {
                    $AllGroupMembers = Get-ADGroupMember -Identity $ADGroup -Server $this.Domain | Sort-Object Name
                }Catch{
                    $Message = "An Unexpected error occurs while getting the AD Group Members using [$($ADGroup)] on Domain [$($This.Domain)]"
                    Write-RTSLog -Message $Message -Type [MessageType]::Error
                    Throw $Message
                }
                ForEach ($Computer in $AllGroupMembers) {
                    $ThisComputerPath = $ThisGroupPath
                    # Build the path if apply
                    if ($this.ComputerNameRegexp.Pattern -notlike '') {
                        # there is a patter defined create the path from it
                        # $RegexResult = $Computer.($this.ComputerNameRegexp.Property) | Select-String -Pattern $This.ComputerNameRegexp.Pattern -AllMatches
                        # if ($RegexResult) {
                        #     $RXgroups = $RegexResult.Matches.Groups | Where-Object {$_.Name -ne 0} # Exclude the group 0 as it's the whole regex match
                        #     if ($This.ComputerNameRegexp.GroupOrder -notlike '') {
                        #         ForEach ($RXGroupName in $This.ComputerNameRegexp.GroupOrder) {
                        #             $Value = $RXGroups | Where-Object Name -eq $RXGroupName | Select-Object -ExpandProperty 'Value'
                        #             if ($Value -notlike '') {
                        #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                        #             }
                        #         }
                        #     }Else{
                        #         # no order defined use the detected order
                        #         ForEach ($RXGroupName in $RXGroups) {
                        #             $Value = $RXGroupName.Value
                        #             if ($Value -notlike ''){
                        #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                        #             }
                        #         }
                        #     }
                        # }
                        $ThisComputerPath = $This.ComputerNameRegexp.GetPath($ThisComputerPath,$Computer.($this.ComputerNameRegexp.Property))
                    }
                    # Get the AD Object
                    $Computer = Get-ADComputer -Identity $Computer
                    # Build the computer object
                    $ComputerObject = [RoyalTSRemoteDesktopConnection]::New()
                    # $ComputerObject.ID = $GroupMember.objectGUID.ToString()
                    $ComputerObject.Name = $Computer.Name
                    $ComputerObject.Description = $Computer.distinguishedName
                    if ($this.DefaultComputerName -notlike '') {
                        $ComputerObject.ComputerName = $this.DefaultComputerName
                    }Else{
                        $ComputerObject.ComputerName = "$($Computer.DNSHostName)"
                    }
                    $ComputerObject.UserName = $this.UserName
                    $ComputerObject.Path = $ThisComputerPath
                    # Add the computer to the list
                    $List.Add($ComputerObject)
                }
            }
            Return $List
        }Else{
            Return $null
        }
    }
}
Class RoyalTSADComputerRule {
    [String] ${Name}
    [String] ${Domain}
    [String] ${Path}
    [String] ${UserName} = $Env:UserName
    [String] ${DefaultComputerName}
    [RoyalTSRegexp] ${ComputerNameRegexp}
    RoyalTSADComputerRule([String] $Name, [String] $Domain) {
        $this.Name = $Name
        $this.Domain = $Domain
    }
    RoyalTSADComputerRule([System.Xml.XmlElement] $ADComputer,[String] ${Domain}) {
        $this.Name = $ADComputer.Name
        $this.Domain = $Domain
        $this.Path = $ADComputer.Path
        $this.ComputerNameRegexp = [RoyalTSRegexp]::new($ADComputer.ComputerNameRegexp)
        $this.UserName = $ADComputer.UserName
        $this.DefaultComputerName = $ADComputer.DefaultComputerName
    }
    [System.Collections.Generic.List[RoyalTSObject]] GetComputers() {
        if ($this.Name -notlike '') {
            $List = [System.Collections.Generic.List[RoyalTSObject]]::new()
            # Get the computers matching the rule
            $Filter = "SamAccountName -like '$($this.Name)'"
            $SplatGetComputer = @{
                Filter = $Filter
                Server = $this.Domain
            }
            $AllComputers = Get-ADComputer @SplatGetComputer
            $RootPath = $this.Path
            ForEach ($Computer in $AllComputers) {
                $ThisComputerPath = $RootPath
                # Build the path if apply
                if ($this.ComputerNameRegexp.Pattern -notlike '') {
                    # there is a patter defined create the path from it
                    # $RegexResult = $Computer.($this.ComputerNameRegexp.Property) | Select-String -Pattern $This.ComputerNameRegexp.Pattern -AllMatches
                    # if ($RegexResult) {
                    #     $RXgroups = $RegexResult.Matches.Groups | Where-Object {$_.Name -ne 0} # Exclude the group 0 as it's the whole regex match
                    #     if ($This.ComputerNameRegexp.GroupOrder -notlike '') {
                    #         ForEach ($RXGroupName in $This.ComputerNameRegexp.GroupOrder) {
                    #             $Value = $RXGroups | Where-Object Name -eq $RXGroupName | Select-Object -ExpandProperty 'Value'
                    #             if ($Value -notlike '') {
                    #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                    #             }
                    #         }
                    #     }Else{
                    #         # no order defined use the detected order
                    #         ForEach ($RXGroupName in $RXGroups) {
                    #             $Value = $RXGroupName.Value
                    #             if ($Value -notlike ''){
                    #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                    #             }
                    #         }
                    #     }
                    # }
                    $ThisComputerPath = $This.ComputerNameRegexp.GetPath($ThisComputerPath,$Computer.($this.ComputerNameRegexp.Property))
                }
                # Build the computer object
                $ComputerObject = [RoyalTSRemoteDesktopConnection]::New()
                # $ComputerObject.ID = $GroupMember.objectGUID.ToString()
                $ComputerObject.Name = $Computer.Name
                $ComputerObject.Description = $Computer.distinguishedName
                if ($this.DefaultComputerName -notlike '') {
                    $ComputerObject.ComputerName = $this.DefaultComputerName
                }Else{
                    $ComputerObject.ComputerName = "$($Computer.DNSHostName)"
                }
                $ComputerObject.UserName = $this.UserName
                $ComputerObject.Path = $ThisComputerPath
                # Add the computer to the list
                $List.Add($ComputerObject)
            }
            Return $List
        }Else{
            Return $Null
        }
    }
}
Class RoyalTSVIComputerRule {
    [String] ${Server}
    [String] ${Name}
    [String] ${Path}
    [String] ${UserName}
    [String] ${Domain}
    [String] ${DefaultComputerName}
    [RoyalTSRegexp] ${ComputerNameRegexp}
    RoyalTSVIComputerRule([System.Xml.XmlElement] $VIComputer,[String] ${Domain}) {
        $this.Server = $VIComputer.Server
        $this.Name = $VIComputer.Name
        $this.Domain = $Domain
        $this.Path = $VIComputer.Path
        $this.UserName = $VIComputer.UserName
        $this.DefaultComputerName = $VIComputer.DefaultComputerName
        $this.ComputerNameRegexp = [RoyalTSRegexp]::new($VIComputer.ComputerNameRegexp)
    }
    [System.Collections.Generic.List[RoyalTSObject]] GetComputers() {
        if ($this.Name -notlike '') {
            $List = [System.Collections.Generic.List[RoyalTSObject]]::new()
            # Connect the VI Server
            Try {
                $VIServer = Connect-VIServer -Server $This.Server -Verbose:$False
            }Catch{
                $Message = "Unable to connect to VIServer [$($This.Server)]: $($_.Exception.Message)"
                Write-RTSLog -Message $Message
                Throw $Message
            }
            # Get the VM matching the name
            $AllVMComputers = Get-VM -Name $this.Name -Server $VIServer -Verbose:$False | Sort-Object 'Name'
            $RootPath = $this.Path
            ForEach ($Computer in $AllVMComputers) {
                $ThisComputerPath = $RootPath
                if ($this.ComputerNameRegexp.Pattern -notlike '') {
                    # there is a patter defined create the path from it
                    # $RegexResult = $Computer.($this.ComputerNameRegexp.Property) | Select-String -Pattern $This.ComputerNameRegexp.Pattern -AllMatches
                    # if ($RegexResult) {
                    #     $RXgroups = $RegexResult.Matches.Groups | Where-Object {$_.Name -ne 0} # Exclude the group 0 as it's the whole regex match
                    #     if ($This.ComputerNameRegexp.GroupOrder -notlike '') {
                    #         ForEach ($RXGroupName in $This.ComputerNameRegexp.GroupOrder) {
                    #             $Value = $RXGroups | Where-Object Name -eq $RXGroupName | Select-Object -ExpandProperty 'Value'
                    #             if ($Value -notlike '') {
                    #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                    #             }
                    #         }
                    #     }Else{
                    #         # no order defined use the detected order
                    #         ForEach ($RXGroupName in $RXGroups) {
                    #             $Value = $RXGroupName.Value
                    #             if ($Value -notlike ''){
                    #                 $ThisComputerPath = "$($ThisComputerPath)\$($Value)"
                    #             }
                    #         }
                    #     }
                    # }
                    $ThisComputerPath = $This.ComputerNameRegexp.GetPath($ThisComputerPath,$Computer.($this.ComputerNameRegexp.Property))
                }
                # Build the computer object
                $ComputerObject = [RoyalTSRemoteDesktopConnection]::New()
                # $ComputerObject.ID = $GroupMember.objectGUID.ToString()
                $ComputerObject.Name = $Computer.Name
                $ComputerObject.Description = $Computer.Guest
                if ($this.DefaultComputerName -notlike '') {
                    $ComputerObject.ComputerName = $this.DefaultComputerName
                }Else{
                    # Search for the ip Address
                    $IPAddress = $Computer.ExtensionData.Guest.IPAddress
                    if ($IPAddress -notlike '') {
                        $ComputerObject.ComputerName = $IPAddress
                    }Else{
                        $ComputerObject.ComputerName = $Computer.ExtensionData.Guest.HostName
                    }
                }
                $ComputerObject.UserName = $this.UserName
                $ComputerObject.Path = $ThisComputerPath
                # Add the computer to the list
                $List.Add($ComputerObject)

            }
            # Disconnect the VI Server
            Disconnect-VIServer -Server $VIServer -Verbose:$False -Force -Confirm:$False -ErrorAction SilentlyContinue | out-null
            Return $List
        }Else{
            Return $Null
        }
    }
}
#endregion Define the Module Class and Enums
#region Define the RoyalTSObjects Class and Enums
Enum RoyalTSObjectType {
    Folder
    Credential
    DynamicCredential
    ToDo
    Information
    CommandTask
    KeySequenceTask
    SecureGateway
    RoyalServer
    RemoteDesktopGateway
    RemoteDesktopConnection
    TerminalConnection
    WebConnection
    VNCConnection
    FileTransferConnection
    TeamViewerConnection
    ExternalApplicationConnection
    PerformanceConnection
    VMwareConnection
    HyperVConnection
    WindowsEventsConnection
    WindowsServicesConnection
    WindowsProcessesConnection
    TerminalServicesConnection
    PowerShellConnection
}
Class RoyalTSJson {
    [System.Collections.Generic.List[RoyalTSObject]] ${Objects} = [System.Collections.Generic.List[RoyalTSObject]]::new()
    RoyalTSJson(){}
    [void] Add([RoyalTSObject] $Object) {
        $this.Objects.Add($Object)
    }
    [String] ToConsole() {
        Return ($this | ConvertTo-Json -Depth 100 | Write-output)
    }
    static [RoyalTSJson] FromRules([System.Xml.XmlElement] ${DynamicRules},[String] $Domain) {
        $RoyalTSJson = [RoyalTSJson]::new()
        # Handle the Rules
        $Rules = $DynamicRules.Rules
        if ($Rules) {
            # Handle the ADGroupRules
            $AllADGroupRules = $Rules.ADGroupRules
            if ($AllADGroupRules) {
                # ForEach($ADGroupRule in $AllADGroupRules.ADGroup) {
                #     # Get AdGroups and their Computers matching this rule and add them to the RoyalTSJson
                #     $ADGroupRuleObject = [RoyalTSADGroupRule]::new($ADGroupRule,$Domain)
                #     ForEach($ComputerObject in $ADGroupRuleObject.GetComputers()) {
                #         $RoyalTSJson.Add($ComputerObject)
                #     }
                # }
                # using the pipeline
                $AllADGroupRules.ADGroup | ForEach-Object {([RoyalTSADGroupRule]::new($_,$Domain)).GetComputers()} | ForEach-Object {$RoyalTSJson.Add($_)}
            }
            # Handle the ComputerNameRules
            $AllComputerNameRules = $Rules.ComputerNameRules
            if ($AllComputerNameRules) {
                # ForEach($ADComputerNameRule in $AllComputerNameRules.ComputerName) {
                #     # Get the computers matching this rule and add them to the RoyalTSJson
                #     $ADComputerRuleObject = [RoyalTSADComputerRule]::new($ADComputerNameRule,$Domain)
                #     ForEach($ComputerObject in $ADComputerRuleObject.GetComputers()) {
                #         $RoyalTSJson.Add($ComputerObject)
                #     }
                # }
                $AllComputerNameRules.ComputerName | ForEach-Object {([RoyalTSADComputerRule]::new($_,$Domain)).GetComputers()} | ForEach-Object {$RoyalTSJson.Add($_)}
            }
            $AllVIComputers = $Rules.VIComputerRules
            if ($AllVIComputers) {
                # ForEach($VIComputer in $AllVIComputers.VIComputer) {
                #     #  Get the VI computers matching this rule and add them to the RoyalTSJson
                #     $VIComputerRuleObject = [RoyalTSVIComputerRule]::new($VIComputer,$Domain)
                #     ForEach ($ComputerObject in $VIComputerRuleObject.GetComputers()) {
                #         $RoyalTSJson.Add($ComputerObject)
                #     }
                # }
                $AllVIComputers.VIComputer | ForEach-Object {([RoyalTSVIComputerRule]::new($_,$Domain)).GetComputers()} | ForEach-Object {$RoyalTSJson.Add($_)}
            }
        }
        # Handle the SingleComputers
        $AllSingleComputers = $DynamicRules.SingleComputers
        if ($AllSingleComputers) {
            ForEach($Computer in $AllSingleComputers.SingleComputer) {
                if ($Computer.Name -Notlike '') {
                    # Get the single computer and add it to the RoyalTSJson
                    # $ComputerObject = [RoyalTSRemoteDesktopConnection]::New()
                    # $ComputerObject.Name = $Computer.Name
                    # $ComputerObject.UserName = $Computer.UserName
                    # $ComputerObject.Path = $Computer.Path
                    # if ($Computer.DefaultComputerName -like '') {
                    #     $ComputerObject.ComputerName = $Computer.ComputerName
                    # }Else{
                    #     $ComputerObject.ComputerName = $Computer.DefaultComputerName
                    # }
                    # $RoyalTSJson.Add($ComputerObject)
                    $RoyalTSJson.Add([RoyalTSRemoteDesktopConnection]::new($Computer))
                }
            }
        }
        Return $RoyalTSJson
    }
    [void] BuildFolders() {
        # Build the folders based on the objects path
    }
}
Class RoyalTSObject {
    [String] ${Id} = [Guid]::NewGuid().ToString()
    [String] ${Name}
    [String] ${Description}
    [String] ${Type}
    RoyalTSObject(){}
    RoyalTSObject([String] $Name, [String] $Description) {
        $this.Name = $Name
        $this.Description = $Description
    }
    RoyalTSObject([String] $Name, [String] $Description, [RoyalTSObjectType] $Type) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Type = $Type
    }
}
Class RoyalTSFolder : RoyalTSObject {
    [System.Collections.Generic.List[RoyalTSObject]] ${Objects} = [System.Collections.Generic.List[RoyalTSObject]]::new()
    RoyalTSFolder(){
        $this.Type = [RoyalTSObjectType]::Folder
    }
    RoyalTSFolder([String] $Name) {
        $this.Name = $Name
        $this.Type = [RoyalTSObjectType]::Folder
    }
    RoyalTSFolder([String] $Name, [String] $Description) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Type = [RoyalTSObjectType]::Folder
    }
    [void] Add([Object] $Object) {
        $this.Objects.Add($Object)
    }
}
Class RoyalTSRemoteDesktopConnection : RoyalTSObject {
    [String] ${ComputerName}
    [String] ${UserName}
    [String] ${Path}
    [Boolean] ${CredentialsFromParent} = $false
    RoyalTSRemoteDesktopConnection(){
        $this.Type = [RoyalTSObjectType]::RemoteDesktopConnection
    }
    RoyalTSRemoteDesktopConnection([String] $Name, [String] $Description, [String] $ComputerName, [String] $UserName, [String] $Path) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Type = [RoyalTSObjectType]::RemoteDesktopConnection
        $this.ComputerName = $ComputerName
        $this.UserName = $UserName
        $this.Path = $Path
        $this.Type = [RoyalTSObjectType]::RemoteDesktopConnection
    }
    RoyalTSRemoteDesktopConnection([System.Xml.XmlElement] $XMLElement) {
        $this.Name = $XMLElement.Name
        $this.Description = $XMLElement.Description
        if ($this.DefaultComputerName -like '') {
            $this.ComputerName = $XMLElement.ComputerName
        }Else{
            $this.ComputerName = $XMLElement.DefaultComputerName
        }
        $this.UserName = $XMLElement.UserName
        $this.Path = $XMLElement.Path
        $this.Type = [RoyalTSObjectType]::RemoteDesktopConnection
    }
}
#endregion Define the RoyalTSObjects Class and Enums
Function Write-RTSLog {
    [CmdletBinding()]
    Param(
        [String] ${Message},
        [MessageType] ${Type} = [MessageType]::Information
    )
    $LogPath = Join-Path -Path $Env:APPDATA -ChildPath 'SPS-RoyalTS\RoyalTSLog.log'
    if (-not (Test-Path -Path $LogPath)) {
        Try {
            New-Item -Path $LogPath -ItemType File -Force | Out-Null
        }Catch {
            Throw "An unexpected error occured while creating the log file $($LogPath): $($_.Exception.Message)"
        }

    }
    $LogMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$($Type.ToString())] $($Message)"
    Try {
        Add-Content -Path $LogPath -Value $LogMessage | Out-Null
    }Catch{
        Throw "An unexpected error occured while writing the log message $($LogMessage) to the log file $($LogPath): $($_.Exception.Message)"
    }
}
Function Show-RTSMessageBox {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,Position=0)]
        [String] ${Message},
        [String] ${Title} = 'SPS RoyalTS Module',
        [System.Windows.Forms.MessageBoxButtons] ${Buttons} = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon] ${Icon} = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
}
Function New-DefaultRoyalTSConfiguration {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    Write-Verbose -Message "Processing the function $($MyInvocation.MyCommand)"
    # Create the configuration file
    $DefaultConfigPath = "$($PSScriptRoot)\DefaultConfig.xml"
    if ((Test-Path -path $DefaultConfigPath -ErrorAction Ignore) -eq $True) {
        Try {
            [XML] $Configuration = Get-Content -Path $DefaultConfigPath -ErrorAction Stop
        }Catch {
            Throw $_
        }
        # Create the path if it does not exist
        if ((Test-Path -Path $Path -ErrorAction Ignore) -eq $False) {New-Item -Path $Path -ItemType File -Force | Out-Null}
        # Save the configuration in the default path
        Try {
            $Configuration.Save($Path)
        }Catch{
            Throw $_
        }
        $MSGBoxResult = Show-RTSMessageBox -Message @"
The default configuration file has been created at [$($Path)]

Open for editing?
"@ -Buttons YesNo -Icon Information
        if ($MSGBoxResult -eq 'Yes') {
            Start-Process -FilePath 'Notepad.exe' -Argumentlist $Path
        }
    }Else{
        $Message = "The default configuration file [$($DefaultConfigPath)} does not exist"
        Throw $Message
    }

}
Function New-RoyalTSDynamicFolder {
    [CmdletBinding()]
    Param(
        [String] ${Name} = 'Default',
        [String] ${ConfigurationFile} = $(Join-Path -Path $Env:APPDATA -ChildPath "SPS-RoyalTS\RoyalTSConfiguration_$($Name).xml"),
        [Parameter(Mandatory)]
        [String] ${Domain}
    )
    Begin {
        Write-Verbose -Message "Starting the function $($MyInvocation.MyCommand)"
        # Check if the configuration file exists
        if (-not (Test-Path -Path $ConfigurationFile)) {
            Write-Verbose -Message "The configuration file [$($ConfigurationFile)] does not exist"
            # Create the configuration file
            Try {
                New-DefaultRoyalTSConfiguration -Path $ConfigurationFile | Out-Null
                Write-Verbose -Message "The configuration file [$($ConfigurationFile)] has been created"
            }
            Catch {
                $Message = "An unexpected error occured while creating the configuration file [$($ConfigurationFile)}: $($_.Exception.Message)"
                Write-RTSLog -Message $Message -Type Error
                Throw $Message
            }
        }
        # Read the configuration file
        [XML] $Configuration = Get-Content -Path $ConfigurationFile -Raw
        # Check if the configuration file has the right structure
        if (-not $Configuration.RoyalTSConfiguration) {
            $Message = "The configuration file [$($ConfigurationFile)] does not have the right structure"
            Write-RTSLog -Message $Message -Type Error
            Throw $Message
        }
        # Check if the configuration file call a remote file if yes load it instead
        if ($Configuration.RoyalTSConfiguration.RemoteFile) {
            Write-Verbose -Message "The configuration file [$($ConfigurationFile)] call a remote file [($($Configuration.RoyalTSConfiguration.RemoteFile))]"
            # Get the remote file
            $RemoteFile = $Configuration.RoyalTSConfiguration.RemoteFile
            # Check if the remote file exists
            if (-not (Test-Path -Path $RemoteFile)) {
                $Message = "The remote file [$($RemoteFile)] declared in [$($ConfigurationFile) - (\RoyalTSConfiguration\RemoteFile)] does not exist"
                Write-RTSLog -Message $Message -Type Error
                Throw $Message
            }
            # Read the remote file
            [XML] $Configuration = Get-Content -Path $RemoteFile -Raw
            $DynamicFolderConfig = $Configuration.RoyalTSConfiguration.DynamicFolder
            # Check if the remote file has the right structure
            if (-not $RemoteConfiguration.RoyalTSConfiguration) {
                $Message = "The remote file [$($RemoteFile)] does not have the right structure (Missing \RoyalTSConfiguration)"
                Write-RTSLog -Message $Message -Type Error
                Throw $Message
            }
        }Else{
            $DynamicFolderConfig = $Configuration.RoyalTSConfiguration.DynamicFolder
        }
    }
    Process {
        Write-Verbose -Message "Processing the function $($MyInvocation.MyCommand)"
        # Check if the configuration file has dynamic folders
        if ($DynamicFolderConfig) {
            # Get the dynamic folders rule
            $RoyalTSObject = [RoyalTSJson]::FromRules($DynamicFolderConfig,$Domain)
            Return $RoyalTSObject
        }Else{
            $Message = "The configuration file $($ConfigurationFile) does not have any dynamic folder entry the function $($MyInvocation.MyCommand) will not do anything"
            Write-RTSLog -Message $Message -Type Warning
            Write-Warning $Message
        }
    }
    End {
        Write-Verbose -Message "Ending the function $($MyInvocation.MyCommand)"
    }
}

#region expose classes and enums
##Expose the classes and enums
# Define the types to export with type accelerators.
$ExportableTypes =@(
    [RoyalTSObjectType],
    [RoyalTSJson],
    [RoyalTSObject],
    [RoyalTSFolder],
    [RoyalTSRemoteDesktopConnection]
)
# Get the internal TypeAccelerators class to use its static methods.
$TypeAcceleratorsClass = [psobject].Assembly.GetType(
    'System.Management.Automation.TypeAccelerators'
)
# Ensure none of the types would clobber an existing type accelerator.
# If a type accelerator with the same name exists, throw an exception.
$ExistingTypeAccelerators = $TypeAcceleratorsClass::Get
foreach ($Type in $ExportableTypes) {
    if ($Type.FullName -in $ExistingTypeAccelerators.Keys) {
        $Message = @(
            "Unable to register type accelerator '$($Type.FullName)'"
            'Accelerator already exists.'
        ) -join ' - '
        Write-Warning -Message $Message
    }
}
# Add type accelerators for every exportable type.
foreach ($Type in $ExportableTypes) {
    $TypeAcceleratorsClass::Add($Type.FullName, $Type)
}
# Remove type accelerators when the module is removed.
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    foreach($Type in $ExportableTypes) {
        $TypeAcceleratorsClass::Remove($Type.FullName)
    }
}.GetNewClosure()
#endregion expose classes and enums
