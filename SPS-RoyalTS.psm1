Add-Type -AssemblyName System.Windows.Forms
#region Define the Module Class and Enums
Enum MessageType {
    Information
    Warning
    Error
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
        Return ($this | ConvertTo-Json -Depth 100 | Write-Host)
    }
    static [RoyalTSJson] FromRules([System.Xml.XmlElement] ${DynamicRules}) {
        $RoyalTSJson = [RoyalTSJson]::new()
        # Handle the Rules
        $Rules = $DynamicRules.Rules
        if ($Rules) {
            # Handle the ADGroupRules
            $AllADGroupRules = $Rules.ADGroupRules
            if ($AllADGroupRules) {
                ForEach($ADGroupRule in $AllADGroupRules.ADGroup) {
                    # Get AdGroups and their Computers matching this rule and add them to the RoyalTSJson
                }
            }
            # Handle the ComputerNameRules
            $AllComputerNameRules = $Rules.ComputerNameRules
            if ($AllComputerNameRules) {
                ForEach($ADGroupRule in $AllComputerNameRules.ComputerName) {
                    # Get the computers matching this rule and add them to the RoyalTSJson
                }
            }
        }
        # Handle the SingleComputers
        $AllSingleComputers = $DynamicRules.SingleComputers
        if ($AllSingleComputers) {
            ForEach($SingleComputer in $AllSingleComputers.SingleComputer) {
                # Get the single computer and add it to the RoyalTSJson
            }
        }
        Return $RoyalTSJson
    }
}
Class RoyalTSObject {
    [String] ${Id} = [Guid]::NewGuid().ToString()
    [String] ${Name}
    [String] ${Description}
    [RoyalTSObjectType] ${Type}
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
    RoyalTSRemoteDesktopConnection(){}
    RoyalTSRemoteDesktopConnection([String] $Name, [String] $Description, [String] $ComputerName, [String] $UserName, [String] $Path) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Type = [RoyalTSObjectType]::RemoteDesktopConnection
        $this.ComputerName = $ComputerName
        $this.UserName = $UserName
        $this.Path = $Path
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
        
    }Else{
        $Message = "The default configuration file [$($DefaultConfigPath)} does not exist"
        Throw $Message
    }
    
}
Function New-RoyalTSDynamicFolder {
    [CmdletBinding()]
    Param(
        [String] ${Name} = 'Default',
        [String] ${ConfigurationFile} = $(Join-Path -Path $Env:APPDATA -ChildPath "SPS-RoyalTS\RoyalTSConfiguration_$($Name).xml")
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
            $Json = [RoyalTSJson]::FromRules($DynamicFolderConfig)
            $Json.ToConsole()
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