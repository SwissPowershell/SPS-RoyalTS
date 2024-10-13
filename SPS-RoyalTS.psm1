Function New-DefaultRoyalTSConfiguration {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    Write-Verbose -Message "Processing the function $($MyInvocation.MyCommand)"
    # Create the configuration file
    $Configuration = @'
<?xml version="1.0" encoding="utf-8"?>
<RoyalTSConfiguration>
<!-- Help for the RoyalTSConfiguration file
    # Remote file
    You can define a remote file instead of using this configuration file the remote will be used as a source

    # Dynamic folder
    You can define a set of rules to create dynamic folder or add single computers
        ## Rules
            ### ADGroupRules
                The AD Group rule intend to add computers based on AD groups
                You can add multiple AD groups
                The fields are as follow :
                    Name : The name of the AD group (mandatory,Accepts wildcards)
                    Domain : The domain of the AD group (mandatory,Accepts wildcards,Accept Tokens)
                    Regexp : a Regularexpression (Optional)
                        Regex : the regular expression to apply (mandatory)
                        RXGroupOrder : if the regular expression contains groups you can define the order of the groups to use as subfolders (Optional)
            ### ComputerNameRules
                The Computer Name rule intend to add computers based on their name
                You can add multiple computer names
                The fields are as follow :
                    Name : The name of the computer (mandatory,Accepts wildcards,Accept Tokens)
                    Domain : The domain of the computer (mandatory,Accepts wildcards,Accept Tokens)
                    Regexp : a Regularexpression (Optional)
                        Regex : the regular expression to apply (mandatory)
                        RXGroupOrder : if the regular expression contains groups you can define the order of the groups to use as subfolders (Optional)
        ## SingleComputers
            You can add some single computers using their name
            the fileds are as follow :
                Name : The name of the computer (mandatory,Accepts wildcards,Accept Tokens)
                Domain : The domain of the computer
                UserName : The username to use to connect to the computer
        
        ! the default file contain one example for each type of rule
    
    ! Token and automatic variable
        the following token can be used in the configuration file
        $Name$ : The name of the dynamic folder
    
-->
    <RemoteFile></RemoteFile>
    <DynamicFolder>
        <Rules>
            <ADGroupRules>
                <ADGroup>
                    <Name></Name>
                    <Domain></Domain>
                    <Regexp>
                        <Regex></Regex>
                        <RXGroupOrder></RXGroupOrder>
                    </Regexp>
                </ADGroup>
            </ADGroupRules>
            <ComputerNameRules>
                <ComputerName>
                    <Name></Name>
                    <Domain></Domain>
                    <Regexp>
                        <Regex></Regex>
                        <RXGroupOrder></RXGroupOrder>
                    </Regexp>
                </ComputerName>
            </ComputerNameRules>
        </Rules>
        <SingleComputers>
            <SingleComputer>
                <Name></Name>
                <Domain></Domain>
                <UserName></UserName>
            </SingleComputer>
        </SingleComputers>
    </DynamicFolder> 
</RoyalTSConfiguration>
'@
    Try {
        Write-Verbose -Message "Creating the configuration file $($Path)"
        # Create the configuration file
        Set-Content -Path $Path -Value $Configuration -Force | Out-Null
    }
    Catch {
        Write-Error -Message "An unexpected error occured while creating the configuration file $($Path): $($_.Exception.Message)"
    }
}
Function New-RoyalTSDynamicFolder {
    [CmdletBinding()]
    Param()
    Begin {
        Write-Verbose -Message "Starting the function $($MyInvocation.MyCommand)"
        # Define the location of the configuration file
        $ConfigurationFile = Join-Path -Path $Env:APPDATA -ChildPath 'SPS-RoyalTS\RoyalTSConfiguration.xml'
        # Check if the configuration file exists
        if (-not (Test-Path -Path $ConfigurationFile)) {
            Write-Verbose -Message "The configuration file $($ConfigurationFile) does not exist"
            # Create the configuration file
            New-DefaultRoyalTSConfiguration -Path $ConfigurationFile | out-null
            Write-Verbose -Message "The configuration file$($ConfigurationFile) has been created"
        }
        # Read the configuration file
        [XML] $Configuration = Get-Content -Path $ConfigurationFile -Raw
        # Check if the configuration file has the right structure
        if (-not $Configuration.RoyalTSConfiguration) {
            Throw "The configuration file $($ConfigurationFile) does not have the right structure"
        }
        # Check if the configuration file call a remote file if yes load it instead
        if ($Configuration.RoyalTSConfiguration.RemoteFile) {
            Write-Verbose -Message "The configuration file $($ConfigurationFile) call a remote file ($($Configuration.RoyalTSConfiguration.RemoteFile))"
            # Get the remote file
            $RemoteFile = $Configuration.RoyalTSConfiguration.RemoteFile
            # Check if the remote file exists
            if (-not (Test-Path -Path $RemoteFile)) {
                Throw "The remote file [$($RemoteFile)] does not exist"
            }
            # Read the remote file
            [XML] $Configuration = Get-Content -Path $RemoteFile -Raw
            # Check if the remote file has the right structure
            if (-not $RemoteConfiguration.RoyalTSConfiguration) {
                Throw "The remote file [$($RemoteFile)] does not have the right structure"
            }
        }
    }
    Process {
        Write-Verbose -Message "Processing the function $($MyInvocation.MyCommand)"
        # Check if the configuration file has dynamic folders
        if ($Configuration.RoyalTSConfiguration.DynamicFolder) {
            # Get the dynamic folders rule
        }Else{
            Write-Warning -Message "The configuration file $($ConfigurationFile) does not have any dynamic folder the function $($MyInvocation.MyCommand) will not do anything"
        }
    }
    End {
        Write-Verbose -Message "Ending the function $($MyInvocation.MyCommand)"
    }
}