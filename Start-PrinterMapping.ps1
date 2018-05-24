######################################################################
##
## (C) 2018 Bechtle GmbH & Co. KG - Tim Bitzer
##
##
## Filename:      Start-PrinterMapping.ps1
##
## Version:       1.0
##
## Release:       Development
##
## Requirements:  - ActiveDirectory Module
##                - PowerShell Version 3
##
## Changes:       2018-04-27    Refactor / Initial creation
##
## 
## This script is provided 'AS-IS'. The author does not provide
## any guarnatee or warranty, stated or implied. Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##	            
##
######################################################################

#Requires -Module ActiveDirectory
#Requires -Version 3

<#
        .SYNOPSIS
        This script is a framework to control and map the printers for
        each user or computer. The printer mapping is based on 
        ActiveDirectory groupmembership and is manly focused for 
        terminal server environments.


        .DESCRIPTION
        For each printer you want to manage with this framework you
        have to create an AD group. The name of the group needs to 
        start with a certain prefix followed by the exact printer 
        name. I you want to determine the default printer aswell a
        second group is needed for the printer with a certain suffix.
        Based on the group membership of either the user or computer 
        (Mode) the printer will be connected. 

         
        .PARAMETER Mode
        A string determining whether the user or the computer is part 
        of the printer groups


        .PARAMETER PrintServer
        A string containing the name of the printserver.


        .PARAMETER PrinterGroupPrefix
        A string containing the prefix of the printer groups created
        for each printer


        .PARAMETER DefaultPrinterGroupSuffix
        A string containing the the suffix if the behavior of the 
        default printer should be managed by this framework. If a user
        or computer is member of multiple default printer groups the 
        last one wins since there can only be one default printer


        .PARAMETER CollectionGroupPrefix
        A string containing the collection group prefix. In addition 
        it´s possible to create a collection group for e.g. a 
        department and make these groups member of the printer groups. 
        

        .PARAMETER ClearPrinters
        A switch determining whether you want to remove all shared
        printers before connecting new ones.


        .PARAMETER LeavePrefix
        A switch determining the actual printer name infact contains
        the prefix.


        .PARAMETER LogFilePath
        A string containing the path where the log file should be
        located. The default value is $env:TEMP (%TEMP%).


        .PARAMETER OutConsole
        A switch parameter determining whether the output should be
        logged to console.


        .PARAMETER UseWMI
        A switch parameter whether WMI should be used instead of the
        'Get-Printer' cmdlet to maintain downward compatibility. Keep 
        in mind that using WMI is significantly slower.


        .PARAMETER DefaultPrinterGUI
        A switch which starts a small GUI with a list of all the 
        printers connected where the user can select the default
        printer himself.



        .EXAMPLE
        Start-PrinterMapping.ps1 -Mode User -PrintServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-" -DefaultPrinterGroupSuffix "_Default" -CollectionGroupPrefix "GRP-PRT-" -ClearPrinters -LeavePrefix

        
        .EXAMPLE
        Start-PrinterMapping.ps1 -Mode Computer -PrinterServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-"


        .EXAMPLE
        Start-PrinterMapping.ps1 -DefaultPrinterGUI
#>

[CmdletBinding()]
Param
(
    [parameter(Mandatory = $true, ParameterSetName = "Default")]
    [ValidateSet("User", "Computer")]
    [string]$Mode,
    [parameter(Mandatory = $true, ParameterSetName = "Default")]
    [string]$PrintServer,
    [parameter(Mandatory = $true, ParameterSetName = "Default")]
    [string]$PrinterGroupPrefix,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [string]$DefaultPrinterGroupSuffix,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [string]$CollectionGroupPrefix,    
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [switch]$ClearPrinters,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [switch]$LeavePrefix,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [String]$LogFilePath = $env:TEMP,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [switch]$OutConsole,
    [parameter(Mandatory = $false, ParameterSetName = "Default")]
    [switch]$UseWMI,
    [parameter(Mandatory = $false, ParameterSetName = "GUI")]
    [switch]$DefaultPrinterGUI
)

#### Scripting definitions
Set-PSDebug -Strict
Set-StrictMode -Version latest

#___________________________________________________________________
#region FUNCTION Start-PrinterMapping
function Start-PrinterMapping
{
    <#
            .SYNOPSIS
            This function handles the printer mapping based on the ActiveDirectory
            group membership of either a computer or user.


            .DESCRIPTION
            For each printer you want to manage with this framework you
            have to create an AD group. The name of the group needs to 
            start with a certain prefix followed by the exact printer 
            name. I you want to determine the default printer aswell a
            second group is needed for the printer with a certain suffix.
            Based on the group membership of either the user or computer 
            (Mode) the printer will be connected. 

            
            .PARAMETER Mode
            A string determining whether the user or the computer is part 
            of the printer groups


            .PARAMETER PrintServer
            A string containing the name of the printserver.


            .PARAMETER PrinterGroupPrefix
            A string containing the prefix of the printer groups created
            for each printer


            .PARAMETER DefaultPrinterGroupSuffix
            A string containing the the suffix if the behavior of the 
            default printer should be managed by this framework. If a user
            or computer is member of multiple default printer groups the 
            last one wins since there can only be one default printer


            .PARAMETER CollectionGroupPrefix
            A string containing the collection group prefix. In addition 
            it´s possible to create a collection group for e.g. a 
            department and make these groups member of the printer groups. 
            

            .PARAMETER ClearPrinters
            A switch determining whether you want to remove all shared
            printers before connecting new ones.


            .PARAMETER LeavePrefix
            A switch determining the actual printer name infact contains
            the prefix.



            .EXAMPLE
            Start-PrinterMapping.ps1 -Mode User -PrintServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-" -DefaultPrinterGroupSuffix "_Default" -CollectionGroupPrefix "GRP-PRT-" -ClearPrinters -LeavePrefix

            
            .EXAMPLE
            Start-PrinterMapping.ps1 -Mode Computer -PrinterServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-"


            .EXAMPLE
            Start-PrinterMapping.ps1 -DefaultPrinterGUI
    #>
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory = $true)]
        [ValidateSet("User", "Computer")]
        [string]$Mode,
        [parameter(Mandatory = $true)]
        [string]$PrintServer,
        [parameter(Mandatory = $true)]
        [string]$PrinterGroupPrefix,
        [parameter(Mandatory = $false)]
        [string]$DefaultPrinterGroupSuffix,
        [parameter(Mandatory = $false)]
        [string]$CollectionGroupPrefix,
        [parameter(Mandatory = $false)]
        [switch]$ClearPrinters,
        [parameter(Mandatory = $false)]
        [switch]$LeavePrefix
    )

    #### Variable definition
    $PrinterGroups = @()
    $CurrentDefaultPrinter = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Default -eq "True"}
    $NewDefaultPrinter = $null


    ### Load & test all requierments
    Start-PreFlightCheck
    

    #### Delete all shared (network) printers
    if ($ClearPrinters)
    {
        Remove-ExistingNetworkPrinters
    }
    
    #### Check if user or computer
    if ($Mode -eq "User")
    {
        try
        {
            ### Get all ActiveDirectory groups of the user
            $Groups = (Get-ADPrincipalGroupMembership $env:USERNAME).Name
        }
        catch
        {
            New-LogMessage -Message "Could not retrieve user groups" -LogLevel ERROR
            New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG 

            Exit
        }
        
    }        
   
    elseif ($ComputerOrUser -eq "Computer")
    {
        ### Retrieve local client name
        $LocalClientName = Get-ClientName

        ### Check if name could be retrieved
        if ($LocalClientName)
        {
            try 
            {
                ### Get all ActiveDirectory groups of the loca client 
                $Groups = (Get-ADPrincipalGroupMembership (Get-ADComputer $LocalClientName)).Name
            }
            catch 
            {
                New-LogMessage -Message "Could not retrieve computer groups" -LogLevel ERROR
                New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG 

                Exit
            }    
        }

        else
        {
            New-LogMessage -Message "Could not retrieve local client name" -LogLevel ERROR

            Exit
        }
    }

    ### Loop each ActiveDirectory group found for user/computer
    foreach ($Group in $Groups)
    {
        ### Check if $CollectionGroupPrefix is set for user/computer collection groups and check if collection group start with $CollectionGroupPrefix
        if ($CollectionGroupPrefix -and $Group.StartsWith($CollectionGroupPrefix, "CurrentCultureIgnoreCase"))
        {
            try
            {
                ### Retrieve group membership of user/computer collection group
                $NestedGroups = (Get-ADPrincipalGroupMembership (Get-ADGroup $Group)).Name
            }

            catch
            {
                New-LogMessage -Message "Could not retrieve user collection groups" -LogLevel ERROR
                New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG 

                Exit
            }
            
            ### Loop each collection group (nested) found for user/computer
            foreach ($NestedGroup in $NestedGroups)
            {
                ### Check if any nested group starts with $PrinterGroupPrefix
                if ($NestedGroup.StartsWith($PrinterGroupPrefix, "CurrentCultureIgnoreCase"))
                {
                    ### Add group to $PrinterGroup list
                    $PrinterGroups += $NestedGroup
                }
            }
        }

        ### Check if any group of the user/computer starts with $PrinterGroupPrefix
        elseif ($Group.StartsWith($PrinterGroupPrefix, "CurrentCultureIgnoreCase"))
        {
            ### Add group to $PrinterGroup list
            $PrinterGroups += $Group
        }
    }


    New-LogMessage -Message "Found $($PrinterGroups.count) relevant groups containing the prefix '$PrinterGroupPrefix'" -LogLevel INFO


    ### Loop each $PrinterGroup found for user/computer
    foreach ($PrinterGroup in $PrinterGroups)
    {
        ### Check if $DefaultPrinterGroupSuffix is set since its an optional parameter
        if ($DefaultPrinterGroupSuffix)
        {
            ### Check if any of the found $PrinterGroups has $DefaultPrinterGroupSuffix at its end
            if ($PrinterGroup.Substring(($PrinterGroup.length - $DefaultPrinterGroupSuffix.Length), $DefaultPrinterGroupSuffix.Length) -eq $DefaultPrinterGroupSuffix)
            {
                ### Chop of the $DefaultPrinterGroupSuffix at the end so its a valid printer name 
                $PrinterGroup = $PrinterGroup.Substring(0, ($PrinterGroup.length - $DefaultPrinterGroupSuffix.Length))

                $NewDefaultPrinter = $PrinterGroup
            }
        }
        
        ### Check if $LeavePrefix is set 
        if ($LeavePrefix)
        {
            ### Printer names actually contain the $PrinterGroupPrefix ($PrinterGroup name = actual printer name)
            $Printer = $PrinterGroup
        }

        else 
        {
            ### Chop of the $PrinterGroupPrefix to retrieve the actual printer name
            $Printer = $PrinterGroup.Substring($PrinterGroupPrefix.Length, ($PrinterGroup.Length - $PrinterGroupPrefix.Length))

            ### if $NewDefaultPrinter exists chop of the $PrinterGroupPrefix for the $NewDefaultPrinter aswell
            if ($NewDefaultPrinter)
            {
                $NewDefaultPrinter = $NewDefaultPrinter.Substring($PrinterGroupPrefix.Length, ($PrinterGroup.Length - $PrinterGroupPrefix.Length))
            }
        }

        ### Connect the printer
        Add-NetworkPrinter -Printer $Printer -PrintServer $PrintServer
    }


    ### Check if the $NewDefaultPrinter is not $null and therefor the default printer is defined by AD group membership 
    if ($NewDefaultPrinter)
    {
        ### Build UNC path for $NewDefaultPrinter
        $NewDefaultPrinterUNCPath = "\\{0}\{1}" -f $PrintServer, $NewDefaultPrinter

        ### Use WMI anyway since the method '.SetDefaultPrinter()' only wokrs with WMI
        $InstalledPrinters = Get-Printer
        

        ### Check if $NewDefaultPrinterUNCPath printer is already connected
        if ($InstalledPrinters.Name -contains $NewDefaultPrinterUNCPath)
        {
            ### Get the $NewDefaultPrinterUNCPath printer as WMI object
            $DefaultPrinter = $InstalledPrinters | Where-Object {$_.Name -eq $NewDefaultPrinterUNCPath}

            try
            {
                ### Set the default printer
                $DefaultPrinter.SetDefaultPrinter() | Out-Null
                New-LogMessage -Message "Default printer has been set to '$NewDefaultPrinterUNCPath' (AD group definition)" -LogLevel INFO
            }
            catch
            {
                New-LogMessage -Message "Could not set the default printer to '$NewDefaultPrinterUNCPath' (AD group definition)" -LogLevel ERROR
                New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG
            }            
        }
    }

    else
    {
        #### Check if the current default printer is a shared printer
        if ($CurrentDefaultPrinter.Shared -eq $true)
        {
            ### Check if 'UseWMI' is set and get all connected printers wither with WMI or PSCmdlet
            if ($Script:UseWMI)
            {
                $InstalledPrinters = Get-WmiObject -Class Win32_Printer
            }

            else
            {
                $InstalledPrinters = Get-Printer
            }

            ### Check if $DefaultPrinter is already connected
            if ($InstalledPrinters.Name -contains $CurrentDefaultPrinter.Name)
            {
                try
                {
                    ### Reset the 'old' default printer
                    $CurrentDefaultPrinter.SetDefaultPrinter() | Out-Null 
                    New-LogMessage -Message "Default printer has been reset to '$($CurrentDefaultPrinter.Name)'" -LogLevel INFO
                }
                catch
                {
                    New-LogMessage -Message "Could not reset the default printer to '$($CurrentDefaultPrinter.Name)'" -LogLevel ERROR
                    New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG
                }                
            }

            else
            {
                New-LogMessage -Message "The previous default printer $($CurrentDefaultPrinter.Name) could not been reset because it is not connected anymore" -LogLevel WARN
            }
        }

        else
        {
            New-LogMessage -Message "Default printer is not a shared printer and therefor was not modified" -LogLevel INFO
        }
    }
}
#endregion
#___________________________________________________________________
#region FUNCTION New-LogMessage
function New-LogMessage
{
    <#
        .SYNOPSIS
        This is function handles the logging


        .DESCRIPTION
        Whith the New-LogMessage function you can write log
        messages to a log file and/or write it to the console

         
        .PARAMETER Message
        A string containing the log message


        .PARAMETER LogLevel
        A string containing log level. Possible values are:
        "INFO", "WARN", "ERROR", "DEBUG"


        .PARAMETER Start
        A switch to place the log file header


        .PARAMETER End
        A switch to place the log file footer



        .EXAMPLE
        New-LogMessage -Message "This is my message" -LogLevel INFO


        .EXAMPLE
        New-LogMessage -Start


        .EXAMPLE
        New-LogMessage -End
#>
    Param
    (
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [String]$Message,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [String]$LogLevel,
        [Parameter(Mandatory = $true, ParameterSetName = 'Start')]
        [Switch]$Start,
        [Parameter(Mandatory = $true, ParameterSetName = 'Start')]
        [String]$GivenParameters,
        [Parameter(Mandatory = $true, ParameterSetName = 'End')]
        [Switch]$End,
        [Parameter(Mandatory = $true, ParameterSetName = 'End')]
        [String]$ExecutionTime
    )

    ### Logfile header
    if ($Start)
    {
        $LogMessage = @" 
#####################################################################################################################################

    Started 'Start-PrinterMapping.ps1' - {0}

    Executed on machine: {1}

    Following parameters were used:
    {2}

_____________________________________________________________________________________________________________________________________
"@ -f (Get-Date), $env:COMPUTERNAME, $GivenParameters
    }

    #### Logfile footer
    elseif ($End)
    {
        $LogMessage = @" 
_____________________________________________________________________________________________________________________________________

    Finished 'Start-PrinterMapping.ps1'

    Script took {0} seconds to execute.

_____________________________________________________________________________________________________________________________________








"@ -f $ExecutionTime
    }

    #### Normal log message
    else
    {
        ### Variable definitions
        $TimeStamp = (Get-Date -f "dd.MM.yyyy HH:mm:ss.fff").ToString()
        $Caller = (Get-PSCallStack)[1].Command


        ### Column size of each entry in LogMessage
        [int]$TimeStampColumnSize = 24
        [int]$CallerColumnSize = 30
        [int]$LevelColumnSize = 5
        [int]$MessageColumnSize = 150
	
	
        ### Shrink $Caller to $CallerColumnSize if needed
        if ($Caller.Length -gt $CallerColumnSize)
        { 
            $Caller = $Caller.Substring(0, $CallerColumnSize) 
        }
	
	
        ### Shrink $Message to $MessageColumnSize if needed
        if ($Message.Length -gt $MessageColumnSize)
        { 
            $Message = $Message.Substring(0, $MessageColumnSize)         
        }

        else
        {
            $MessageColumnSize = $Message.Length
        }
	
	
        ### Fix format of LogMessage
        [string]$LogMessage = [string]::Format("{0,-" + $TimeStampColumnSize + "}`t{1,-" + $CallerColumnSize + "}`t{2,-" + $LevelColumnSize + "}`t{3,-" + $MessageColumnSize + "}", $TimeStamp, $Caller, $LogLevel, $Message)
    }


    ### Out LogFile
    $LogFileName = "PrinterMapping_{0}.txt" -f $env:USERNAME
    $LogFileAbsolutePath = (Join-Path -Path $Script:LogFilePath -ChildPath $LogFileName)
    $LogMessage | Out-File -FilePath $LogFileAbsolutePath -Append 


    ### Out console
    if ($Script:OutConsole)
    {
        switch ($LogLevel)
        {
            "INFO"
            {
                $ForeGroundColor = "White"
            }

            "WARN"
            {
                $ForeGroundColor = "Yellow"   
            }

            "ERROR"
            {
                $ForeGroundColor = "Red"   
            }

            default
            {
                $ForeGroundColor = "White"    
            }
        }

        Write-Host $LogMessage -ForegroundColor $ForeGroundColor
       
    }
}
#endregion
#___________________________________________________________________
#region FUNCTION Add-NetworkPrinter
function Add-NetworkPrinter
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [String]$Printer,
        [Parameter(Mandatory = $true)]
        [String]$PrintServer
    )

    ### Build UNC path printer name
    $UNCPrinterPath = "\\{0}\{1}" -f $PrintServer, $Printer

    ### Check if 'UseWMI' is set and get all connected printers wither with WMI or PSCmdlet
    if ($Script:UseWMI)
    {
        $InstalledPrinters = Get-WmiObject -Class Win32_Printer
    }

    else
    {
        $InstalledPrinters = Get-Printer
    }
    
    ### Check if printer already exists
    if ($InstalledPrinters.Name -contains $UNCPrinterPath)
    {
        New-LogMessage -Message "Printer '$UNCPrinterPath' already exists" -LogLevel INFO
    }

    else
    {
        try
        {
            ### Install/connect printer
            $WNet = New-Object -ComObject WScript.Network
            $WNet.AddWindowsPrinterConnection($UNCPrinterPath)

            New-LogMessage -Message "Printer '$UNCPrinterPath' successfully added" -LogLevel INFO
        }
        catch
        {
            New-LogMessage -Message "Could not add printer '$UNCPrinterPath'" -LogLevel ERROR
            New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG
        }
    }
}
#endregion
#___________________________________________________________________
#region FUNCTION Get-ClientName
function Get-ClientName 
{
    ### Check if client or server OS
    if ((Get-WmiObject Win32_OperatingSystem).Name -match "Windows Server") 
    {
        New-LogMessage -Message "Running on server OS" -LogLevel INFO

        ### Query user session with query.exe and create an object out of it
        $Session = c:\windows\system32\query.exe session $env:USERNAME | ForEach-Object {$_.Trim() -replace '\s+', ','} | ConvertFrom-Csv

        ### Build registry path
        $RegKey = 'HKCU:\Volatile Environment\{0}' -f $Session.ID
        
        try 
        {
            ### Retrieve session information
            $SessionInfo = Get-ItemProperty -Path $RegKey
            New-LogMessage -Message "Successfully retrieved session information" -LogLevel INFO
        }
        catch 
        {
            New-LogMessage -Message "Could not retrieve session information" -LogLevel ERROR
            New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG
        }
        
        ### Check if current session is not a console session   
        if ($SessionInfo.SESSIONNAME -ne 'Console')
        {
            $ClientName = $SessionInfo.CLIENTNAME
            New-LogMessage -Message "Successfully retrieved local clientname. Client name is '$ClientName'" -LogLevel INFO
        }
        else
        {
            New-LogMessage -Message "Current session is a console session" -LogLevel WARN
            $ClientName = $null
        }
    }

    ### Works only for Citrix VDIs
    else
    {
        New-LogMessage -Message "Running on client OS" -LogLevel INFO
        
        ### Regkey 
        $RegKey = "HKLM:\SOFTWARE\Citrix\Ica\Session"

        ### Check if reg key exists
        if (Test-Path -Path $RegKey)
        {
            ### Get client name
            $ClientName = (Get-ItemProperty -Path $RegKey).ClientName
            New-LogMessage -Message "Successfully retrieved local clientname. Client name is '$ClientName'" -LogLevel INFO
        }

        else
        {
            New-LogMessage -Message "Client name cannot be retrieved. This is probably not a Citrix VDI." -LogLevel WARN
            $ClientName = $null
        }
    }

    return $ClientName
}
#endregion
#___________________________________________________________________
#region FUNCTION Start-PreFlightCheck
function Start-PreFlightCheck
{
    try 
    {
        ### Load ActiveDirectory module
        Import-Module -Name "ActiveDirectory" -ErrorAction Stop
        New-LogMessage -Message "Successfully loaded ActiveDirectory module" -LogLevel INFO
    }

    catch 
    {
        New-LogMessage -Message "Could not load ActiveDirectory module. RSAT needs to be installed" -LogLevel ERROR
        New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG 

        Exit
    }
}
#endregion
#___________________________________________________________________
#region FUNCTION Remove-ExistingNetworkPrinters 
function Remove-ExistingNetworkPrinters 
{
    ### Check if 'UseWMI' is set and get all connected printers wither with WMI or PSCmdlet
    if ($Script:UseWMI)
    {
        $InstalledPrinters = Get-WmiObject -Class Win32_Printer
    }

    else
    {
        $InstalledPrinters = Get-Printer
    }

    ### Create Wscript object to remove printers
    $WNet = New-Object -ComObject WScript.Network

    foreach ($Printer in $Printers)
    {
        ### Check if printer is a network (shared) printer
        if ($Printer.Shared -eq $true)
        {    
            try
            {
                ### Remove printer
                $WNet.RemovePrinterConnection($Printer.Name)
                New-LogMessage -Message "Printer '$($Printer.Name)' has been removed" -LogLevel INFO
            }    
            catch
            {
                New-LogMessage -Message "Could not remove printer '$($Printer.Name)'" -LogLevel ERROR
                New-LogMessage -Message $_.Exception.Message -LogLevel DEBUG
            }                
        }        
    } 


    #TODO Clean all registry values aswell?!
    
}
#endregion
#___________________________________________________________________
#region FUNCTION Start-DefaultPrinterGUI
#TODO Cleanup and notaion
function Start-DefaultPrinterGUI
{
    Add-Type -AssemblyName PresentationFramework

    $inputXAMLDefaultPrinter = @"
    <Window x:Class="DeleteScheduledJobs.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DeleteScheduledJobs"
        mc:Ignorable="d"
        Title="Default Printer" Height="296.997" Width="350">
        <Grid>
            <ListBox x:Name="listBox" Height="198" Margin="10,10,12,0" VerticalAlignment="Top"/>
            <Button x:Name="button" Content="Select Default Printer" Margin="110,230,102,0" Height="20" VerticalAlignment="Top"/>
        </Grid>
    </Window>
"@
    $inputXAMLDefaultPrinter = $inputXAMLDefaultPrinter -replace 'mc:Ignorable="d"', ''
    [XML]$XAMLDefaultPrinter = $inputXAMLDefaultPrinter
    
    $XAMLDefaultPrinter.Window.RemoveAttribute("x:Class") 
    $ReaderDefaultPrinter = New-Object System.Xml.XmlNodeReader $XAMLDefaultPrinter

    $FormDefaultPrinter = [Windows.Markup.XamlReader]::Load($ReaderDefaultPrinter)
    $FormDefaultPrinter.Topmost = $true


    $LBDefaultPrinter = $FormDefaultPrinter.FindName('listBox')
    $LBDefaultPrinter.SelectionMode = "Single"
    Get-InstalledPrinters

    $BTNSetDefaultPrinter = $FormDefaultPrinter.FindName('button')
    $BTNSetDefaultPrinter.Add_Click( {Invoke-BTNSetDefaultPrinter_Click})


    $FormDefaultPrinter.ShowDialog() | Out-Null
}

Function Get-InstalledPrinters
{
    $Printers = Get-Printer

    $LBDefaultPrinter.Items.Clear()
    if ($Printers -eq $Null)
    {
        $LBDefaultPrinter.Items.Add("No printers found!") | Out-Null
    }
    
    else
    {
        foreach ($Printer in $Printers)
        {
            $LBDefaultPrinter.Items.Add($Printer.Name) | Out-Null
        }
    }
}

Function Invoke-BTNSetDefaultPrinter_Click
{
    $Printers = Get-Printer
    foreach ($Printer in $Printers)
    {
        if ($Printer.Name -eq $LBDefaultPrinter.SelectedItem)
        {
            $WSNet = New-Object -ComObject WScript.Network
            $WSNet.SetDefaultPrinter($Printer.Name)
        }
    }

    $FormDefaultPrinter.Close()
}
#endregion
#___________________________________________________________________
#region MAIN


### Variable definitions
if ($OutConsole)
{
    $Script:OutConsole = $true
}
else
{
    $Script:OutConsole = $false
}

$Script:LogFilePath = $LogFilePath
$Script:UseWMI = $UseWMI
$GivenParameters = ""
$ExcludedParameters = ""
$ExcludeParameters = @("LogFilePath", "OutConsole")
$StartTime = Get-Date


if ($DefaultPrinterGUI)
{
    ### Start GUI
    Start-DefaultPrinterGUI
}

### Dynamic parameter parsing
else
{    
    foreach ($Parameter in $PSBoundParameters.GetEnumerator())
    {
        ### if switch parameter, $Parameter.Value is not necessarily needed
        switch ($Parameter.Value.GetType())
        {
            "Switch"
            {
                $CurrentParameter = ("-{0} " -f $Parameter.Key)
            }
            
            "String"
            {
                $CurrentParameter = ("-{0} ""{1}"" " -f $Parameter.Key, $Parameter.Value)
            }

            Default
            {
                $CurrentParameter = ("-{0} {1} " -f $Parameter.Key, $Parameter.Value)
            }
        }

        if ($ExcludeParameters -contains $Parameter.Key)
        {
            $ExcludedParameters += $CurrentParameter
        }

        else
        {
            $GivenParameters += $CurrentParameter
        }
    }

    New-LogMessage -Start -GivenParameters ($GivenParameters + $ExcludedParameters)

    ### Execute main function with dynamic parsed parameter
    "Start-PrinterMapping $GivenParameters" | Invoke-Expression

    New-LogMessage -End -ExecutionTime ([Math]::Round((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds, 2)).ToString()
}
#endregion
#___________________________________________________________________

