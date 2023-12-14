<#
  .SYNOPSIS
  PowerShell script for testing SAP Connector for Microsoft .NET

  .DESCRIPTION
  You can use the script to validate installation of the SAP Connector for Microsoft .NET.

  .PARAMETER Hostname
  Hostname of the SAP system

  .PARAMETER SystemNumber
  System number of the SAP instance

  .PARAMETER Client
  Client number

  .PARAMETER Username
  SAP username for making the connection

  .PARAMETER Password
  Password for the SAP user

  .INPUTS
  None. 

  .OUTPUTS
  None. 

  .EXAMPLE
  PS> .\Update-Month.ps1

  .EXAMPLE
  PS> .\Test-NetConnectorSAP.ps1 `
        -Hostname my.sap.com `
        -SystemNumber 00 `
        -Client 100 `
        -Username USER01 
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Hostname,
    [Parameter(Mandatory=$true)]
    [string]$SystemNumber,
    [Parameter(Mandatory=$true)]
    [string]$Client,
    [Parameter(Mandatory=$true)]
    [string]$Username,
    [Parameter(Mandatory=$true)]
    [securestring]$Password
)

Write-Host
Write-Host "You can use this script to test connectivity to the SAP system using .NET Connector for SAP"

Write-Host
Write-Host "Checking PowerShell version"
if ($PSVersionTable.PSVersion.Major -eq 7) {
    Write-Host "    Use PowerShell 5 to use this script. Using PowerShell 7 is not supported!" -ForegroundColor Yellow
    break
}
elseif ($PSVersionTable.PSVersion.Major -eq 5) {
    Write-Host "    PowerShell version 5 detected" -ForegroundColor Green
}
else {
    Write-Host "    Use PowerShell 5 to use this script" -ForegroundColor Yellow
}

Write-Host
Write-Host "Checking if the GAC module is installed"
$gacModule = Get-Module -ListAvailable -Name Gac
if ($gacModule) {
    Write-Host "    The Gac module is installed" -ForegroundColor Green
}
else {
    Write-Host "    The Gac module is not installed" -ForegroundColor Yellow
    Write-Host "    Install the module using following command:"
    Write-Host "    Install-Module -Name Gac"
    break
}
Write-Host
Write-Host "Checking if SAP .NET Connector is registered in GAC"

# Get all assemblies that match 'sapnco*'
$assemblies = Get-GacAssembly -Name 'sapnco*' | Select Name, Version, FullName, ProcessorArchitecture
Write-Host "    Found following libraries:"
$assemblies | Format-Table 

# Check the PS Architecture
if([System.Environment]::Is64BitProcess) {
    Write-Host "    PowerShell is running in 64-bit mode"
    Write-Host "    Excluding incompatible assemblies"
    $assemblies = $assemblies | Where-Object { $_.ProcessorArchitecture -eq 'Amd64' }
}
else {
    Write-Host "    PowerShell is running in 32-bit mode"
    Write-Host "    Excluding incompatible assemblies"
    $assemblies = $assemblies | Where-Object { $_.ProcessorArchitecture -eq 'X86' }
}

# Check the number of unique versions
$uniqueVersions = $assemblies | Select -Unique Version

if (@($uniqueVersions).Count -eq 0) {
    Write-Host "    No SAP .NET Connector libraries could be found" -ForegroundColor Yellow
    break
}
elseif (@($uniqueVersions).Count -eq 1) {
    $sapncoAssembly = $assemblies | Where-Object { $_.Name -eq 'sapnco' } | Select -ExpandProperty FullName
    $sapnco_utilsAssembly = $assemblies | Where-Object { $_.Name -eq 'sapnco_utils' } | Select -ExpandProperty FullName
    Write-Host "    Found SAP .NET Connector libraries in GAC" -ForegroundColor Green
}
else {
    Write-Host "    Warning: Multiple versions of the assemblies are installed." -ForegroundColor Yellow
    
    # Create a list of unique versions with index
    $index = 1
    $versionList = @{}
    $uniqueVersions | ForEach-Object { 
        Write-Host "    ${index}: Version $($_.Version)"
        $versionList.Add($index, $_.Version)
        $index++
    }
    Write-Host
    # Prompt user to choose a version
    $choice = Read-Host "   Please choose a version number (default is 1)"
    if ([string]::IsNullOrEmpty($choice)) {
        $choice = 1
    } else {
        try {
            $choice = [int]$choice
        } catch {
            Write-Host "    Invalid input. Defaulting to version 1."
            $choice = 1
        }
    }

    $selectedVersion = $versionList[$choice]

    # Assign the assemblies based on selected version
    $sapncoAssembly = $assemblies | Where-Object { $_.Name -eq 'sapnco' -and $_.Version -eq $selectedVersion } | Select -ExpandProperty FullName
    $sapnco_utilsAssembly = $assemblies | Where-Object { $_.Name -eq 'sapnco_utils' -and $_.Version -eq $selectedVersion } | Select -ExpandProperty FullName
}

Write-Host
Write-Host "Importing SAP .NET Connector using:"
Write-Host "    sapnco assembly: $sapncoAssembly"
Write-Host "    sapnco_utils assembly: $sapnco_utilsAssembly"
Write-Host 
# Load the SAP .NET Connector assembly
# Ensure the location of the sapnco and sapnco_utils are correct. They are located in the C:\Windows\Microsoft.NET\assembly\GAC_64\ directory
# Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_64\sapnco\v4.0_3.0.0.42__50436dca5c7f7d23\sapnco.dll'
# Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_64\sapnco_utils\v4.0_3.0.0.42__50436dca5c7f7d23\sapnco_utils.dll'

Add-Type -AssemblyName $sapncoAssembly
Add-Type -AssemblyName $sapnco_utilsAssembly

# Check if the sapnco assembly was loaded
$assembly = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like $sapncoAssembly }
if ($assembly -eq $null) {
    Write-Host "Failed to load assembly sapnco. Please check if the libraris still exists" -ForegroundColor Yellow
    break
}
else {
    Write-Host "    Assembly sapnco loaded successfully." -ForegroundColor Green
}

# Check if the assembly was loaded
$assembly = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like $sapnco_utilsAssembly }
if ($assembly -eq $null) {
    Write-Host "Failed to load assembly sapnco_utils. Please check if the libraris still exists" -ForegroundColor Yellow
    break
}
else {
    Write-Host "    Assembly sapnco_utils loaded successfully." -ForegroundColor Green
}

Write-Host

# Checking sapnco version
$Version = [SAP.Middleware.Connector.SAPConnectorInfo]::get_Version()
$PatchLevel = [SAP.Middleware.Connector.SAPConnectorInfo]::get_KernelPatchLevel()
$SAPRelease = [SAP.Middleware.Connector.SAPConnectorInfo]::get_SAPRelease()

Write-Host "SAP .NET Connector verion:" $Version
Write-Host "    Patch Level:" $PatchLevel
Write-Host "    SAP Release:" $SAPRelease

Write-Host

# Set up your SAP system parameters
$parameters = New-Object SAP.Middleware.Connector.RfcConfigParameters
$parameters.Add("NAME", "Connection Test")
$parameters.Add("USER", $Username)
$parameters.Add("PASSWD", [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)))
$parameters.Add("CLIENT", $Client)
$parameters.Add("LANG", "EN")
$parameters.Add("ASHOST", $Hostname)
$parameters.Add("SYSNR", $SystemNumber)

Write-Host "Connecting to SAP system"

Try {
    $destination = [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($parameters)
    $ping = $destination.Ping()
}
Catch {
    Write-Host "    Couldn't connect to the SAP system" -ForegroundColor Yellow
    Write-Host "Exception" $_.Exception.Message "occured"
    Write-Host $_.Exception
    break
}

Write-Host "    Connection to SAP system succesfull" -ForegroundColor Green
Write-Host
# Call funtion module
Try {
    # Establish the RFC connection
        Write-Host "Calling SAP function module"
        [SAP.Middleware.Connector.IRfcFunction]$rfcFunction = $destination.Repository.CreateFunction("RFC_SYSTEM_INFO")
        $rfcFunction.Invoke($destination)
    
        [SAP.Middleware.Connector.IRfcStructure]$Export =
        $rfcFunction.GetStructure("RFCSI_EXPORT")
    
        #-Get information---------------------------------------------
        Write-Host "    SAP Host:" $Export.GetValue("RFCHOST")
        Write-Host "    SAP System ID:" $Export.GetValue("RFCSYSID")
        Write-Host "    SAP Database:" $Export.GetValue("RFCDBSYS")
        Write-Host "    SAP Database Host:" $Export.GetValue("RFCDBHOST")
        Write-Host "    Calling SAP function module succesfull" -ForegroundColor Green
    }
Catch {
    Write-Host "Exception" $_.Exception.Message "occured"
    Write-Host $_.Exception
}