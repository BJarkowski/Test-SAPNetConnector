# Test-SAPNetConnector
PowerShell script for testing SAP Connector for Microsoft .NET

## Prerequisites
This script is intended to use only in PowerShell 5. <br>
You have to install GAC module prior to using this script. <br>
`Install-Module -Name Gac`

## Running the test
Run following command to run the test:

```
.\Test-NetConnectorSAP.ps1 `
    -Hostname <hostname> `
    -SystemNumber <system_number> `
    -Client <client_number> `
    -Username <username> `
    -Password <password> 
```
Where:
- hostname - hostname or IP address of the SAP system
- system_number - system number of the SAP instance
- client - client number 
- username - user name 
- password - password

For example:

```
.\Test-NetConnectorSAP.ps1 `
    -Hostname my.sap.com `
    -SystemNumber 00 `
    -Client 100 `
    -Username USER01 
```
If you don't provide a parameter value in the command line, the script will ask you for missing input during runtime. This is especially useful if you don't want to provide password in the command line.

## Scope of the test
The test scripts uses GAC to module to list available SAP Connector for .NET libraries. It automatically checks the architecture (32-bit vs 64-bit) based on the PowerShell session. In case there is more than one version of the libraries installed the script will ask you to choose one. <br>
Warning! Once libraries are loaded to PowerShell session there is no way to unload them. If you want to run the script against a different libraries ensure you start a new PowerShell session.

## Sample output
```
You can use this script to test connectivity to the SAP system using .NET Connector for SAP

Checking PowerShell version
    PowerShell version 5 detected

Checking if the GAC module is installed
    The Gac module is installed

Checking if SAP .NET Connector is registered in GAC
    Found following libraries:

Name         Version  FullName                                                                         ProcessorArchitecture
----         -------  --------                                                                         ---------------------
sapnco       3.0.0.42 sapnco, Version=3.0.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23                       Amd64
sapnco_utils 3.0.0.42 sapnco_utils, Version=3.0.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23                 Amd64
sapnco       3.1.0.42 sapnco, Version=3.1.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23                         X86
sapnco_utils 3.1.0.42 sapnco_utils, Version=3.1.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23                   X86


    PowerShell is running in 64-bit mode
    Excluding incompatible assemblies
    Found SAP .NET Connector libraries in GAC

Importing SAP .NET Connector using:
    sapnco assembly: sapnco, Version=3.0.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23
    sapnco_utils assembly: sapnco_utils, Version=3.0.0.42, Culture=neutral, PublicKeyToken=50436dca5c7f7d23

    Assembly sapnco loaded successfully.
    Assembly sapnco_utils loaded successfully.

SAP .NET Connector verion: 3.0.25.0
    Patch Level: 1210
    SAP Release: 722

Connecting to SAP system
    Connection to SAP system succesfull

Calling SAP function module
    SAP Host: S41909-A
    SAP System ID: A4H
    SAP Database: HDB
    SAP Database Host: nwhana
    Calling SAP function module succesfull
```