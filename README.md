# Start-PrinterMapping

## Description

This script is a framework to control and map the printers for
each user or computer. The printer mapping is based on 
ActiveDirectory groupmembership and is manly focused for 
terminal server environments.

## Example

```powershell
Start-PrinterMapping.ps1 -Mode User -PrintServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-" -DefaultPrinterGroupSuffix "_Default" -CollectionGroupPrefix "GRP-PRT-" -ClearPrinters -LeavePrefix

Start-PrinterMapping.ps1 -Mode Computer -PrinterServer "PrintServer01" -PrinterGroupPrefix "CTX-PRT-"

Start-PrinterMapping.ps1 -DefaultPrinterGUI
```

## Requirements

1. ActiveDirectory PowerShell module is requiered
2. Active Directory groups must be created for each printer with a certain prefix
