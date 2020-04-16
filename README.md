# Get-OutlookDump

PowerShell script for dumping local Outlook inbox and searching for goodies.

## Installation

Simply download the script to the host and run it. Notice the dot (".") at the beginning of the command.

```
.'C:\<FAKEPATH>\Get-OutlookDump.ps1'
```

Get help by running:

```
 Get-Help Get-OutlookDump -full
```

## Example Usage

```
 PS C:>Get-OutlookDump -Search Password
 
   My secret password is: 12345
   -------------------------------------------------
   Found in mail number: 18


 PS C:>Get-OutlookDump -Index 18
 
   <Outputs raw email message>

```
