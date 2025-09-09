# Get-Group-Policy-Processing-Time-V2
Gets the average, minimum, and maximum Group Policy processing time on computers in Microsoft Active Directory.

By default, it builds a list of all computers where "Server" is in the OperatingSystem property unless the ComputerName or InputFile parameter is used.

The script must be run from an elevated PowerShell session.
	
Process each server, looking in the Microsoft-Windows-GroupPolicy/Operational event log for all Event ID 8001 events.
	
Displays the Average, Minimum, and Maximum processing times.
	
All events with a processing time of zero are ignored. A 0 time means a local account was used for login.
	
Display the results on the console and create two text files, by default, in the folder where the script is run.
	
Optionally, can specify the output folder.
	
Unless the InputFile parameter is used, it needs the ActiveDirectory module.
	
The script has been tested with PowerShell versions 2, 3, 4, 5, and 5.1.
The script has been tested with Microsoft Windows Server 2008 R2, 2012, 2012 R2, 2016, and Windows 10 Creators Update.
	
There is a bug with Get-WinEvent and PowerShell versions later than 2, or a culture other than en-US, where the Message property is not returned.

There are two workarounds:
	1. PowerShell.exe -Version 2
	2. Add this line to the script: 
	[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
