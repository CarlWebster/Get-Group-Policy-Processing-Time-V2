<#
.SYNOPSIS
	Gets the average, minimum and maximum Group Policy processing time on 
	computers in Microsoft Active Directory.
.DESCRIPTION
	By default, builds a list of all computers where "Server" is in the 
	OperatingSystem property unless the ComputerName or InputFile 
	parameter is used.

	The script must be run from an elevated PowerShell session.
	
	Process each server looking in the Microsoft-Windows-GroupPolicy/Operational 
	for all Event ID 8001.
	
	Displays the Avergage, Minimum and Maximum processing times.
	
	All events where processing time is 0 are ignored. A 0 time means a local account was used for login.
	
	Display the results on the console and creates two text files, by default, 
	in the folder where the script is run.
	
	Optionally, can specify the output folder.
	
	Unless the InputFile parameter is used, needs the ActiveDirectory module.
	
	The script has been tested with PowerShell versions 2, 3, 4, 5, and 5.1.
	The script has been tested with Microsoft Windows Server 2008 R2, 2012, 
	2012 R2, and 2016 and Windows 10 Creators Update.
	
	There is a bug with Get-WinEvent and PowerShell versions later than 2 or culture other than en-US,
	the Message property is not returned.

	There are two work-arounds:
	1. PowerShell.exe -Version 2
	2. Add this line to the script: 
	[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
	
.PARAMETER MaxSeconds
	Specifies the number of seconds to use for the cutoff for GPO processing time.
	Any value greater than or equal to MaxSeconds is recorded along with the user name and server name.
	Default is 30.
.PARAMETER ComputerName
	Computer name used to restrict the computer search.
	Script surrounds ComputerName with "*".
	
	For example, if "RDS" is entered, the script uses "*RDS*".
	
	This allows the script to reduce the number of servers searched.
	
	If both ComputerName and InputFile are used, ComputerName is used to filter
	the list of computer names read from InputFile.
	
	Alias is CN
.PARAMETER InputFile
	Specifies the optional input file containing computer account names to search.
	
	Computer account names can be either the NetBIOS or Fully Qualified Domain Name.
	
	ServerName and ServerName.domain.tld are both valid.
	
	If both ComputerName and InputFile are used, ComputerName is used to filter
	the list of computer names read from InputFile.
	
	The computer names contained in the input file are not validated.
	
	Using this parameter causes the script to not check for or load the ActiveDirectory module.
	
	Alias is IF
.PARAMETER OrganizationalUnit
	Restricts the retrieval of computer accounts to a specific OU tree. 
	Must be entered in Distinguished Name format. i.e. OU=XenDesktop,DC=domain,DC=tld. 
	
	The script retrieves computer accounts from the top level OU and all sub-level OUs.
	
	Alias OU
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -Folder \\ServerName\Share
	
	Saves the two output text files in \\ServerName\Share.
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -ComputerName XEN
	
	Retrieves all Folder Redirection errors for the last 30 days.
	
	The script will only search computers with "XEN" in the DNSHostName.
	
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -ComputerName RDS -Folder \\FileServer\ShareName
	
	The script will only search computers with "RDS" in the DNSHostName.
	
	Output file will be saved in the path \\FileServer\ShareName

.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -ComputerName CTX -Folder \\FileServer\ShareName -InputFile c:\Scripts\computers.txt
	
	The script will only search computers with "CTX" in the entries contained in the computers.txt file.

	Output file will be saved in the path \\FileServer\ShareName

	InputFile causes the script to not check for or use the ActiveDirectory module.

.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -OU "ou=RDS Servers,dc=domain,dc=tld"
	
	Gathers GPO time in all computers found in the "ou=RDS Servers,dc=domain,dc=tld" OU tree.
	
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -MaxSeconds 10
	
	When the total group policy processing time is greater than or equal to 10 seconds,
	the time, user name and server name are recorded in LongGPOTimes.txt.
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV2.ps1 -MaxSeconds 17 -Folder c:\LogFiles
	
	When the total group policy processing time is greater than or equal to 17 seconds,
	the time, user name and server name are recorded in LongGPOTimes.txt.
	
	Saves the two output text files in C:\LogFiles.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.
	The script creates two text files:
		LongGPOTimes.txt
		GPOAvgMinMaxTimes.txt
		
	By default, the two text files are stored in the folder where the script is run.
.NOTES
	NAME: Get-GPOTimeV2.ps1
	VERSION: 2.0
	AUTHOR: Carl Webster
	LASTEDIT: May 18, 2017
#>


#Created by Carl Webster, CTP and independent consultant 05-Mar-2016
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com

[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Int]$MaxSeconds = 30,

	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("CN")]
	[string]$ComputerName,
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("IF")]
	[string]$InputFile="",
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("OU")]
	[string]$OrganizationalUnit="",
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[string]$Folder=""
	
	)

#Version 1.1 24-Mar-2016
#	Allows you to specify the maximum number of seconds group policy processing should take. Any number greater than or equal to that number is recorded in LongGPOTimes.txt.
#	Allows you to specify an output folder.
#	Records the long GPO times in an text file.
#	Records the Average, Minimum and Maximum processing time to GPOAvgMinMaxTimes.txt.
#	GPOAvgMinMaxTimes.txt is a cumulative file and records the Average, Minimum and Maximum times for each run of the script.
#
#Version 2.0
#	Remove dependence on XenApp 6.x cmdlets

Set-StrictMode -Version 2
	
Write-Host "$(Get-Date): Setting up script"

If(![String]::IsNullOrEmpty($InputFile))
{
	Write-Host "$(Get-Date): Validating input file"
	If(!(Test-Path $InputFile))
	{
		Write-Error "Input file specified but $InputFile does not exist. Script cannot continue."
		Exit
	}
}

If($Folder -ne "")
{
	Write-Host "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Host "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder. Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#test to see if OrganizationalUnit is valid
If(![String]::IsNullOrEmpty($OrganizationalUnit))
{
	Write-Host "$(Get-Date): Validating Organnization Unit"
	try 
	{
		$results = Get-ADOrganizationalUnit -Identity $OrganizationalUnit
	} 
	
	catch
	{
		#does not exist
		Write-Error "Organization Unit $OrganizationalUnit does not exist.`n`nScript cannot continue`n`n"
		Exit
	}	
}

If($Folder -eq "")
{
	$pwdpath = $pwd.Path
}
Else
{
	$pwdpath = $Folder
}

If($pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If($ModuleFound -ne $ModuleName) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function ElevatedSession
{
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose "$(Get-Date): This is an elevated PowerShell session"
		Return $True
	}
	Else
	{
		Write-Host "" -Foreground White
		Write-Host "$(Get-Date): This is NOT an elevated PowerShell session" -Foreground White
		Write-Host "" -Foreground White
		Return $False
	}
}

#only check for the ActiveDirectory module if an InputFile was not entered
If([String]::IsNullOrEmpty($InputFile) -and !(Check-LoadedModule "ActiveDirectory"))
{
	Write-Host "Unable to run script, no ActiveDirectory module"
	Exit
}

If(!(ElevatedSession))
{
	Write-Host "Unable to run script. Rerun script in an elevated session."
	Exit
}

Function ProcessComputer 
{
	Param([string]$TmpComputerName)
	
	If(Test-Connection -ComputerName $TmpComputerName -quiet -EA 0)
	{
		try
		{
			Write-Host "$(Get-Date): `tRetrieving Group Policy event log entries"
			$GPTime = Get-WinEvent -logname Microsoft-Windows-GroupPolicy/Operational `
			-computername $TmpComputerName | `
			Where {$_.id -eq "8001"} | `
			Select message
		}
		
		catch
		{
			Write-Host "$(Get-Date): `tComputer $($TmpComputerName) had error being accessed"
			Continue
		}
		
		If($? -and $Null -ne $GPTime)
		{
			ForEach($GPT in $GPTime)
			{
				$tmparray = $GPT.Message.ToString().Split(" ")
				[int]$GPOTime = $tmparray[8]
				If($GPOTime -ne 0)
				{
					$Script:TimeArray += $GPOTime
				}
				If($GPOTime -ge $MaxSeconds)
				{
					$obj = New-Object -TypeName PSObject
					$obj | Add-Member -MemberType NoteProperty -Name MaxSeconds	-Value $GPOTime
					$obj | Add-Member -MemberType NoteProperty -Name User		-Value $tmparray[6]
					$obj | Add-Member -MemberType NoteProperty -Name Server		-Value $TmpComputerName
					$Script:LongGPOsArray += $obj
				}
			}
		}
	}
	Else
	{
		Write-Host "`tComputer $($TmpComputerName) is not online"
		Out-File -FilePath $Script:FileName2 -Append -InputObject "Computer $($TmpComputerName) was not online $(Get-Date)"
	}
}

$startTime = Get-Date
[string]$Script:FileName1 = "$($pwdpath)\LongGPOTimes.txt"
[string]$Script:FileName2 = "$($pwdpath)\GPOAvgMinMaxTimes.txt"
#make sure filename1 contains the current date only
Out-File -FilePath $Script:FileName1 -InputObject (Get-Date)

If(![String]::IsNullOrEmpty($ComputerName) -and [String]::IsNullOrEmpty($InputFile))
{
	#computer name but no input file
	Write-Host "$(Get-Date): Retrieving list of computers from Active Directory"
	$testname = "*$($ComputerName)*"
	If(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		$Computers = Get-AdComputer -filter {DNSHostName -like $testname} -SearchBase $OrganizationalUnit -SearchScope Subtree -properties DNSHostName, Name -EA 0 | Sort Name
	}
	Else
	{
		$Computers = Get-AdComputer -filter {DNSHostName -like $testname} -properties DNSHostName, Name -EA 0 | Sort Name
	}
}
ElseIf([String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
{
	#input file but no computer name
	Write-Host "$(Get-Date): Retrieving list of computers from Input File"
	$Computers = Get-Content $InputFile
}
ElseIf(![String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
{
	#both computer name and input file
	Write-Host "$(Get-Date): Retrieving list of computers from Input File"
	$testname = "*$($ComputerName)*"
	$Computers = Get-Content $InputFile | ? {$_ -like $testname}
}
Else
{
	Write-Host "$(Get-Date): Retrieving list of computers from Active Directory"
	If(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		$Computers = Get-AdComputer -filter {OperatingSystem -like "*server*"} -SearchBase $OrganizationalUnit -SearchScope Subtree -properties DNSHostName, Name -EA 0 | Sort Name
	}
	Else
	{
		$Computers = Get-AdComputer -filter {OperatingSystem -like "*server*"} -properties DNSHostName, Name -EA 0 | Sort Name
	}
}

If($? -and $Null -ne $Computers)
{
	If($Computers -is [array])
	{
		Write-Host "Found $($Computers.Count) servers to process"
	}
	Else
	{
		Write-Host "Found 1 server to process"
	}

	$Script:TimeArray = @()
	$Script:LongGPOsArray = @()

	If(![String]::IsNullOrEmpty($InputFile))
	{
		ForEach($Computer in $Computers)
		{
			$TmpComputerName = $Computer
			Write-Host "Testing computer $($TmpComputerName)"
			ProcessComputer $TmpComputerName
		}
	}
	Else
	{
		ForEach($Computer in $Computers)
		{
			$TmpComputerName = $Computer.DNSHostName
			Write-Host "Testing computer $($TmpComputerName)"
			ProcessComputer $TmpComputerName
		}
	}

	Write-Host "$(Get-Date): Output long GPO times to file"
	#first sort array by seconds, longest to shortest
	$Script:LongGPOsArray = $Script:LongGPOsArray | Sort MaxSeconds -Descending
	Out-File -FilePath $Script:FileName1 -InputObject $Script:LongGPOsArray

	If(Test-Path "$($Script:FileName1)")
	{
		Write-Host "$(Get-Date): $($Script:FileName1) is ready for use"
	}

	$Avg = ($Script:TimeArray | Measure-Object -Average -minimum -maximum)
	Write-Host "Average: " $Avg.Average
	Write-Host "Minimum: " $Avg.Minimum
	Write-Host "Maximum: " $Avg.Maximum

	Write-Host "$(Get-Date): Output GPO Avg/Min/Max times to file"
	Out-File -FilePath $Script:FileName2 -Append -InputObject " "
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Average: $($Avg.Average) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Minimum: $($Avg.Minimum) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Maximum: $($Avg.Maximum) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject " "

	If(Test-Path "$($Script:FileName2)")
	{
		Write-Host "$(Get-Date): $($Script:FileName2) is ready for use"
	}

	Write-Host "$(Get-Date): Script started: $($StartTime)"
	Write-Host "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Host "$(Get-Date): Elapsed time: $($Str)"
	$runtime = $Null
}
Else
{
	If(![String]::IsNullOrEmpty($ComputerName) -and [String]::IsNullOrEmpty($InputFile))
	{
		#computer name but no input file
		Write-Host "Unable to retrieve a list of computers from Active Directory"
	}
	ElseIf([String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
	{
		#input file but no computer name
		Write-Host "Unable to retrieve a list of computers from the Input File $InputFile"
	}
	ElseIf(![String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
	{
		#computer name and input file
		Write-Host "Unable to retrieve a list of matching computers from the Input File $InputFile"
	}
	Else
	{
		Write-Host "Unable to retrieve a list of computers from Active Directory"
	}
}



<#
If($? -and $Null -ne $servers)
{
	If($servers -is [Array])
	{
		[int]$Total = $servers.count
	}
	Else
	{
		[int]$Total = 1
	}
	Write-Host "$(Get-Date): Found $($Total) XenApp servers"
	$Script:TimeArray = @()
	$Script:LongGPOsArray = @()
	$cnt = 0
	ForEach($server in $servers)
	{
		$cnt++
		Write-Host "$(Get-Date): Processing server $($Server.ServerName) $($Total - $cnt) left"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			try
			{
				$GPTime = Get-WinEvent -logname Microsoft-Windows-GroupPolicy/Operational `
				-computername $server.servername | Where {$_.id -eq "8001"} | Select message
			}
			
			catch
			{
				Write-Host "$(Get-Date): `tServer $($Server.ServerName) had error being accessed"
				Continue
			}
			
			If($? -and $Null -ne $GPTime)
			{
				ForEach($GPT in $GPTime)
				{
					$tmparray = $GPT.Message.ToString().Split(" ")
					[int]$GPOTime = $tmparray[8]
					If($GPOTime -ne 0)
					{
						$Script:TimeArray += $GPOTime
					}
					If($GPOTime -ge $MaxSeconds)
					{
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name MaxSeconds	-Value $GPOTime
						$obj | Add-Member -MemberType NoteProperty -Name User		-Value $tmparray[6]
						$obj | Add-Member -MemberType NoteProperty -Name Server		-Value $server.servername
						$Script:LongGPOsArray += $obj
					}
					
				}
			}
		}
		Else
		{
			Write-Host "$(Get-Date): `tServer $($Server.ServerName) is not online"
		}
	}
	
	Write-Host "$(Get-Date): Output long GPO times to file"
	#first sort array by seconds, longest to shortest
	$Script:LongGPOsArray = $Script:LongGPOsArray | Sort MaxSeconds -Descending
	Out-File -FilePath $Script:FileName1 -InputObject $Script:LongGPOsArray

	If(Test-Path "$($Script:FileName1)")
	{
		Write-Host "$(Get-Date): $($Script:FileName1) is ready for use"
	}

	$Avg = ($Script:TimeArray | Measure-Object -Average -minimum -maximum)
	Write-Host "Average: " $Avg.Average
	Write-Host "Minimum: " $Avg.Minimum
	Write-Host "Maximum: " $Avg.Maximum

	Write-Host "$(Get-Date): Output GPO Avg/Min/Max times to file"
	Out-File -FilePath $Script:FileName2 -Append -InputObject " "
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Average: $($Avg.Average) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Minimum: $($Avg.Minimum) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject "$(Get-Date): Maximum: $($Avg.Maximum) seconds"
	Out-File -FilePath $Script:FileName2 -Append -InputObject " "

	If(Test-Path "$($Script:FileName2)")
	{
		Write-Host "$(Get-Date): $($Script:FileName2) is ready for use"
	}
}
ElseIf($? -and $Null -eq $servers)
{
	Write-Warning "Server information could not be retrieved"
}
Else
{
	Write-Warning "No results returned for Server information"
}

Write-Host "$(Get-Date): Script started: $($StartTime)"
Write-Host "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Host "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null

#>
