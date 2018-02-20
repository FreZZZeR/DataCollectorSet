<#
Windows PowerShell Performance Monitor Data Collection Set Module
This module contains a set of wrapper scripts that enable a user to start, stop, add and remove Data Collector Set in Performance Monitor
Some changes were added in module
#>
#
#FUNCTIONS
#
function Start-DataCollectorSet 
{
<#
.SYNOPSIS
   Start a Data Collector Set (DCS) in local or remote computer
.DESCRIPTION
   PowerShell version 4 or higher
   Start a DCS in local or remote computer. If DCS is already working it will be restarted.
.PARAMETER ComputerName
   Local or remote computer name. Use FQDN, NET-BIOS name or "localhost" for local computer
.PARAMETER DCSName
   Data Collector Set name
.PARAMETER LogFile
   Path to log-file
.EXAMPLE
   Start-DataCollectorSet -ComputerName server-test1 -DCSName Proccessor_Time -LogFile \\server-test1\logs\ps\dcs\dcs.log
.EXAMPLE
   Start-DataCollectorSet -CN server-test1 -DCSName Disk_Time
   Log will be written to default path to %userprofile%\appdata\local\temp\PowerShell_Module_DCS_$CurrentDate.log
.EXAMPLE
   Start-DCS server-test2 Memory
   Short variant using the function (only for PowerShell version 5 and later)
#>
	[CmdletBinding ()]
	[Alias("Start-DCS")]
	Param (
	[PARAMETER (Mandatory=$true, ValueFromPipeline=$true, Position=0)][Alias("CN")]$ComputerName,
	[PARAMETER (Mandatory=$true, Position=1)][string]$DCSName,
	[PARAMETER (Mandatory=$false)][IO.FileInfo]$LogFile
	)
	Process 
	{
		Write-Log -Message "Command $($myinvocation.MyCommand.Name) has been run by $env:USERDNSDOMAIN\$env:USERNAME on computer $env:COMPUTERNAME" -Path $LogFile
		foreach ($Computer in $ComputerName) 
		{
			If ($Computer -eq "") 
			{
				$Computer = "$env:COMPUTERNAME"
			}
			If ([int]::Parse(($(Get-WmiObject Win32_OperatingSystem -ComputerName $Computer).Version).Split("\.")[0]) -ge "10") 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to connect to computer `"$Computer`"..."
				Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status() -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be restarted."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be restarted." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status() -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
						$PerfMonDataCollectorSet.Start($false)
						While ($PerfMonDataCollectorSet.Status() -eq "0") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is starting..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is starting..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -Level Success  -Path $LogFile
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" will be started."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be started." -Path $LogFile
						$PerfMonDataCollectorSet.Start($false)
						While ($PerfMonDataCollectorSet.Status() -eq "0") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is starting..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is starting..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -Level Success  -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
			Else 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to connect to computer `"$Computer`"..."
				Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be restarted."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be restarted." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
						$PerfMonDataCollectorSet.Start($false)
						While ($PerfMonDataCollectorSet.Status -eq "0") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is starting..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is starting..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -Level Success -Path $LogFile
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" will be started."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be started." -Path $LogFile
						$PerfMonDataCollectorSet.Start($false)
						While ($PerfMonDataCollectorSet.Status -eq "0") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is starting..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is starting..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been started on computer `"$Computer`"." -Level Success -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
		}
	}
}
#
function Stop-DataCollectorSet 
{
<#
.SYNOPSIS
   Stop a Data Collector Set (DCS) in local or remote computer
.DESCRIPTION
   PowerShell version 4 or higher
   Stop a DCS in local or remote computer.
   Logging to file temporary is not available
.PARAMETER ComputerName
   Local or remote computer name. Use FQDN, NET-BIOS name or "localhost" for local computer
.PARAMETER DCSName
   Data Collector Set name
.PARAMETER LogFile
   Path to log-file
.EXAMPLE
   Stop-DataCollectorSet -Computer server-test1 -DCSName Proccessor_Time -LogFile \\server-test1\logs\ps\dcs\dcs.log
.EXAMPLE
   Stop-DataCollectorSet -Computer server-test1 -DCSName Disk_Time
   Log will be written to default path to %userprofile%\appdata\local\temp\PowerShell_Module_DCS_$CurrentDate.log
.EXAMPLE
   Stop-DCS server-test2 Memory
   Short variant using the function (only for PowerShell version 5 and later)
#>
	[CmdletBinding ()]
	[Alias("Stop-DCS")]
	Param (
	[PARAMETER (Mandatory=$true, ValueFromPipeline=$true, Position=0)][Alias("CN")]$ComputerName,
	[PARAMETER (Mandatory=$true, Position=1)][string]$DCSName,
	[PARAMETER (Mandatory=$false, HelpMessage="Temporary is not available")][AllowNull()][IO.FileInfo]$LogFile
	)
	Process 
	{
		Write-Log -Message "Command $($myinvocation.MyCommand.Name) has been run by $env:USERDNSDOMAIN\$env:USERNAME on computer $env:COMPUTERNAME" -Path $LogFile
		foreach ($Computer in $ComputerName) 
		{
			If ($Computer -eq "") 
			{
				$Computer = "$env:COMPUTERNAME"
			}
			If ([int]::Parse(($(Get-WmiObject Win32_OperatingSystem -ComputerName $Computer).Version).Split("\.")[0]) -ge "10") 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to connect to computer `"$Computer`"..."
				Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status() -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status() -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped on computer `"$Computer`"." -Level Success -Path $LogFile
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" is not working now on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" is not working now on computer `"$Computer`"." -Level Success -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error  -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
			Else 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to connect to computer `"$Computer`"..."
				Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped on computer `"$Computer`"." -Level Success -Path $LogFile
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" is not working now on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" is not working now on computer `"$Computer`"." -Level Success -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
		}
	}
}
#
function Add-DataCollectorSet 
{
<#
.SYNOPSIS
   Add a Data Collector Set (DCS) in local or remote computer from xml-file.
.DESCRIPTION
   PowerShell version 4 or higher
   Add a DCS in local or remote computer. If the DCS is already present, it will be stopped, removed and added again, when -Force flag is present.
   Logging to file temporary is not available
.PARAMETER ComputerName
   Local or remote computer name. Use FQDN, NET-BIOS name or "localhost" for local computer
.PARAMETER DCSName
   Data Collector Set name
.PARAMETER DCSXMLTemplate
   Path to xml-template
.PARAMETER LogFile
   Path to log-file
.EXAMPLE
   Add-DataCollectorSet -ComputerName server-test1 -DCSName Proccessor_Time -DCSXMLTemplate "C:\test.xml" -LogFile \\server-test1\logs\ps\dcs\dcs.log
.EXAMPLE
   Add-DataCollectorSet -CN server-test1 -DCSName Disk_Time -XML "C:\test.xml"
   Log will be written to default path to %userprofile%\appdata\local\temp\PowerShell_Module_DCS_$CurrentDate.log
.EXAMPLE
   Add-DCS server-test2 Memory "C:\test.xml"
   Short variant using the function (only for PowerShell version 5 and later)
#>
	[CmdletBinding ()]
	[Alias("Add-DCS")]
	Param (
	[PARAMETER (Mandatory=$true, ValueFromPipeline=$true, Position=0)][Alias("CN")]$ComputerName,
	[PARAMETER (Mandatory=$true, Position=1)][string]$DCSName,
	[PARAMETER (Mandatory=$true, Position=2)][string][Alias("XML")][IO.FileInfo]$DCSXMLTemplate,
	[PARAMETER (Mandatory=$false)][switch]$Force,
	[PARAMETER (Mandatory=$false)][IO.FileInfo]$LogFile
	)
	Process 
	{
		Write-Log -Message "Command $($myinvocation.MyCommand.Name) has been run by $env:USERDNSDOMAIN\$env:USERNAME on computer $env:COMPUTERNAME" -Path $LogFile
		$XMLData = Get-Content -Path $DCSXMLTemplate
		foreach ($Computer in $ComputerName) 
		{
			If ($Computer -eq "") 
			{
				$Computer = "$env:COMPUTERNAME"
			}
			If ([int]::Parse(($(Get-WmiObject Win32_OperatingSystem -ComputerName $Computer).Version).Split("\.")[0]) -ge "10") 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..."
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" was found."
					Write-Log -Message "Data Collector Set `"$DCSName`" was found." -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" is already present."
					Write-Log -Message "Data Collector Set `"$DCSName`" is already present." -Path $LogFile
					If ($Force) 
					{
						Write-Host "Checking Data Collector Set `"$DCSName`" status..."
						Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
						If ($PerfMonDataCollectorSet.Status() -eq "1") 
						{
							Write-Host "Data Collector Set `"$DCSName`" is working now."
							Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" will be stopped and then removed."
							Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped and then removed." -Path $LogFile
							$PerfMonDataCollectorSet.Stop($false)
							While ($PerfMonDataCollectorSet.Status() -eq "1") 
							{
								Start-Sleep -Milliseconds 500
								Write-Host "Data Collector Set `"$DCSName`" is stopping..."
								Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
							}
							Write-Host "Data Collector Set `"$DCSName`" has been stopped."
							Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" removing..."
							Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
							$PerfMonDataCollectorSet.Delete()
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been removed."
								Write-Log -Message "Data Collector Set `"$DCSName`" has been removed." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
								$PerfMonDataCollectorSet.SetXml($XMLData)
								If ($? -eq $true) 
								{
									Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
									Write-Host "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..."
									Write-Log -Message "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..." -Path $LogFile
									$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
									If ($? -eq $true)
									{
										Write-Host "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -Level Success -Path $LogFile
									}
									Else 
									{
										Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -Level Error -Path $LogFile
									}
								}
								Else 
								{
									Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
								}
							}
							Else 
							{
								Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -Level Error -Path $LogFile
							}
						}
						Else 
						{
							Write-Host "Data Collector Set `"$DCSName`" removing..."
							Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
							$PerfMonDataCollectorSet.Delete()
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been removed."
								Write-Log -Message "Data Collector Set `"$DCSName`" has been removed." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
								$PerfMonDataCollectorSet.SetXml($XMLData)
								If ($? -eq $true) 
								{
									Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
									Write-Host "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..."
									Write-Log -Message "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..." -Path $LogFile
									$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
									If ($? -eq $true)
									{
										Write-Host "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -Level Success -Path $LogFile
									}
									Else 
									{
										Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -Level Error -Path $LogFile
									}
								}
								Else 
								{
									Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
								}
							}
							Else 
							{
								Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -Level Error -Path $LogFile
							}
						}
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" has been left on computer `"$Computer`"." -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been left on computer `"$Computer`"." -Level Warning -Path $LogFile
						Write-Host "Use -Force flag for rewriting Data Collector Set `"$DCSName`"" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Use -Force flag for rewriting Data Collector Set `"$DCSName`"" -Level Warning -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be added."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be added." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" adding..."
						Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
						Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
						$PerfMonDataCollectorSet.SetXml($XMLData)
						If ($? -eq $true) 
						{
							Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
							Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" committing data..."
							Write-Log -Message "Data Collector Set `"$DCSName`" committing data..." -Path $LogFile
							$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been added on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
								Write-Log -Message "Data Collector Set `"$DCSName`" has been added on computer `"$Computer`"." -Level Success -Path $LogFile
							}
							Else 
							{
								Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added!" -Level Error -Path $LogFile
							}
						}
						Else 
						{
							Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
						}
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established!" -Level Error -Path $LogFile
					}
				}
			}
			Else 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..."
				Write-Log -Message "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" was found."
					Write-Log -Message "Data Collector Set `"$DCSName`" was found." -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" is already present."
					Write-Log -Message "Data Collector Set `"$DCSName`" is already present." -Path $LogFile
					If ($Force) 
					{
						Write-Host "Checking Data Collector Set `"$DCSName`" status..."
						Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
						If ($PerfMonDataCollectorSet.Status -eq "1") 
						{
							Write-Host "Data Collector Set `"$DCSName`" is working now."
							Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" will be stopped and then removed."
							Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped and then removed." -Path $LogFile
							$PerfMonDataCollectorSet.Stop($false)
							While ($PerfMonDataCollectorSet.Status -eq "1") 
							{
								Start-Sleep -Milliseconds 500
								Write-Host "Data Collector Set `"$DCSName`" is stopping..."
								Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
							}
							Write-Host "Data Collector Set `"$DCSName`" has been stopped."
							Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" removing..."
							Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
							$PerfMonDataCollectorSet.Delete()
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been removed."
								Write-Log -Message "Data Collector Set `"$DCSName`" has been removed." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
								$PerfMonDataCollectorSet.SetXml($XMLData)
								If ($? -eq $true) 
								{
									Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
									Write-Host "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..."
									Write-Log -Message "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..." -Path $LogFile
									$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
									If ($? -eq $true)
									{
										Write-Host "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -Level Success -Path $LogFile
									}
									Else 
									{
										Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -Level Error -Path $LogFile
									}
								}
								Else 
								{
									Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
								}
							}
							Else 
							{
								Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -Level Error -Path $LogFile
							}
						}
						Else 
						{
							Write-Host "Data Collector Set `"$DCSName`" removing..."
							Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
							$PerfMonDataCollectorSet.Delete()
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been removed."
								Write-Log -Message "Data Collector Set `"$DCSName`" has been removed." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
								Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
								Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
								$PerfMonDataCollectorSet.SetXml($XMLData)
								If ($? -eq $true) 
								{
									Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
									Write-Host "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..."
									Write-Log -Message "Data Collector Set `"$DCSName`" committing data to computer `"$Computer`"..." -Path $LogFile
									$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
									If ($? -eq $true)
									{
										Write-Host "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has been added to computer `"$Computer`"." -Level Success -Path $LogFile
									}
									Else 
									{
										Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
										Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added to computer `"$Computer`"!" -Level Error -Path $LogFile
									}
								}
								Else 
								{
									Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
									Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
								}
							}
							Else 
							{
								Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`" on computer `"$Computer`"!" -Level Error -Path $LogFile
							}
						}
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" has been left on computer `"$Computer`"." -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Data Collector Set `"$DCSName`" has been left on computer `"$Computer`"." -Level Warning -Path $LogFile
						Write-Host "Use -Force flag for rewriting Data Collector Set `"$DCSName`"" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Use -Force flag for rewriting Data Collector Set `"$DCSName`"" -Level Warning -Path $LogFile
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be added."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be added." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" adding..."
						Write-Log -Message "Data Collector Set `"$DCSName`" adding..." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" adding XML-data..."
						Write-Log -Message "Data Collector Set `"$DCSName`" adding XML-data..." -Path $LogFile
						$PerfMonDataCollectorSet.SetXml($XMLData)
						If ($? -eq $true) 
						{
							Write-Host "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object."
							Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has been added to COM-object." -Path $LogFile
							Write-Host "Data Collector Set `"$DCSName`" committing data..."
							Write-Log -Message "Data Collector Set `"$DCSName`" committing data..." -Path $LogFile
							$null = $PerfMonDataCollectorSet.Commit("$DCSName", $Computer, 0x0003)
							If ($? -eq $true)
							{
								Write-Host "Data Collector Set `"$DCSName`" has been added on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
								Write-Log -Message "Data Collector Set `"$DCSName`" has been added on computer `"$Computer`"." -Level Success -Path $LogFile
							}
							Else 
							{
								Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been added!" -ForegroundColor Red -BackgroundColor DarkBlue
								Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been added!" -Level Error -Path $LogFile
							}
						}
						Else 
						{
							Write-Host "Error! XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "XML-data for Data Collector Set `"$DCSName`" has NOT been added to COM-object!" -Level Error -Path $LogFile
						}
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established!" -Level Error -Path $LogFile
					}
				}
			}
		}
	}
}
#
function Remove-DataCollectorSet 
{
<#
.SYNOPSIS
   Remove a Data Collector Set (DCS) in local or remote computer.
.DESCRIPTION
   PowerShell version 4 or higher
   Remove a DCS in local or remote computer. If the DCS is working at the moment, it will be stopped and then removed.
   Logging to file temporary is not available
.PARAMETER ComputerName
   Local or remote computer name. Use FQDN, NET-BIOS name or "localhost" for local computer
.PARAMETER DCSName
   Data Collector Set name
.PARAMETER LogFile
   Path to log-file
.EXAMPLE
   Remove-DataCollectorSet -ComputerName server-test1 -DCSName Proccessor_Time -LogFile \\server-test1\logs\ps\dcs\dcs.log
.EXAMPLE
   Remove-DataCollectorSet -CN server-test1 -DCSName Disk_Time
   Log will be written to default path to %userprofile%\appdata\local\temp\PowerShell_Module_DCS_$CurrentDate.log
.EXAMPLE
   Remove-DCS server-test2 Memory
   Short variant using the function (only for PowerShell version 5 and later)
#>
	[CmdletBinding ()]
	[Alias("Remove-DCS")]
	Param (
	[PARAMETER (Mandatory=$true, ValueFromPipeline=$true, Position=0)][Alias("CN")]$ComputerName,
	[PARAMETER (Mandatory=$true, Position=1)][string]$DCSName,
	[PARAMETER (Mandatory=$false)][IO.FileInfo]$LogFile
	)
	Process 
	{
		Write-Log -Message "Command $($myinvocation.MyCommand.Name) has been run by $env:USERDNSDOMAIN\$env:USERNAME on computer $env:COMPUTERNAME" -Path $LogFile
		foreach ($Computer in $ComputerName) 
		{
			If ($Computer -eq "") 
			{
				$Computer = "$env:COMPUTERNAME"
			}
			If ([int]::Parse(($(Get-WmiObject Win32_OperatingSystem -ComputerName $Computer).Version).Split("\.")[0]) -ge "10") 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..."
				Write-Log -Message "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" was found."
					Write-Log -Message "Data Collector Set `"$DCSName`" was found." -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status() -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be stopped and then removed."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped and then removed." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status() -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" removing..."
						Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
						$PerfMonDataCollectorSet.Delete()
						If ($? -eq $true)
						{
							Write-Host "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -Level Success -Path $LogFile
						}
						Else 
						{
							Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`"!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`"!" -Level Error -Path $LogFile
						}
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" removing..."
						Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
						$PerfMonDataCollectorSet.Delete()
						If ($? -eq $true)
						{
							Write-Host "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -Level Success -Path $LogFile
						}
						Else 
						{
							Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been removed on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been removed on computer `"$Computer`"!" -Level Error -Path $LogFile
						}
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
			Else 
			{
				$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
				Write-Host "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..."
				Write-Log -Message "Trying to find Data Collector Set `"$DCSName`" on computer `"$Computer`"..." -Path $LogFile
				$PerfMonDataCollectorSet.Query($DCSName, $Computer)
				If ($? -eq $true) 
				{
					Write-Host "Successfully! Connection is established!"
					Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
					Write-Host "Data Collector Set `"$DCSName`" was found."
					Write-Log -Message "Data Collector Set `"$DCSName`" was found." -Path $LogFile
					Write-Host "Checking Data Collector Set `"$DCSName`" status..."
					Write-Log -Message "Checking Data Collector Set `"$DCSName`" status..." -Path $LogFile
					If ($PerfMonDataCollectorSet.Status -eq "1") 
					{
						Write-Host "Data Collector Set `"$DCSName`" is working now."
						Write-Log -Message "Data Collector Set `"$DCSName`" is working now." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" will be stopped and then removed."
						Write-Log -Message "Data Collector Set `"$DCSName`" will be stopped and then removed." -Path $LogFile
						$PerfMonDataCollectorSet.Stop($false)
						While ($PerfMonDataCollectorSet.Status -eq "1") 
						{
							Start-Sleep -Milliseconds 500
							Write-Host "Data Collector Set `"$DCSName`" is stopping..."
							Write-Log -Message "Data Collector Set `"$DCSName`" is stopping..." -Path $LogFile
						}
						Write-Host "Data Collector Set `"$DCSName`" has been stopped."
						Write-Log -Message "Data Collector Set `"$DCSName`" has been stopped." -Path $LogFile
						Write-Host "Data Collector Set `"$DCSName`" removing..."
						Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
						$PerfMonDataCollectorSet.Delete()
						If ($? -eq $true)
						{
							Write-Host "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -Level Success -Path $LogFile
						}
						Else 
						{
							Write-Host "Error! Can NOT remove Data Collector Set `"$DCSName`"!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "Can NOT remove Data Collector Set `"$DCSName`"!" -Level Error -Path $LogFile
						}
					}
					Else 
					{
						Write-Host "Data Collector Set `"$DCSName`" removing..."
						Write-Log -Message "Data Collector Set `"$DCSName`" removing..." -Path $LogFile
						$PerfMonDataCollectorSet.Delete()
						If ($? -eq $true)
						{
							Write-Host "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -ForegroundColor Green -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has been removed on computer `"$Computer`"." -Level Success -Path $LogFile
						}
						Else 
						{
							Write-Host "Error! Data Collector Set `"$DCSName`" has NOT been removed on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
							Write-Log -Message "Data Collector Set `"$DCSName`" has NOT been removed on computer `"$Computer`"!" -Level Error -Path $LogFile
						}
					}
				}
				Else 
				{
					Write-Host "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!"  -ForegroundColor Yellow -BackgroundColor DarkBlue
					Write-Log -Message "Data Collector Set `"$DCSName`" is not found or Connection is not established with computer `"$Computer`"!" -Level Warning -Path $LogFile
					Write-Host "Trying to connect to computer `"$Computer`"..."
					Write-Log -Message "Trying to connect to computer `"$Computer`"..." -Path $LogFile
					$PerfMonDataCollectorSet.Query("System\System Diagnostics", $Computer)
					If ($? -eq $true) 
					{
						Write-Host "Successfully! Connection is established!"
						Write-Log -Message "Successfully! Connection is established!" -Path $LogFile
						Write-Host "Check the Data Collector Set name!" -ForegroundColor Yellow -BackgroundColor DarkBlue
						Write-Log -Message "Check the Data Collector Set name!" -Level Warning -Path $LogFile
						Write-Host "Error! Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Probably Data Collector Set `"$DCSName`" is not presented on computer `"$Computer`"!" -Level Error -Path $LogFile
					}
					Else 
					{
						Write-Host "Error! Connection is NOT established with computer `"$Computer`"!" -ForegroundColor Red -BackgroundColor DarkBlue
						Write-Log -Message "Connection is NOT established with computer `"$Computer`"!" -Level Error -Path $LogFile
					}
				}
			}
		}
	}
}
#
function Write-Log 
{
<#
.SYNOPSIS
   Writing messages to file.
.DESCRIPTION
   PowerShell version 4 or higher
   This function writes messages to log file with severity level.
   Alias for this function "wl"
.PARAMETER Message
   A message to log-file
.PARAMETER Path
   Log-file path
.PARAMETER Level
   Severity level ("Success", "Information", "Warning", "Error")
.EXAMPLE
   Write-Log -Message "This message will be written to $Path with date-time before text with severity level $Level" -Level "Error" -Path "C:\test.log"
   Full using
.EXAMPLE
   Write-Log "Test message will be written to $Path with severity Error" Error
   Using without naming parameters with severity level Error
.EXAMPLE
   wl "Test message"
   Short variant using the function (only for PowerShell version 5 and later)
#>
	[CmdletBinding ()]
	[Alias("wl")]
	Param (
	[PARAMETER(Mandatory=$true, Position=0, ValueFromPipeline=$true)][ValidateNotNullOrEmpty()]$Message,
	[PARAMETER(Mandatory=$false,Position=1)][ValidateSet("Success", "Information", "Warning", "Error")][String]$Level="Information",
	[PARAMETER(Mandatory=$false)][IO.FileInfo]$Path
	)
	Process 
	{
		If (!$Path) 
		{
			$Date = Get-Date -UFormat %Y.%m.%d
			$Path = "$env:TEMP\PowerShell_Module_DCS_$Date.log"
		}
		$DateWrite = Get-Date -Format FileDateTime
		$Line = "{0} ***{1}*** {2}" -f $DateWrite, $Level.ToUpper(), $Message
		Add-Content -Path $Path -Value $Line
	}
}
#
Export-ModuleMember -Function "Start-DataCollectorSet", "Stop-DataCollectorSet", "Add-DataCollectorSet", "Remove-DataCollectorSet", "Write-Log" -Alias "Start-DCS", "Stop-DCS", "Add-DCS", "Remove-DCS", "wl"
