#
# This is a PowerShell Unit Test file.
# You need a unit test framework such as Pester to run PowerShell Unit tests. 
# You can download Pester from http://go.microsoft.com/fwlink/?LinkID=534084
#

Describe "Test Function Start-DataCollectorSet" 
{
	Context "Function Exists" 
	{
		It "Should Return Status 1" 
		{
			Import-Module .\PowerShellModule_DataCollectorSet.psm1
			$DCSName = "Test"
			$Computer = "localhost"
			Start-DCS -CN $Computer -DCSName $DCSName -Force
			$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
			$PerfMonDataCollectorSet.Query($DCSName, $Computer)
			$Result = $PerfMonDataCollectorSet.Status()
			$Result | Should Be "1"
		}
	}
}