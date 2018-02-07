#
# This is a PowerShell Unit Test file.
#
Describe "Test Function Start-DataCollectorSet" 
{
	It "Should Return Status 1" 
	{
		$DCSName = "Test"
		$Computer = "localhost"
		Start-DataCollectorSet -CN $Computer -DCSName $DCSName -Force
		$PerfMonDataCollectorSet = New-Object -ComObject Pla.DataCollectorSet
		$PerfMonDataCollectorSet.Query($DCSName, $Computer)
		$Result = $PerfMonDataCollectorSet.Status()
		$Result | Should Be "1"
	}
}