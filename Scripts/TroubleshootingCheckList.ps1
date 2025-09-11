param(
    $ComputerName
)

$ErrorSummary = New-Object System.Collections.ArrayList

function Check-LocalTime{
    $properties = @("Day",
    "DayOfWeek",
    "Hour",
    "Minute",
    "Month",
    "Quarter",
    "Second",
    "WeekInMonth",
    "Year")
    
    $localComputerDateTime = Get-WmiObject -Class Win32_LocalTime -Property $properties
    $remoteComputerDateTime = Get-WmiObject -Class Win32_LocalTime -Property $properties -ComputerName $ComputerName

    if($localComputerDateTime -ne $remoteComputerDateTime){
        $ErrorSummary.Add("The date or time of the remote computer don't match the local machine.")
    }
}

function Check-DiskSpace{
    $freeSpaceOnCDrive = Get-Counter -Counter "\LogicalDisk(c:)\% Free Space" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue
    
    if($freeSpaceOnCDrive -lt 15){
        $ErrorSummary.Add("The C: drive on $ComputerName has $freeSpaceOnCDrive free space on the C: drive. Consider getting more storage.")
    }
}

Check-LocalTime
Check-DiskSpace

Write-Host $ErrorSummary