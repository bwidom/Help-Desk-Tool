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

function Check-PhysicalDiskIdleTime{
    $diskIdleTime = Get-Counter -Counter "\PhysicalDisk(_total)\% Idle Time" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue

    if($diskIdleTime -lt 20){
        $ErrorSummary.Add("Your disk system is saturated. You should consider replacing the current disk system with a faster one.")
    }
}

function Check-AverageDiskReadTime{
    $averageDiskReadSeconds = Get-Counter -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Read" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue
}

Check-LocalTime
Check-DiskSpace
Check-PhysicalDiskIdleTime

ForEach($message in $ErrorSummary){
    Write-Host $message
}