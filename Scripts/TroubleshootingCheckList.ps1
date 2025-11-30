<#
.SYNOPSIS

Troubleshoot remote or local computer for performance issues.

.DESCRIPTION

A script that collects performance counters and other metrics from remote computers or 
the local computer and compares them to benchmarks to troubleshoot performance issues 
and attempt to identify resource bottlenecks. Performance counter benchmarks based on 
Microsoft Learn document for troubleshooting: 
https://learn.microsoft.com/en-us/training/modules/monitor-troubleshoot-windows-client-performance/4-monitor-windows-client-performance

.PARAMETER ComputerName

Name of computer to troubleshoot.

.EXAMPLE

PS> .\troubleshootingchecklist.ps1 -ComputerName Client1
The processor has an excessive amount of hardware interruptions. There could be a hardware issue.

#>

param(
    $ComputerName = "localhost"
)

$ErrorSummary = New-Object System.Collections.ArrayList

class TroubleshootingCheck{
    [string]$Counter    
    [string]$CorrectAmount
    [string]$ComputerAmount
    [string]$Message    

    TroubleshootingCheck([string]$counter, [string]$correctAmount, [string]$computerAmount, [string]$message){
        $this.counter = $counter
        $this.computerAmount = $computerAmount
        $this.correctAmount = $correctAmount
        $this.message = $message        
    }
}

#Check if remote computer has same date and time as local computer.
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
    #$ErrorSummary.Add("The date or time of the remote computer don't match the local machine.") | Out-Null
}

#Check that C drive has more than 15% free space
$freeSpaceOnCDrive = Get-Counter -Counter "\LogicalDisk(c:)\% Free Space" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue 
    
if($freeSpaceOnCDrive -lt 15){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(
            '\LogicalDisk(c:)\% Free Space',            
            '15%',
            "$freeSpaceonCDrive", 
            'The C: drive is low on free space. Consider getting more storage.'           
        )
    ) | Out-Null
}


#Check if amount of time disk is idle is below 20%
$diskIdleTime = Get-Counter -Counter "\PhysicalDisk(_total)\% Idle Time" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue
if($diskIdleTime -lt 20){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            'Disk Idle Time',
            '20%',
            "$diskIdleTime",            
            'Your disk system is saturated. You should consider replacing the current disk system with a faster one.'
        )        
    ) | Out-Null
}


#Check if average disk read time is longer than 25 milliseconds
$averageDiskReadSeconds = Get-Counter -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Read" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue
if($averageDiskReadSeconds -gt .025){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\PhysicalDisk(_total)\Avg. Disk sec/Read',
            '25 milliseconds',
            "$averageDiskReadSeconds",            
            'The disk system is experiencing latency when reading from disk.'
        )        
    ) | Out-Null    
}


#Check if average disk write time is longer than 25 milliseconds
$averageDiskWriteSeconds = Get-Counter -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Write" -ComputerName $ComputerName -MaxSamples 1 | Select-Object -ExpandProperty CounterSamples |Select-Object -ExpandProperty CookedValue
if($averageDiskWriteSeconds -gt .025){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\PhysicalDisk(_total)\Avg. Disk sec/Write',
            '25 milliseconds',
            "$averageDiskWriteSeconds",            
            'The disk system is experiencing latency when writing to disk.'
        )        
    ) | Out-Null    
}


#Check if amount of memory that file system cache uses is above 300 MB.
$memoryCacheBytes = (Get-counter "\Memory\Cache Bytes" -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue)/1MB
if($memoryCacheBytes -gt 300){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Cache Bytes',
            '300 MB',
            "$memoryCacheBytes",            
            'Memory cache is high. There could be a disk bottleneck.'
        )        
    ) | Out-Null    
}

#Check if the ratio of committed bytes to the commit limit is above 80%
$memoryPercentCommittedBytes = Get-Counter "\Memory\% Committed Bytes In Use" -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($memoryPercentCommittedBytes -gt 80){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\% Committed Bytes In Use',
            '80%',
            "$memoryPercentCommittedBytes",            
            'The percent of committed byes is high. This indicates insufficient memory.'
        )        
    ) | Out-Null    
}

#Check if available memory is less than 5% of total physical memory
$memoryAvailableMBytes = Get-Counter "\Memory\Available MBytes" -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
$totalPhysicalMemory = Get-WmiObject Win32_ComputerSystem -Property TotalPhysicalMemory -ComputerName $ComputerName | Select-Object -ExpandProperty TotalPhysicalMemory
if(($memoryAvailableMBytes * 1048576) -lt ($totalPhysicalMemory * 0.05)){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Available MBytes',
            '5%',
            "$($memoryAvailableMBytes * 1048576)",            
            'The amount of available memory is insufficient for running processes.'
        ) 
    ) | Out-Null    
}

#Check of number of page table entries not in use by system is less than 5000
$memoryFreePageTableEntries = Get-Counter '\Memory\Free System Page Table Entries' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($memoryFreePageTableEntries -lt 5000){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Free System Page Table Entries',
            '5000',
            "$memoryFreePageTableEntries",            
            'The amount of free page table entries is low. There could be a memory leak.'
        ) 
    ) | Out-Null      
}


#Check if the pool non-paged bytes is greater than 175 MB
$memoryNonPagedBytes = Get-Counter '\Memory\Pool Nonpaged Bytes'-ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($memoryNonPagedBytes -gt 175MB){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Pool Nonpaged Bytes',
            '175MB',
            "$memoryNonPagedBytes",            
            'There amount of non-paged bytes is high. There could be a memory leak.'
        ) 
    ) | Out-Null       
}

#Check if pool paged bytes is greater than 250 MB.
$memoryPagedBytes = Get-Counter '\Memory\Pool Paged Bytes' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($memoryPagedBytes -gt 250MB){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Pool Paged Bytes',
            '250MB',
            "$memoryPagedBytes",            
            'There amount of paged bytes is high. There could be a memory leak.'
        ) 
    ) | Out-Null     
}

#Check if amount of pages per second is greater than 1000
$memoryPagesPerSecond = Get-Counter '\Memory\Pages/sec' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($memoryPagesPerSecond -gt 1000){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Memory\Pages/sec',
            '1000',
            "$memoryPagesPerSecond",            
            'There is excessive paging. There may be a memory leak.'
        ) 
    ) | Out-Null    
}

#Check if percent of processor time executing non idle threads is greater than 85%.
$processorPercentTimeBusy = Get-Counter '\processor(_total)\% processor time' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($processorPercentTimeBusy -gt 85){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\processor(_total)\% processor time',
            '85%',
            "$processorPercentTimeBusy",            
            'The processor is overwhelmed. The system may need a faster processor.'
        ) 
    ) | Out-Null    
}

#Check if the percent of time processor spends in user mode is greater then 85%.
$processorPercentUserMode = Get-Counter '\Processor(_total)\% User Time' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($processorPercentUserMode -gt 85){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Processor(_total)\% User Time',
            '85%',
            "$processorPercentUserMode",            
            'The processor is overwhelmed with applications. Consider ending process intensive applications.'
        ) 
    ) | Out-Null    
}

#Check if the percent of time processor spends handling hardware interrupts is greater then 15%.
$processorPercentInterrupts = Get-Counter '\Processor(_total)\% Interrupt Time' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($processorPercentInterrupts -gt 15){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Processor(_total)\% Interrupt Time',
            '15%',
            "$processorPercentInterrupts",            
            'The processor has an excessive amount of hardware interruptions. There could be a hardware issue.'
        ) 
    ) | Out-Null    
}

#Check if traffic through network interface is more than 70% used.
$netInterfaceTotalBytesPerSecond = Get-WmiObject 'Win32_NetworkAdapter' -Filter {Speed != null} | Select-Object -ExpandProperty Speed
$netInterfaceBytesUsedPerSecond = Get-Counter '\Network Interface(*)\Bytes Total/sec' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if((($netInterfaceBytesUsedPerSecond/$netInterfaceTotalBytesPerSecond) * 100) -gt 70){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Network Interface(*)\Bytes Total/sec',
            '70',
            "$(($netInterfaceBytesUsedPerSecond/$netInterfaceTotalBytesPerSecond) * 100)",            
            'The network interface is oversaturated.'
        ) 
    ) | Out-Null    
}

#Check if network output queue is greater than 2 packets.
$netInterfaceOutputQueue = Get-Counter '\Network Interface(*)\Output Queue Length' -ComputerName $ComputerName | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
if($netInterfaceOutputQueue -gt 2){
    $ErrorSummary.Add(
        [TroubleshootingCheck]::new(            
            '\Network Interface(*)\Output Queue Length',
            '2',
            "$netInterfaceOutputQueue",            
            'The network interface output queue is full. The network interface is oversaturated.'
        ) 
    ) | Out-Null    
}

$ErrorSummary | Out-GridView

if($ErrorSummary.Count -eq 0){
    Write-Host 'No issues detected'
}