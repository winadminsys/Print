Get-WmiObject Win32_Printer -ComputerName EMAPPA2144DEV  | Where {$_.Name -eq "MyPrinter"} | ForEach { $_.CancelAllJobs() }
