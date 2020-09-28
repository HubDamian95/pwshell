$printserver = "printingServerDNSName" 
Get-WMIObject -class Win32_Printer -computer $printserver | Select Name,DriverName,PortName | Export-CSV -path 'D:\log.csv'