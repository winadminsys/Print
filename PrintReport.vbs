Import-Module -Name ServerManager
#Import-module "PrintManagement"

# Report Name
$FileName = "C:\Temp\PrintReport.htm"

# Server List
$PrintServers = @("myprintserver1","myprintserver2")

# Variables
$StatsData =  @()
[string]$bodycloud = ""
$Treshold = 100

# Print Jobs by servers  
Foreach ($PrintServer in $PrintServers)
{
    $Stats = @()
    $Stats = Get-WMIObject Win32_PerfFormattedData_Spooler_PrintQueue -ComputerName $PrintServer | Where {$_.Name -like "Impression_Gen_*"}
    $StatsData  += $Stats

    $bodycloud += "<h2>Print Jobs for : $PrintServer</h2>"
    $bodycloud += "<table border=0>`n"
    $bodycloud += "<tr><th>ID</th> `
    <th>Name</th> `
    <th>Jobs in the queue</th> `
    <th>TotalJobsPrinted</th> `
    <th>TotalPagesPrinted</th> `
    <th>MaxJobsSpooling</th> `
    <th>JobErrors</th> `
    <th>Server</th></tr>"    
    
    $ID = 1
    Foreach ($item in $Stats)
    {   
        $LineColor = ""
        If ( ($item.jobs -gt $Treshold) -or ($item.JobErrors -gt $Treshold) )
        {
            $LineColor = "red"
        }
        Else
        {
            $LineColor = "#00CC66"
        }
        
        $bodycloud += "<tr style='background-color:$LineColor;'>"
        $bodycloud += "<td>" + $ID++ `
        + "</td><td>" + $item.Name + "</td>" `
        + "</td><td>" + $item.jobs + "</td>" `
        + "</td><td>" + $item.TotalJobsPrinted + "</td>" `
        + "</td><td>" + $item.TotalPagesPrinted + "</td>" `
        + "</td><td>" + $item.MaxJobsSpooling + "</td>" `
        + "</td><td>" + $item.JobErrors + "</td>" `
        + "</td><td>" + $item.__SERVER + "</td>"
        $bodycloud += "</tr>"
    } 
    $bodycloud += "</table>`n"   
}

# Print Jobs Global 
$bodycloud += "<h2>Global Print Jobs</h2>"
$bodycloud += "<table border=0>`n"
$bodycloud += "<tr><th>ID</th> `
<th>Name</th> `
<th>Jobs in the queue</th> `
<th>TotalJobsPrinted</th> `
<th>TotalPagesPrinted</th> `
<th>MaxJobsSpooling</th> `
<th>JobErrors</th> `
<th>Server</th></tr>"  

$ID = 1
Foreach ($item in $StatsData)
{
    $LineColor = ""
    If ( ($item.jobs -gt $Treshold) -or ($item.JobErrors -gt $Treshold) )
    {
        $LineColor = "red"
    }
    Else
    {
        $LineColor = "#00CC66"
    }
        
    $bodycloud += "<tr style='background-color:$LineColor;'>"
    $bodycloud += "<td>" + $ID++ `
    + "</td><td>" + $item.Name + "</td>" `
    + "</td><td>" + $item.jobs + "</td>" `
    + "</td><td>" + $item.TotalJobsPrinted + "</td>" `
    + "</td><td>" + $item.TotalPagesPrinted + "</td>" `
    + "</td><td>" + $item.MaxJobsSpooling + "</td>" `
    + "</td><td>" + $item.JobErrors + "</td>" `
    + "</td><td>" + $item.__SERVER + "</td>"
    $bodycloud += "</tr>"
}
$bodycloud += "</table>`n"

# Generate html report
[string]$emailbody = @("
<html>
<head>
<title>Print Overview : $(Get-Date)</title>
	<STYLE type='text/css'>
		h1 {font-family:Calibri; font-size:25} 
		h2 {font-family:Calibri; font-size:20}

	table, th, td {
		border: 1px solid #000000;
		font-family:Calibri; font-size:13;
		text-align: left;
		border-collapse: collapse; 
	}

	th {
		font-weight: bold;
		background-color: #acf;
	}

	td,th {
		padding: 4px 5px; 
	}

	.odd {
		background-color: #def; 
	}

	.odd td {
		border-bottom: 1px solid #cef; 
	}
	</STYLE>
</head>
<body>
<p>
<center><h1>Print Overview : $(Get-Date)</h1></center>
</p>
<br />
$bodycloud
</body>
</html>
") 

# Generate html file
$emailbody | Out-File $FileName

# Open File
Invoke-Item $FileName

# Get-WmiObject win32_printer -ComputerName $PrintServer | Where {$_.Name -like "Impression_Gen_*"}
# Get-WmiObject Win32_TcpIpPrinterPort -ComputerName $PrintServer | Where {$_.Name -like "Impression_Gen_*"}
