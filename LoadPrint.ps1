Param(
  
  [parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [ValidateRange(1,2000)]
  [int]$NumberOfJobs,

  [parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [ValidateSet("Parallel","Sequential")]
  [string]$Mode,

  [parameter(Mandatory=$true)]
  [ValidateNotNullOrEmpty()]
  [ValidateSet("SimpleText","CustomFile")]
  [string]$PrintMethod
  
)

# Create folder if not exist
$CustomFolderDirectory = "c:\Temp\"
$CustomFolderName = "CustomFiles"
$CustomFolderPath = $CustomFolderDirectory+$CustomFolderName
$PrinterName = "\\myserver\myprinter"  

If (!(Test-Path $CustomFolderPath))
{
    New-Item -ItemType Directory -Name $CustomFolderName -Path $CustomFolderDirectory -Force | out-null
}

Switch ($Mode) 
{
    ###################################
    # Start all jobs in parrallel mode 
    ###################################
    "Parallel" 
    {
        # Launch Process
        For($i=1; $i -le $NumberOfJobs; $i++)
        {

            # Print random file
            If ($PrintMethod -eq "CustomFile")
            {
                If ((Get-ChildItem "$CustomFolderPath").count -ne 0) 
                {

                    $RandomFile = Get-ChildItem "$CustomFolderPath" | Get-Random
                     
                    $RandomFileCalcSize = ($RandomFile | Measure-Object -property length -sum) 
                    $RandomFileSize = "{0:N2}" -f ($RandomFileCalcSize.sum / 1MB) + " MB"   

                    $ScriptBlock = 
                    {
                           param
                            (
                                $RandomFile,
                                $Content
                            )     
                            
                            get-content $RandomFile.FullName | Out-Printer $PrinterName      
                    }

                    write-host "$(get-date) --- $i --- Printing $RandomFile on $PrinterName --- size $RandomFileSize"
                    Start-Job -scriptblock $ScriptBlock -ArgumentList $RandomFile, $Content  | Out-Null

                }
                Else
                {
                    write-host "$CustomFolderPath folder is empty ! Please insert files !" -ForegroundColor Red
                    break
                }
            }

            # Print Simple File
            If ($PrintMethod -eq "SimpleText")
            {
               $ScriptBlock =
                {
                    "Hello, World" | out-printer $PrinterName | Out-Null
            
                    # Other Print File method
                    #Start-Process -FilePath $RandomFile.FullName –Verb Print -PassThru| %{sleep 10;$_} | kill
                }

                write-host "$(get-date) --- $i --- Printing SimpleText on $PrinterName"
                Start-Job -scriptblock $ScriptBlock | Out-Null
            }
               

        } 

        $JobList = Get-Job | Wait-Job
        $JobList
        $JobList | Remove-job

    }

    ###################################
    # Start job one by one
    ###################################
    "Sequential" 
    {
        # Print random file
        # Launch Process
        For($i=1; $i -le $NumberOfJobs; $i++)
        {
            If ($PrintMethod -eq "CustomFile")
            {
                If ((Get-ChildItem "$CustomFolderPath").count -ne 0) 
                {
                    #Select and print random file in the folder

                    $RandomFile = Get-ChildItem "$CustomFolderPath" | Get-Random
                    $Content = get-content $RandomFile.FullName
                     
                    $RandomFileCalcSize = ($RandomFile | Measure-Object -property length -sum) 
                    $RandomFileSize = "{0:N2}" -f ($RandomFileCalcSize.sum / 1MB) + " MB"   

                    write-host "$(get-date) --- $i --- Printing $RandomFile on $PrinterName --- size $RandomFileSize"

                    $Content | Out-Printer $PrinterName 

                    #get-content -path $RandomFile.FullName | Out-Printer "\\myserver\myprinter"
                }
                Else
                {
                    write-host "$CustomFolderPath folder is empty ! Please insert files !" -ForegroundColor Red
                    break
                }
            }

            #Print Simple File
            If ($PrintMethod -eq "SimpleText")
            {
                write-host "$(get-date) --- $i --- Printing SimpleText on $PrinterName"
                "Hello, World" | out-printer $PrinterName | Out-Null

                Start-Sleep -Seconds 2

                #Other Print File method
                #write-host "$i --- Printing on \\myserver\myprinter"
                #Start-Process -FilePath $RandomFile.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill 
            }              
            
        }

    }
}

