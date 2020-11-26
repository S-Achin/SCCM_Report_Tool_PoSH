#region SCCMSQLCONNECTION
$Global:SCCMSQLSERVER = "ServerName" 
$Global:DBNAME = "DataBaseName"
$ReprtsPath = "\\Server\share$\"
Try 
{ 
$SQLConnection = New-Object System.Data.SQLClient.SQLConnection 
$SQLConnection.ConnectionString ="server=$SCCMSQLSERVER;database=$DBNAME;Integrated Security=True;"
$SQLConnection.Open() 

#endregion SCCMSQLCONNECTION
#region Menu
DO{

clear   
        write-host  -ForegroundColor black -BackgroundColor White "                        SCCM Report Automation Posh Tool                        "
        write-host ""
        write-host -ForegroundColor White  "================================Menu================================="
        write-host -ForegroundColor Yellow   "1 - Please press 1 to get Hostname Against the Username."
        write-host -ForegroundColor Yellow   "2 - Please Press 2 to get Computer Hardware report."
        write-host -ForegroundColor Yellow   "3 - Please Press 3 to get IP and MAC details of Machines."
        write-host -ForegroundColor Yellow   "4 - Please Press 4 to get SCCM Client Status Report."
        write-host -ForegroundColor Yellow   "X - Please Press 'X' to go out"
        write-host -ForegroundColor white   "================================END=================================="
        write-host  -ForegroundColor black -BackgroundColor White "                                                                              "
#endregion menu

#region selection
$selection = Read-Host "Please make a selection"
$ok = $selection -match '^[1234X]+$'
if ($ok  -eq $False) 
{write-host "Invalid selection" }
if ($ok -eq $true)
{

switch ($selection)
 {
     '1' {
            $user = Read-Host " please provide the Username"
            $query = "select sys.name0,
                      sys.User_Name0
                      from v_R_System sys
                      where sys.User_Name0 in ('$user')
                      "
            $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
            $SQLCommand.CommandText = $query
            $SQLCommand.Connection = $SQLConnection
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SQLCommand                  
            $SQLDataset = New-Object System.Data.DataSet 
            $SqlAdapter.fill($SQLDataset) | out-null 
            $data = $SQLDataset.Tables[0]
            $data | Out-GridView
            Pause
                                 
         
         } 
     '2' 
         {    $TestList = Test-Path "$ReprtsPath\Hardware_report_query.csv"
               if($TestList -eq $true)
               {         
              $Date = Get-Date -Format yyyy.MM.dd-HH.mm.ss
              Rename-Item -Path "$ReprtsPath\Hardware_report_query.csv" -NewName $date-Hardware_report_query.csv -ErrorAction SilentlyContinue
              }
              else {}
                Write-Host -ForegroundColor DarkCyan "Fetching Data from SCCM server Databasa. please wait.......................... "
                     $query = "select distinct sys.name0 as 'Hostname',
                                sys.User_Name0 as 'User Name',
                                sys.Last_Logon_Timestamp0 as 'Last Login time' ,
                                sys.AD_Site_Name0 as 'AD Site',
                                sys.SMS_Unique_Identifier0 as 'GUID' ,
                                OS.Caption0 as 'Operating System',
                                CS.Model0 as 'Model',
                                CS.Manufacturer0 as 'Manufacturer',
                                CS.SystemType0 as 'Architecture',
                                CS.Domain0 as 'Domain' ,
                                PC.SerialNumber0 as 'Serial Number'
                                from v_R_System sys
                                JOIN v_GS_OPERATING_SYSTEM OS
                                on OS.ResourceID = sys.ResourceID
                                join v_GS_COMPUTER_SYSTEM CS
                                on CS.ResourceID = sys.ResourceID
                                join v_GS_PC_BIOS PC
                                on PC.ResourceID = sys.ResourceID"
            $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
            $SQLCommand.CommandText = $query
            $SQLCommand.Connection = $SQLConnection
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SQLCommand                  
            $SQLDataset = New-Object System.Data.DataSet 
            $SqlAdapter.fill($SQLDataset) | out-null 
            $SQLDataset.Tables[0] | select  Hostname , User*, Domain, Last*, AD*, GUID, Operatin*, model, Serial*, Manufacturer, Architecture | Export-Csv  "$ReprtsPath\Hardware_report_query.csv"
            Write-Host -ForegroundColor Green "Report has been saved on path $ReprtsPath\SCCM Posh Report"
         } 
     '3' 
         {    $TestList = Test-Path "$ReprtsPath\Machine_MAC_IP_Details.csv"
               if($TestList -eq $true)
               {         
              $Date = Get-Date -Format yyyy.MM.dd-HH.mm.ss
              Rename-Item -Path "$ReprtsPath\Machine_MAC_IP_Details.csv" -NewName $date-Machine_MAC_IP_Details.csv -ErrorAction SilentlyContinue
              }
              else {}
                Write-Host -ForegroundColor DarkCyan "Fetching Data from SCCM server Databasa. please wait.......................... "
                     $query = "SELECT distinct       dbo.v_R_System.Netbios_Name0, dbo.v_R_System.User_Name0, dbo.v_R_System.Last_Logon_Timestamp0, dbo.v_RA_System_IPAddresses.IP_Addresses0, 
                                                     dbo.v_RA_System_MACAddresses.MAC_Addresses0, dbo.v_RA_System_MACAddresses.ResourceID
                                       FROM          dbo.v_RA_System_MACAddresses INNER JOIN
                                                     dbo.v_RA_System_IPAddresses ON dbo.v_RA_System_MACAddresses.ResourceID = dbo.v_RA_System_IPAddresses.ResourceID INNER JOIN
                                                     dbo.v_R_System ON dbo.v_RA_System_IPAddresses.ResourceID = dbo.v_R_System.ResourceID "
                    
            $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
            $SQLCommand.CommandText = $query
            $SQLCommand.Connection = $SQLConnection
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SQLCommand                  
            $SQLDataset = New-Object System.Data.DataSet 
            $SqlAdapter.fill($SQLDataset) | out-null 
            $SQLDataset.Tables[0] | Select Netbios_Name0, User_Name0, Last_Logon_Timestamp0, IP_Addresses0, MAC_Addresses0, ResourceID | Export-Csv  "$ReprtsPath\Machine_MAC_IP_Details.csv"
            Write-Host -ForegroundColor Green "Report has been saved over $ReprtsPath\Machine_MAC_IP_Details.csv"
            
         }
      '4'
         {    $TestList = Test-Path "$ReprtsPath\Machine_MAC_IP_Details.csv"
               if($TestList -eq $true)
               {         
              $Date = Get-Date -Format yyyy.MM.dd-HH.mm.ss
              Rename-Item -Path "$ReprtsPath\Machine_MAC_IP_Details.csv" -NewName $date-Machine_MAC_IP_Details.csv -ErrorAction SilentlyContinue
              }
              else {}
                Write-Host -ForegroundColor DarkCyan "Fetching Data from SCCM server Databasa. please wait.......................... "
                     $query = "SELECT distinct       dbo.v_R_System.Netbios_Name0, dbo.v_R_System.User_Name0, dbo.v_R_System.Last_Logon_Timestamp0, dbo.v_RA_System_IPAddresses.IP_Addresses0, 
                                                     dbo.v_RA_System_MACAddresses.MAC_Addresses0, dbo.v_RA_System_MACAddresses.ResourceID
                                       FROM          dbo.v_RA_System_MACAddresses INNER JOIN
                                                     dbo.v_RA_System_IPAddresses ON dbo.v_RA_System_MACAddresses.ResourceID = dbo.v_RA_System_IPAddresses.ResourceID INNER JOIN
                                                     dbo.v_R_System ON dbo.v_RA_System_IPAddresses.ResourceID = dbo.v_R_System.ResourceID "
                    
            $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
            $SQLCommand.CommandText = $query
            $SQLCommand.Connection = $SQLConnection
            $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SQLCommand                  
            $SQLDataset = New-Object System.Data.DataSet 
            $SqlAdapter.fill($SQLDataset) | out-null 
            $SQLDataset.Tables[0] | Select Netbios_Name0, User_Name0, Last_Logon_Timestamp0, IP_Addresses0, MAC_Addresses0, ResourceID | Export-Csv  "$ReprtsPath\Machine_MAC_IP_Details.csv"
            Write-Host -ForegroundColor Green "Report has been saved over $ReprtsPath\Machine_MAC_IP_Details.csv"
         }    
     'q' 
         {
         Pause
         }
 }
 }
 
 
$Retry = Read-Host "Would you like to run the script again? Y/N" 
}
While ($Retry -eq 'Y')

Write-Host "Have a good day!"
Pause
}
Catch 
{
[System.Windows.Forms.MessageBox]::Show("Failed to connect SQL Server:") 
}

 #endregion 
 
