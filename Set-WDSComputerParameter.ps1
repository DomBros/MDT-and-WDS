Function Set-WDSComputerParameter
{
    <#
            .SYNOPSIS
            Function inserts row to ComputerSpecification.csv rebuild CustomSettings.ini

            .DESCRIPTION
            Function inserts new row to ComputerSpecification.csv if computer with specify SerialNumber doesn't exist,
            replace parameters if SerialNumber exist and rebuild CustomSettings.ini file with new parameters.

            .PARAMETER Ticket
            Specify Ticket in which note will be placed 'Computer re-formated by WDS'.

            .PARAMETER ServiceTag
            Computer ServiceTag.

            .PARAMETER ComputerName
            Computer new name.

            .PARAMETER WindowsSerialNumber
            Windows activation key.

            .PARAMETER ComputerOU
            Specify computer OU if you like to, if not it will be calculated.

            .PARAMETER UserLogin
            Specify UserLogin (like Test\test) who will become local admin.

            .PARAMETER UserPSWD
            User temporary password to set profile things.

            .PARAMETER ForUK
            Add computer to UK OUs in AD.

            .PARAMETER ComputerParameterCSV
            Path to the computers configuration CSV file.

            .PARAMETER CustomSettingsINI
            Path to the INI WDS configuration file.

            .PARAMETER Show
            Switch to mark that you want Out-GridView of new data to pop-up at the end to ensure that everthing is OK

            .EXAMPLE
            Set-WDSComputerParameter -ServiceTag 4HLJ9K -ComputerName NBTEST

            .EXAMPLE
            Set-WDSComputerParameter -ServiceTag 4HLJ9K -ComputerName NBTEST -WindowsSerialNumber 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX'

            .EXAMPLE
            Set-WDSComputerParameter -Ticket 'IISTS-9999' -ServiceTag 4HLJ9K -ComputerName NBTEST -UserLogin 'Test\jnowak'
    #>


    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([void])]

    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'IISTS-')]
        [ValidateLength(7,12)]
        [string]
        $Ticket,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateLength(2,49)]
        [string]
        $ServiceTag,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateLength(2,15)]
        $ComputerName,

        [Parameter(Mandatory = $false, Position = 3, HelpMessage = 'AD login starting with Test\')]
        [string]
        $UserLogin,

        [Parameter(Mandatory = $false, Position = 4)]
        [string]
        $WindowsSerialNumber,

        [Parameter(Mandatory = $false, Position = 5, HelpMessage = 'If empty default will be calculated')]
        [string]
        $ComputerOU,

        [Parameter(Mandatory = $false, Position = 6)]
        [string]
        $UserPSWD,
        
        [Parameter(Mandatory = $false, Position = 7, HelpMessage = 'Use it to add COmputer in UK OU')]
        [switch]
        $ForUK,

        [Parameter(Mandatory = $false, Position = 10)]
        [ValidateScript({
                    Test-Path $_
        })]
        [string]
        $ComputerParameterCSV = '\\DeployServer\deploymentshare$\Control\ComputerSpecification.csv',
        
        [Parameter(Mandatory = $false, Position = 11)]
        [ValidateScript({
                    Test-Path $_
        })]
        [string]
        $CustomSettingsINI = '\\DeployServer\deploymentshare$\Control\CustomSettings.ini',

        [Parameter(Mandatory = $false, Position = 12, HelpMessage = 'Show INI and CSV content')]
        [switch]
        $Show

    )
    Try
    {   
        # Preparing
        $CustomSettingsDir = Split-Path $CustomSettingsINI
        $ComputerParameterDir = Split-Path $ComputerParameterCSV
        $CustomSettingsBase = Get-Content -Path "$CustomSettingsDir\CustomSettings.bak" -ReadCount 0
        $CustomSettingsINInewtemp = "$CustomSettingsDir\CustomSettings.new"
        $CustomSettingsBlock = "$CustomSettingsDir\CustomSettings.block"
        $CsvDelimiter = ';'
        [string]$Date = Get-Date -UFormat '%Y%m%d%H%M'

        # Create Backup folder
        $BackupFolder = $ComputerParameterDir + '\Arch'
        if(-not (Test-Path $BackupFolder))
        {
            $null = New-Item -Path $ComputerParameterDir -Name 'Arch' -ItemType Directory
        }

        While(Test-Path -Path $CustomSettingsBlock)
        {
            $Waiting = 5
            Write-Output "Someone is also editing WDS parameters file, waiting $Waiting seconds ..."
            Start-Sleep -Seconds $Waiting
        } 

        # Blocking file CustomSettings.ini and ComputerSpecification.csv
        Write-Output 'Blocking CustomSettings.ini'
        $null = New-Item -Path $CustomSettingsBlock -ItemType File
        
        $Watch = New-Object -TypeName System.Diagnostics.StopWatch
        $Watch.Start() 

        # import the CSV file
        $Params = Import-Csv -Path $ComputerParameterCSV -Delimiter $CsvDelimiter
        if($Params[0].ServiceTag -eq 'ServiceTag')
        {
            $Params = $Params[1..($Params.count - 1)]
        }

        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Backing up old ComputerSpecification.csv ($Took s)"
        Copy-Item -Path $ComputerParameterCSV -Destination "$BackupFolder\$Date-ComputerSpecification.csv" -Force
        
        # for each row
        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Creating new ComputerSpecification.csv ($Took s)"
        if($Params.ServiceTag -contains $ServiceTag)
        {
            $Output = foreach($row in $Params)
            {
                # replace computer parameters if ServiceTag exist
                if($row.ServiceTag -eq $ServiceTag)
                {
                    if($WindowsSerialNumber)
                    {
                        $row.Serial = $WindowsSerialNumber.ToUpper()
                    }
                    
                    if($ComputerName)
                    {
                        $row.ComputerName = $ComputerName.ToUpper()
                    }
                    
                    if($UserLogin)
                    {
                        $row.UserLogin = $UserLogin.ToUpper()
                    }
                    
                    if($UserPSWD)
                    {
                        $row.UserPSWD = $UserPSWD
                    }
                    
                    if($Ticket)
                    {
                        $row.Ticket = $Ticket.ToUpper()
                    }

                    if($ComputerOU)
                    {
                        $row.ComputerOU = $ComputerOU
                    }

                    if(-not $ComputerOU)
                    {
                        if($row.ComputerName -like 'PL*')
                        {
                            $row.ComputerOU = 'OU=Desktops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'NB*')
                        {
                            $row.ComputerOU = 'OU=Laptops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'v*')
                        {
                            $row.ComputerOU = 'OU=VMs,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'UK*')
                        {
                            $row.ComputerOU = 'OU=Desktops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'NB*' -and $ForUK)
                        {
                            $row.ComputerOU = 'OU=Laptops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'v*' -and $ForUK)
                        {
                            $row.ComputerOU = 'OU=VMs,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                    }
                    $row
                }
                else
                {
                    # rebuild existing parameters
                    $row.ServiceTag = ($row.ServiceTag).ToUpper()
                    $row.Serial = ($row.Serial).ToUpper()
                    $row.ComputerName = ($row.ComputerName).ToUpper()
                    $row.UserLogin = ($row.UserLogin).ToUpper()
                    $row.UserPSWD = $row.UserPSWD
                    $row.Ticket = ($row.Ticket).ToUpper()
                    if($row.ComputerOU)
                    {
                        $row.ComputerOU = $row.ComputerOU
                    }
                    if(-not $row.ComputerOU)
                    {
                        if($row.ComputerName -like 'PL*')
                        {
                            $row.ComputerOU = 'OU=Desktops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'NB*')
                        {
                            $row.ComputerOU = 'OU=Laptops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'v*')
                        {
                            $row.ComputerOU = 'OU=VMs,OU=ComputersPL,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'UK*')
                        {
                            $row.ComputerOU = 'OU=Desktops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'NB*' -and $ForUK)
                        {
                            $row.ComputerOU = 'OU=Laptops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                        if($row.ComputerName -like 'v*' -and $ForUK)
                        {
                            $row.ComputerOU = 'OU=VMs,OU=ComputersUK,DC=test,DC=org,DC=pl'
                        }
                    }
                    $row
                }
            }
            $Params = $Output
        }
        else
        {
            # add new row if computer ServiceTag doesn't exist
            $ServiceTag = $ServiceTag.ToUpper()
            $WindowsSerialNumber = $WindowsSerialNumber.ToUpper()
            $ComputerName = $ComputerName.ToUpper()
            $Ticket = $Ticket.ToUpper()
            $UserLogin = $UserLogin.ToUpper()
            $UserPSWD = $UserPSWD
            if($ComputerOU)
            {
                $ComputerOU = $ComputerOU
            }
            if(-not $ComputerOU)
            {
                if($ComputerName -like 'PL*')
                {
                    $ComputerOU = 'OU=Desktops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                }
                if($ComputerName -like 'NB*')
                {
                    $ComputerOU = 'OU=Laptops,OU=ComputersPL,DC=test,DC=org,DC=pl'
                }
                if($ComputerName -like 'v*')
                {
                    $ComputerOU = 'OU=VMs,OU=ComputersPL,DC=test,DC=org,DC=pl'
                }
                if($ComputerName -like 'UK*')
                {
                    $ComputerOU = 'OU=Desktops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                }
                if($ComputerName -like 'NB*' -and $ForUK)
                {
                    $ComputerOU = 'OU=Laptops,OU=ComputersUK,DC=test,DC=org,DC=pl'
                }
                if($ComputerName -like 'v*' -and $ForUK)
                {
                    $ComputerOU = 'OU=VMs,OU=ComputersUK,DC=test,DC=org,DC=pl'
                }
            }
            
            $newRow = New-Object -TypeName PSCustomObject -Property @{
                ServiceTag   = $ServiceTag
                Serial       = $WindowsSerialNumber
                ComputerName = $ComputerName
                UserLogin    = $UserLogin
                UserPSWD     = $UserPSWD
                ComputerOU   = $ComputerOU
                Ticket       = $Ticket
            }
            $Params += $newRow
        }

        $Params | Export-Csv -Path $ComputerParameterCSV -Delimiter ';' -NoTypeInformation

        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Creating new values for CustomSettings.ini ($Took s)"

# build CustomSettings depends on computers SerialNumber
        $SectionsAll = @"
"@

        foreach($item in $Params)
        {
            if($item.ServiceTag)
            {
                if($item.Serial)
                {
                    $PRODUCTKEY = @"
PRODUCTKEY=$($item.Serial)
SKIPTASKSEQUENCE=NO
TASKSEQUENCEID=WIN10FROM7
SKIPDEPLOYMENTTYPE=YES
DEPLOYMENTTYPE=NEWCOMPUTER
"@
                }
                else
                {
                    $PRODUCTKEY = @"
SKIPTASKSEQUENCE=NO
TASKSEQUENCEID=WIN10FROM8
SKIPDEPLOYMENTTYPE=YES
DEPLOYMENTTYPE=NEWCOMPUTER
"@
                }

                if($item.ComputerName)
                {
                    $OSDCOMPUTERNAME = @"
OSDCOMPUTERNAME=$($item.ComputerName)
"@
                }
                else
                {
                    $OSDCOMPUTERNAME = $null
                }

                if($item.UserLogin)
                {
                    $ADMINISTRATORS002 = @"
ADMINISTRATORS002=$($item.UserLogin)
"@
                }
                else
                {
                    $ADMINISTRATORS002 = $null
                }

                if($item.ComputerOU)
                {
                    $MACHINEOBJECTOU = @"
MACHINEOBJECTOU=$($item.ComputerOU)
"@
                }
                else
                {
                    $MACHINEOBJECTOU = $null
                }

                $Section = @"
[$($item.ServiceTag)]
;$item
$OSDCOMPUTERNAME
$PRODUCTKEY
$MACHINEOBJECTOU
$ADMINISTRATORS002


"@

                $SectionsAll += $Section
            }
        }
        
        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Creating new temp CustomSettings.new ($Took s)"
        $CustomSettingsBase | Out-File -FilePath $CustomSettingsINInewtemp -Append -Encoding ascii
        $SectionsAll | Out-File -FilePath $CustomSettingsINInewtemp -Append -Encoding ascii

        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Overwriting old CustomSettings.ini ($Took s)"
        
        $TookA = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Move-Item -Path $CustomSettingsINInewtemp -Destination $CustomSettingsINI -Force
        $TookB = '{0:N2}' -f $Watch.Elapsed.TotalSeconds

        Write-Output "Operation start $TookA unavailable till $TookB End"

        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        Write-Output "Unblocking CustomSettings.ini ($Took s)"
        Remove-Item -Path $CustomSettingsBlock -Force -ErrorAction SilentlyContinue
        
        $Watch.Stop()
        $Took = '{0:N2}' -f $Watch.Elapsed.TotalSeconds
        
        if($Show)
        {
            $csv_objects = Import-Csv -Path $ComputerParameterCSV -Delimiter $CsvDelimiter
            $csv_objects | Out-GridView
            $ini_content = Get-Content -Path $CustomSettingsINI -ReadCount 0
            $ini_content | Out-GridView
            $csv_content = Get-Content -Path $ComputerParameterCSV -ReadCount 0
            $csv_content | Out-GridView
        }
    }
    Catch
    {
        # cleaning up after error
        Write-Log -Error -Message 'Deleting CustomSettings.new'
        Remove-Item -Path $CustomSettingsINInewtemp -Force -ErrorAction SilentlyContinue

        Write-Log -Error -Message 'Unblocking CustomSettings.ini'
        Remove-Item -Path $CustomSettingsBlock -Force -ErrorAction SilentlyContinue
        
        $Error
    }
    Finally
    {
        # ending
        Write-Output "Operation took: $Took seconds."
    }
}
