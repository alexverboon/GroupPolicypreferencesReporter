<#
.Synopsis
   Backup-GPO
.DESCRIPTION
   This script creates a backup of all  Group Policy objects within the Active Directory domain and stores the 
   backup and the GPO report in the specified output folder
.EXAMPLE
  Backup-GPO -Path
.NOTES
    v1.0, 26.11.2020, Alex Verboon
#>
    [CmdletBinding()]
    Param
    (
        # Backup Path
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $GPOBackupPath="C:\DATA\GPOBackup"
    )

Begin{
    Function Get-TimeStamp{
    Get-Date -Format "ddMMyyyy:HHmmss"
}

}
Process{
  
    $DateTimeStamp = (Get-Date -Format "ddMMyyyy_HHmm")
    Write-output "Timestamp: $DateTimeStamp"

    # GPO Backup Path
    $GPOBackupExport  = "$GPOBackupPath\Backup"
    # GPO Report Path
    $GPOReportPath    = "$GPOBackupPath\Reports"
    # GPO Backup logs
    $GPOLogs           = "$GPOBackupPath\Logs"

    If (-not (Test-Path -Path "$GPOBackupPath" -PathType Container )){
        New-Item -Path "$GPOBackupPath" -ItemType "Directory"
        New-Item -Path "$GPOBackupPath\Backup" -ItemType "Directory"
        New-Item -Path "$GPOBackupPath\Reports" -ItemType "Directory"
        New-Item -Path "$GPOBackupPath\Logs" -ItemType "Directory"
    }
    Else{
        If (-not (Test-Path -Path "$GPOBackupExport" -PathType Container )){
             New-Item -Path "$GPOBackupExport" -ItemType "Directory"
        }
        If (-not (Test-Path -Path "$GPOReportPath" -PathType Container )){
             New-Item -Path "$GPOReportPath" -ItemType "Directory"
        }
        If (-not (Test-Path -Path "$GPOLogs" -PathType Container )){
             New-Item -Path "$GPOBackupPath\Logs" -ItemType "Directory"
        }
    }

    $runLogFile = "$GPOLogs\GPOBackup_$DateTimeStamp.LOG"

    Write-output "GPO Backup Path: $GPOBackupPath"
    Write-Output "GPO Export Path: $GPOBackupExport"
    Write-Output "GPO Reports: $GPOReportPath"
    Write-Output "Log Directory: $GPOLogs"

    Write-Output "$(Get-TimeStamp) Starting Group Policy Backup" | out-file $runLogFile -Append

    Try{
        # Create daily GPO backup folder
        $GPOBackupExportToday = "$GPOBackupExport\$DateTimeStamp"
        New-Item -Path "$GPOBackupExport\$DateTimeStamp" -ItemType "Directory" | Out-Null
    }Catch{
        Write-Output "$(Get-TimeStamp) Error creating daily backup folder" | out-file $runLogFile -Append
        Write-Error "Error creating daily backup folder"
        
    }

    Try{
        # Create Daily Report backup folder
        $GPOReportPathToday = "$GPOReportPath\$DateTimeStamp"
        New-Item -Path "$GPOReportPath\$DateTimeStamp" -ItemType "Directory" | Out-Null
    }Catch{
        Write-Output "$(Get-TimeStamp) Error creating daily report folder" | out-file $runLogFile -Append
        Write-Error "Error creating daily report backup folder"
    }

    Write-output "Daily GPO Backup Folder: $GPOBackupExportToday"
    Write-output "Daily GPO Report Backup folder: $GPOReportPathToday"

    Write-Output "Retrieving all GPO objects"
    # Gather all GPOs
    $AllGPO = Get-GPO -All 
    #$AllGPO = Get-GPO -All | Where-Object {$_.DisplayName -like "Workplace*"}
    Write-Output "Total GPOs found: $($AllGPO.count)"

    # Generate a Report and a Backup of each GPO Object
    ForEach ($gpo in $AllGPO)
    {
        Write-Output "Generating GPO Report for $($gpo.DisplayName)"
        Get-GPOReport -Name "$($gpo.DisplayName)" -ReportType Html -Path "$GPOReportPathToday\$($gpo.DisplayName).html" 
        Write-Output "Generating Backup for $($gpo.DisplayName)"
        If (Test-path -Path "$GPOBackupExportToday\$($gpo.DisplayName)")
        {
            # no action required
        }
        Else{
            Try{
            New-Item -Path  "$GPOBackupExportToday\$($gpo.DisplayName)" -ItemType "directory" | Out-Null
            }
            Catch{
                 Write-Output "$(Get-TimeStamp) Error creating backup folder for $($gpo.DisplayName)" | out-file $runLogFile -Append
            }
        }

        Try{
        Backup-GPO -Name "$($gpo.DisplayName)" -Path "$GPOBackupExportToday\$($gpo.DisplayName)" -Comment "$($gpo.DisplayName)"
        }
        Catch{
                Write-Output "$(Get-TimeStamp) Error with GPO Backup $($gpo.DisplayName)" | out-file $runLogFile -Append
        }
    }
        Write-Output "$(Get-TimeStamp) Group Policy Backup Completed" | out-file $runLogFile -Append
}
End{}
