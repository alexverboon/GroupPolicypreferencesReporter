<#
.SYNOPSIS
    Analyzes Group Policy Preferences (GPP) from GPO backup files and generates an interactive HTML report.

.DESCRIPTION
    This script parses Group Policy Preferences XML files from a GPO backup location, extracts preference data
    from 18 different preference types (Drive Mappings, Registry Settings, Printers, Shortcuts, Local Groups,
    Local Users, Files, Folders, INI Files, Network Shares, Data Sources, Devices, Environment Variables,
    Internet Settings, Power Options, Folder Options, Services, and Scheduled Tasks), and generates a
    comprehensive interactive HTML report with search and filter capabilities.

    The report includes:
    - Overview statistics (total domains, GPOs, preference records)
    - Preference type distribution (categorized by Windows Settings and Control Panel Settings)
    - Detailed collapsible sections for each preference type with search filters
    - Optional CSV export for each preference type

.PARAMETER GPOBackupRoot
    The root directory path containing the GPO backup files with sysvol structure.
    Default: "C:\Temp\GPOBackup"

.PARAMETER OutputDir
    The output directory where HTML report and optional CSV files will be saved.
    Default: "c:\Temp\GPOBackup\GPP-Report"

.PARAMETER ExportCSV
    Switch parameter to enable CSV export for each preference type.
    If specified, CSV files will be generated alongside the HTML report.
    Default: $false (CSV export disabled)

.PARAMETER ExportJSON
    Switch parameter to enable JSON export of all collected data.
    If specified, a JSON file will be generated with all preference data.
    Default: $false (JSON export disabled)

.EXAMPLE
    # Generate HTML report only (default)
    .\Create-GPPReport.ps1

.EXAMPLE
    # Generate HTML report with CSV exports
    .\Create-GPPReport.ps1 -ExportCSV

.EXAMPLE
    # Specify custom backup and output directories
    .\Create-GPPReport.ps1 -GPOBackupRoot "D:\GPOBackup" -OutputDir "C:\Reports"

.EXAMPLE
    # Generate report with CSV exports to custom location
    .\Create-GPPReport.ps1 -GPOBackupRoot "D:\GPOBackup" -OutputDir "C:\Reports" -ExportCSV

.EXAMPLE
    # Generate report with JSON export
    .\Create-GPPReport.ps1 -ExportJSON

.EXAMPLE
    # Generate report with both CSV and JSON exports
    .\Create-GPPReport.ps1 -ExportCSV -ExportJSON

.NOTES
    Author:         Alex Verboon
    Created:        2026-01-17
    Last Modified:  2026-01-17
    Version:        1.1
    
    Requirements:
    - PowerShell 5.1 or higher
    - Read access to GPO backup folder structure
    - Write access to output directory
    
    Output Files:
    - HTML Report: Group Policy Preferences Report_YYYYMMDD_HHMMSS.html
    - CSV Files (if -ExportCSV): Multiple CSV files for each preference type
    - JSON File (if -ExportJSON): GPP_Complete_Data_YYYYMMDD_HHMMSS.json
    
    License:        MIT
    Repository:     https://github.com/averboon/GroupPolicypreferencesReporter

.LINK
    https://docs.microsoft.com/en-us/previous-versions/windows/desktop/grouppolicy/group-policy-preferences

#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Root directory containing GPO backup files")]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$GPOBackupRoot = "C:\\Temp\\GPOBackup\\GPoBackup2",
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for reports")]
    [string]$OutputDir = "c:\\Temp\\GPOBackup\\GPP-Report",
    
    [Parameter(Mandatory=$false, HelpMessage="Export CSV files for each preference type")]
    [switch]$ExportCSV = $false,
    
    [Parameter(Mandatory=$false, HelpMessage="Export JSON file with all data")]
    [switch]$ExportJSON = $false
)

#region Initialization
$ErrorActionPreference = 'Stop'
$script:ErrorCount = 0
$script:WarningCount = 0

Write-Host "Initializing GPP Report Generator..." -ForegroundColor Cyan

try {
    # Validate GPO backup root
    if (-not (Test-Path $GPOBackupRoot)) {
        throw "Backup root not found: $GPOBackupRoot"
    }
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputDir)) {
        Write-Host "Creating output directory: $OutputDir" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $OutputDir -ErrorAction Stop | Out-Null
    }
    
    Write-Host "GPO Backup Root: $GPOBackupRoot" -ForegroundColor Green
    Write-Host "Output Directory: $OutputDir" -ForegroundColor Green
}
catch {
    Write-Error "Initialization failed: $($_.Exception.Message)"
    exit 1
}
#endregion

#region Data Structures and Helper Functions

$results = [ordered]@{
    Drives = New-Object System.Collections.Generic.List[object]
    Registry = New-Object System.Collections.Generic.List[object]
    Printers = New-Object System.Collections.Generic.List[object]
    Shortcuts = New-Object System.Collections.Generic.List[object]
    Groups = New-Object System.Collections.Generic.List[object]
    GroupsDetail = New-Object System.Collections.Generic.List[object]
    Users = New-Object System.Collections.Generic.List[object]
    Files = New-Object System.Collections.Generic.List[object]
    Folders = New-Object System.Collections.Generic.List[object]
    FoldersACL = New-Object System.Collections.Generic.List[object]
    IniFiles = New-Object System.Collections.Generic.List[object]
    NetworkShares = New-Object System.Collections.Generic.List[object]
    DataSources = New-Object System.Collections.Generic.List[object]
    Devices = New-Object System.Collections.Generic.List[object]
    EnvironmentVariables = New-Object System.Collections.Generic.List[object]
    InternetSettings = New-Object System.Collections.Generic.List[object]
    PowerOptions = New-Object System.Collections.Generic.List[object]
    FolderOptions = New-Object System.Collections.Generic.List[object]
    Services = New-Object System.Collections.Generic.List[object]
    ScheduledTasks = New-Object System.Collections.Generic.List[object]
    OtherTypes = New-Object System.Collections.Generic.List[object]
}

function Get-GpoMetadata([string]$preferenceFilePath) {
    # The path structure is: .../{GPO-GUID}/DomainSysvol/GPO/{User|Machine}/Preferences/{Type}/
    # We need to find gpreport.xml which is at .../{GPO-GUID}/gpreport.xml
    
    $reportFile = $null
    $pathParts = $preferenceFilePath -split "[\\\/]"
    
    # Find the DomainSysvol index and go back one level to the GUID folder
    $dsIndex = [array]::IndexOf($pathParts, "DomainSysvol")
    if ($dsIndex -gt 0) {
        $gpoGuidFolder = $pathParts[$dsIndex - 1]
        # Check if previous folder is GUID-like
        if ($gpoGuidFolder -match "^\{[A-F0-9\-]{36}\}$") {
            # Reconstruct path up to GUID folder
            $pathToGuid = $pathParts[0..($dsIndex - 2)] -join "\"
            $reportFile = Join-Path "$pathToGuid\$gpoGuidFolder" "gpreport.xml"
        }
    }
    
    # Fallback: walk up the tree
    if (-not $reportFile -or -not (Test-Path $reportFile)) {
        $current = Split-Path $preferenceFilePath
        $maxDepth = 10
        $depth = 0
        while ($current -and $current.Length -gt 3 -and $depth -lt $maxDepth) {
            $test = Join-Path $current "gpreport.xml"
            if (Test-Path $test) {
                $reportFile = $test
                break
            }
            $current = Split-Path $current
            $depth++
        }
    }
    
    $domain = "UnknownDomain"
    $gpoName = "UnknownGPO"
    
    if ($reportFile -and (Test-Path $reportFile)) {
        try {
            [xml]$xml = Get-Content -Path $reportFile -Raw -ErrorAction Stop
            
            # Extract Name from DocumentElement's direct Name child
            foreach ($child in $xml.DocumentElement.ChildNodes) {
                if ($child.LocalName -eq 'Name') {
                    $gpoName = $child.InnerText.Trim()
                    break
                }
            }
            
            # Extract Domain from Identifier/Domain (with namespace handling)
            foreach ($child in $xml.DocumentElement.ChildNodes) {
                if ($child.LocalName -eq 'Identifier') {
                    foreach ($sub in $child.ChildNodes) {
                        if ($sub.LocalName -eq 'Domain') {
                            $domain = $sub.InnerText.Trim()
                            break
                        }
                    }
                    break
                }
            }
        } catch {
            $script:WarningCount++
            Write-Verbose "Failed to parse gpreport.xml at $reportFile : $($_.Exception.Message)"
            # Continue with UnknownDomain/UnknownGPO defaults
        }
    }
    
    return @{ Domain = $domain; GPO = $gpoName }
}

function Add-Item([string]$type, [hashtable]$data) {
    if (-not $results.Contains($type)) { return }
    $results[$type].Add([pscustomobject]$data)
}

function Show-ProgressBar($current, $total, $label) {
    $percent = [int](($current / $total) * 100)
    $filled = [int]($percent / 5)
    $empty = 20 - $filled
    $bar = "[" + ("█" * $filled) + ("░" * $empty) + "]"
    Write-Host "`r$label $bar $percent% ($current/$total)" -NoNewline -ForegroundColor Cyan
}

function EncodeHtml([string]$s) {
    if ($null -eq $s) { return '' }
    return $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
}

# Security check helpers (centralized for clarity)
function Test-CpasswordPresent($node, $props) {
    # Check cpassword attribute on node or its Properties element
    if ($node -and $node.PSObject.Properties['cpassword']) {
        $val = $node.PSObject.Properties['cpassword'].Value
        if (-not [string]::IsNullOrEmpty($val)) { return $true }
    }
    if ($props -and $props.PSObject.Properties['cpassword']) {
        $val = $props.PSObject.Properties['cpassword'].Value
        if (-not [string]::IsNullOrEmpty($val)) { return $true }
    }
    if ($node -and $node.Attributes -and $node.Attributes['cpassword']) {
        $val = $node.Attributes['cpassword'].Value
        if (-not [string]::IsNullOrEmpty($val)) { return $true }
    }
    if ($props -and $props.Attributes -and $props.Attributes['cpassword']) {
        $val = $props.Attributes['cpassword'].Value
        if (-not [string]::IsNullOrEmpty($val)) { return $true }
    }
    return $false
}

function Get-DriveSecurityFinding($driveNode, $props) {
    if (Test-CpasswordPresent $driveNode $props) {
        return 'Drive mapping with stored credentials - Password in SYSVOL (AES-encrypted)'
    }
    return ''
}

function Get-RegistrySecurityFinding($props) {
    if (-not $props) { return '' }
    $regKey = ($props.key | ForEach-Object { $_.ToLower() })
    $regName = ($props.name | ForEach-Object { $_.ToLower() })
    # DefaultPassword and DefaultUserName are merged at risky config aggregation level
    if ($regKey -match 'winlogon' -and ($regName -eq 'defaultpassword' -or $regName -eq 'defaultusername')) {
        return 'Insecure: Plaintext credentials in registry'
    }
    if ($regKey -match 'firewall' -and ($regName -match 'enabled' -or $regName -match 'firewall')) {
        if ($props.value -eq '0') { return 'Critical: Windows Firewall disabled' }
    }
    if ($regKey -match 'uac' -and $regName -match 'enableuac') {
        if ($props.value -eq '0') { return 'Critical: User Account Control (UAC) disabled' }
    }
    return ''
}

function Get-PrinterSecurityFinding($printerNode, $props) {
    return ''
}

function Get-GroupSecurityFinding($props) {
    if (-not $props) { return '' }
    if ($props.groupName -eq 'Administrators' -or $props.groupName -match 'Admin') {
        return 'Local admin group managed via GPP - High privilege modification risk'
    }
    return ''
}

function Get-UserSecurityFinding($userNode, $props) {
    if (Test-CpasswordPresent $userNode $props) {
        return 'Local user with stored password in SYSVOL (AES-encrypted) - Credential exposure risk'
    }
    return ''
}

function Get-DataSourceSecurityFinding($dsNode, $props) {
    if (Test-CpasswordPresent $dsNode $props) {
        return 'Data Source stored credentials - Password in SYSVOL (AES-encrypted)'
    }
    return ''
}

function Get-FolderAclSecurityFinding($ace) {
    if (-not $ace) { return '' }
    $principal = ($ace.principal | ForEach-Object { $_.ToLower() })
    $rights = ($ace.rights | ForEach-Object { $_.ToLower() })
    if (($principal -match 'everyone|authenticated users|system' -or $principal -eq 's-1-1-0') -and ($rights -match 'full|modify')) {
        return "Overly permissive ACL: $($ace.principal) has $($ace.rights) rights"
    }
    return ''
}

function Get-TaskSecurityFinding($taskNode, $props) {
    if (Test-CpasswordPresent $taskNode $props) {
        return 'Scheduled task with stored credentials - Password in SYSVOL (AES-encrypted)'
    }
    return ''
}

function Get-ServiceSecurityFinding($serviceNode, $props) {
    if (Test-CpasswordPresent $serviceNode $props) {
        return 'Service with stored credentials - Password in SYSVOL (AES-encrypted)'
    }
    return ''
}

#endregion

#region Main Processing

$xmlFiles = Get-ChildItem -Path $GPOBackupRoot -Recurse -Filter *.xml | Where-Object { $_.FullName -match "Preferences" }
$totalFiles = $xmlFiles.Count
$fileIndex = 0

Write-Host "Starting to process $totalFiles preference files..." -ForegroundColor Cyan

foreach ($file in $xmlFiles) {
    $fileIndex++
    $fileName = Split-Path -Leaf $file.FullName
    Show-ProgressBar $fileIndex $totalFiles "Processing:"
    
    try {
        $config = if ($file.FullName -match "\\User\\") { "User" } elseif ($file.FullName -match "\\Machine\\") { "Machine" } else { "Unknown" }
        $metadata = Get-GpoMetadata $file.FullName
        $domain = $metadata.Domain
        $gpo = $metadata.GPO
    $type = switch -Regex ($file.Name) {
        '^Drives.xml$' { 'Drives' }
        '^Registry.xml$' { 'Registry' }
        '^Printers.xml$' { 'Printers' }
        '^Shortcuts.xml$' { 'Shortcuts' }
        '^Groups.xml$' { 'Groups' }
        '^Files.xml$' { 'Files' }
        '^Folders.xml$' { 'Folders' }
        '^IniFiles.xml$' { 'IniFiles' }
        '^NetworkShares.xml$' { 'NetworkShares' }
        '^DataSources.xml$' { 'DataSources' }
        '^Devices.xml$' { 'Devices' }
        '^EnvironmentVariables.xml$' { 'EnvironmentVariables' }
        '^InternetSettings.xml$' { 'InternetSettings' }
        '^PowerOptions.xml$' { 'PowerOptions' }
        '^FolderOptions.xml$' { 'FolderOptions' }
        '^Services.xml$' { 'Services' }
        '^ScheduledTasks.xml$' { 'ScheduledTasks' }
        default { 'OtherTypes' }
    }
        try {
            $xml = [xml](Get-Content -Path $file.FullName -Raw -ErrorAction Stop)
        } catch {
            $script:ErrorCount++
            Write-Warning "Failed to parse XML file: $($file.FullName) - $($_.Exception.Message)"
            Add-Item 'OtherTypes' @{ Domain=$domain; GPO=$gpo; Config=$config; PreferenceType=$type; XmlPath=$file.FullName; Note="XML load failed: $($_.Exception.Message)" }
            continue
        }

    switch ($type) {
        'Drives' {
            $nodes = @($xml.Drives.Drive)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.ChildNodes | Where-Object { $_.Name -eq 'Properties' } | Select-Object -First 1
                if ($props) {
                    $securityFinding = Get-DriveSecurityFinding $n $props
                    $hasCpassword = Test-CpasswordPresent $n $props
                    Add-Item 'Drives' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        DriveLetter=$props.letter; Path=$props.path; Label=$props.label;
                        UserName=$props.userName; Persistent=$props.persistent; UseLetter=$props.useLetter;
                        Action=$props.action; SecurityFinding=$securityFinding; HasCpassword=$hasCpassword
                    }
                }
            }
        }
        'Registry' {
            # Handle both direct Registry nodes and nested ones within Collections
            $nodes = @()
            $nodes += @($xml.RegistrySettings.Registry)
            $nodes += @($xml.RegistrySettings.Collection.Registry)
            $nodes += @($xml.Registry)
            
            # Function to recursively find all Registry elements within nested Collections
            function Find-RegistryNodes($element) {
                $result = @()
                if ($element.Registry) {
                    $result += @($element.Registry)
                }
                if ($element.Collection) {
                    foreach ($c in @($element.Collection)) {
                        $result += Find-RegistryNodes $c
                    }
                }
                return $result
            }
            
            # Find all registry nodes including nested ones
            if ($xml.RegistrySettings) {
                $nodes += Find-RegistryNodes $xml.RegistrySettings
            }
            
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    $securityFinding = Get-RegistrySecurityFinding $props
                    Add-Item 'Registry' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        Hive=$props.hive; RegistryKey=$props.key; ValueName=$props.name;
                        ValueData=$props.value; ValueType=$props.type; Action=$props.action;
                        DisplayDecimal=$props.displayDecimal; Default=$props.default; SecurityFinding=$securityFinding
                    }
                }
            }
        }
        'Printers' {
            $nodes = @($xml.Printers.SharedPrinter) + @($xml.Printers.TcpIpPrinter) + @($xml.Printers.LocalPrinter)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    $securityFinding = Get-PrinterSecurityFinding $n $props
                    Add-Item 'Printers' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        PrinterName=$n.name; Path=$props.path; Comment=$props.comment;
                        Location=$props.location; Port=$props.port; Action=$props.action;
                        SecurityFinding=$securityFinding
                    }
                }
            }
        }
        'Shortcuts' {
            $nodes = @($xml.Shortcuts.Shortcut)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'Shortcuts' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        ShortcutName=$n.name; TargetPath=$props.targetPath;
                        ShortcutPath=$props.shortcutPath; Arguments=$props.arguments;
                        Comment=$props.comment; TargetType=$props.targetType;
                        Action=$props.action
                    }
                }
            }
        }
        'Groups' {
            $nodes = @($xml.Groups.Group)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    $securityFinding = Get-GroupSecurityFinding $props
                    $members = @()
                    if ($props.Members.Member) {
                        $members = @($props.Members.Member)
                    }
                    # Add group record to GroupsDetail
                    Add-Item 'GroupsDetail' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        GroupName=$props.groupName; Description=$props.description;
                        GroupSID=$props.groupSID; SID=$n.sid; GroupAction=$props.action; Changed=$n.changed; SecurityFinding=$securityFinding
                    }
                    # If no members, create one record for the group without member info
                    if ($members.Count -eq 0) {
                        Add-Item 'Groups' @{
                            Domain=$domain; GPO=$gpo; Config=$config;
                            GroupName=$props.groupName; Description=$props.description;
                            MemberName=''; MemberAction=''; MemberSID='';
                            GroupAction=$props.action; SecurityFinding=$securityFinding
                        }
                    } else {
                        # Create a record for each member
                        foreach ($member in $members) {
                            Add-Item 'Groups' @{
                                Domain=$domain; GPO=$gpo; Config=$config;
                                GroupName=$props.groupName; Description=$props.description;
                                MemberName=$member.name; MemberAction=$member.action; MemberSID=$member.sid;
                                GroupAction=$props.action; SecurityFinding=$securityFinding
                            }
                        }
                    }
                }
            }
            # Handle User elements (local users)
            $userNodes = @($xml.Groups.User)
            foreach ($u in $userNodes) {
                if ($null -eq $u) { continue }
                $props = $u.Properties
                if ($props) {
                    $securityFinding = Get-UserSecurityFinding $u $props
                    $hasCpassword = Test-CpasswordPresent $u $props
                    Add-Item 'Users' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        UserName=$props.userName; FullName=$props.fullName;
                        Description=$props.description; Action=$props.action;
                        NeverExpires=$props.neverExpires; AccountDisabled=$props.acctDisabled;
                        ChangeLogon=$props.changeLogon; NoChange=$props.noChange;
                        Changed=$u.changed; SecurityFinding=$securityFinding; HasCpassword=$hasCpassword
                    }
                }
            }
        }
        'Files' {
            $nodes = @($xml.Files.File)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'Files' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        FileName=$n.name; SourcePath=$props.fromPath;
                        TargetPath=$props.targetPath; Action=$props.action;
                        ReadOnly=$props.readOnly; Archive=$props.archive;
                        Hidden=$props.hidden; Suppress=$props.suppress
                    }
                }
            }
        }
        'Folders' {
            $nodes = @($xml.Folders.Folder)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'Folders' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        FolderName=$n.name; FolderPath=$props.path;
                        Action=$props.action; DeleteFolder=$props.deleteFolder;
                        DeleteSubFolders=$props.deleteSubFolders; DeleteFiles=$props.deleteFiles;
                        DeleteReadOnly=$props.deleteReadOnly; ReadOnly=$props.readOnly;
                        Archive=$props.archive; Hidden=$props.hidden
                    }
                    # Extract ACL information if present
                    if ($props.ACL -and $props.ACL.ACE) {
                        $aces = @($props.ACL.ACE)
                        foreach ($ace in $aces) {
                            $securityFinding = Get-FolderAclSecurityFinding $ace
                            Add-Item 'FoldersACL' @{
                                Domain=$domain; GPO=$gpo; Config=$config;
                                FolderName=$n.name; FolderPath=$props.path;
                                ACEType=$ace.type; Principal=$ace.principal;
                                Rights=$ace.rights; Inheritance=$ace.inheritance;
                                FolderAction=$props.action; Changed=$n.changed; SecurityFinding=$securityFinding
                            }
                        }
                    }
                }
            }
        }
        'IniFiles' {
            $nodes = @($xml.IniFiles.Ini)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'IniFiles' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        IniName=$n.name; IniPath=$props.path; Section=$props.section;
                        Property=$props.property; Value=$props.value; Action=$props.action;
                        Changed=$n.changed
                    }
                }
            }
        }
        'NetworkShares' {
            $nodes = @($xml.NetworkShareSettings.NetShare)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'NetworkShares' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        ShareName=$props.name; SharePath=$props.path; Comment=$props.comment;
                        Action=$props.action; AllRegular=$props.allRegular; AllHidden=$props.allHidden;
                        AllAdminDrive=$props.allAdminDrive; LimitUsers=$props.limitUsers;
                        ABE=$props.abe; Changed=$n.changed
                    }
                }
            }
        }
        'DataSources' {
            $nodes = @($xml.DataSources.DataSource)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    $securityFinding = Get-DataSourceSecurityFinding $n $props
                    $hasCpassword = Test-CpasswordPresent $n $props
                    Add-Item 'DataSources' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        DataSourceName=$n.name; DSN=$props.dsn; Driver=$props.driver;
                        Description=$props.description; UserName=$props.username;
                        Action=$props.action; UserDSN=$props.userDSN; UserContext=$n.userContext;
                        Changed=$n.changed; SecurityFinding=$securityFinding; HasCpassword=$hasCpassword
                    }
                }
            }
        }
        'Devices' {
            $nodes = @($xml.Devices.Device)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'Devices' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        DeviceName=$n.name; DeviceAction=$props.deviceAction;
                        DeviceClass=$props.deviceClass; DeviceType=$props.deviceType;
                        DeviceClassGUID=$props.deviceClassGUID; DeviceTypeID=$props.deviceTypeID;
                        UserContext=$n.userContext; RemovePolicy=$n.removePolicy; Changed=$n.changed
                    }
                }
            }
        }
        'EnvironmentVariables' {
            $nodes = @($xml.EnvironmentVariables.EnvironmentVariable)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'EnvironmentVariables' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        VariableName=$props.name; VariableValue=$props.value;
                        Action=$props.action; User=$props.user; Partial=$props.partial;
                        Changed=$n.changed; RemovePolicy=$n.removePolicy; BypassErrors=$n.bypassErrors
                    }
                }
            }
        }
        'InternetSettings' {
            # Parse IE settings (IE7, IE8, IE9, IE10, IE11, etc.)
            $ieNodes = @($xml.InternetSettings.ChildNodes | Where-Object { $_.LocalName -match '^IE' })
            foreach ($ieNode in $ieNodes) {
                if ($null -eq $ieNode) { continue }
                $ieName = $ieNode.name
                $ieVersion = $ieNode.LocalName
                
                # Parse each Reg setting within Properties
                if ($ieNode.Properties -and $ieNode.Properties.Reg) {
                    foreach ($reg in @($ieNode.Properties.Reg)) {
                        Add-Item 'InternetSettings' @{
                            Domain=$domain; GPO=$gpo; Config=$config;
                            IEVersion=$ieVersion; IEName=$ieName;
                            SettingId=$reg.id; SettingName=$reg.name;
                            RegistryKey=$reg.key; RegistryHive=$reg.hive;
                            ValueType=$reg.type; Value=$reg.value;
                            Disabled=$reg.disabled
                        }
                    }
                }
            }
            # If no IE nodes found, record the file path for reference
            if ($ieNodes.Count -eq 0) {
                Add-Item 'InternetSettings' @{ Domain=$domain; GPO=$gpo; Config=$config; XmlFile=$file.FullName }
            }
        }
        'PowerOptions' {
            $nodes = @($xml.PowerOptions.GlobalPowerOptionsV2)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'PowerOptions' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        PowerPlanName=$n.name; Action=$props.action;
                        NameGuid=$props.nameGuid; DefaultPlan=$props.default;
                        RequireWakePwdAC=$props.requireWakePwdAC; RequireWakePwdDC=$props.requireWakePwdDC;
                        TurnOffHDAC=$props.turnOffHDAC; TurnOffHDDC=$props.turnOffHDDC;
                        SleepAfterAC=$props.sleepAfterAC; SleepAfterDC=$props.sleepAfterDC;
                        AllowHybridSleepAC=$props.allowHybridSleepAC; AllowHybridSleepDC=$props.allowHybridSleepDC;
                        HibernateAC=$props.hibernateAC; HibernateDC=$props.hibernateDC;
                        LidCloseAC=$props.lidCloseAC; LidCloseDC=$props.lidCloseDC;
                        PBActionAC=$props.pbActionAC; PBActionDC=$props.pbActionDC;
                        StrtMenuActionAC=$props.strtMenuActionAC; StrtMenuActionDC=$props.strtMenuActionDC;
                        LinkPwrMgmtAC=$props.linkPwrMgmtAC; LinkPwrMgmtDC=$props.linkPwrMgmtDC;
                        ProcStateMinAC=$props.procStateMinAC; ProcStateMinDC=$props.procStateMinDC;
                        ProcStateMaxAC=$props.procStateMaxAC; ProcStateMaxDC=$props.procStateMaxDC;
                        DisplayOffAC=$props.displayOffAC; DisplayOffDC=$props.displayOffDC;
                        AdaptiveAC=$props.adaptiveAC; AdaptiveDC=$props.adaptiveDC;
                        CritBatActionAC=$props.critBatActionAC; CritBatActionDC=$props.critBatActionDC;
                        LowBatteryLvlAC=$props.lowBatteryLvlAC; LowBatteryLvlDC=$props.lowBatteryLvlDC;
                        CritBatteryLvlAC=$props.critBatteryLvlAC; CritBatteryLvlDC=$props.critBatteryLvlDC;
                        LowBatteryNotAC=$props.lowBatteryNotAC; LowBatteryNotDC=$props.lowBatteryNotDC;
                        LowBatteryActionAC=$props.lowBatteryActionAC; LowBatteryActionDC=$props.lowBatteryActionDC;
                        Changed=$n.changed; UID=$n.uid
                    }
                }
            }
            # If no GlobalPowerOptionsV2 nodes found, record the file path for reference
            if ($nodes.Count -eq 0) {
                Add-Item 'PowerOptions' @{ Domain=$domain; GPO=$gpo; Config=$config; XmlFile=$file.FullName }
            }
        }
        'FolderOptions' {
            $nodes = @($xml.FolderOptions.OpenWith)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    Add-Item 'FolderOptions' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        OpenWithName=$n.name; FileExtension=$props.fileExtension;
                        ApplicationPath=$props.applicationPath; Default=$props.default;
                        Action=$props.action; Image=$n.image; UserContext=$n.userContext;
                        RemovePolicy=$n.removePolicy; Changed=$n.changed; UID=$n.uid
                    }
                }
            }
            # If no OpenWith nodes found, record the file path for reference
            if ($nodes.Count -eq 0) {
                Add-Item 'FolderOptions' @{ Domain=$domain; GPO=$gpo; Config=$config; XmlFile=$file.FullName }
            }
        }
        'ScheduledTasks' {
            # Parse TaskV2 and ImmediateTaskV2 nodes
            $taskNodes = @($xml.ScheduledTasks.TaskV2) + @($xml.ScheduledTasks.ImmediateTaskV2)
            foreach ($n in $taskNodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                $task = $props.Task
                if ($task) {
                    # Extract basic info
                    $author = ''
                    $description = ''
                    $enabled = ''
                    $hidden = ''
                    $priority = ''
                    $execTimeLimit = ''
                    $command = ''
                    $arguments = ''
                    $workingDir = ''
                    $runLevel = ''
                    $triggerType = ''
                    
                    if ($task.RegistrationInfo) {
                        $author = $task.RegistrationInfo.Author
                        $description = $task.RegistrationInfo.Description
                    }
                    
                    if ($task.Settings) {
                        $enabled = $task.Settings.Enabled
                        $hidden = $task.Settings.Hidden
                        $priority = $task.Settings.Priority
                        $execTimeLimit = $task.Settings.ExecutionTimeLimit
                    }
                    
                    if ($task.Principals -and $task.Principals.Principal) {
                        $runLevel = $task.Principals.Principal.RunLevel
                    }
                    
                    if ($task.Actions -and $task.Actions.Exec) {
                        $command = $task.Actions.Exec.Command
                        $arguments = $task.Actions.Exec.Arguments
                        $workingDir = $task.Actions.Exec.WorkingDirectory
                    }
                    
                    if ($task.Triggers) {
                        if ($task.Triggers.CalendarTrigger) { $triggerType = 'Calendar' }
                        elseif ($task.Triggers.EventTrigger) { $triggerType = 'Event' }
                        elseif ($task.Triggers.LogonTrigger) { $triggerType = 'Logon' }
                        elseif ($task.Triggers.TimeTrigger) { $triggerType = 'Time' }
                        elseif ($task.Triggers.RegistrationTrigger) { $triggerType = 'Registration' }
                        else { $triggerType = 'None' }
                    }
                    
                    Add-Item 'ScheduledTasks' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        TaskName=$props.name; Action=$props.action;
                        Author=$author; Description=$description;
                        RunAs=$props.runAs; LogonType=$props.logonType;
                        Enabled=$enabled; Hidden=$hidden; Priority=$priority;
                        ExecutionTimeLimit=$execTimeLimit; RunLevel=$runLevel;
                        TriggerType=$triggerType; Command=$command;
                        Arguments=$arguments; WorkingDirectory=$workingDir;
                        Changed=$n.changed; UID=$n.uid;
                        SecurityFinding = Get-TaskSecurityFinding $n $props; HasCpassword = (Test-CpasswordPresent $n $props)
                    }
                }
            }
        }
        'Services' {
            $nodes = @($xml.NTServices.NTService)
            foreach ($n in $nodes) {
                if ($null -eq $n) { continue }
                $props = $n.Properties
                if ($props) {
                    $securityFinding = Get-ServiceSecurityFinding $n $props
                    $hasCpassword = Test-CpasswordPresent $n $props
                    $serviceUsername = $props.logonAsServiceAccount
                    Add-Item 'Services' @{
                        Domain=$domain; GPO=$gpo; Config=$config;
                        ServiceName=$props.serviceName; DisplayName=$n.name;
                        StartupType=$props.startupType; Timeout=$props.timeout;
                        AccountName=$serviceUsername;
                        Changed=$n.changed; UID=$n.uid; Action=$n.action; SecurityFinding=$securityFinding; HasCpassword=$hasCpassword
                    }
                }
            }
        }
        default {
            Add-Item 'OtherTypes' @{ Domain=$domain; GPO=$gpo; Config=$config; PreferenceType=$type; XmlPath=$file.FullName }
        }
    }
    }
    catch {
        $script:ErrorCount++
        Write-Warning "Error processing file $($file.FullName): $($_.Exception.Message)"
    }
}

#endregion

#region Export CSV Files

Write-Host "" -ForegroundColor Cyan
Write-Host "Exporting CSV files..." -ForegroundColor Cyan
if ($ExportCSV) {
    try {
        $csvFiles = @('Drives','Registry','Printers','Shortcuts','GroupsDetail','Groups','Users','Files','Folders','FoldersACL','IniFiles','NetworkShares','DataSources','Devices','EnvironmentVariables','InternetSettings','PowerOptions','FolderOptions','Services','ScheduledTasks','OtherTypes')
        $totalCsvFiles = $csvFiles.Count
        
        $csvIndex = 0
        foreach ($csvType in $csvFiles) {
            $csvIndex++
        $csvFileName = switch($csvType) {
            'Drives' { 'GPP_DriveMapping_Details.csv' }
            'Registry' { 'GPP_Registry_Details.csv' }
            'Printers' { 'GPP_Printers_Details.csv' }
            'Shortcuts' { 'GPP_Shortcuts_Details.csv' }
            'GroupsDetail' { 'GPP_Groups_Details.csv' }
            'Groups' { 'GPP_GroupMemberships_Details.csv' }
            'Users' { 'GPP_Users_Details.csv' }
            'Files' { 'GPP_Files_Details.csv' }
            'Folders' { 'GPP_Folders_Details.csv' }
            'FoldersACL' { 'GPP_FoldersACL_Details.csv' }
            'IniFiles' { 'GPP_IniFiles_Details.csv' }
            'NetworkShares' { 'GPP_NetworkShares_Details.csv' }
            'DataSources' { 'GPP_DataSources_Details.csv' }
            'Devices' { 'GPP_Devices_Details.csv' }
            'EnvironmentVariables' { 'GPP_EnvironmentVariables_Details.csv' }
            'InternetSettings' { 'GPP_InternetSettings_Details.csv' }
            'PowerOptions' { 'GPP_PowerOptions_Details.csv' }
            'FolderOptions' { 'GPP_FolderOptions_Details.csv' }
            'Services' { 'GPP_Services_Details.csv' }
            'ScheduledTasks' { 'GPP_ScheduledTasks_Details.csv' }
            'OtherTypes' { 'GPP_OtherTypes.csv' }
        }
        
            Show-ProgressBar $csvIndex $totalCsvFiles "Exporting:"
            try {
                $results[$csvType] | Export-Csv -Path (Join-Path $OutputDir $csvFileName) -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                $script:ErrorCount++
                Write-Warning "Failed to export CSV for $csvType`: $($_.Exception.Message)"
            }
        }
        Write-Host "" -ForegroundColor Cyan
        Write-Host "CSV export completed!" -ForegroundColor Green
    }
    catch {
        $script:ErrorCount++
        Write-Error "CSV export failed: $($_.Exception.Message)"
    }
} else {
    Write-Host "Skipping CSV export (disabled)" -ForegroundColor Yellow
}

#endregion

#region Calculate Statistics

# Calculate overview statistics and preference type distribution for both JSON and HTML
$allRecords = @()
$results.Keys | ForEach-Object {
    if ($results[$_] -is [System.Collections.Generic.List[Object]]) {
        $allRecords += $results[$_]
    }
}

# Count all GPOs from gpreport.xml files
$allGpos = @()
$gpoReportFiles = Get-ChildItem -Path $GPOBackupRoot -Recurse -Filter "gpreport.xml" -ErrorAction SilentlyContinue
foreach ($reportFile in $gpoReportFiles) {
    try {
        [xml]$xml = Get-Content -Path $reportFile.FullName -Raw
        $gpoName = "UnknownGPO"
        $domain = "UnknownDomain"
        
        foreach ($child in $xml.DocumentElement.ChildNodes) {
            if ($child.LocalName -eq 'Name') {
                $gpoName = $child.InnerText.Trim()
            }
            if ($child.LocalName -eq 'Identifier') {
                foreach ($sub in $child.ChildNodes) {
                    if ($sub.LocalName -eq 'Domain') {
                        $domain = $sub.InnerText.Trim()
                        break
                    }
                }
            }
        }
        
        if ($gpoName -ne "UnknownGPO" -and $domain -ne "UnknownDomain") {
            $allGpos += [PSCustomObject]@{ Domain = $domain; GPO = $gpoName }
        }
    } catch {
        # Skip if gpreport.xml fails to parse
    }
}

$uniqueDomains = @($allRecords | Select-Object -ExpandProperty Domain -Unique | Where-Object {$_}).Count
$uniqueGPOs = @($allGpos | Select-Object Domain, GPO -Unique).Count
$gposWithPrefs = @($allRecords | Select-Object Domain, GPO -Unique | Where-Object {$_.Domain}).Count

# Count risky configurations (non-empty SecurityFinding entries)
$riskyConfigs = 0
@($results.Drives, $results.Registry, $results.Printers, $results.Groups, $results.GroupsDetail, $results.Users, $results.DataSources, $results.FoldersACL, $results.ScheduledTasks, $results.Services) | ForEach-Object {
    $riskyConfigs += @($_ | Where-Object { $_.SecurityFinding -and $_.SecurityFinding -ne '' }).Count
}

# Build preference type breakdown
$prefTypeStats = @()
$prefTypeMap = @(
    @{Type='Drives'; Data=$results.Drives; Name='Drive Mappings'; Id='drives'; Category='ControlPanel'}
    @{Type='Registry'; Data=$results.Registry; Name='Registry Settings'; Id='registry'; Category='Windows'}
    @{Type='Printers'; Data=$results.Printers; Name='Printers'; Id='printers'; Category='ControlPanel'}
    @{Type='Shortcuts'; Data=$results.Shortcuts; Name='Shortcuts'; Id='shortcuts'; Category='Windows'}
    @{Type='Groups'; Data=$results.Groups; Name='Local Groups'; Id='groups'; Category='ControlPanel'}
    @{Type='Users'; Data=$results.Users; Name='Local Users'; Id='users'; Category='ControlPanel'}
    @{Type='Files'; Data=$results.Files; Name='Files'; Id='files'; Category='Windows'}
    @{Type='Folders'; Data=$results.Folders; Name='Folders'; Id='folders'; Category='Windows'}
    @{Type='IniFiles'; Data=$results.IniFiles; Name='INI Files'; Id='inifiles'; Category='Windows'}
    @{Type='NetworkShares'; Data=$results.NetworkShares; Name='Network Shares'; Id='networkshares'; Category='ControlPanel'}
    @{Type='DataSources'; Data=$results.DataSources; Name='Data Sources'; Id='datasources'; Category='ControlPanel'}
    @{Type='Devices'; Data=$results.Devices; Name='Devices'; Id='devices'; Category='ControlPanel'}
    @{Type='EnvironmentVariables'; Data=$results.EnvironmentVariables; Name='Environment Variables'; Id='env'; Category='Windows'}
    @{Type='InternetSettings'; Data=$results.InternetSettings; Name='Internet Settings'; Id='inet'; Category='ControlPanel'}
    @{Type='PowerOptions'; Data=$results.PowerOptions; Name='Power Options'; Id='power'; Category='ControlPanel'}
    @{Type='FolderOptions'; Data=$results.FolderOptions; Name='Folder Options'; Id='folderopts'; Category='ControlPanel'}
    @{Type='Services'; Data=$results.Services; Name='Services'; Id='services'; Category='ControlPanel'}
    @{Type='ScheduledTasks'; Data=$results.ScheduledTasks; Name='Scheduled Tasks'; Id='tasks'; Category='ControlPanel'}
)

$allowedPrefCategories = @('Windows','ControlPanel')

foreach ($pref in ($prefTypeMap | Where-Object { $allowedPrefCategories -contains $_.Category })) {
    $gpoCount = @($pref.Data | Select-Object Domain, GPO -Unique | Where-Object {$_.Domain}).Count
    if ($gpoCount -gt 0) {
        $prefTypeStats += [PSCustomObject]@{
            Type = $pref.Type
            Name = $pref.Name
            Id = $pref.Id
            Category = $pref.Category
            RecordCount = $pref.Data.Count
            GPOCount = $gpoCount
        }
    }
}

$summary = @(
    @{Name='Drive Mappings'; Data=$results.Drives},
    @{Name='Registry Settings'; Data=$results.Registry},
    @{Name='Printers'; Data=$results.Printers},
    @{Name='Shortcuts'; Data=$results.Shortcuts},
    @{Name='Local Groups'; Data=$results.GroupsDetail},
    @{Name='Group Memberships'; Data=$results.Groups},
    @{Name='Local Users'; Data=$results.Users},
    @{Name='Files'; Data=$results.Files},
    @{Name='Folders'; Data=$results.Folders},
    @{Name='Folders ACL'; Data=$results.FoldersACL},
    @{Name='INI Files'; Data=$results.IniFiles},
    @{Name='Network Shares'; Data=$results.NetworkShares},
    @{Name='Data Sources'; Data=$results.DataSources},
    @{Name='Devices'; Data=$results.Devices},
    @{Name='Environment Variables'; Data=$results.EnvironmentVariables},
    @{Name='Internet Settings'; Data=$results.InternetSettings},
    @{Name='Power Options'; Data=$results.PowerOptions},
    @{Name='Folder Options'; Data=$results.FolderOptions},
    @{Name='Services'; Data=$results.Services},
    @{Name='Scheduled Tasks'; Data=$results.ScheduledTasks},
    @{Name='Other Types'; Data=$results.OtherTypes}
)

$totalRecords = (($summary | ForEach-Object { $_.Data.Count }) | Measure-Object -Sum).Sum

#endregion

#region Export JSON

if ($ExportJSON) {
    try {
        Write-Host "" -ForegroundColor Cyan
        Write-Host "Exporting JSON data..." -ForegroundColor Cyan
        $jsonFileName = "GPP_Complete_Data_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        $jsonPath = Join-Path $OutputDir $jsonFileName
    
    # Build comprehensive JSON structure with metadata, overview, distribution, and raw data
    $jsonData = [ordered]@{
        ReportMetadata = [ordered]@{
            ReportGeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            GPOBackupFolder = $GPOBackupRoot
            OutputDirectory = $OutputDir
        }
        Overview = [ordered]@{
            TotalDomains = $uniqueDomains
            TotalGPOs = $uniqueGPOs
            GPOsWithPreferences = $gposWithPrefs
            TotalPreferenceRecords = $totalRecords
            RiskyConfigurations = $riskyConfigs
        }
        PreferenceTypeDistribution = @{
            WindowsSettings = @($prefTypeStats | Where-Object { $_.Category -eq 'Windows' } | ForEach-Object {
                [ordered]@{
                    Type = $_.Type
                    Name = $_.Name
                    GPOCount = $_.GPOCount
                    RecordCount = $_.RecordCount
                }
            })
            ControlPanelSettings = @($prefTypeStats | Where-Object { $_.Category -eq 'ControlPanel' } | ForEach-Object {
                [ordered]@{
                    Type = $_.Type
                    Name = $_.Name
                    GPOCount = $_.GPOCount
                    RecordCount = $_.RecordCount
                }
            })
        }
        PreferenceData = $results
    }
    
        $jsonData | ConvertTo-Json -Depth 10 -ErrorAction Stop | Out-File $jsonPath -Encoding UTF8 -ErrorAction Stop
        Write-Host "JSON export completed: $jsonPath" -ForegroundColor Green
    }
    catch {
        $script:ErrorCount++
        Write-Error "JSON export failed: $($_.Exception.Message)"
    }
} else {
    Write-Host "Skipping JSON export (disabled)" -ForegroundColor Yellow
}

#endregion

#region Generate HTML Report

Write-Host "Generating HTML report..." -ForegroundColor Cyan

try {

function Build-Section($title,$id,$data,$columns,$filterId,$tableId) {
    if (-not $data -or $data.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<details><summary>$title <span style='color:#106ebe;font-size:14px;'>($($data.Count))</span></summary>")
    [void]$sb.Append("<div class='section' id='$id'>")
    if ($filterId -and $tableId) {
        [void]$sb.Append("<div class='filter'><input id='$filterId' onkeyup=`"filterTable('$filterId','$tableId')`" placeholder='Filter...'></div>")
    }
    [void]$sb.Append("<table id='$tableId'><tr>")
    foreach ($c in $columns) {
        if ($c -eq 'SecurityFinding') {
            [void]$sb.Append("<th class='security-header'></th>")
        } else {
            [void]$sb.Append("<th>$c</th>")
        }
    }
    [void]$sb.Append("</tr>")
    foreach ($row in ($data | Select-Object -First 500)) {
        [void]$sb.Append("<tr>")
        foreach ($c in $columns) {
            $val = EncodeHtml $row.$c
            if ($c -eq 'CPassword' -and $val) {
                [void]$sb.Append("<td class='cpassword-cell'></td>")
            } elseif ($c -eq 'SecurityFinding' -and $val) {
                [void]$sb.Append("<td class='security-cell' title='$val'>⚠️</td>")
            } elseif ($c -eq 'SecurityFinding') {
                [void]$sb.Append("<td class='security-cell'></td>")
            } else {
                [void]$sb.Append("<td>$val</td>")
            }
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</table></div></details>")
    return $sb.ToString()
}

function Build-SubSection($title,$data,$columns,$filterId,$tableId) {
    if (-not $data -or $data.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<h3 style='margin-top:20px;color:#0078d4;'>$title <span style='color:#106ebe;font-size:13px;'>($($data.Count))</span></h3>")
    if ($filterId -and $tableId) {
        [void]$sb.Append("<div class='filter'><input id='$filterId' onkeyup=`"filterTable('$filterId','$tableId')`" placeholder='Filter...'></div>")
    }
    [void]$sb.Append("<table id='$tableId'><tr>")
    foreach ($c in $columns) {
        if ($c -eq 'SecurityFinding') {
            [void]$sb.Append("<th class='security-header'></th>")
        } else {
            [void]$sb.Append("<th>$c</th>")
        }
    }
    [void]$sb.Append("</tr>")
    foreach ($row in ($data | Select-Object -First 500)) {
        [void]$sb.Append("<tr>")
        foreach ($c in $columns) {
            $val = EncodeHtml $row.$c
            if ($c -eq 'SecurityFinding' -and $val) {
                [void]$sb.Append("<td class='security-cell' title='$val'>⚠️</td>")
            } elseif ($c -eq 'SecurityFinding') {
                [void]$sb.Append("<td class='security-cell'></td>")
            } else {
                [void]$sb.Append("<td>$val</td>")
            }
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</table>")
    return $sb.ToString()
}

function Build-RiskyGroupedSection($data) {
    if (-not $data -or $data.Count -eq 0) { return "<p>No risky configurations detected.</p>" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div style='max-width:100%;'>")
    $grouped = $data | Group-Object SecurityFinding
    foreach ($group in ($grouped | Sort-Object @{ Expression = { $_.Name -notmatch 'stored credentials' }; Ascending = $true }, @{ Expression = 'Name'; Ascending = $true })) {
        $finding = EncodeHtml $group.Name
        [void]$sb.Append("<details><summary>$finding <span style='color:#106ebe;font-size:14px;'>($($group.Count))</span></summary>")
        [void]$sb.Append("<div class='section'>")
        
        # Check if any row in this group has a cpassword value
        # Also skip cpassword column for ACL and admin group findings
        $isCpasswordNeeded = ($group.Group | Where-Object { $_.HasCpassword -eq $true } | Measure-Object | Select-Object -ExpandProperty Count) -gt 0
        $isAclOrAdminFinding = $group.Name -match 'permissive ACL|admin group'
        $hasCpasswordColumn = $isCpasswordNeeded -and -not $isAclOrAdminFinding
        
        # Build header with or without cpassword column
        if ($hasCpasswordColumn) {
            [void]$sb.Append("<table><tr><th>Domain</th><th>GPO</th><th>Type</th><th>Item</th><th>UserName</th><th>Has cpassword</th></tr>")
        } else {
            [void]$sb.Append("<table><tr><th>Domain</th><th>GPO</th><th>Type</th><th>Item</th><th>UserName</th></tr>")
        }
        
        foreach ($row in $group.Group) {
            $domain = EncodeHtml $row.Domain
            $gpo = EncodeHtml $row.GPO
            $type = EncodeHtml $row.Type
            $item = EncodeHtml $row.Item
            $username = EncodeHtml $row.UserName
            
            if ($hasCpasswordColumn) {
                $hasCpassword = if ($row.HasCpassword -eq $true) { 'Yes' } elseif ($row.HasCpassword -eq $false) { 'No' } else { '' }
                [void]$sb.Append("<tr><td>$domain</td><td>$gpo</td><td>$type</td><td>$item</td><td>$username</td><td>$hasCpassword</td></tr>")
            } else {
                [void]$sb.Append("<tr><td>$domain</td><td>$gpo</td><td>$type</td><td>$item</td><td>$username</td></tr>")
            }
        }
        [void]$sb.Append("</table></div></details>")
    }
    [void]$sb.Append("</div>")
    return $sb.ToString()
}

$report = Join-Path $OutputDir "Group Policy Preferences Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

$html = @"
<!DOCTYPE html>
<html><head><meta charset='UTF-8'><title>Group Policy Preferences Report</title>
<style>body{font-family:Segoe UI,Arial,sans-serif;background:#f5f5f5;margin:0;position:relative;} .github-corner{position:fixed;top:0;right:0;z-index:9999;} .github-corner svg{fill:#0078d4;color:#fff;width:80px;height:80px;} .github-corner:hover svg{fill:#106ebe;} .github-corner .octo-arm{transform-origin:130px 106px;} .github-corner:hover .octo-arm{animation:octocat-wave 560ms ease-in-out;} @keyframes octocat-wave{0%,100%{transform:rotate(0)}20%,60%{transform:rotate(-25deg)}40%,80%{transform:rotate(10deg)}} .header{background:linear-gradient(135deg,#0078d4,#106ebe);color:#fff;padding:26px;text-align:center;} .container{max-width:1650px;margin:0 auto;padding:18px;display:flex;flex-direction:column;gap:24px;} .stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin:18px 0;} .card{background:#fff;padding:18px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.12);} .card h3{margin:0;color:#0078d4;} .value{font-size:30px;font-weight:700;color:#0078d4;} .pref-section{display:flex;gap:24px;margin:18px 0;} .pref-section-col{flex:1;} .pref-section-col h3{margin:0 0 12px 0;color:#0078d4;font-size:16px;font-weight:600;} .pref-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;} .pref-card{background:#fff;padding:14px;border-radius:6px;border-left:4px solid #0078d4;box-shadow:0 1px 3px rgba(0,0,0,.1);} .pref-card h4{margin:0 0 8px 0;color:#0078d4;font-size:13px;text-transform:uppercase;} .pref-card .pref-value{font-size:20px;font-weight:700;color:#106ebe;display:flex;align-items:baseline;gap:6px;} .pref-card .pref-sub{font-size:11px;font-weight:600;color:#666;} table{width:100%;border-collapse:collapse;margin:12px 0;font-size:13px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.1);table-layout:auto;} th{background:#0078d4;color:#fff;padding:9px 7px;text-align:left;position:sticky;top:0;word-break:break-word;white-space:nowrap;overflow-wrap:anywhere;} th.security-header{background:#0078d4;} td{padding:8px 7px;border-bottom:1px solid #e5e5e5;word-break:break-word;overflow-wrap:anywhere;} tr:nth-child(even){background:#f8f9fa;} tr:hover{background:#e9f3ff;} .security-cell{background-color:#fff3cd !important;color:#856404;font-weight:600;cursor:pointer;} #driveTable td:nth-child(5),#driveTable td:nth-child(6),#driveTable td:nth-child(10){max-width:200px;} #regTable td:nth-child(7){max-width:350px;} #printerTable td:nth-child(5),#printerTable td:nth-child(10){max-width:200px;} #shortcutTable td:nth-child(5),#shortcutTable td:nth-child(6){max-width:200px;} #filesTable td:nth-child(5),#filesTable td:nth-child(6){max-width:200px;} #foldersTable td:nth-child(4),#foldersTable td:nth-child(5){max-width:200px;} #folderaclTable td:nth-child(2),#folderaclTable td:nth-child(4){max-width:250px;} #inifilesTable td:nth-child(4),#inifilesTable td:nth-child(5){max-width:200px;} #networksharesTable td:nth-child(4),#networksharesTable td:nth-child(5){max-width:200px;} #datasourcesTable td:nth-child(5),#datasourcesTable td:nth-child(6),#datasourcesTable td:nth-child(10){max-width:200px;} #usersTable td:nth-child(10){max-width:200px;} #devicesTable td:nth-child(4){max-width:200px;} #inetTable td:nth-child(6){max-width:250px;} #tasksTable td:nth-child(5),#tasksTable td:nth-child(12){max-width:200px;} #tasksTable td:nth-child(12){max-width:200px;} #powerTable td:nth-child(4){max-width:200px;} #folderoptTable td:nth-child(4){max-width:200px;} #servicesTable td:nth-child(9){max-width:200px;} details{margin:20px 0;background:#fff;border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,.1);} details[open]{box-shadow:0 2px 8px rgba(0,0,0,.15);} summary{cursor:pointer;padding:16px;background:#f8f9fa;border-radius:6px;font-weight:600;color:#0078d4;user-select:none;display:flex;justify-content:space-between;align-items:center;} summary:hover{background:#e9f3ff;} details[open]>summary{background:#0078d4;color:#fff;border-radius:6px 6px 0 0;} .section{display:block;padding:0;width:100%;} .filter input{padding:7px 9px;width:280px;border:1px solid #ccc;border-radius:4px;} .filter{padding:12px 16px;border-top:1px solid #e5e5e5;} .main-section{background:#fff;padding:24px;border-radius:8px;margin-bottom:30px;box-shadow:0 2px 8px rgba(0,0,0,.1);} .main-section h1{color:#0078d4;margin:0 0 20px 0;padding-bottom:12px;border-bottom:3px solid #0078d4;font-size:24px;}</style>
<script>function filterTable(i,t){const v=document.getElementById(i).value.toUpperCase();const r=document.getElementById(t).getElementsByTagName('tr');for(let x=1;x<r.length;x++){const d=r[x].getElementsByTagName('td');let s=false;for(let j=0;j<d.length;j++){if(d[j].textContent.toUpperCase().indexOf(v)>-1){s=true;break;}}r[x].style.display=s?'':'none';}}</script>
</head><body>
<a href='https://github.com/alexverboon/GroupPolicypreferencesReporter' target='_blank' class='github-corner' aria-label='View source on GitHub'><svg viewBox='0 0 250 250' aria-hidden='true'><path d='M0,0 L115,115 L130,115 L142,142 L250,250 L250,0 Z'></path><path d='M128.3,109.0 C113.8,99.7 119.0,89.6 119.0,89.6 C122.0,82.7 120.5,78.6 120.5,78.6 C119.2,72.0 123.4,76.3 123.4,76.3 C127.3,80.9 125.5,87.3 125.5,87.3 C122.9,97.6 130.6,101.9 134.4,103.2' fill='currentColor' style='transform-origin:130px 106px;' class='octo-arm'></path><path d='M115.0,115.0 C114.9,115.1 118.7,116.5 119.8,115.4 L133.7,101.6 C136.9,99.2 139.9,98.4 142.2,98.6 C133.8,88.0 127.5,74.4 143.8,58.0 C148.5,53.4 154.0,51.2 159.7,51.0 C160.3,49.4 163.2,43.6 171.4,40.1 C171.4,40.1 176.1,42.5 178.8,56.2 C183.1,58.6 187.2,61.8 190.9,65.4 C194.5,69.0 197.7,73.2 200.1,77.6 C213.8,80.2 216.3,84.9 216.3,84.9 C212.7,93.1 206.9,96.0 205.4,96.6 C205.1,102.4 203.0,107.8 198.3,112.5 C181.9,128.9 168.3,122.5 157.7,114.1 C157.9,116.9 156.7,120.9 152.7,124.9 L141.0,136.5 C139.8,137.7 141.6,141.9 141.8,141.8 Z' fill='currentColor' class='octo-body'></path></svg></a>
<div class='header'><h1>Group Policy Preferences Report</h1><p>GPO Backup Folder: $GPOBackupRoot</p><p>Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p></div>
<div class='container'>
<div class='main-section'>
<h1>Overview</h1>
<div class='stat-grid'>
<div class='card'><h3>Total Domains</h3><div class='value'>$uniqueDomains</div></div>
<div class='card'><h3>Total GPOs</h3><div class='value'>$uniqueGPOs</div></div>
<div class='card'><h3>GPOs with Preferences</h3><div class='value'>$gposWithPrefs</div></div>
<div class='card'><h3>Total Preference Records</h3><div class='value'>$totalRecords</div></div>
<div class='card'><h3>Risky Configurations</h3><div class='value'>$riskyConfigs</div></div>
</div>
</div>

<div class='main-section'>
<h1>Preference Type Distribution</h1>
<div class='pref-section'>
<div class='pref-section-col'>
<h3>Windows Settings (Preferences)</h3>
<div class='pref-grid'>
"@

foreach ($stat in ($prefTypeStats | Where-Object { $_.Category -eq 'Windows' })) {
    $html += @"
<div class='pref-card'>
<h4><a href='#$($stat.Id)' style='color:#0078d4;text-decoration:none;'>$($stat.Name)</a></h4>
<div class='pref-value'>$($stat.GPOCount)<span class='pref-sub'>GPOs | $($stat.RecordCount) settings</span></div>
</div>
"@
}

$html += @"
</div>
</div>
<div class='pref-section-col'>
<h3>Control Panel Settings (Preferences)</h3>
<div class='pref-grid'>
"@

foreach ($stat in ($prefTypeStats | Where-Object { $_.Category -eq 'ControlPanel' })) {
    $html += @"
<div class='pref-card'>
<h4><a href='#$($stat.Id)' style='color:#0078d4;text-decoration:none;'>$($stat.Name)</a></h4>
<div class='pref-value'>$($stat.GPOCount)<span class='pref-sub'>GPOs | $($stat.RecordCount) settings</span></div>
</div>
"@
}

$html += @"
</div>
</div>
</div>
</div>

<div class='main-section'>
<h1>Details</h1>
"@

$html += Build-Section 'Drive Mappings' 'drives' $results.Drives @('Domain','GPO','Config','DriveLetter','Path','Label','UserName','Action','HasCpassword','SecurityFinding') 'driveFilter' 'driveTable'
# Filter out DefaultPassword/DefaultUserName from Registry display (they're merged in Risky Configs)
$registryForDisplay = @($results.Registry | Where-Object {
    $valueName = $_.ValueName | ForEach-Object { $_.ToLower() }
    -not ($_.RegistryKey -match 'winlogon' -and ($valueName -eq 'defaultpassword' -or $valueName -eq 'defaultusername'))
})
$html += Build-Section 'Registry Settings' 'registry' $registryForDisplay @('Domain','GPO','Config','Hive','RegistryKey','ValueName','ValueData','ValueType','Action','DisplayDecimal','SecurityFinding') 'regFilter' 'regTable'
$html += Build-Section 'Printers' 'printers' $results.Printers @('Domain','GPO','Config','PrinterName','Path','Comment','Location','Port','Action','SecurityFinding') 'printerFilter' 'printerTable'
$html += Build-Section 'Shortcuts' 'shortcuts' $results.Shortcuts @('Domain','GPO','Config','ShortcutName','TargetPath','ShortcutPath','Arguments','Comment','TargetType','Action') 'shortcutFilter' 'shortcutTable'
$html += Build-Section 'Local Groups' 'groups' $results.Groups @('Domain','GPO','Config','GroupName','GroupAction','MemberName','MemberAction','MemberSID','Description','SecurityFinding') 'groupsFilter' 'groupsTable'
$html += Build-Section 'Local Users' 'users' $results.Users @('Domain','GPO','Config','UserName','FullName','Description','Action','NeverExpires','AccountDisabled','HasCpassword','SecurityFinding') 'usersFilter' 'usersTable'
$html += Build-Section 'Files' 'files' $results.Files @('Domain','GPO','Config','FileName','SourcePath','TargetPath','Action','ReadOnly','Archive','Hidden','Suppress','Changed') 'filesFilter' 'filesTable'
$html += Build-Section 'Folders' 'folders' $results.Folders @('Domain','GPO','Config','FolderName','FolderPath','Action','DeleteFolder','DeleteSubFolders','DeleteFiles','DeleteReadOnly','Archive','Hidden','Changed') 'foldersFilter' 'foldersTable'
$html += Build-Section 'Folders ACL' 'foldersacl' $results.FoldersACL @('Domain','GPO','Config','FolderName','FolderPath','Principal','ACEType','Rights','Inheritance','FolderAction','SecurityFinding') 'folderaclFilter' 'folderaclTable'
$html += Build-Section 'INI Files' 'inifiles' $results.IniFiles @('Domain','GPO','Config','IniName','IniPath','Section','Property','Value','Action','Changed') 'inifilesFilter' 'inifilesTable'
$html += Build-Section 'Network Shares' 'networkshares' $results.NetworkShares @('Domain','GPO','Config','ShareName','SharePath','Comment','Action','AllRegular','AllHidden','AllAdminDrive','LimitUsers','ABE','Changed') 'networksharesFilter' 'networksharesTable'
$html += Build-Section 'Data Sources' 'datasources' $results.DataSources @('Domain','GPO','Config','DataSourceName','DSN','Driver','Description','UserName','Action','HasCpassword','SecurityFinding') 'datasourcesFilter' 'datasourcesTable'
$html += Build-Section 'Devices' 'devices' $results.Devices @('Domain','GPO','Config','DeviceName','DeviceAction','DeviceClass','DeviceType','DeviceClassGUID','DeviceTypeID','UserContext','RemovePolicy','Changed') 'devicesFilter' 'devicesTable'
$html += Build-Section 'Environment Variables' 'env' $results.EnvironmentVariables @('Domain','GPO','Config','VariableName','VariableValue','Action','User','Partial','RemovePolicy','BypassErrors','Changed') 'envFilter' 'envTable'
$html += Build-Section 'Internet Settings' 'inet' $results.InternetSettings @('Domain','GPO','Config','IEVersion','SettingId','SettingName','RegistryKey','ValueType','Value','Disabled') 'inetFilter' 'inetTable'
$html += Build-Section 'Scheduled Tasks' 'tasks' $results.ScheduledTasks @('Domain','GPO','Config','TaskName','Action','Author','Description','RunAs','Enabled','TriggerType','Command','HasCpassword','SecurityFinding') 'tasksFilter' 'tasksTable'
$html += Build-Section 'Power Options' 'power' $results.PowerOptions @('Domain','GPO','Config','PowerPlanName','DefaultPlan','SleepAfterAC','SleepAfterDC','HibernateAC','HibernateDC','DisplayOffAC','DisplayOffDC','LidCloseAC','LidCloseDC','ProcStateMinAC','ProcStateMinDC','Changed') 'powerFilter' 'powerTable'
$html += Build-Section 'Folder Options' 'folderopts' $results.FolderOptions @('Domain','GPO','Config','OpenWithName','FileExtension','ApplicationPath','Action','Default','UserContext','RemovePolicy','Changed') 'folderoptFilter' 'folderoptTable'
$html += Build-Section 'Services' 'services' $results.Services @('Domain','GPO','Config','ServiceName','DisplayName','StartupType','Timeout','AccountName','Action','HasCpassword','SecurityFinding') 'servicesFilter' 'servicesTable'
$html += Build-Section 'Other Types' 'other' $results.OtherTypes @('Domain','GPO','Config','PreferenceType','XmlPath','Note') 'otherFilter' 'otherTable'

# Aggregate all risky configurations
$riskyConfigsList = @()
$registryItems = @()
$mergedRegistryItems = @()

# First, collect registry items to potentially merge DefaultPassword and DefaultUserName
@($results.Drives, $results.Registry, $results.Printers, $results.Groups, $results.GroupsDetail, $results.Users, $results.DataSources, $results.FoldersACL, $results.ScheduledTasks, $results.Services) | ForEach-Object {
    $_ | Where-Object { $_.SecurityFinding -and $_.SecurityFinding -ne '' -and $_.SecurityFinding -notmatch 'Windows Firewall' } | ForEach-Object {
        if ($_ -in $results.Registry) {
            $registryItems += $_
        } else {
            $type = ''
            if ($_ -in $results.Drives) { $type = 'Drive Mapping' }
            elseif ($_ -in $results.Printers) { $type = 'Printer' }
            elseif ($_ -in $results.Groups) { $type = 'Group Member' }
            elseif ($_ -in $results.GroupsDetail) { $type = 'Group' }
            elseif ($_ -in $results.Users) { $type = 'User' }
            elseif ($_ -in $results.DataSources) { $type = 'DataSource' }
            elseif ($_ -in $results.FoldersACL) { $type = 'Folder ACL' }
            elseif ($_ -in $results.ScheduledTasks) { $type = 'Scheduled Task' }
            elseif ($_ -in $results.Services) { $type = 'Service' }
            
            $itemName = ''
            $username = ''
            $hasCpassword = ''
            
            if ($_.DriveLetter) { 
                $itemName = $_.DriveLetter
                $username = $_.UserName
                $hasCpassword = $_.HasCpassword
            }
            elseif ($_.PrinterName) { $itemName = $_.PrinterName }
            elseif ($_.GroupName) { $itemName = if ($_.MemberName) { "$($_.GroupName) / $($_.MemberName)" } else { $_.GroupName } }
            elseif ($_.DataSourceName) { 
                $itemName = $_.DataSourceName
                $username = $_.UserName
                $hasCpassword = $_.HasCpassword
            }
            elseif ($_.UserName) { 
                $itemName = $_.UserName
                $hasCpassword = $_.HasCpassword
            }
            elseif ($_.FolderName) { $itemName = "$($_.FolderName) / $($_.Principal)" }
            elseif ($_.TaskName) { 
                $itemName = $_.TaskName
                $username = $_.RunAs
                $hasCpassword = $_.HasCpassword
            }
            elseif ($_.ServiceName) { 
                $itemName = $_.ServiceName
                $username = $_.AccountName
                $hasCpassword = $_.HasCpassword
            }
            
            $riskyConfigsList += [pscustomobject]@{
                Domain = $_.Domain
                GPO = $_.GPO
                Type = $type
                Item = $itemName
                UserName = $username
                HasCpassword = $hasCpassword
                SecurityFinding = $_.SecurityFinding
            }
        }
    }
}

# Process registry items and merge DefaultPassword/DefaultUserName
$regByKey = $registryItems | Group-Object { "$($_.Domain)_$($_.GPO)_$($_.RegistryKey)" }
foreach ($group in $regByKey) {
    $defaultPass = $group.Group | Where-Object { $_.ValueName -eq 'DefaultPassword' }
    $defaultUser = $group.Group | Where-Object { $_.ValueName -eq 'DefaultUserName' }
    
    if ($defaultPass -and $defaultUser) {
        # Both exist - merge them
        $userValue = $defaultUser[0].ValueData
        $passValue = $defaultPass[0].ValueData
        # Check if DefaultPassword has a non-empty value
        $hasPassword = -not [string]::IsNullOrWhiteSpace($passValue)
        $finding = "Insecure: Plaintext credentials in registry (DefaultUserName and DefaultPassword)"
        $mergedRegistryItems += [pscustomobject]@{
            Domain = $defaultPass[0].Domain
            GPO = $defaultPass[0].GPO
            Type = 'Registry'
            Item = $defaultPass[0].RegistryKey
            UserName = $userValue
            HasCpassword = $hasPassword
            SecurityFinding = $finding
        }
    } else {
        # Process individually
        $group.Group | ForEach-Object {
            $riskyConfigsList += [pscustomobject]@{
                Domain = $_.Domain
                GPO = $_.GPO
                Type = 'Registry'
                Item = "$($_.RegistryKey)\$($_.ValueName)"
                UserName = ''
                HasCpassword = $false
                SecurityFinding = $_.SecurityFinding
            }
        }
    }
}

# Add merged registry items to the list
$riskyConfigsList += $mergedRegistryItems

$html += @"
</div>

<div class='main-section'>
<h1>Risky Configurations</h1>
"@

$html += Build-RiskyGroupedSection $riskyConfigsList

$html += @"
</div>

</div>
</body></html>
"@
    $html | Out-File $report -Encoding UTF8 -ErrorAction Stop
    Write-Host "HTML report generation completed!" -ForegroundColor Green
    Write-Host "Report generated at: $report" -ForegroundColor Green
}
catch {
    $script:ErrorCount++
    Write-Error "HTML report generation failed: $($_.Exception.Message)"
    exit 1
}

#endregion

#region Summary and Completion

Write-Host "" -ForegroundColor Green
Write-Host "✓ Processing complete!" -ForegroundColor Green

if ($script:ErrorCount -gt 0) {
    Write-Host "⚠ Completed with $($script:ErrorCount) error(s)" -ForegroundColor Yellow
}
if ($script:WarningCount -gt 0) {
    Write-Host "⚠ Completed with $($script:WarningCount) warning(s)" -ForegroundColor Yellow
}

Write-Host "" -ForegroundColor Cyan
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Total files processed: $totalFiles" -ForegroundColor White
Write-Host "  Total records collected: $totalRecords" -ForegroundColor White
Write-Host "  Errors encountered: $($script:ErrorCount)" -ForegroundColor $(if ($script:ErrorCount -gt 0) { 'Yellow' } else { 'Green' })

if ($script:ErrorCount -gt 0) {
    exit 1
}

#endregion
