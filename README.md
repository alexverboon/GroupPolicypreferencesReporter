# Group Policy Preferences Reporter

A PowerShell script that analyzes Group Policy Preferences (GPP) from GPO backup files and generates an interactive HTML report with detailed insights.

## Overview

This repository contains a powerful automation tool for analyzing Group Policy Preferences across your environment. The `Create-GPPReport.ps1` script parses preference XML files from GPO backups, categorizes them across different preference types, and produces an interactive HTML report with search and filter capabilities.

### Preference Types Supported (18 Types)

**Windows Settings:**

- Registry Settings
- Shortcuts
- Files
- Folders
- INI Files
- Environment Variables

**Control Panel Settings:**

- Drive Mappings
- Printers
- Local Groups
- Local Users
- Network Shares
- Data Sources
- Devices
- Internet Settings
- Power Options
- Folder Options
- Services
- Scheduled Tasks

## Requirements

- **PowerShell 5.1** or higher
- **Read Access** to GPO backup folder structure
- **Write Access** to output directory
- GPO backup files with standard sysvol structure

## Creating GPO Backups

Before running the report, you need to create GPO backups. You can do this in two ways:

### Option 1: Using Group Policy Management Console (GPMC)

1. Open Group Policy Management Console
2. Right-click on the GPO you want to backup
3. Select "Back Up..."
4. Choose the backup location

### Option 2: Using Backup-GPO.ps1 Script (Included)

This repository includes a `Backup-GPO.ps1` script for automated GPO backups:

```powershell
.\Backup-GPO.ps1 -GPOBackupPath "C:\Temp\GPOBackup"
```

## Usage

### Basic Usage (Generate HTML Report Only)

```powershell
.\Create-GPPReport.ps1 -GPOBackupRoot "C:\Temp\GPOBackup"-OutputDir "C:\Temp\Reports"
```

### With CSV Export

```powershell
.\Create-GPPReport.ps1 -GPOBackupRoot "C:\Temp\GPOBackup"-OutputDir "C:\Temp\Reports" -ExportCSV
```

### With JSON Export

```powershell
.\Create-GPPReport.ps1 -GPOBackupRoot "C:\Temp\GPOBackup"-OutputDir "C:\Temp\Reports" -ExportJSON
```

## Author

Alex Verboon

## Contributing

Contributions, issues, and feature requests are welcome!

## References

- [Group Policy Preferences Documentation](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/grouppolicy/group-policy-preferences)

- [MS14-025: Vulnerability in Group Policy Preferences could allow elevation of privilege: May 13, 2014](https://support.microsoft.com/en-us/topic/ms14-025-vulnerability-in-group-policy-preferences-could-allow-elevation-of-privilege-may-13-2014-60734e15-af79-26ca-ea53-8cd617073c30)

- [Unsecured Credentials: Group Policy Preferences](https://attack.mitre.org/techniques/T1552/006/)

## Changelog

### Version 1.1 (2026-01-17)

- Merged DefaultPassword and DefaultUserName registry checks into single Risky Configuration entry
- Renamed ODBC findings to "Data Source stored credentials"
- Removed AutoAdminLogon check from security findings
- Enhanced Risky Configurations display with conditional "Has cpassword" column
- Excluded ACL and admin group findings from cpassword column display
- Improved DefaultPassword validation to detect plaintext credentials

### Version 1.0 (2026-01-17)

- Initial release
- Support for 18 preference types
- Interactive HTML report with search filters
- Optional CSV or JSON export functionality
