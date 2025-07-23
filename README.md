# Mephisto's Active Directory User Password Change Script

## Overview

This PowerShell script allows administrators to efficiently manage password resets for multiple Active Directory (AD) users. The tool is interactive and flexible, offering features like user selection, password generation, change logging, and reporting.

## Features

- Retrieve users directly from Active Directory or load them via a CSV file.
- Interactive interface to select/deselect user accounts for password change.
- Enforced strong password generation with a mix of upper/lowercase, numbers, and symbols.
- Flexible password length configuration (single length or ranged).
- Export final password reset report to CSV.
- Supports verbose logging for audit and tracking.
- Fails gracefully with console output if file write operations fail.

## Usage

### 1. Start the Script

Run the script using PowerShell with administrative privileges:

```powershell
.\ADPWChange.ps1
```

### 2. Choose Data Source

You can either:
- Connect to Active Directory to retrieve users directly, **or**
- Provide a CSV containing a list of users and selection states.

CSV Format:
```csv
AccountName,LastPWChange,UserEnabled,Included
jdoe,2023-01-10,True,y
asmith,2023-02-01,True,n
```

### 3. Review and Select Users

Navigate using the interactive interface:
- `+ / -`: Next / Previous user
- `x`: Toggle inclusion
- `?`: Search by account name
- `#`: Jump to index
- `S`: Summary
- `W`: Export user list to CSV
- `.`: Proceed to reset passwords

### 4. Configure Password Policy

- Choose to apply a fixed password length or a range (8 to 100 characters).
- The script ensures at least one character from each class (uppercase, lowercase, number, symbol).

### 5. Logging

- Enable verbose logging to output password changes to a timestamped file in the script's directory.

### 6. Export Results

Once complete, the script can write the results to a uniquely named CSV file (based on timestamp).

## Requirements

- Windows PowerShell 5.1+
- Active Directory PowerShell Module
- Sufficient permissions to read/write AD user account data

## Example

```powershell
Would you like to connect to Active Directory to get a list of users? (y,n): y
Enable verbose change logging? (y,n): y
Would you like to write changes to Active Directory? (y,n): y
```

## Notes

- This script must be run by a user with sufficient privileges to perform password resets.
- Ensure the Active Directory module is installed: `Install-WindowsFeature RSAT-AD-PowerShell`

**Author**: Mephistopheles  
**GitHub**: [https://github.com/mephistopheles](https://github.com/mephistopheles)  
**Support**: Contributions, issues, and PRs welcome.
