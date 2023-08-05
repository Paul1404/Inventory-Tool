# 🖥️ System Information Collection Script

This PowerShell script collects various system information and appends it to an inventory file in either JSON or CSV format. The system information collected includes:

- 💻 Computer Information
- 🎛️ CPU Information
- 💾 RAM Information
- 🖥️ Motherboard Information
- 💿 Drive Information

## 🌟 Features

- Custom PowerShell functions for fetching system information.
- Uses CIM (Common Information Model) instances for gathering data.
- Progress bar display while data is being fetched.
- Creates inventory files if they don't already exist.
- Handles both JSON and CSV output formats.
- Exception handling and colorful console messages for better user interaction.

## 🚀 Usage

Just run the script in PowerShell. It will fetch the system information and either create a new inventory file or append the data to an existing one.

At the start of the script, you are prompted to select the output format. Choose `1` for JSON or `2` for CSV. If you don't select any option and press Enter, the default JSON format is chosen.

🛠️ Custom Functions

This script depends on several custom functions, including:

- Get-ComputerInfo
- Get-CPUInfo
- Get-RAMInfo
- Get-MotherboardInfo
- Get-DriveInfo
- Update-InventoryFile
- Write-Info
- Write-Success
- Write-Error
- Write-InlineProgress

Ensure these functions are defined and accessible before running this script.

📋 Requirements

The script uses PowerShell's CIM cmdlets which are built into PowerShell version 3 and later, and are available on all Windows operating systems starting with Windows Server 2012 and Windows 8.

🗂️ Inventory File Format

The format of the inventory file can be either JSON or CSV, as specified at the beginning of the script execution.

The inventory file includes the following fields for each system:

- Timestamp
- ComputerName
- CPU
- TotalRAM
- Motherboard
- Drive

For instance, a JSON entry might look like this:

```json
{
    "TimeStamp": "2023-08-06 12:34:56",
    "ComputerName": "DESKTOP-12345",
    "CPU": "Intel(R) Core(TM) i7-9700K CPU @ 3.60GHz",
    "TotalRAM": "16.00 GB",
    "Motherboard": "ROG MAXIMUS XI HERO (WI-FI)",
    "Drive": "931.51 GB SSD, 1863.02 GB HDD"
}
```

## 📤 Output

After successful execution, the script will append the system information to the inventory file in the chosen format and output a success message.

## ❗ Error Handling

In case of any error, the script will output an error message with the details of the exception that occurred.

## ⚠️ Note

The script should be run with the necessary permissions to access the system information and to create/write files in the script's directory.
