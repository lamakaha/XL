# xlwings System Information Collector

This tool collects detailed system information to help diagnose COM communication errors between Excel and Python when using xlwings. It's designed to gather all relevant configuration parameters, library versions, and system settings that could potentially impact the COM communication.

## Purpose

When experiencing intermittent COM errors with xlwings, it can be difficult to identify the root cause, especially when the issues only occur on certain machines. This script helps by:

1. Collecting comprehensive system information from both working and problematic machines
2. Exporting the data in a CSV format that's easy to compare
3. Gathering information about all components that could affect COM communication

## Requirements

- Windows operating system
- Python 3.6 or higher
- The following Python packages (will be used if available):
  - xlwings
  - pywin32
  - psutil (optional, for hardware information)

## Usage

1. Download both `xlwings_system_info.py` and this README file to your computer.

2. Run the script from the command line:
   ```
   python xlwings_system_info.py [optional_output_filename.csv]
   ```

3. If you don't specify an output filename, the script will create one with the format:
   ```
   xlwings_system_info_{hostname}_{timestamp}.csv
   ```

4. Run this script on both the problematic machine and a working machine.

5. Compare the CSV files to identify differences that might be causing the COM errors.

## What Information is Collected

The script collects the following categories of information:

### Operating System
- Windows version and build
- System architecture
- Hostname and user information

### Python Environment
- Python version and implementation
- Python executable path
- Python architecture (32-bit vs 64-bit)
- Pip version

### Library Versions
- xlwings version
- pywin32 version and components
- comtypes version
- Other Excel-related libraries (pandas, numpy, openpyxl, etc.)

### Excel Information
- Excel version and build
- Installation paths
- Registry settings related to Excel
- Excel security settings
- Excel startup paths
- Excel template paths
- Excel automation security settings
- Excel product code

### Office Versions and Patches
- Office version and build
- Office update channel
- Office SKU information
- Office Click-to-Run configuration
- Office update history
- Windows updates related to Office
- Office patch levels

### COM Configuration
- DCOM settings from registry
- COM+ applications
- COM-related environment variables

### .NET Framework
- Installed .NET Framework versions

### Hardware Information
- Memory usage and capacity
- CPU information
- Disk space

### Office Add-ins
- Installed Excel add-ins (XLA, XLAM, XLL)
- Add-in paths and versions
- COM add-ins (from registry and COM interface)
- VSTO add-ins
- Add-in load behavior and connection status

### xlwings-specific Configuration
- xlwings.conf file contents
- xlwings add-in installation status and path
- xlwings version and dependencies
- xlwings UDF modules
- xlwings PRO license status
- xlwings settings
- xlwings registry entries
- xlwings server status

### Environment Variables
- PATH and other relevant environment variables

## CSV Output Format

The script generates a CSV file with three columns:
- **Group**: The category of the parameter (e.g., python, excel, xlwings)
- **Parameter**: The specific parameter name
- **Value**: The value of the parameter

This format makes it easier to:
- Sort and filter by parameter groups
- Compare specific categories between machines
- Identify differences without repetitive prefixes

## Comparing Results

When comparing the CSV files from different machines:

1. **Version Differences** (look in these groups: python, excel, office, library_versions):
   - Python version and architecture (32-bit vs 64-bit)
   - Excel version and build number
   - Office patch level and update history
   - xlwings version
   - pywin32 version and components
   - Windows version and build

2. **Missing Components** (look in these groups: library_versions, office_addins, xlwings):
   - Required libraries (pywin32, comtypes, etc.)
   - Excel add-ins
   - xlwings add-in installation status

3. **Configuration Differences** (look in these groups: com, excel, office, environment_variables):
   - COM and DCOM settings
   - Excel security settings
   - Excel automation security level
   - Office update channel
   - xlwings settings
   - Environment variables

4. **Add-in Conflicts** (look in these groups: office_addins, com_addin, reg_addin):
   - Installed Excel add-ins
   - COM add-ins and their connection status
   - Add-in load behavior settings
   - VSTO add-ins

5. **System Differences** (look in these groups: hardware, dotnet_framework, os):
   - Hardware resources (memory, CPU)
   - .NET Framework version
   - Office installation type (Click-to-Run vs MSI)

6. **Registry Settings** (look in these groups: registry, user_registry):
   - Excel-related registry entries
   - Office configuration in registry
   - COM registration settings

Common issues that can cause COM errors include:

- **Architecture Mismatch**: Mismatched Python and Excel architectures (32-bit vs 64-bit)
- **Version Incompatibility**: Outdated or incompatible pywin32 version
- **Add-in Conflicts**: Conflicting Excel add-ins or COM add-ins
- **Resource Limitations**: Insufficient system resources (memory, CPU)
- **Permission Issues**: DCOM permission issues or restricted user accounts
- **Security Software**: Antivirus or security software interference
- **Office Updates**: Different Office patch levels or update channels
- **Registry Problems**: Corrupted registry entries for COM components
- **Excel Settings**: Different Excel security or automation settings
- **xlwings Configuration**: Mismatched xlwings settings or add-in versions

## Troubleshooting Tips

If you identify differences between working and non-working machines, try these steps:

1. **Version Alignment**:
   - Update Python, xlwings, and pywin32 to match the versions on working machines
   - Ensure Office/Excel is on the same patch level and update channel
   - Consider downgrading to a known working version if updates caused the issue

2. **Architecture Compatibility**:
   - Ensure Python and Excel architectures match (both 32-bit or both 64-bit)
   - Check that pywin32 matches your Python architecture
   - Reinstall Python and dependencies if architecture issues are detected

3. **Add-in Management**:
   - Disable non-essential Excel add-ins, especially COM add-ins
   - Test with a clean Excel profile (rename the XLSTART folder temporarily)
   - Check for add-in load order conflicts

4. **COM Registration**:
   - Re-register COM components with `regsvr32` (especially pythoncom*.dll)
   - Run `python -m pip install --upgrade --force-reinstall pywin32`
   - Run `python -m pywin32_postinstall -install` as administrator

5. **Excel Settings**:
   - Check Excel Trust Center settings (especially COM add-in trust)
   - Set Excel automation security to low temporarily for testing
   - Disable protected view for Excel files

6. **System Resources**:
   - Close other applications to free up memory
   - Increase virtual memory allocation
   - Check for memory leaks in long-running processes

7. **Permission and Security**:
   - Run Excel and Python with administrator privileges
   - Check Windows User Account Control (UAC) settings
   - Temporarily disable antivirus or add exceptions for Python and Excel

8. **xlwings-specific**:
   - Reinstall the xlwings add-in using `xlwings addin install`
   - Check xlwings.conf settings for compatibility
   - Try using xlwings without UDFs as a test

9. **Office Repair**:
   - Run Office repair from Control Panel
   - Reset Excel settings (delete the Excel15.xlb or Excel16.xlb file)
   - Consider a clean Office installation if problems persist

10. **Registry Cleanup**:
    - Use Registry Editor to check for corrupted COM entries
    - Compare registry settings with working machines
    - Consider exporting/importing specific registry keys from working machines

## Privacy Note

This script collects detailed system information. The CSV file may contain:
- Computer name and username
- File paths that include user names
- System configuration details

Review the CSV file before sharing it to ensure you're comfortable with the information being disclosed.

## License

This tool is provided as-is under the MIT License.
