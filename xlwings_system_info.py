#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlwings_system_info.py - Collect system information for xlwings COM error troubleshooting

This script collects detailed system information that could be relevant to diagnosing
COM communication errors between Excel and Python when using xlwings.

Usage:
    python xlwings_system_info.py [output_filename.csv]

If no output filename is provided, it will default to:
    xlwings_system_info_{hostname}_{timestamp}.csv
"""

import os
import sys
import csv
import platform
import socket
import datetime
import subprocess
import ctypes
import winreg
import traceback
import logging
from pathlib import Path
import importlib.metadata as metadata

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def get_python_info():
    """Get Python version and installation details."""
    info = {
        "python_version": platform.python_version(),
        "python_implementation": platform.python_implementation(),
        "python_compiler": platform.python_compiler(),
        "python_build": " ".join(platform.python_build()),
        "python_executable": sys.executable,
        "python_path": sys.path,
        "python_64bit": platform.architecture()[0] == "64bit",
    }

    # Get pip version
    try:
        pip_version = subprocess.check_output(
            [sys.executable, "-m", "pip", "--version"],
            universal_newlines=True
        ).strip()
        info["pip_version"] = pip_version
    except Exception as e:
        info["pip_version"] = f"Error getting pip version: {str(e)}"

    return info

def get_os_info():
    """Get operating system information."""
    info = {
        "os_name": platform.system(),
        "os_release": platform.release(),
        "os_version": platform.version(),
        "os_platform": platform.platform(),
        "os_machine": platform.machine(),
        "os_processor": platform.processor(),
        "hostname": socket.gethostname(),
        "username": os.environ.get("USERNAME", "Unknown"),
        "domain": os.environ.get("USERDOMAIN", "Unknown"),
    }

    # Get Windows version details
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        info["windows_major_version"] = kernel32.GetVersion() & 0xFF
        info["windows_minor_version"] = (kernel32.GetVersion() >> 8) & 0xFF
        info["windows_build"] = kernel32.GetVersion() >> 16
    except Exception as e:
        info["windows_version_details"] = f"Error getting Windows version details: {str(e)}"

    return info

def get_package_version(package_name):
    """Get the version of an installed package."""
    try:
        return metadata.version(package_name)
    except metadata.PackageNotFoundError:
        return "Not installed"
    except Exception as e:
        return f"Error: {str(e)}"

def get_library_versions():
    """Get versions of relevant libraries."""
    libraries = [
        "xlwings", "pywin32", "comtypes", "numpy", "pandas",
        "openpyxl", "win32com", "pythoncom", "pywintypes",
        "pytz", "psutil", "pyxll", "pyexcel", "xlrd", "xlwt"
    ]

    versions = {}
    for lib in libraries:
        versions[f"{lib}_version"] = get_package_version(lib)

    # Special case for pywin32 components
    try:
        import win32api
        versions["win32api_file_version"] = win32api.__file__
    except Exception:
        versions["win32api_file_version"] = "Error importing win32api"

    try:
        import pythoncom
        versions["pythoncom_file"] = pythoncom.__file__
    except Exception:
        versions["pythoncom_file"] = "Error importing pythoncom"

    return versions

def get_excel_info():
    """Get Excel version and installation details."""
    info = {}

    # Try to get Excel version from registry
    try:
        excel_paths = []
        excel_versions = []

        # Check common registry paths for Office/Excel
        reg_paths = [
            r"SOFTWARE\Microsoft\Office",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs",
            r"SOFTWARE\Microsoft\Office\16.0\Excel",
            r"SOFTWARE\Microsoft\Office\15.0\Excel",
            r"SOFTWARE\Microsoft\Office\14.0\Excel",
            r"SOFTWARE\Wow6432Node\Microsoft\Office"
        ]

        for reg_path in reg_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                    info[f"registry_{reg_path.replace('\\', '_')}"] = "Key exists"

                    # Try to enumerate subkeys
                    try:
                        i = 0
                        while True:
                            subkey_name = winreg.EnumKey(key, i)
                            if "office" in subkey_name.lower() or "excel" in subkey_name.lower():
                                info[f"registry_{reg_path.replace('\\', '_')}_{subkey_name}"] = "Found"

                                # Try to get Excel path
                                try:
                                    with winreg.OpenKey(key, subkey_name) as subkey:
                                        try:
                                            path, _ = winreg.QueryValueEx(subkey, "Path")
                                            if path and "excel" in path.lower():
                                                excel_paths.append(path)
                                        except:
                                            pass
                                except:
                                    pass

                                # Try to get version
                                try:
                                    with winreg.OpenKey(key, subkey_name) as subkey:
                                        try:
                                            version, _ = winreg.QueryValueEx(subkey, "Version")
                                            if version:
                                                excel_versions.append(version)
                                        except:
                                            pass
                                except:
                                    pass
                            i += 1
                    except WindowsError:
                        # No more subkeys
                        pass

                    # Try to get values directly from the key
                    try:
                        i = 0
                        while True:
                            name, value, _ = winreg.EnumValue(key, i)
                            # Only include relevant values to avoid too much noise
                            if any(term in name.lower() for term in ["version", "build", "update", "patch", "excel", "office"]):
                                info[f"registry_{reg_path.replace('\\', '_')}_{name}"] = str(value)
                            i += 1
                    except WindowsError:
                        # No more values
                        pass
            except Exception as e:
                info[f"registry_{reg_path.replace('\\', '_')}"] = f"Key not found or error: {str(e)}"

        # Check for Office Click-to-Run configuration
        try:
            c2r_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, c2r_path) as key:
                for value_name in ["ClientCulture", "Platform", "ProductReleaseIds", "UpdateChannel", "UpdatesEnabled", "VersionToReport"]:
                    try:
                        value, _ = winreg.QueryValueEx(key, value_name)
                        info[f"office_c2r_{value_name}"] = str(value)
                    except:
                        pass
        except:
            pass

        # Check for Office update history
        try:
            update_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Updates"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, update_path) as key:
                i = 0
                try:
                    while True:
                        name, value, _ = winreg.EnumValue(key, i)
                        info[f"office_updates_{name}"] = str(value)
                        i += 1
                except WindowsError:
                    # No more values
                    pass
        except:
            pass

        # Check current user registry for Excel settings
        try:
            excel_user_paths = [
                r"Software\Microsoft\Office\16.0\Excel",
                r"Software\Microsoft\Office\15.0\Excel",
                r"Software\Microsoft\Office\14.0\Excel",
                r"Software\Microsoft\Office\Excel"
            ]

            for path in excel_user_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                        info[f"user_registry_{path.replace('\\', '_')}"] = "Key exists"

                        # Try to get security settings
                        try:
                            with winreg.OpenKey(key, "Security") as security_key:
                                i = 0
                                try:
                                    while True:
                                        name, value, _ = winreg.EnumValue(security_key, i)
                                        info[f"excel_security_{name}"] = str(value)
                                        i += 1
                                except WindowsError:
                                    # No more values
                                    pass
                        except:
                            pass

                        # Try to get options
                        try:
                            with winreg.OpenKey(key, "Options") as options_key:
                                for option in ["DontUpdateLinks", "DisableAutoRepublish", "DisableLivePreview"]:
                                    try:
                                        value, _ = winreg.QueryValueEx(options_key, option)
                                        info[f"excel_option_{option}"] = str(value)
                                    except:
                                        pass
                        except:
                            pass
                except:
                    pass
        except Exception as e:
            info["excel_user_registry_error"] = str(e)

        info["excel_paths_from_registry"] = "; ".join(excel_paths) if excel_paths else "Not found"
        info["excel_versions_from_registry"] = "; ".join(excel_versions) if excel_versions else "Not found"

        # Try to find Excel executable
        excel_exe_paths = []
        common_paths = [
            r"C:\Program Files\Microsoft Office",
            r"C:\Program Files (x86)\Microsoft Office",
            r"C:\Program Files\Microsoft Office 15",
            r"C:\Program Files\Microsoft Office 16",
            r"C:\Program Files\Microsoft Office\root\Office16",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16",
        ]

        for base_path in common_paths:
            if os.path.exists(base_path):
                for root, dirs, files in os.walk(base_path):
                    for file in files:
                        if file.lower() == "excel.exe":
                            excel_exe_paths.append(os.path.join(root, file))

                            # Try to get file version info
                            try:
                                import win32api
                                info["excel_exe_file_version"] = win32api.GetFileVersionInfo(os.path.join(root, file), '\\')
                                version_info = win32api.GetFileVersionInfo(os.path.join(root, file), '\\StringFileInfo\\040904B0\\FileVersion')
                                info["excel_exe_file_version_string"] = version_info
                            except:
                                pass

        info["excel_exe_paths"] = "; ".join(excel_exe_paths) if excel_exe_paths else "Not found"

    except Exception as e:
        info["excel_registry_error"] = str(e)

    # Try to get Excel version using COM
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        info["excel_version_com"] = excel.Version
        info["excel_build_com"] = excel.Build

        # Get more detailed Excel information
        try:
            info["excel_product_code"] = excel.ProductCode
        except:
            pass

        try:
            info["excel_path"] = excel.Path
        except:
            pass

        try:
            info["excel_library_path"] = excel.LibraryPath
        except:
            pass

        try:
            info["excel_template_path"] = excel.TemplatePath
        except:
            pass

        try:
            info["excel_startup_path"] = excel.StartupPath
        except:
            pass

        try:
            info["excel_alt_startup_path"] = excel.AltStartupPath
        except:
            pass

        try:
            info["excel_user_name"] = excel.UserName
        except:
            pass

        try:
            info["excel_automation_security"] = excel.AutomationSecurity
        except:
            pass

        excel.Quit()
    except Exception as e:
        info["excel_version_com_error"] = str(e)

    # Try to get Office patch information using PowerShell
    try:
        ps_command = "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | Where-Object {$_.DisplayName -like '*Microsoft Office*' -or $_.DisplayName -like '*Microsoft 365*'} | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | ConvertTo-Csv -NoTypeInformation"
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.stdout:
            info["office_installed_products"] = result.stdout.replace('\r\n', '; ')
    except Exception as e:
        info["office_powershell_error"] = str(e)

    # Try to get Office update history using PowerShell
    try:
        ps_command = "Get-ItemProperty HKLM:\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Updates\\* | Select-Object -Property * | ConvertTo-Csv -NoTypeInformation"
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.stdout:
            info["office_update_history"] = result.stdout.replace('\r\n', '; ')
    except Exception as e:
        info["office_update_history_error"] = str(e)

    return info

def get_com_info():
    """Get COM-related configuration."""
    info = {}

    # Check DCOM configuration in registry
    try:
        dcom_keys = [
            r"SOFTWARE\Microsoft\Ole",
            r"SOFTWARE\Microsoft\COM3",
            r"SOFTWARE\Microsoft\DCOM",
        ]

        for key_path in dcom_keys:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                    i = 0
                    try:
                        while True:
                            name, value, type = winreg.EnumValue(key, i)
                            info[f"dcom_{key_path.replace('\\', '_')}_{name}"] = str(value)
                            i += 1
                    except WindowsError:
                        # No more values
                        pass
            except Exception as e:
                info[f"dcom_{key_path.replace('\\', '_')}_error"] = str(e)
    except Exception as e:
        info["dcom_registry_error"] = str(e)

    # Check COM+ applications
    try:
        result = subprocess.run(
            ["dcomcnfg", "/s"],
            capture_output=True,
            text=True,
            timeout=5
        )
        info["dcomcnfg_output"] = result.stdout if result.stdout else "No output"
        info["dcomcnfg_error"] = result.stderr if result.stderr else "No error"
    except Exception as e:
        info["dcomcnfg_error"] = str(e)

    return info

def get_environment_variables():
    """Get relevant environment variables."""
    relevant_vars = [
        "PATH", "PYTHONPATH", "COMSPEC", "DCOM_MACHINE_LAUNCH_ACCESS",
        "TEMP", "TMP", "APPDATA", "LOCALAPPDATA", "PROGRAMDATA",
        "PROCESSOR_ARCHITECTURE", "NUMBER_OF_PROCESSORS", "PROCESSOR_IDENTIFIER",
        "PATHEXT", "COMPUTERNAME", "USERNAME", "USERPROFILE", "SYSTEMROOT",
        "SYSTEMDRIVE", "WINDIR", "XLSTART", "XLSTARTUP"
    ]

    env_vars = {}
    for var in relevant_vars:
        env_vars[f"env_{var}"] = os.environ.get(var, "Not set")

    return env_vars

def get_dotnet_framework_version():
    """Get .NET Framework version."""
    info = {}

    try:
        # Check registry for installed .NET versions
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\NET Framework Setup\NDP") as key:
            i = 0
            try:
                while True:
                    subkey_name = winreg.EnumKey(key, i)
                    if subkey_name.startswith("v"):
                        info[f"dotnet_{subkey_name}"] = "Installed"

                        # For .NET 4.x, check additional details
                        if subkey_name.startswith("v4"):
                            try:
                                with winreg.OpenKey(key, f"{subkey_name}\\Full") as subkey:
                                    release, _ = winreg.QueryValueEx(subkey, "Release")
                                    if release:
                                        # Convert release number to version
                                        if release >= 528040:
                                            info["dotnet_4_version"] = "4.8 or later"
                                        elif release >= 461808:
                                            info["dotnet_4_version"] = "4.7.2"
                                        elif release >= 461308:
                                            info["dotnet_4_version"] = "4.7.1"
                                        elif release >= 460798:
                                            info["dotnet_4_version"] = "4.7"
                                        elif release >= 394802:
                                            info["dotnet_4_version"] = "4.6.2"
                                        elif release >= 394254:
                                            info["dotnet_4_version"] = "4.6.1"
                                        elif release >= 393295:
                                            info["dotnet_4_version"] = "4.6"
                                        elif release >= 379893:
                                            info["dotnet_4_version"] = "4.5.2"
                                        elif release >= 378675:
                                            info["dotnet_4_version"] = "4.5.1"
                                        elif release >= 378389:
                                            info["dotnet_4_version"] = "4.5"
                                        else:
                                            info["dotnet_4_version"] = f"4.0 or earlier (release {release})"
                            except Exception:
                                pass
                    i += 1
            except WindowsError:
                # No more subkeys
                pass
    except Exception as e:
        info["dotnet_registry_error"] = str(e)

    return info

def get_hardware_info():
    """Get system hardware information."""
    info = {}

    # Get memory information
    try:
        import psutil
        vm = psutil.virtual_memory()
        info["total_memory_gb"] = round(vm.total / (1024**3), 2)
        info["available_memory_gb"] = round(vm.available / (1024**3), 2)
        info["memory_percent_used"] = vm.percent

        # CPU information
        info["cpu_count_physical"] = psutil.cpu_count(logical=False)
        info["cpu_count_logical"] = psutil.cpu_count(logical=True)
        info["cpu_percent"] = psutil.cpu_percent(interval=1)

        # Disk information
        disk = psutil.disk_usage('/')
        info["disk_total_gb"] = round(disk.total / (1024**3), 2)
        info["disk_free_gb"] = round(disk.free / (1024**3), 2)
        info["disk_percent_used"] = disk.percent
    except ImportError:
        info["psutil_error"] = "psutil module not installed"
    except Exception as e:
        info["hardware_info_error"] = str(e)

    return info

def get_office_addins():
    """Get information about installed Office/Excel add-ins."""
    info = {}

    # Check common add-in locations
    addin_paths = [
        os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Excel", "XLSTART"),
        os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "AddIns"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "Office16", "Library"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "root", "Office16", "Library"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft Office", "Office16", "Library"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft Office", "root", "Office16", "Library"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "root", "Office16", "XLSTART"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft Office", "root", "Office16", "XLSTART"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "Office16", "XLSTART"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft Office", "Office16", "XLSTART"),
    ]

    addin_files = []
    for path in addin_paths:
        if os.path.exists(path):
            for file in os.listdir(path):
                if file.endswith((".xla", ".xlam", ".xll", ".xlb")):
                    addin_files.append(os.path.join(path, file))

                    # Try to get file version info for add-ins
                    if file.endswith((".xll")):
                        try:
                            import win32api
                            version_info = win32api.GetFileVersionInfo(os.path.join(path, file), '\\StringFileInfo\\040904B0\\FileVersion')
                            info[f"addin_version_{file}"] = version_info
                        except:
                            pass

    info["excel_addins"] = "; ".join(addin_files) if addin_files else "None found"

    # Check registry for COM add-ins
    try:
        com_addin_paths = [
            r"Software\Microsoft\Office\Excel\Addins",
            r"Software\Microsoft\Office\16.0\Excel\Addins",
            r"Software\Microsoft\Office\15.0\Excel\Addins",
            r"Software\Microsoft\Office\14.0\Excel\Addins",
            r"Software\Wow6432Node\Microsoft\Office\Excel\Addins",
            r"Software\Wow6432Node\Microsoft\Office\16.0\Excel\Addins",
            r"Software\Wow6432Node\Microsoft\Office\15.0\Excel\Addins",
            r"Software\Wow6432Node\Microsoft\Office\14.0\Excel\Addins"
        ]

        registry_addins = []

        # Check HKCU
        for path in com_addin_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                    i = 0
                    try:
                        while True:
                            addin_name = winreg.EnumKey(key, i)
                            registry_addins.append(f"HKCU\\{path}\\{addin_name}")

                            # Get addin details
                            try:
                                with winreg.OpenKey(key, addin_name) as addin_key:
                                    for value_name in ["FriendlyName", "Description", "LoadBehavior"]:
                                        try:
                                            value, _ = winreg.QueryValueEx(addin_key, value_name)
                                            info[f"reg_addin_{addin_name}_{value_name}"] = str(value)
                                        except:
                                            pass
                            except:
                                pass

                            i += 1
                    except WindowsError:
                        # No more subkeys
                        pass
            except:
                pass

        # Check HKLM
        for path in com_addin_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                    i = 0
                    try:
                        while True:
                            addin_name = winreg.EnumKey(key, i)
                            registry_addins.append(f"HKLM\\{path}\\{addin_name}")

                            # Get addin details
                            try:
                                with winreg.OpenKey(key, addin_name) as addin_key:
                                    for value_name in ["FriendlyName", "Description", "LoadBehavior"]:
                                        try:
                                            value, _ = winreg.QueryValueEx(addin_key, value_name)
                                            info[f"reg_addin_{addin_name}_{value_name}"] = str(value)
                                        except:
                                            pass
                            except:
                                pass

                            i += 1
                    except WindowsError:
                        # No more subkeys
                        pass
            except:
                pass

        info["registry_com_addins"] = "; ".join(registry_addins) if registry_addins else "None found"
    except Exception as e:
        info["registry_com_addins_error"] = str(e)

    # Try to get add-ins via COM
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        # Regular add-ins
        com_addins = []
        for addin in excel.AddIns:
            try:
                com_addins.append(f"{addin.Name} - {addin.Path} - {'Installed' if addin.Installed else 'Not Installed'}")
            except:
                pass

        info["excel_com_addins"] = "; ".join(com_addins) if com_addins else "None found"

        # COM add-ins (different from regular add-ins)
        try:
            com_addins_detailed = []
            for com_addin in excel.COMAddIns:
                try:
                    com_addins_detailed.append(f"{com_addin.Description} - {com_addin.progID} - {'Connected' if com_addin.Connect else 'Disconnected'}")
                    info[f"com_addin_{com_addin.progID}_description"] = com_addin.Description
                    info[f"com_addin_{com_addin.progID}_connected"] = "Yes" if com_addin.Connect else "No"
                except:
                    pass

            info["excel_detailed_com_addins"] = "; ".join(com_addins_detailed) if com_addins_detailed else "None found"
        except Exception as e:
            info["excel_detailed_com_addins_error"] = str(e)

        excel.Quit()
    except Exception as e:
        info["excel_com_addins_error"] = str(e)

    # Try to get VSTO add-ins from registry
    try:
        vsto_paths = [
            r"Software\Microsoft\VSTO\SolutionMetadata",
            r"Software\Microsoft\VSTO_DRMClient"
        ]

        vsto_addins = []

        for path in vsto_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                    i = 0
                    try:
                        while True:
                            addin_name = winreg.EnumKey(key, i)
                            if "excel" in addin_name.lower():
                                vsto_addins.append(addin_name)
                            i += 1
                    except WindowsError:
                        # No more subkeys
                        pass
            except:
                pass

        info["vsto_addins"] = "; ".join(vsto_addins) if vsto_addins else "None found"
    except Exception as e:
        info["vsto_addins_error"] = str(e)

    return info

def get_xlwings_specific_info():
    """Get xlwings-specific configuration."""
    info = {}

    # Check for xlwings.conf file
    config_paths = [
        os.path.join(os.getcwd(), "xlwings.conf"),
        os.path.join(os.path.expanduser("~"), ".xlwings", "xlwings.conf"),
    ]

    for path in config_paths:
        if os.path.exists(path):
            try:
                with open(path, "r") as f:
                    content = f.read()
                info[f"xlwings_conf_{path.replace('\\', '_')}"] = content
            except Exception as e:
                info[f"xlwings_conf_{path.replace('\\', '_')}_error"] = str(e)

    # Check for xlwings add-in
    try:
        import xlwings
        info["xlwings_version"] = xlwings.__version__
        info["xlwings_path"] = xlwings.__path__[0]

        # Get xlwings dependencies
        try:
            info["xlwings_dependencies"] = str(xlwings.__dependencies__)
        except:
            pass

        # Check if xlwings add-in is installed
        addin_paths = [
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Excel", "XLSTART", "xlwings.xlam"),
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "AddIns", "xlwings.xlam"),
            os.path.join(xlwings.__path__[0], "addin", "xlwings.xlam")
        ]

        for path in addin_paths:
            if os.path.exists(path):
                info["xlwings_addin_installed"] = True
                info["xlwings_addin_path"] = path

                # Try to get file modification time
                try:
                    mod_time = os.path.getmtime(path)
                    info["xlwings_addin_modified"] = datetime.datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass

                break
        else:
            info["xlwings_addin_installed"] = False

        # Check for UDF modules
        try:
            udf_modules = []
            if hasattr(xlwings, "UDF_MODULES"):
                udf_modules = xlwings.UDF_MODULES
            info["xlwings_udf_modules"] = str(udf_modules)
        except:
            pass

        # Check for xlwings PRO license
        try:
            if hasattr(xlwings, "pro"):
                info["xlwings_pro_available"] = True
                try:
                    info["xlwings_license_key"] = "Present (details hidden for security)"
                except:
                    info["xlwings_license_key"] = "Not found"
            else:
                info["xlwings_pro_available"] = False
        except:
            pass

        # Try to get xlwings settings
        try:
            settings = {}
            if hasattr(xlwings, "settings"):
                for setting in dir(xlwings.settings):
                    if not setting.startswith("_"):
                        try:
                            value = getattr(xlwings.settings, setting)
                            if not callable(value):
                                settings[setting] = str(value)
                        except:
                            pass
            info["xlwings_settings"] = str(settings)
        except:
            pass

    except ImportError:
        info["xlwings_error"] = "xlwings module not installed"
    except Exception as e:
        info["xlwings_info_error"] = str(e)

    # Check registry for xlwings-related entries
    try:
        xlwings_reg_paths = [
            r"Software\Python\XLWings",
            r"Software\XLWings"
        ]

        for path in xlwings_reg_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                    i = 0
                    try:
                        while True:
                            name, value, _ = winreg.EnumValue(key, i)
                            info[f"xlwings_registry_{name}"] = str(value)
                            i += 1
                    except WindowsError:
                        # No more values
                        pass
            except:
                pass
    except Exception as e:
        info["xlwings_registry_error"] = str(e)

    # Check for Python UDF server
    try:
        # Check if xlwings server is running
        import psutil
        python_processes = [p for p in psutil.process_iter() if "python" in p.name().lower()]
        for proc in python_processes:
            try:
                cmd_line = proc.cmdline()
                if any("xlwings" in arg.lower() for arg in cmd_line) and any("serve" in arg.lower() for arg in cmd_line):
                    info["xlwings_server_running"] = True
                    info["xlwings_server_pid"] = proc.pid
                    info["xlwings_server_cmdline"] = str(cmd_line)
                    break
            except:
                pass
        else:
            info["xlwings_server_running"] = False
    except:
        pass

    return info

def get_office_patches():
    """Get information about Office patches and updates."""
    info = {}

    # Try to get Office patch information using PowerShell
    try:
        # Get Windows Update history for Office updates
        ps_command = """
        Get-WmiObject -Class Win32_QuickFixEngineering |
        Where-Object { $_.Description -like '*Office*' -or $_.HotFixID -like '*KB*' } |
        Select-Object HotFixID, Description, InstalledOn |
        ConvertTo-Csv -NoTypeInformation
        """
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=15
        )
        if result.stdout:
            info["windows_office_updates"] = result.stdout.replace('\r\n', '; ')
    except Exception as e:
        info["windows_office_updates_error"] = str(e)

    # Try to get Office patch level from registry
    try:
        office_versions = ["16.0", "15.0", "14.0"]
        for version in office_versions:
            patch_path = f"SOFTWARE\\Microsoft\\Office\\{version}\\Common\\ProductVersion"
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, patch_path) as key:
                    i = 0
                    try:
                        while True:
                            name, value, _ = winreg.EnumValue(key, i)
                            info[f"office_{version}_patch_{name}"] = str(value)
                            i += 1
                    except WindowsError:
                        # No more values
                        pass
            except:
                pass
    except Exception as e:
        info["office_patch_registry_error"] = str(e)

    # Try to get Office update channel information
    try:
        channel_paths = [
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"SOFTWARE\Microsoft\Office\16.0\Common\ProductVersion",
            r"SOFTWARE\Microsoft\Office\15.0\Common\ProductVersion"
        ]

        for path in channel_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                    for value_name in ["UpdateChannel", "UpdateBranch", "UpdateUrl", "VersionToReport"]:
                        try:
                            value, _ = winreg.QueryValueEx(key, value_name)
                            info[f"update_channel_{path.replace('\\', '_')}_{value_name}"] = str(value)
                        except:
                            pass
            except:
                pass
    except Exception as e:
        info["office_update_channel_error"] = str(e)

    # Try to get Office C2R configuration
    try:
        ps_command = """
        $c2r = New-Object -ComObject Microsoft.Office.ClickToRun.ClickToRun
        $config = @{
            'InstallPath' = $c2r.InstallPath
            'Platform' = $c2r.Platform
            'ProductReleaseIds' = $c2r.ProductReleaseIds
            'VersionToReport' = $c2r.VersionToReport
            'ClientCulture' = $c2r.ClientCulture
        }
        $config | ConvertTo-Json
        """
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.stdout and not "Exception" in result.stdout:
            info["office_c2r_config"] = result.stdout.replace('\r\n', ' ')
    except Exception as e:
        info["office_c2r_config_error"] = str(e)

    # Try to get Office SKU information
    try:
        ps_command = """
        Get-ItemProperty HKLM:\\Software\\Microsoft\\Office\\ClickToRun\\Configuration |
        Select-Object ProductReleaseIds |
        ConvertTo-Csv -NoTypeInformation
        """
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.stdout:
            info["office_sku_info"] = result.stdout.replace('\r\n', '; ')
    except Exception as e:
        info["office_sku_info_error"] = str(e)

    return info

def collect_all_info():
    """Collect all system information."""
    all_info = {}

    # Add timestamp and collection info
    all_info["info_collected_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    all_info["hostname"] = socket.gethostname()

    # Collect all information
    collectors = [
        ("os_info", get_os_info),
        ("python_info", get_python_info),
        ("library_versions", get_library_versions),
        ("excel_info", get_excel_info),
        ("com_info", get_com_info),
        ("environment_variables", get_environment_variables),
        ("dotnet_framework", get_dotnet_framework_version),
        ("hardware_info", get_hardware_info),
        ("office_addins", get_office_addins),
        ("office_patches", get_office_patches),
        ("xlwings_specific", get_xlwings_specific_info),
    ]

    for prefix, collector_func in collectors:
        try:
            logger.info(f"Collecting {prefix}...")
            info = collector_func()
            for key, value in info.items():
                all_info[f"{prefix}_{key}"] = value
        except Exception as e:
            logger.error(f"Error collecting {prefix}: {e}")
            all_info[f"{prefix}_error"] = str(e)
            all_info[f"{prefix}_traceback"] = traceback.format_exc()

    return all_info

def save_to_csv(info, filename=None):
    """Save collected information to a CSV file."""
    if not filename:
        hostname = socket.gethostname()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"xlwings_system_info_{hostname}_{timestamp}.csv"

    logger.info(f"Saving information to {filename}...")

    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Group", "Parameter", "Value"])

            # Process and sort the data
            processed_data = []
            for key, value in info.items():
                # Handle complex values like lists
                if isinstance(value, (list, dict, tuple)):
                    value = str(value)

                # Split the key into group and parameter
                if "_" in key:
                    # Find the first underscore that's not part of a prefix like "office_c2r_"
                    parts = key.split("_", 1)
                    group = parts[0]
                    param = parts[1]

                    # Handle special cases with common prefixes
                    special_prefixes = ["office_c2r", "excel_com", "excel_exe", "excel_option",
                                       "excel_security", "com_addin", "reg_addin", "update_channel",
                                       "xlwings_addin", "xlwings_conf", "xlwings_registry",
                                       "office_updates", "registry_com", "user_registry"]

                    for prefix in special_prefixes:
                        if key.startswith(prefix + "_"):
                            group = prefix
                            param = key[len(prefix) + 1:]
                            break
                else:
                    group = "general"
                    param = key

                processed_data.append((group, param, value))

            # Sort by group then parameter
            for group, param, value in sorted(processed_data):
                writer.writerow([group, param, value])

        logger.info(f"Information saved to {filename}")
        return filename
    except Exception as e:
        logger.error(f"Error saving to CSV: {e}")
        return None

def main():
    """Main function."""
    logger.info("Starting xlwings system information collection...")

    # Get output filename from command line if provided
    output_filename = None
    if len(sys.argv) > 1:
        output_filename = sys.argv[1]

    # Collect all information
    all_info = collect_all_info()

    # Save to CSV
    saved_file = save_to_csv(all_info, output_filename)

    if saved_file:
        print(f"\nSystem information collected and saved to: {saved_file}")
        print(f"Total parameters collected: {len(all_info)}")
    else:
        print("\nError saving system information. Check the logs for details.")

if __name__ == "__main__":
    main()
