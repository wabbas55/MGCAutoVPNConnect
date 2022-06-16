import cx_Freeze
import sys

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("mentor_vpn_connect.py", base=base, icon="icon2.ico")]

# Shortcuts on Desktop and Start Menu
shortcut_table = [
    ("DesktopShortcut",                             # Shortcut
     "DesktopFolder",                               # Directory_
     "MGC VPN Connect",                          # Name
     "TARGETDIR",                                   # Component_
     "[TARGETDIR]mentor_vpn_connect.exe",           # Target
     None,                                          # Arguments
     None,                                          # Description
     None,                                          # Hotkey
     None,                                          # Icon
     None,                                          # IconIndex
     None,                                          # ShowCmd
     'TARGETDIR'                                    # WkDir
     ),
    ("StartMenuShortcut",                           # Shortcut
     "ProgramMenuFolder",                           # Directory_
     "MGC VPN Connect",                             # Name
     "TARGETDIR",                                   # Component_
     "[TARGETDIR]mentor_vpn_connect.exe",           # Target
     None,                                          # Arguments
     None,                                          # Description
     None,                                          # Hotkey
     None,                                          # Icon
     None,                                          # IconIndex
     None,                                          # ShowCmd
     'TARGETDIR'                                    # WkDir
     )
    ]
    
# Now create the table dictionary
msi_data = {"Shortcut": shortcut_table}

# Change some default MSI options and specify the use of the above defined tables
bdist_msi_options = {'data': msi_data}

cx_Freeze.setup(
    name = "MGC VPN Connect",
    options = {"build_exe": {"packages":["tkinter"], "include_files":["icon2.png", "credentials.txt"]}, "bdist_msi": bdist_msi_options},
    version = "1.0",
    description = "An app to connect to Mentor VPN automatically.",
    executables = executables
    )