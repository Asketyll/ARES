# 🚀 Quick Install
powershell
# One-line PowerShell installation
Invoke-WebRequest -Uri "https://github.com/Asketyll/ARES/raw/main/dist/latest/AresInstaller.exe" -OutFile "$env:TEMP\AresInstaller.exe"; Start-Process "$env:TEMP\AresInstaller.exe" -Verb RunAs

#📥 Manual Downloads
AresInstaller.exe - Complete installer
MVBA Source Code - VBA source files

#📁 Project Structure
ARES/
├── MVBA/                  # MicroStation VBA project
├── installer/             # Windows installer source
├── license-validator/     # License DLL source (coming soon)
├── tools/                 # PowerShell utilities (coming soon)
└── dist/                  # Compiled binaries

#✨ Features
- Auto Lengths: Automatic length calculation for linked graphical elements
- License Management: AES-256 encrypted license validation
- Multi-language: French/English interface support
- Configuration: Centralized settings management
- Error Handling: Comprehensive logging and recovery

#📋 System Requirements
- Windows 7/10/11
- .NET Framework 4.7.2+
- MicroStation Connect Edition or OpenCities Map PowerView
- Administrator privileges for installation

#🔧 Installation
The installer automatically:

Creates C:\ARES\ directory structure
Downloads latest components
Registers COM components
Installs license generation tools
Configures MicroStation integration

#📖 Usage
Run AresInstaller.exe as Administrator
Choose language (English/Français)
Click Install and wait for completion
Load C:\ARES\ARES.mvba in MicroStation

#🤝 Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

#📄 License
AGPL-3.0 - See LICENSE file for details.