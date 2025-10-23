# 🚀 Quick Install
```powershell
# One-line PowerShell installation
Invoke-WebRequest -Uri "https://github.com/Asketyll/ARES/releases/download/installer-v1.0.0/AresInstaller.exe" -OutFile "$env:TEMP\AresInstaller.exe"; Start-Process "$env:TEMP\AresInstaller.exe" -Verb RunAs
```

# 📥 Manual Downloads

- [AresInstaller.exe](https://github.com/Asketyll/ARES/releases/tag/installer-v1.0.0) - Complete installer
- [MVBA Source Code](https://github.com/Asketyll/ARES/releases/latest) - VBA source files

## 📁 Project Structure
```
ARES/
├── MVBA/                  # MicroStation VBA project
├── installer/             # Windows installer source
├── license-validator/     # License DLL source
└── tools/                 # PowerShell utilities
```

## ✨ Features

- **Auto Lengths**: Automatic length calculation for linked graphical elements
- **License Management**: AES-256 encrypted license validation
- **Multi-language**: French/English interface support
- **Configuration**: Centralized settings management
- **Error Handling**: Comprehensive logging and recovery

## 📋 System Requirements

- Windows 7/10/11
- .NET Framework 4.7.2+
- MicroStation Connect Edition or OpenCities Map PowerView
- Administrator privileges for installation

## 🔧 Installation

The installer automatically:

- Creates `C:\ARES\` directory structure
- Downloads latest components
- Registers COM components

## 📖 Usage

1. Run `AresInstaller.exe` as Administrator
2. Choose language (English/Français)
3. Click Install and wait for completion
4. Load `C:\ARES\ARES.mvba` in MicroStation

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

AGPL-3.0 - See [LICENSE](./LICENSE) file for details.