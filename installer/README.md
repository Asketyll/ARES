# ARES Installer

Windows installer application for ARES MicroStation Add-in.

## Features

- Multi-language support (English/French)
- Administrator privilege enforcement
- Automatic download from GitHub releases
- COM component registration
- Progress tracking and logging

## Building
```bash
# Restore NuGet packages
nuget restore AresInstaller.sln

# Build with MSBuild
msbuild AresInstaller.sln /p:Configuration=Release
```

# Development

Built with:

- .NET Framework 4.7.2
- C# Windows Forms
- No external dependencies (except Newtonsoft.Json for license validator)