# CLAUDE.md - AI Assistant Guide for ARES

**Last Updated:** 2025-12-18
**Project:** ARES - MicroStation Add-in for Automated Length Calculation
**License:** AGPL-3.0

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Architecture & Structure](#architecture--structure)
3. [Development Environment](#development-environment)
4. [Code Conventions & Standards](#code-conventions--standards)
5. [Key Components & Patterns](#key-components--patterns)
6. [Development Workflow](#development-workflow)
7. [Testing & Quality Assurance](#testing--quality-assurance)
8. [Security Considerations](#security-considerations)
9. [Deployment](#deployment)
10. [AI Assistant Guidelines](#ai-assistant-guidelines)

---

## Project Overview

### What is ARES?

ARES is a professional MicroStation add-in that provides automated length calculation for linked graphical elements. It's built as a multi-layered solution combining:

- **VBA Application** (8,782 lines): Core functionality in MicroStation Visual Basic for Applications
- **C# License Validator**: COM DLL for RSA/AES-based license validation
- **C# Installer**: Windows Forms application for automated deployment
- **PowerShell Tools**: License generation and management utilities

### Key Features

- Automatic length calculation for linked graphical elements
- Enterprise license management with AES-256 encryption
- Multi-language support (English/French)
- Centralized configuration management
- Comprehensive error handling and logging
- Network license deployment

### Target Environment

- **OS:** Windows 7/10/11
- **.NET Framework:** 4.7.2+
- **MicroStation:** Connect Edition, OpenCities Map PowerView
- **MVBA:** Version 7.1
- **PowerShell:** 5.1+ (for license tools)

---

## Architecture & Structure

### Directory Layout

```
/home/user/ARES/
├── MVBA/                      # MicroStation VBA Application (8,782 lines)
│   ├── Command/               # User command interface
│   ├── Components/            # Functional components (lengths, links, graphics, cells)
│   ├── Configuration/         # Config management, language support, file dialogs
│   ├── Core/                  # Boot loader, constants, error handling
│   ├── EventHandlers/         # DGN open/close, element changes, idle events
│   ├── LengthsFeature/        # Auto-lengths GUI and logic
│   ├── Security/              # License validation, encryption, UUID generation
│   └── Tests/                 # Unit testing framework
│
├── installer/                 # Windows Installer (C# WinForms)
│   ├── AresInstaller/         # Main installer project
│   │   ├── Forms/             # Language selection, product selection, main form
│   │   └── Resources/         # Icons, translations
│   └── AresInstaller.sln      # Visual Studio solution
│
├── license-validator/         # COM DLL for License Validation (C#)
│   ├── AresLicenseValidator/  # Main validator project
│   │   ├── Services/          # Core validation logic
│   │   ├── Models/            # License data structures
│   │   └── Interfaces/        # COM interface contracts
│   └── AresLicenseValidator.sln  # Visual Studio solution
│
└── tools/                     # PowerShell Utilities
    └── Generate-ARESLicense.ps1  # License generation tool
```

### VBA Project Structure (28 files)

| Directory | Files | Purpose |
|-----------|-------|---------|
| **Command/** | 1 module | User command interface layer |
| **Components/** | 7 modules | Length calculation, link management, graphics interaction, cell handling, custom properties |
| **Configuration/** | 4 modules, 2 classes | Settings management, MicroStation variables, language manager, file dialogs |
| **Core/** | 2 modules, 2 classes | Constants, boot loader, error handler, element processing |
| **EventHandlers/** | 3 classes | DGN file events, element change tracking, idle event processing |
| **LengthsFeature/** | 1 class, 2 forms | Auto-lengths GUI and business logic |
| **Security/** | 4 modules | License manager, user validation, UUID generation, AES-256 encryption |
| **Tests/** | 1 module | Unit testing framework |

---

## Development Environment

### Required Tools

**For VBA Development:**
- MicroStation Connect Edition (or OpenCities Map PowerView)
- MVBA 7.1 IDE (built into MicroStation)
- Text editor (for VBA source files)

**For C# Development:**
- Visual Studio 2017+ (2019 or 2022 recommended)
- .NET Framework 4.7.2 SDK
- NuGet package manager

**For PowerShell Development:**
- PowerShell 5.1+
- PowerShell ISE or VS Code with PowerShell extension
- Administrator privileges (for license deployment)

### Dependencies

**NuGet Packages (Installer):**
```xml
<package id="Costura.Fody" version="6.0.0" />
<package id="Fody" version="6.9.3" />
<package id="Newtonsoft.Json" version="13.0.4" />
```

**NuGet Packages (License Validator):**
```xml
<package id="Newtonsoft.Json" version="13.0.3" />
```

### Build System

**VBA:**
- Manual compilation in MicroStation VBA IDE
- Output: `ARES.mvba` file

**C# Projects:**
- MSBuild via Visual Studio solutions
- Post-build COM registration (license-validator):
  ```
  regasm.exe "$(TargetPath)" /tlb /codebase
  ```
- Fody IL weaving for assembly embedding (installer)

---

## Code Conventions & Standards

### VBA Conventions

**File Headers:**
Every VBA file must start with a standard header:
```vb
' Module: ModuleName
' Description: Brief description of purpose
' License: This project is licensed under the AGPL-3.0.
' Dependencies: List of dependencies (or "None")
Option Explicit
```

**Naming Conventions:**
- **Constants:** `UPPERCASE_WITH_UNDERSCORES` with `ARES_` prefix
  - Examples: `ARES_DEFAULT_GRAPHIC_GROUP_ID`, `ARES_VAR_DELIMITER`
- **Global Objects:** `PascalCase`
  - Examples: `ChangeHandler`, `ErrorHandler`, `ElementInProcesse`, `ARESConfig`
- **Private Members:** `mPascalCase` (m prefix)
  - Examples: `mLogFilePath`, `mbLicenseChecked`
- **Functions/Subs:** `PascalCase`
  - Examples: `ValidateLicenseOnLoad()`, `InitializeErrorHandler()`
- **Parameters:** `PascalCase`
  - Examples: `ByVal Description As String`, `ByVal Number As Long`
- **Local Variables:** `PascalCase` (can use type prefixes like `str`, `b`, `l`, `i`)
  - Examples: `strErrorMsg`, `mbLicenseValid`, `FileNum`

**Error Handling Pattern:**
```vb
Public Function SomeFunction() As Boolean
    On Error GoTo ErrorHandler

    ' Function logic here
    SomeFunction = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "ModuleName.SomeFunction", "ERROR"
    SomeFunction = False
End Function
```

**Option Explicit:**
- ALWAYS use `Option Explicit` at the top of every module
- All variables MUST be declared before use

**Comments:**
```vb
' === SECTION HEADERS ===
' Use triple equals for major section headers

' Regular comments for inline documentation
' Explain WHY, not WHAT (code should be self-documenting)
```

### C# Conventions

**Naming:**
- **Classes:** `PascalCase` (e.g., `LicenseValidator`, `AresInstaller`)
- **Interfaces:** `IPascalCase` (e.g., `IAresLicenseValidator`)
- **Methods:** `PascalCase` (e.g., `ValidateLicense()`)
- **Private Fields:** `_camelCase` or `camelCase`
- **Constants:** `UPPERCASE_WITH_UNDERSCORES` or `PascalCase` for const

**COM Interop:**
```csharp
[ComVisible(true)]
[Guid("unique-guid-here")]
[ClassInterface(ClassInterfaceType.None)]
public class AresLicenseValidator : IAresLicenseValidator
{
    // Implementation
}
```

**Error Handling:**
```csharp
try
{
    // Operation
}
catch (Exception ex)
{
    // Log and handle
    LastError = ex.Message;
    return false;
}
```

### PowerShell Conventions

**Function Names:** `Verb-Noun` format (e.g., `Generate-ARESLicense`)

**Parameters:**
```powershell
param(
    [Parameter(Mandatory=$false)]
    [string]$Company,

    [Parameter(Mandatory=$false)]
    [string[]]$AuthorizedUsers
)
```

**Comments:**
```powershell
# Single-line comments for inline documentation

<#
.SYNOPSIS
    Brief description
.DESCRIPTION
    Detailed description
.PARAMETER ParameterName
    Parameter description
#>
```

---

## Key Components & Patterns

### 1. Bootstrap Pattern (BootLoader.bas)

**Purpose:** Initialize the VBA project when loaded into MicroStation

**Entry Point:** `OnProjectLoad()` - Auto-called when MVBA project loads

**Initialization Sequence:**
1. Initialize global error handler (`ErrorHandler`)
2. Validate license (critical - exit if fails)
3. Initialize DGN event handlers (`moOpenClose`)
4. Initialize idle event handler

**Global Objects Created:**
- `ErrorHandler` - Centralized error logging
- `ChangeHandler` - Element change tracking
- `ElementInProcesse` - Track elements being processed
- `ARESConfig` - Configuration management

### 2. Error Handling Pattern

**Centralized Handler:** `ErrorHandlerClass.cls`

**Features:**
- Log file rotation (max 1 MB)
- Structured error messages with timestamp
- Debug mode popup display
- Fail-safe logging (errors in error handler don't crash)

**Usage Pattern:**
```vb
ErrorHandler.HandleError "Description", ErrNumber, "ModuleName.FunctionName", "ERROR"
```

**Log Format:**
```
[2025-12-18 14:30:22] ERROR in ModuleName.FunctionName (123): Error description here
```

### 3. Configuration Management

**ARES_MS_VAR_Class.cls:** Manages MicroStation configuration variables

**Pattern:**
- Get/Set wrapper around MicroStation's ActiveWorkspace.ConfigurationVariables
- Automatic persistence (write to MS config on set)
- Type-safe accessors
- Default value handling with `ARES_NAVD` constant ("Not a Variable Defined")

**Usage:**
```vb
' Store configuration
ARES_MS_VAR.SetVariable "ARES_SettingName", "value"

' Retrieve configuration
Dim value As String
value = ARES_MS_VAR.GetVariable("ARES_SettingName")
If value = ARES_NAVD Then
    ' Variable not defined, use default
End If
```

### 4. License Validation Architecture

**Three-Layer Approach:**

**Layer 1: VBA Interface** (`LicenseManager.bas`)
```vb
Public Function ValidateLicense() As Boolean
    ' Creates COM object and validates
    Set oValidator = CreateObject("ARES.LicenseValidator")
    ValidateLicense = oValidator.ValidateLicense()
End Function
```

**Layer 2: COM Wrapper** (`AresLicenseValidator.cs`)
```csharp
[ComVisible(true)]
public class AresLicenseValidator : IAresLicenseValidator
{
    public bool ValidateLicense()
    {
        return _validator.ValidateLicense();
    }
}
```

**Layer 3: Core Logic** (`Services/LicenseValidator.cs`)
- RSA-2048 signature verification
- AES-256 decryption
- Environment hash validation
- User authorization checking
- Network path license file search

**License File Search Order:**
1. Network drives: Z: → Y: → X: → ... → K:
2. UNC paths: `\\server\shared\ARES_Licenses\`
3. Local fallback: `C:\ARES\`

### 5. Event Handler Pattern

**DGNOpenClose.cls:** Handles file open/close events
```vb
Private WithEvents m_WorkspaceEvents As Workspace

Private Sub m_WorkspaceEvents_DgnFileOpened(...)
    ' Initialize on file open
End Sub

Private Sub m_WorkspaceEvents_DgnFileClosing(...)
    ' Cleanup on file close
End Sub
```

**ElementChangeHandler.cls:** Tracks element modifications
```vb
Private WithEvents m_ElementAddEvent As ElementAddEvent
Private WithEvents m_ElementModifyEvent As ElementModifyEvent
```

### 6. Language Manager Pattern

**LangManager.bas:** Multi-language support

**Features:**
- Language detection from MicroStation settings
- Fallback to English if translation missing
- Dictionary-based translations

**Usage:**
```vb
Dim msg As String
msg = GetTranslation("ErrorKeyName")
MsgBox msg, vbExclamation
```

### 7. Constants Organization

**ARESConstants.bas:** Central constant definitions

**Categories:**
- System constants (graphic groups, element types)
- String delimiters
- Configuration constants
- Error values
- Custom property names

**Pattern:**
```vb
' Always prefix with ARES_ to avoid conflicts
Public Const ARES_CONSTANT_NAME As DataType = Value

' Include usage comment
' Used in ModuleName for specific purpose
```

---

## Development Workflow

### Git Workflow

**Branch Strategy:**
- **Main Branch:** Production-ready code
- **Development Branches:** Feature work in `claude/` prefixed branches
- **PR Pattern:** Merge via pull requests from development branches

**Commit Message Format:**
```
type: Brief description

Examples:
feat: Add automatic persistence to MicroStation configuration
fix: Use sanitized path for all extraction operations
refactor: Simplify path validation by inlining security checks
docs: Improve Markdown formatting across all README files
```

**Commit Types:**
- `feat:` - New features
- `fix:` - Bug fixes
- `refactor:` - Code restructuring without behavior change
- `docs:` - Documentation updates
- `test:` - Test additions/modifications
- `chore:` - Build process, dependency updates

### Development Process

**For VBA Changes:**
1. Edit source files in `MVBA/` directory structure
2. Test in MicroStation VBA IDE
3. Export changes back to source files
4. Commit with descriptive message
5. Create pull request

**For C# Changes:**
1. Open solution in Visual Studio
2. Make changes and build
3. Test COM registration (for license-validator)
4. Commit changes
5. Create pull request

**For Installer Changes:**
1. Update installer code
2. Test installation process
3. Verify COM registration
4. Update version numbers if needed
5. Commit and create pull request

### Building the Project

**MVBA:**
```
1. Open MicroStation VBA IDE (Alt+F11)
2. File → Import all .bas, .cls, .frm files from MVBA/
3. Maintain directory structure as modules/classes
4. Build → Compile
5. File → Save As → ARES.mvba
```

**Installer:**
```bash
cd installer
nuget restore AresInstaller.sln
msbuild AresInstaller.sln /p:Configuration=Release
```

**License Validator:**
```bash
cd license-validator
nuget restore AresLicenseValidator.sln
msbuild AresLicenseValidator.sln /p:Configuration=Release
regasm bin/Release/AresLicenseValidator.dll /tlb /codebase
```

---

## Testing & Quality Assurance

### VBA Unit Testing

**Framework:** `UnitTesting.bas` module

**Pattern:**
```vb
Public Sub RunAllTests()
    ' Initialize test counter
    TestsRun = 0
    TestsPassed = 0

    ' Run individual test functions
    Test_SomeFeature
    Test_AnotherFeature

    ' Report results
    ReportTestResults
End Sub

Private Sub Test_SomeFeature()
    On Error GoTo TestFailed

    ' Arrange
    Dim expected As String
    expected = "expected value"

    ' Act
    Dim actual As String
    actual = SomeFunctionToTest()

    ' Assert
    If actual = expected Then
        RecordTestPass "Test_SomeFeature"
    Else
        RecordTestFail "Test_SomeFeature", "Expected: " & expected & ", Got: " & actual
    End If
    Exit Sub

TestFailed:
    RecordTestFail "Test_SomeFeature", Err.Description
End Sub
```

**Configuration Variable:**
```vb
' Enable test mode
ARES_MS_VAR.SetVariable "ARES_UnitTesting", "True"
```

### Manual Testing Checklist

**Before Committing:**
- [ ] Code compiles without errors
- [ ] No warnings in VBA IDE
- [ ] Error handlers in place for all functions
- [ ] Constants properly prefixed with `ARES_`
- [ ] Option Explicit present in all modules
- [ ] File headers complete with dependencies
- [ ] Comments explain WHY, not WHAT

**Before Creating PR:**
- [ ] All unit tests pass
- [ ] Manual testing in MicroStation completed
- [ ] License validation works (if security changes)
- [ ] Configuration persistence works (if config changes)
- [ ] Multi-language support intact (if UI changes)
- [ ] Error logging functional
- [ ] No regression in existing features

### Integration Testing

**License Validation:**
```vb
' Test valid license
mbValid = LicenseManager.ValidateLicense()
Debug.Print "Valid License: " & mbValid

' Test invalid license (rename license file temporarily)
mbValid = LicenseManager.ValidateLicense()
Debug.Print "Invalid License: " & mbValid

' Check error message
Debug.Print "Last Error: " & LicenseManager.LastError
```

**Configuration Persistence:**
```vb
' Write value
ARES_MS_VAR.SetVariable "TestVar", "TestValue"

' Close and reopen MicroStation

' Read value
Dim value As String
value = ARES_MS_VAR.GetVariable("TestVar")
Debug.Print "Persisted Value: " & value
```

---

## Security Considerations

### License Security Architecture

**Encryption Layers:**
1. **RSA-2048 Signature:** Prevents license tampering
2. **AES-256 Encryption:** Protects sensitive license data
3. **Environment Hash:** Binds license to specific hardware/domain
4. **User Authorization:** Restricts access to specific users

**Critical Security Files:**
- `modCspAES256.bas` - AES encryption implementation
- `Services/LicenseValidator.cs` - RSA signature verification
- `Generate-ARESLicense.ps1` - License generation with private key

### Security Best Practices

**For VBA Development:**
- NEVER hardcode sensitive data (keys, passwords)
- Use license validator for authorization checks
- Log security events via ErrorHandler
- Validate all user inputs
- Use `On Error GoTo` to prevent data leaks via error messages

**For C# Development:**
- Keep RSA private key secure (never commit to git)
- Public key is embedded in DLL code (safe to include)
- Use strong-name signing for COM DLL
- Validate all file paths before operations
- Sanitize user inputs

**For PowerShell Tools:**
- Require administrator privileges for license generation
- Protect private key files (filesystem ACLs)
- Use secure network paths for license distribution
- Validate all user inputs
- Log all license generation operations

### .gitignore Security

**Currently Excluded:**
```
.claude/              # AI assistant config
docs/                 # Local documentation
CLAUDE.md             # (NOTE: This should be REMOVED from .gitignore!)
*.snk                 # Strong-name keys
*.log                 # Log files
*.txt                 # Temporary files
```

**Important:** Private keys should NEVER be committed to repository.

---

## Deployment

### Installation Targets

**Default Installation Path:** `C:\ARES\`

**Directory Structure Created:**
```
C:\ARES\
├── ARES.mvba                    # Main VBA application
├── AresLicenseValidator.dll     # COM DLL
├── config/                      # Configuration files
└── logs/                        # Application logs
```

**Registry Entries:**
```
HKEY_CLASSES_ROOT\ARES.LicenseValidator
HKEY_CLASSES_ROOT\Interface\{GUID}
HKEY_CLASSES_ROOT\CLSID\{GUID}
```

### Network License Deployment

**License File Location Options:**
1. **Network Drives:** `Z:\ARES_Licenses\ares_license.json`
2. **UNC Paths:** `\\server\shared\ARES_Licenses\ares_license.json`
3. **Local Fallback:** `C:\ARES\ares_license.json`

**Search Order:** Z → Y → X → W → V → U → T → S → R → Q → P → O → N → M → L → K → UNC → Local

**License File Format:**
```json
{
  "company": "Company Name",
  "domain": "DOMAIN",
  "installed_by": "DOMAIN\\username",
  "installation_date": "2025-12-18 14:30:22",
  "license_key": "UUID-FORMAT-KEY",
  "environment_hash": "HASH",
  "authorized_users": ["DOMAIN\\user1", "DOMAIN\\user2"],
  "max_users": 5,
  "signature": "BASE64_RSA_SIGNATURE",
  "version": "1.0"
}
```

### Release Process

**Version Numbering:** Semantic Versioning (MAJOR.MINOR.PATCH)

**Release Checklist:**
1. Update version numbers in AssemblyInfo.cs files
2. Build all projects in Release configuration
3. Test installer end-to-end
4. Test license validation with real license
5. Test in MicroStation environment
6. Create GitHub release tag
7. Upload installer executable
8. Update README with new version info

**GitHub Release Assets:**
- `AresInstaller.exe` - Complete installer
- `ARES.mvba` - VBA source (if separate release)
- Release notes (changes, fixes, new features)

---

## AI Assistant Guidelines

### When Working with This Codebase

**DO:**
- ✅ Always read existing files before modifying
- ✅ Follow established naming conventions strictly
- ✅ Add proper error handling to all functions
- ✅ Update file headers with correct dependencies
- ✅ Use `Option Explicit` in all VBA modules
- ✅ Prefix all constants with `ARES_`
- ✅ Test license validation after security changes
- ✅ Maintain multi-language support
- ✅ Log errors via ErrorHandler.HandleError
- ✅ Comment WHY, not WHAT
- ✅ Use existing patterns (singleton, factory, event-driven)
- ✅ Keep backward compatibility unless explicitly breaking change
- ✅ Update relevant README files when adding features

**DON'T:**
- ❌ Create new files without reading project structure first
- ❌ Modify encryption/security code without understanding impact
- ❌ Hardcode sensitive data (keys, paths, credentials)
- ❌ Remove error handlers from existing code
- ❌ Change constant values without understanding usage
- ❌ Break COM interface contracts (breaking change)
- ❌ Modify license validation logic without security review
- ❌ Add dependencies without updating documentation
- ❌ Commit private keys or license files
- ❌ Skip testing after changes

### Understanding the Project

**Key Questions to Ask:**
1. What component am I modifying? (VBA/Installer/Validator/Tools)
2. Does this change affect license validation?
3. Does this change affect MicroStation integration?
4. Are there dependencies on other modules?
5. Do I need to update configuration management?
6. Does this affect multi-language support?
7. Is error handling properly implemented?
8. Have I tested in the actual environment?

### Common Tasks

**Adding a New VBA Module:**
1. Create file in appropriate subdirectory
2. Add standard file header with dependencies
3. Use `Option Explicit`
4. Follow naming conventions
5. Add error handling
6. Update BootLoader if initialization needed
7. Add unit tests in UnitTesting.bas
8. Update MVBA/README.md if structural change

**Modifying Configuration:**
1. Check if variable exists in ARES_MS_VAR_Class
2. Use GetVariable/SetVariable pattern
3. Handle ARES_NAVD (not defined) case
4. Test persistence across MicroStation sessions
5. Update documentation

**Adding a Translation:**
1. Locate LangManager.bas
2. Add key-value pairs for both English and French
3. Use GetTranslation("KeyName") in code
4. Test language switching
5. Update relevant UI elements

**Modifying License Validation:**
⚠️ **HIGH RISK - Requires Security Review**
1. Understand current validation flow
2. Don't break COM interface
3. Test with valid and invalid licenses
4. Test network license paths
5. Verify RSA signature validation still works
6. Test in offline scenario
7. Get security review before merging

### Debugging Tips

**VBA Debugging:**
```vb
' Use Debug.Print for output
Debug.Print "Variable value: " & variableName

' Use Immediate Window (Ctrl+G) for runtime evaluation
?variableName
?ErrorHandler.GetLogFilePath()

' Set breakpoints in VBA IDE
' Step through with F8
' Watch window for variable inspection
```

**C# Debugging:**
- Use Visual Studio debugger with breakpoints
- For COM debugging: Attach to MicroStation process
- Check Windows Event Viewer for COM errors
- Use Fusion Log Viewer for assembly binding issues

**License Issues:**
```vb
' Check last error
Debug.Print LicenseManager.LastError

' Verify file exists
Debug.Print Dir("Z:\ARES_Licenses\ares_license.json")

' Test COM object creation
Set obj = CreateObject("ARES.LicenseValidator")
Debug.Print (obj Is Nothing) ' Should be False
```

### File Organization Reference

**VBA File Types:**
- `.bas` - Standard modules (functions, subs, global vars)
- `.cls` - Class modules (objects with methods/properties)
- `.frm` - Forms (GUI elements with code-behind)

**VBA Module Organization:**
```
Public Constants    ' Top of file
Public Variables    ' After constants
Private Variables   ' After public vars
Public Functions    ' Main functions first
Private Functions   ' Helper functions last
```

### Performance Considerations

**VBA Performance:**
- Minimize COM calls (batch operations)
- Avoid excessive MicroStation API calls in loops
- Use early binding when possible
- Cache frequently accessed configuration values
- Disable screen updates during batch operations

**License Validation:**
- License check on bootload (cached)
- Not re-checked on every operation
- Network path search is sequential (optimize order)

---

## Quick Reference

### Important File Locations

| File | Location | Purpose |
|------|----------|---------|
| Boot Loader | `/MVBA/Core/BootLoader.bas` | Project initialization |
| Constants | `/MVBA/Core/ARESConstants.bas` | All ARES constants |
| Error Handler | `/MVBA/Core/ErrorHandlerClass.cls` | Centralized error logging |
| License Manager | `/MVBA/Security/LicenseManager.bas` | VBA license interface |
| Config Manager | `/MVBA/Configuration/ARES_MS_VAR_Class.cls` | MicroStation variables |
| Language Manager | `/MVBA/Configuration/LangManager.bas` | Translation management |
| License Validator | `/license-validator/AresLicenseValidator/Services/LicenseValidator.cs` | Core validation logic |
| Installer Main | `/installer/AresInstaller/Program.cs` | Installer entry point |
| License Tool | `/tools/Generate-ARESLicense.ps1` | License generation |

### Key Constants

```vb
ARES_DEFAULT_GRAPHIC_GROUP_ID  = 0
ARES_MSDETYPE_ERROR            = 44
ARES_VAR_DELIMITER             = "|"
ARES_NAVD                      = "NaVD"  ' Not a Variable Defined
ARES_RND_ERROR_VALUE           = 255
ARES_CELL_INDEX_ERROR_VALUE    = -1
ARES_NAME_LIBRARY_TYPE         = "ARES"
ARES_NAME_ITEM_TYPE            = "ARESAutoLengthObject"
```

### Global Objects

```vb
ChangeHandler       ' ElementChangeHandler instance
ErrorHandler        ' ErrorHandlerClass instance
ElementInProcesse   ' ElementInProcesseClass instance
ARESConfig          ' ARESConfigClass instance
```

### Common Patterns

**Error Handling:**
```vb
On Error GoTo ErrorHandler
' code
Exit Sub/Function
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "Module.Function", "ERROR"
```

**Configuration:**
```vb
ARES_MS_VAR.SetVariable "VarName", "Value"
Dim val As String
val = ARES_MS_VAR.GetVariable("VarName")
If val = ARES_NAVD Then val = "DefaultValue"
```

**Translation:**
```vb
Dim msg As String
msg = GetTranslation("KeyName")
```

**License Check:**
```vb
If Not LicenseManager.ValidateLicense() Then
    MsgBox "License validation failed: " & LicenseManager.LastError
    Exit Sub
End If
```

---

## Additional Resources

### Documentation Files

- `/README.md` - Project overview and quick install
- `/MVBA/README.md` - VBA development guide
- `/installer/README.md` - Installer build instructions
- `/license-validator/README.md` - License validator guide
- `/tools/README.md` - PowerShell tools comprehensive guide

### External References

- MicroStation VBA API Documentation
- .NET Framework 4.7.2 Documentation
- Newtonsoft.Json Documentation
- PowerShell 5.1 Documentation

---

## Changelog

### 2025-12-18
- Initial creation of CLAUDE.md
- Comprehensive documentation of codebase structure
- Development workflows documented
- Code conventions standardized
- Security considerations outlined
- AI assistant guidelines established

---

**For Questions or Updates:**
This file should be updated whenever significant architectural changes occur, new patterns are introduced, or development workflows change. Keep it current to maximize AI assistant effectiveness.
