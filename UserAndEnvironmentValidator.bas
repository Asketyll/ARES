' Module: UserAndEnvironmentValidator
' Description: This module provides functions to verify if user(s) and environment are allowed.

' Dependencies: modCspAES256

Option Explicit

' API function declarations
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare PtrSafe Function GetModuleBaseNameA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

' Main function to check authorization
Public Function IsAuthorized() As Boolean
    On Error GoTo ErrorHandler
    IsAuthorized = False

    Dim userList As Variant
    userList = GetUserList()

    If IsNull(userList) Then Exit Function

    Dim user As Variant
    For Each user In userList
        If IsMicroStationRunning And IsValidUser(CStr(user)) Then
            IsAuthorized = True
            Exit Function
        End If
    Next user

    Exit Function

ErrorHandler:
    IsAuthorized = False
End Function

' Check if MicroStation is running
Private Function IsMicroStationRunning() As Boolean
    On Error GoTo ErrorHandler
    IsMicroStationRunning = False

    Dim processName As String
    Dim expectedVersion As String
    processName = DecryptStringAES(mycrypto.ENCRYPTED_PROCESS_NAME, SHA256(Base64Decode(mycrypto.AES_KEY)))
    expectedVersion = DecryptStringAES(mycrypto.ENCRYPTED_EXPECTED_VERSION, SHA256(Base64Decode(mycrypto.AES_KEY)))

    If IsNull(processName) Or IsNull(expectedVersion) Then Exit Function
    
    If Application.Name = Left(processName, Len(processName) - 4) Then
        If Application.Version = expectedVersion Then
            If Application.Visible And Application.IsInitialized And Application.IsRegistered And Application.IsSerialized Then
                If IsValidProcessID(Application.ProcessID) Then
                    IsMicroStationRunning = True
                End If
            End If
        End If
    End If

    Exit Function

ErrorHandler:
    IsMicroStationRunning = False
End Function

' Check if the user is valid
Private Function IsValidUser(ID As String) As Boolean
    On Error GoTo ErrorHandler
    IsValidUser = (ID = Application.UserName)
    Exit Function

ErrorHandler:
    IsValidUser = False
End Function

' Check if the ProcessID is valid
Private Function IsValidProcessID(ID As Long) As Boolean
    On Error GoTo ErrorHandler
    IsValidProcessID = False

    Dim hProcess As Long
    Dim hModule As Long
    Dim moduleName As String
    Dim bufferSize As Long
    Dim bytesCopied As Long

    hProcess = OpenProcess(&H400 Or &H10, False, ID)

    If hProcess <> 0 Then
        If EnumProcessModules(hProcess, hModule, Len(hModule), bytesCopied) Then
            bufferSize = 1024
            moduleName = Space(bufferSize)
            If GetModuleBaseNameA(hProcess, hModule, moduleName, bufferSize) <> 0 Then
                moduleName = Trim(moduleName)
                If Left(moduleName, Len(moduleName) - 1) = DecryptStringAES(ENCRYPTED_PROCESS_NAME, SHA256(Base64Decode(mycrypto.AES_KEY))) Then
                    IsValidProcessID = True
                    Exit Function
                End If
            End If
        End If
    End If

    Exit Function

ErrorHandler:
    IsValidProcessID = False
End Function

' Function to get the list of users
Private Function GetUserList() As Variant
    Dim decryptedList As String
    Dim userArray() As String

    On Error GoTo ErrorHandler

    decryptedList = DecryptStringAES(mycrypto.ENCRYPTED_USER_LIST, SHA256(Base64Decode(mycrypto.AES_KEY)))

    If IsNull(decryptedList) Then Exit Function

    userArray = Split(decryptedList, "|")
    GetUserList = userArray
    Exit Function

ErrorHandler:
    GetUserList = Null
End Function
