' Module: Crypto
' Description: This module provides function to verify if user(s) and environment are allowed.
' It includes functions to encrypt and decrypt using the Triple DES algorithm.
' It uses a predefined initialization vector (IV) and Triple DES key for encryption and decryption processes.
' The initial code structure for Triple DES algorithm was inspired by the example provided at: https://gist.github.com/motoraku/97ad730891e59159d86c

Option Explicit

' API function declarations
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare PtrSafe Function GetModuleBaseNameA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

' Key example
'Public Const INITIALIZATION_VECTOR = "12345678"  ' Always 8 characters
'Public Const TRIPLE_DES_KEY = "abcdefghijklmnop" ' Always 16 characters

' Encrypted constants example
'Private Const ENCRYPTED_PROCESS_NAME As String = "replace these with the actual encrypted values"
'Private Const ENCRYPTED_EXPECTED_VERSION As String = "replace these with the actual encrypted values"
'Private Const ENCRYPTED_USER_LIST As String = "replace these with the actual encrypted values"

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
    processName = DecryptStringTripleDES(MyCrypto.ENCRYPTED_PROCESS_NAME)
    expectedVersion = DecryptStringTripleDES(MyCrypto.ENCRYPTED_EXPECTED_VERSION)

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
                If Left(moduleName, Len(moduleName) - 1) = DecryptStringTripleDES(ENCRYPTED_PROCESS_NAME) Then
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
    Dim encryptedList As String
    Dim decryptedList As String
    Dim userArray() As String

    On Error GoTo ErrorHandler

    encryptedList = MyCrypto.ENCRYPTED_USER_LIST
    decryptedList = DecryptStringTripleDES(encryptedList)

    If IsNull(decryptedList) Then Exit Function

    userArray = Split(decryptedList, "|")
    GetUserList = userArray
    Exit Function

ErrorHandler:
    GetUserList = Null
End Function

' Function to encrypt strings using Triple DES
Private Function EncryptStringTripleDES(plain_string As String) As Variant
    Dim encryption_object As Object
    Dim plain_byte_data() As Byte
    Dim encrypted_byte_data() As Byte
    Dim encrypted_base64_string As String

    EncryptStringTripleDES = Null

    On Error GoTo ErrorHandler

    plain_byte_data = CreateObject("System.Text.UTF8Encoding").GetBytes_4(plain_string)

    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.Key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(INITIALIZATION_VECTOR)
    encrypted_byte_data = _
            encryption_object.CreateEncryptor().TransformFinalBlock(plain_byte_data, 0, UBound(plain_byte_data) + 1)

    encrypted_base64_string = BytesToBase64(encrypted_byte_data)

    EncryptStringTripleDES = encrypted_base64_string

    Exit Function

ErrorHandler:
    EncryptStringTripleDES = Null
End Function

' Function to decrypt strings using Triple DES
Private Function DecryptStringTripleDES(encrypted_string As String) As Variant
    Dim encryption_object As Object
    Dim encrypted_byte_data() As Byte
    Dim plain_byte_data() As Byte
    Dim plain_string As String

    DecryptStringTripleDES = Null

    On Error GoTo ErrorHandler

    encrypted_byte_data = Base64toBytes(encrypted_string)

    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.Key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(MyCrypto.TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(MyCrypto.INITIALIZATION_VECTOR)
    plain_byte_data = encryption_object.CreateDecryptor().TransformFinalBlock(encrypted_byte_data, 0, UBound(encrypted_byte_data) + 1)

    plain_string = CreateObject("System.Text.UTF8Encoding").GetString(plain_byte_data)

    DecryptStringTripleDES = plain_string

    Exit Function

ErrorHandler:
    DecryptStringTripleDES = Null
End Function

' Function to convert bytes to base64 string
Private Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function

' Function to convert base64 string to bytes
Private Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").CreateElement("b64")
         .DataType = "bin.base64"
         .Text = varStr
         Base64toBytes = .nodeTypedValue
    End With
End Function
Sub EncryptConstants()
    Dim processName As String
    Dim expectedVersion As String
    Dim userList As String

    processName = "exemple1"
    expectedVersion = "exemple1.0"
    userList = "user1|user8"

    Debug.Print "Encrypted Process Name: " & EncryptStringTripleDES(processName)
    Debug.Print "Encrypted Expected Version: " & EncryptStringTripleDES(expectedVersion)
    Debug.Print "Encrypted User List: " & EncryptStringTripleDES(userList)
End Sub
