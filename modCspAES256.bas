' Module: modCspAES256
' Description: This module provides functions for AES256 encryption and decryption.
' Modified code of https://github.com/EszopiCoder/vba-crypto with MIT License
' License: This project is licensed under the AGPL-3.0.
'
'********************************************************************************
' MIT License
'
' Copyright (c) 2019 EszopiCoder
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files, to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
' Contact: pharm.coder@gmail.com
'
'********************************************************************************

Option Explicit

' CipherMode Constants
Private Const CBC = 1 ' Cipher Block Chaining (Default)
Private Const ECB = 2 ' Electronic Codebook
Private Const OFB = 3 ' Output Feedback
Private Const CFB = 4 ' Cipher Feedback
Private Const CTS = 5 ' Cipher Text Stealing

' PaddingMode Constants
Private Const None = 1
Private Const PKCS7 = 2 ' Default
Private Const Zeros = 3
Private Const ANSIX923 = 4
Private Const ISO10126 = 5

' Encoding Constants
Public Const Base64 = 0 ' Default
Public Const HexAES = 1

' Test function to encrypt and decrypt a string using AES
Private Sub TestEncryptAES()
    Dim StrEncrypted As String, StrDecrypted As String
    
    ' Encrypt string and hash key with SHA256 algorithm
    StrEncrypted = EncryptStringAES("MapPowerView.exe", SHA256(Base64Decode(mycrypto.AES_KEY)))
    Debug.Print "Encrypted string: " & StrEncrypted
    
    ' Decrypt string and hash key with SHA256 algorithm
    Debug.Print "IV: " & GetDecryptStringIV(StrEncrypted)
    StrDecrypted = DecryptStringAES(StrEncrypted, SHA256(Base64Decode(mycrypto.AES_KEY)))
    Debug.Print "Decrypted string: " & StrDecrypted
    
End Sub

' Function to generate SHA256 hash of a string
Public Function SHA256(StrText As String) As String
    On Error GoTo ErrorHandler

    Dim ObjUTF8 As Object, ObjSHA256 As Object
    Dim BytesText() As Byte, BytesHash() As Byte
    
    Set ObjUTF8 = CreateObject("System.Text.UTF8Encoding")
    Set ObjSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    BytesText = ObjUTF8.GetBytes_4(StrText)
    BytesHash = ObjSHA256.ComputeHash_2((BytesText))
    
    SHA256 = BytesToBase64(BytesHash)
    
    Set ObjUTF8 = Nothing
    Set ObjSHA256 = Nothing

    Exit Function

ErrorHandler:
    SHA256 = ""
End Function

' Function to get CSP (Cryptographic Service Provider) information
Private Function GetCSPInfo(ObjCSP As Object) As String
    On Error GoTo ErrorHandler
    
    Dim StrCipherMode As String, StrPaddingMode As String

    ' Display block size, key size, mode, and padding information
    Select Case ObjCSP.Mode
        Case CBC
            StrCipherMode = "Mode: Cipher Block Chaining (CBC)"
        Case ECB
            StrCipherMode = "Mode: Electronic Codebook (ECB)"
        Case OFB
            StrCipherMode = "Mode: Output Feedback (OFB)"
        Case CFB
            StrCipherMode = "Mode: Cipher Feedback (CFB)"
        Case CTS
            StrCipherMode = "Mode: Cipher Text Stealing (CTS)"
        Case Else
            StrCipherMode = "Mode: Undefined"
    End Select
    
    Select Case ObjCSP.Padding
        Case None
            StrPaddingMode = "Padding: None"
        Case PKCS7
            StrPaddingMode = "Padding: PKCS7"
        Case Zeros
            StrPaddingMode = "Padding: Zeros"
        Case ANSIX923
            StrPaddingMode = "Padding: ANSIX923"
        Case ISO10126
            StrPaddingMode = "Padding: ISO10126"
        Case Else
            StrPaddingMode = "Padding: Undefined"
    End Select

    GetCSPInfo = ObjCSP & vbNewLine & _
        "Block Size: " & ObjCSP.BlockSize & " bits" & vbNewLine & _
        "Key Size: " & ObjCSP.keySize & " bits" & vbNewLine & _
        StrCipherMode & vbNewLine & StrPaddingMode

    Set ObjCSP = Nothing

    Exit Function

ErrorHandler:
    GetCSPInfo = "Error: Unable to get CSP information"
End Function

' Function to encrypt a string using AES
Public Function EncryptStringAES(StrText As String, StrKey As String, _
    Optional Encoding As Integer = Base64) As Variant
    On Error GoTo ErrorHandler
    
    Dim ObjCSP As Object
    Dim ByteIV() As Byte
    Dim ByteText() As Byte
    Dim ByteEncrypted() As Byte
    Dim ByteEncryptedIV() As Byte
    Dim StrEncryptedIV As String
    Dim RandomIV As String
    Dim i As Long

    EncryptStringAES = Null
    
    Set ObjCSP = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' Check arguments
    If StrText = Null Or Len(StrText) <= 0 Then Err.Raise vbObjectError + 513, "strText", "Argument 'strText' cannot be null"
    If StrKey = Null Or Len(StrKey) <= 0 Then Err.Raise vbObjectError + 514, "strKey", "Argument 'strKey' cannot be null"
    
    RandomIV = GenerateRandomIV()
    ByteIV = Base64toBytes(RandomIV)
    
    ' Encryption Settings:
    ObjCSP.Padding = Zeros
    ObjCSP.Key = Base64toBytes(StrKey) ' Convert SHA256 hash to bytes
    ObjCSP.IV = ByteIV
    
    ' Convert from string to bytes (strText and strIV)
    ByteText = CreateObject("System.Text.UTF8Encoding").GetBytes_4(StrText)
    
    ' Encrypt byte data
    ByteEncrypted = ObjCSP.CreateEncryptor().TransformFinalBlock(ByteText, 0, UBound(ByteText) + 1)
    
    ' Concatenate byteEncrypted and byteIV
    ReDim ByteEncryptedIV(UBound(ByteIV) + UBound(ByteEncrypted) + 1)
    For i = 0 To UBound(ByteIV)
        ByteEncryptedIV(i) = ByteIV(i)
    Next i
    For i = 0 To UBound(ByteEncrypted)
        ByteEncryptedIV(UBound(ByteIV) + 1 + i) = ByteEncrypted(i)
    Next i
    
    ' Convert from bytes to encoded string
    Select Case Encoding
        Case Base64
            StrEncryptedIV = BytesToBase64(ByteEncryptedIV)
        Case HexAES
            StrEncryptedIV = BytesToHex(ByteEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "Invalid encoding type"
    End Select
    
    ' Return IV and encrypted string
    EncryptStringAES = StrEncryptedIV
    
    Set ObjCSP = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error: AES encryption failed" & vbNewLine & Err.Description
    EncryptStringAES = Null
End Function

' Function to decrypt a string using AES
Public Function DecryptStringAES(StrEncryptedIV As String, StrKey As String, _
    Optional Encoding As Integer = Base64) As Variant
    On Error GoTo ErrorHandler

    Dim ObjCSP As Object
    Dim ByteEncryptedIV() As Byte
    Dim ByteIV(0 To 15) As Byte
    Dim StrIV As String
    Dim ByteEncrypted() As Byte
    Dim ByteText() As Byte
    Dim StrText As String
    Dim i As Integer

    DecryptStringAES = Null

    On Error GoTo FunctionError
    
    Set ObjCSP = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' Convert from encoded string to bytes
    Select Case Encoding
        Case Base64
            ByteEncryptedIV = Base64toBytes(StrEncryptedIV)
        Case HexAES
            ByteEncryptedIV = HextoBytes(StrEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "ERROR: Invalid encoding type"
    End Select
    
    ' Check arguments (Part 1)
    If StrEncryptedIV = Null Or Len(StrEncryptedIV) <= 0 Then
        Err.Raise vbObjectError + 513, "strEncryptedIV", "Argument 'strEncryptedIV' cannot be null"
    End If
    If StrKey = Null Or Len(StrKey) <= 0 Then
        Err.Raise vbObjectError + 514, "strKey", "Argument 'strKey' cannot be null"
    End If
    
    ' Extract IV from strEncrypted
    For i = LBound(ByteIV) To UBound(ByteIV)
        ByteIV(i) = ByteEncryptedIV(i)
    Next i
    StrIV = CreateObject("System.Text.UTF8Encoding").GetString(ByteIV)
    
    ' Check arguments (Part 2)
    If StrIV = Null Or Len(StrIV) <= 0 Then Err.Raise vbObjectError + 515, "strIV", "Argument 'strIV' cannot be null"
    
    ' Extract encrypted text from strEncryptedIV
    ReDim ByteEncrypted(UBound(ByteEncryptedIV) - UBound(ByteIV) - 1)
    For i = LBound(ByteEncrypted) To UBound(ByteEncrypted)
        ByteEncrypted(i) = ByteEncryptedIV(UBound(ByteIV) + i + 1)
    Next i
    
    ' Decryption Settings:
    ObjCSP.Padding = Zeros
    ObjCSP.Key = Base64toBytes(StrKey) ' Convert SHA256 hash to bytes
    ObjCSP.IV = ByteIV
    
    ' Decrypt byte data
    ByteText = ObjCSP.CreateDecryptor().TransformFinalBlock(ByteEncrypted, 0, UBound(ByteEncrypted) + 1)
    
    ' Convert from bytes to string
    StrText = CreateObject("System.Text.UTF8Encoding").GetString(ByteText)
    ' Remove padding
    StrText = RemovePadding(StrText, ObjCSP.Padding)
    
    ' Return decrypted string
    DecryptStringAES = StrText
    
    Set ObjCSP = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error: AES decryption failed" & vbNewLine & Err.Description
    DecryptStringAES = Null
End Function

' Function to get the IV from an encrypted string
Private Function GetDecryptStringIV(StrEncryptedIV As String, _
    Optional Encoding As Integer = Base64) As String
    On Error GoTo ErrorHandler

    Dim ByteEncryptedIV() As Byte
    Dim ByteIV(0 To 15) As Byte
    Dim StrIV As String
    Dim i As Integer
    
    ' Convert encoded string to bytes
    Select Case Encoding
        Case Base64
            ByteEncryptedIV = Base64toBytes(StrEncryptedIV)
        Case HexAES
            ByteEncryptedIV = HextoBytes(StrEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "ERROR: Invalid encoding type"
    End Select

    ' Check arguments
    If StrEncryptedIV = Null Or Len(StrEncryptedIV) <= 0 Then
        Err.Raise vbObjectError + 513, "strEncryptedIV", "Argument 'strEncryptedIV' cannot be null"
    End If

    ' Extract IV from strEncrypted
    For i = LBound(ByteIV) To UBound(ByteIV)
        ByteIV(i) = ByteEncryptedIV(i)
    Next i

    ' Convert bytes to Base64 string
    StrIV = BytesToBase64(ByteIV)

    ' Return IV
    GetDecryptStringIV = StrIV

    Exit Function

ErrorHandler:
    MsgBox "Error: GetDecryptStringIV failed" & vbNewLine & Err.Description
    GetDecryptStringIV = ""
End Function

' Internal Base64 Conversion Functions
Private Function BytesToBase64(VarBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = VarBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function

Private Function Base64toBytes(VarStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = VarStr
         Base64toBytes = .nodeTypedValue
    End With
End Function

' Internal Hex Conversion Functions
Private Function BytesToHex(VarBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("hex")
        .DataType = "bin.hex"
        .nodeTypedValue = VarBytes
        BytesToHex = Replace(.Text, vbLf, "")
    End With
End Function

Private Function HextoBytes(VarStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("hex")
         .DataType = "bin.hex"
         .Text = VarStr
         HextoBytes = .nodeTypedValue
    End With
End Function

'********************************************************************************
' END MIT License
' the code in MIT License is modified
'********************************************************************************

' Function to remove padding from a string based on the specified padding mode
Private Function RemovePadding(StrText As String, PaddingMode As Integer) As String
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    i = Len(StrText)

    ' Remove padding based on the specified padding mode
    Select Case PaddingMode
        Case Zeros
            Do While i > 0 And Asc(Mid(StrText, i, 1)) = 0
                i = i - 1
            Loop
        Case PKCS7
            Dim paddingSize As Integer
            paddingSize = Asc(Mid(StrText, i, 1))
            If paddingSize > 0 And paddingSize <= 16 Then
                i = i - paddingSize
            End If
        Case ANSIX923
            ' ANSIX923 padding is similar to PKCS7 but the last byte of padding contains the padding length
            paddingSize = Asc(Mid(StrText, i, 1))
            If paddingSize > 0 And paddingSize <= 16 Then
                i = i - paddingSize
            End If
        Case ISO10126
            ' ISO10126 padding is similar to PKCS7 but the last byte of padding is random
            paddingSize = Asc(Mid(StrText, i, 1))
            If paddingSize > 0 And paddingSize <= 16 Then
                i = i - paddingSize
            End If
        Case None
            ' No padding to remove
    End Select

    RemovePadding = Left(StrText, i)

    Exit Function

ErrorHandler:
    RemovePadding = StrText
End Function

' Function to generate a random IV for AES encryption
Private Function GenerateRandomIV() As String
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim RandomIV(15) As Byte
    Dim RandomString As String

    Randomize
    For i = 0 To 15
        RandomIV(i) = CByte(Int(256 * Rnd()))
    Next i

    RandomString = BytesToBase64(RandomIV)
    GenerateRandomIV = RandomString

    Exit Function

ErrorHandler:
    GenerateRandomIV = ""
End Function

' Function to decode a Base64 string
Public Function Base64Decode(base64String As String) As String
    On Error GoTo ErrorHandler

    Dim Bytes() As Byte
    Dim DecodedString As String

    Bytes = Base64toBytes(base64String)
    DecodedString = CreateObject("System.Text.UTF8Encoding").GetString(Bytes)
    Base64Decode = DecodedString

    Exit Function

ErrorHandler:
    Base64Decode = ""
End Function
