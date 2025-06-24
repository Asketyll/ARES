' Module: UUID
' Description: This module provides functions to generate a Version 1 UUID.

' Function to generate a Version 1 UUID
Public Function GenerateV1() As String
    ' Declare variables
    Dim macAddress As String
    Dim timePart As String
    Dim uuid As String

    ' Get the MAC address and time-based part
    macAddress = GetMacAddress()
    timePart = GetTimePart()

    ' Combine MAC address and time-based part to form a UUID
    uuid = FormatUUID(macAddress, timePart)

    ' Return the generated UUID
    GenerateV1 = uuid
End Function

' Function to get the MAC address using WMI
Private Function GetMacAddress() As String
    ' Declare variables
    Dim objWMIService As Object
    Dim colNetworkAdapters As Object
    Dim objNetworkAdapter As Object
    Dim macAddress As String

    ' Initialize the WMI service
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    ' Query for network adapters
    Set colNetworkAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

    ' Loop through the network adapters to find the MAC address
    For Each objNetworkAdapter In colNetworkAdapters
        If Not IsNull(objNetworkAdapter.macAddress) Then
            macAddress = objNetworkAdapter.macAddress
            Exit For
        End If
    Next objNetworkAdapter

    ' Return the MAC address without colons
    GetMacAddress = Replace(macAddress, ":", "")
End Function

' Function to generate a time-based part for the UUID
Private Function GetTimePart() As String
    ' Declare variables
    Dim currentTime As Double
    Dim secondsToday As Long
    Dim timeHex As String

    ' Get the current date and time
    currentTime = Now

    ' Calculate the number of seconds since midnight
    secondsToday = CLng((currentTime - Int(currentTime)) * 24 * 60 * 60)

    ' Convert the seconds to a hexadecimal string
    timeHex = Hex(secondsToday)

    ' Ensure the hex string is 8 characters long
    GetTimePart = Right("00000000" & timeHex, 8)
End Function

' Function to format the UUID string
Private Function FormatUUID(macAddress As String, timePart As String) As String
    ' Declare variables
    Dim uuid As String
    Dim randomPart As String
    Dim i As Integer

    ' Generate a random part for the UUID
    Randomize
    randomPart = ""
    For i = 1 To 12
        randomPart = randomPart & Hex(Int(Rnd * 16))
    Next i

    ' Format the UUID as per Version 1 standard
    uuid = Mid(macAddress, 1, 8) & "-" & _
           Mid(macAddress, 9, 4) & "-" & _
           "1" & Mid(timePart, 2, 3) & "-" & _
           Mid(randomPart, 1, 4) & "-" & _
           Mid(randomPart, 5, 12)

    ' Set the version (4 bits) and variant (2-3 bits) for UUID Version 1
    Mid(uuid, 15, 1) = "1" ' Set version to 1
    Mid(uuid, 20, 1) = "8" ' Set variant to RFC 4122 (binary: 10)

    ' Return the formatted UUID
    FormatUUID = uuid
End Function
