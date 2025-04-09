' Module: Auto_Lengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.

' Dependencies: Config, Length, ARES_VAR

' Constants
Private Const ARES_LENGTH_RND_DEFAULT As Byte = 2
Public Const ARES_LENGTH_TRIGGER_ID As String = "Xx_"
Private Const ARES_LENGTH_TRIGGER_DEFAULT As String = "(" & ARES_LENGTH_TRIGGER_ID & "m)"

Public Sub MAJLengths(ByVal NewElement As Element)
    On Error GoTo ErrorHandler

    ' Declare variables
    Dim linkedElements() As Element
    Dim lengths() As Double
    Dim triggersList() As String
    Dim rounding As Byte
    Dim i As Long

    ' Retrieve linked elements
    linkedElements = Link.GetLink(NewElement)
    ReDim lengths(UBound(linkedElements))

    ' Set default rounding if not already set
    rounding = GetRoundingSetting()

    ' Retrieve the list of triggers
    triggersList = Split(GetTriggers(), ARES_LENGTH_TRIGGER_DELIMITER)

    ' Calculate lengths for each linked element
    For i = LBound(linkedElements) To UBound(linkedElements)
        lengths(i) = Length.GetLength(linkedElements(i), rounding)
    Next i

    Exit Sub

ErrorHandler:
    ' Handle errors
    ShowStatus "An error occurred while updating lengths."
End Sub

Private Function GetRoundingSetting() As Byte
    ' Retrieve the rounding setting from configuration
    Dim rounding As Byte

    If Config.GetVar(ARES_VAR.LENGTH_ROUND) = "" Then
        If Not Config.SetVar(ARES_VAR.LENGTH_ROUND, ARES_LENGTH_RND_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.LENGTH_ROUND & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.LENGTH_ROUND & " défini à " & ARES_LENGTH_RND_DEFAULT & " par défaut"
        End If
    End If

    rounding = CByte(Config.GetVar(ARES_VAR.LENGTH_ROUND))
    GetRoundingSetting = rounding
End Function

Private Function GetTriggers() As String
    ' Retrieve the list of triggers from configuration
    If Config.GetVar(ARES_VAR.LENGTH_TRIGGER) = "" Then
        If Not SetTriggers(ARES_LENGTH_TRIGGER_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.LENGTH_TRIGGER & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.LENGTH_TRIGGER & " défini à " & ARES_LENGTH_TRIGGER_DEFAULT & " par défaut"
        End If
    End If

    GetTriggers = Config.GetVar(ARES_VAR.LENGTH_TRIGGER)
End Function

Private Function SetTriggers(ByVal trigger As String) As Boolean
    ' Set the triggers in configuration
    If Config.GetVar(ARES_VAR.LENGTH_TRIGGER) = "" Then
        SetTriggers = Config.SetVar(ARES_VAR.LENGTH_TRIGGER, trigger)
    Else
        SetTriggers = AddTrigger(trigger)
    End If
End Function

Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    ' Add a new trigger to the existing list of triggers
    Dim currentTriggers As String
    currentTriggers = Config.GetVar(ARES_VAR.LENGTH_TRIGGER)
    AddTrigger = Config.SetVar(ARES_VAR.LENGTH_TRIGGER, currentTriggers & ARES_LENGTH_TRIGGER_DELIMITER & newTrigger)
End Function
