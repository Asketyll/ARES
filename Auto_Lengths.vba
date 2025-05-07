' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.

' Dependencies: Config, Length, ARES_VAR

Option Explicit

' Private variables for class properties
Private pNewElement As Element
Private pLinkedElements() As Element
Private pLengths() As Double
Private pRounding As Byte
Private pTriggersList As String
Private WithEvents frm As SelectElements

' Constants for default values
Private Const ARES_LENGTH_RND_DEFAULT As Byte = 2
Private Const ARES_LENGTH_TRIGGER_DEFAULT As String = "(" & ARES_LENGTH_TRIGGER_ID & "m)"

' Initialize the class with the new element
Public Sub Initialize(ByVal NewElement As Element)
    Set pNewElement = NewElement
    pRounding = GetRoundingSetting()
    pTriggersList = GetTriggers()
    pLinkedElements = Link.GetLink(NewElement)
    ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))
    CalculateLengths
End Sub

' Calculate lengths for all linked elements
Private Sub CalculateLengths()
    Dim i As Long
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        pLengths(i) = Length.GetLength(pLinkedElements(i), pRounding)
    Next i
End Sub

' Update lengths in the new element based on linked elements
Public Sub UpdateLengths()
    On Error GoTo ErrorHandler

    Dim Results() As String
    Dim count As Long
    Dim j As Long
    Dim selectedElement As Element
    
    ' If there's only one linked element, update the length directly
    If UBound(pLinkedElements) - LBound(pLinkedElements) = 0 Then
        Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(UBound(pLinkedElements))), pTriggersList)
    Else
        count = 0
        ' Count non-zero lengths
        For j = LBound(pLengths) To UBound(pLengths)
            If pLengths(j) <> 0 Then
                count = count + 1
            End If
        Next j

        ' If there's only one non-zero length, update it directly
        If count = 1 Then
            For j = LBound(pLengths) To UBound(pLengths)
                If pLengths(j) <> 0 Then
                    Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(j)), pTriggersList)
                    Exit For
                End If
            Next j
        Else
            ' If there are multiple non-zero lengths, show a selection form
            Set selectedElement = ShowElementSelectionForm
            If Not selectedElement Is Nothing Then
                Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(Length.GetLength(selectedElement, pRounding)), pTriggersList)
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    ShowStatus "An error occurred while updating lengths."
End Sub

' Show a form to select an element from the linked elements
Private Function ShowElementSelectionForm() As Element
    Dim frm As New SelectElements
    Dim i As Long

    ' Add non-zero lengths to the form's list box
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        If pLengths(i) <> 0 Then
            frm.ListBox1.AddItem CStr(pLengths(i))
            frm.ListBox1.List(frm.ListBox1.ListCount - 1, 1) = i
        End If
    Next i

    frm.SetLinkedElements pLinkedElements
    frm.Show vbModeless

    ' Wait for the ElementSelected event
    Do While frm.Visible
        DoEvents
    Loop

    ' Return the selected element
    Set ShowElementSelectionForm = frm.selectedElement

    Unload frm
    Set frm = Nothing
End Function

' Event handler for when an element is selected in the form
Private Sub frm_ElementSelected(ByVal selectedElement As Element)
    Set ShowElementSelectionForm = selectedElement
    frm.Hide
End Sub

' Get the rounding setting from the configuration
Private Function GetRoundingSetting() As Byte
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

' Get the trigger list from the configuration
Private Function GetTriggers() As String
    If Config.GetVar(ARES_VAR.LENGTH_TRIGGER) = "" Then
        If Not SetTriggers(ARES_LENGTH_TRIGGER_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.LENGTH_TRIGGER & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.LENGTH_TRIGGER & " défini à " & ARES_LENGTH_TRIGGER_DEFAULT & " par défaut"
        End If
    End If
    GetTriggers = Config.GetVar(ARES_VAR.LENGTH_TRIGGER)
End Function

' Set the trigger list in the configuration
Private Function SetTriggers(ByVal trigger As String) As Boolean
    If Config.GetVar(ARES_VAR.LENGTH_TRIGGER) = "" Then
        SetTriggers = Config.SetVar(ARES_VAR.LENGTH_TRIGGER, trigger)
    Else
        SetTriggers = AddTrigger(trigger)
    End If
End Function

' Add a new trigger to the existing trigger list
Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    Dim currentTriggers As String
    currentTriggers = Config.GetVar(ARES_VAR.LENGTH_TRIGGER)
    AddTrigger = Config.SetVar(ARES_VAR.LENGTH_TRIGGER, currentTriggers & ARES_LENGTH_TRIGGER_DELIMITER & newTrigger)
End Function
