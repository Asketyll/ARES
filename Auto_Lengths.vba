' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.

' Dependencies: Config, Length, ARES_VAR, AutoLengths_GUI_SelectElements

Option Explicit

Private pNewElement As Element
Private pLinkedElements() As Element
Private pLengths() As Double
Private pRounding As Byte
Private pTriggersList As String

' Constants for default values
Private Const ARES_LENGTH_RND_DEFAULT As Byte = 1
Private Const ARES_LENGTH_TRIGGER_DEFAULT As String = "(Xx_m)"

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
            ShowElementSelectionForm
        End If
    End If
    Exit Sub

ErrorHandler:
    ShowStatus "An error occurred while updating lengths."
End Sub

' Show a form to select an element from the linked elements
Private Sub ShowElementSelectionForm()
    Dim frm As New AutoLengths_GUI_SelectElements
    Dim i As Long

    ' Add non-zero lengths to the form's list box
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        If pLengths(i) <> 0 Then
            frm.ListBox1.AddItem CStr(pLengths(i))
            frm.ListBox1.List(frm.ListBox1.ListCount - 1, 1) = i
        End If
    Next i

    ' Pass the reference to the current instance of AutoLengths
    Set frm.AutoLengthsInstance = Me
    
    frm.SetLinkedElements pLinkedElements
    'pNewElement -> NewElement -> AfterChange is unintialized by MS at the end of ElementChangeHandler ClassModule
    'even if you maintain the instance with elementChangeHandler As ElementChangeHandler
    Set frm.SetMasterElement = ActiveModelReference.GetElementByID(pNewElement.ID)
    'to be reviewed later
    frm.Show vbModeless
End Sub

' Method to be called when an element is selected in the form
Public Sub OnElementSelected(ByVal selectedElement As Element, ByVal MasterElement As Element)
    Dim Results() As String
    
    Results = StringsInEl.GetSetTextsInEl(MasterElement, CStr(Length.GetLength(selectedElement, pRounding)), pTriggersList)
End Sub

' Get the rounding setting from the configuration
Private Function GetRoundingSetting() As Byte
    If Config.GetVar(ARES_VAR.ARES_LENGTH_ROUND) = "" Then
        If Not Config.SetVar(ARES_VAR.ARES_LENGTH_ROUND, ARES_LENGTH_RND_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.ARES_LENGTH_ROUND & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.ARES_LENGTH_ROUND & " défini à " & ARES_LENGTH_RND_DEFAULT & " par défaut"
        End If
    End If
    GetRoundingSetting = CByte(Config.GetVar(ARES_VAR.ARES_LENGTH_ROUND))
End Function

' Get the trigger list from the configuration
Private Function GetTriggers() As String
    If Config.GetVar(ARES_VAR.ARES_LENGTH_TRIGGER) = "" Then
        If Not SetTriggers(ARES_LENGTH_TRIGGER_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.ARES_LENGTH_TRIGGER & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.ARES_LENGTH_TRIGGER & " défini à " & ARES_LENGTH_TRIGGER_DEFAULT & " par défaut"
        End If
    End If
    GetTriggers = Config.GetVar(ARES_VAR.ARES_LENGTH_TRIGGER)
End Function

' Set the trigger list in the configuration
Private Function SetTriggers(ByVal trigger As String) As Boolean
    If Config.GetVar(ARES_VAR.ARES_LENGTH_TRIGGER) = "" Then
        SetTriggers = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER, trigger)
    Else
        SetTriggers = AddTrigger(trigger)
    End If
End Function

' Add a new trigger to the existing trigger list
Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    Dim currentTriggers As String
    currentTriggers = Config.GetVar(ARES_VAR.ARES_LENGTH_TRIGGER)
    AddTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER, currentTriggers & ARES_VAR_DELIMITER & newTrigger)
End Function
