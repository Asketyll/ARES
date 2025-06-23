' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.

' Dependencies: Config, Length, ARES_VAR, AutoLengths_GUI_SelectElements, LangManager

Option Explicit

Private pNewElement As Element
Private pLinkedElements() As Element
Private pLengths() As Double

' Initialize the class with the new element
Public Sub Initialize(ByVal NewElement As Element)
    Set pNewElement = NewElement
    pLinkedElements = Link.GetLink(NewElement)
    ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))
    CalculateLengths
End Sub

' Calculate lengths for all linked elements
Private Sub CalculateLengths()
    Dim i As Long
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        pLengths(i) = Length.GetLength(pLinkedElements(i), ARES_VAR.ARES_LENGTH_ROUND.Value)
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
        Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(UBound(pLinkedElements))), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
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
                    Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(j)), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
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
    ShowStatus GetTranslation("AutoLengthsUpdateError")
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
    Set frm.SetMasterElement = ActiveModelReference.GetElementByID(pNewElement.Id)
    'to be reviewed later
    frm.Show vbModeless
End Sub

' Method to be called when an element is selected in the form
Public Sub OnElementSelected(ByVal selectedElement As Element, ByVal MasterElement As Element)
    Dim Results() As String
    
    Results = StringsInEl.GetSetTextsInEl(MasterElement, CStr(Length.GetLength(selectedElement, ARES_VAR.ARES_LENGTH_ROUND.Value)), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
End Sub

' Set the trigger list in the configuration
Public Function SetTrigger(ByVal trigger As String) As Boolean
    SetTrigger = False
    
    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = "" Then
        SetTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, trigger)
    Else
        SetTrigger = AddTrigger(trigger)
    End If
End Function

' Add a new trigger to the existing trigger list
Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    Dim currentTriggers As String
    currentTriggers = ARES_VAR.ARES_LENGTH_TRIGGER.Value
    AddTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, currentTriggers & ARES_VAR.ARES_VAR_DELIMITER & newTrigger)
End Function

' Reset the trigger list in the configuration
Public Function ResetTrigger() As Boolean
    ResetTrigger = False
    
    ARES_VAR.ResetMSVar ARES_VAR.ARES_LENGTH_TRIGGER
    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = ARES_VAR.ARES_LENGTH_TRIGGER.Default Then
        ResetTrigger = True
    Else
        ResetTrigger = False
    End If
End Function
