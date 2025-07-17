' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, Length, ARES_VAR, AutoLengths_GUI_SelectElements, LangManager

Option Explicit

' Variable to store the new element
Private pNewElement As Element

' Array to store the linked elements
Private pLinkedElements() As Element

' Array to store the lengths of the linked elements
Private pLengths() As Double

' Method to initialize the class with the new element
Public Sub Initialize(ByVal NewElement As Element)
    On Error GoTo ErrorHandler
    Static ElType(5) As Long
    ElType(0) = 12 'Complexe Line
    ElType(1) = 4 'Brised Line
    ElType(2) = 16 'Arc
    ElType(3) = 3 'Line
    ElType(4) = 6 'Shape
    ElType(5) = 14 'Complexe Shape
    
    Set pNewElement = NewElement
    pLinkedElements = Link.GetLink(NewElement, , ElType)
    ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))
    CalculateLengths

    Exit Sub

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsInitError") & Err.Description
End Sub

' Method to calculate lengths for all linked elements
Private Sub CalculateLengths()
    On Error GoTo ErrorHandler

    Dim i As Long
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        pLengths(i) = Length.GetLength(pLinkedElements(i), ARES_VAR.ARES_LENGTH_ROUND.Value)
    Next i

    Exit Sub

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsCalculationError") & Err.Description
End Sub

' Method to update lengths in the new element based on linked elements
Public Sub UpdateLengths()
    On Error GoTo ErrorHandler

    HandleElementSelection

    Exit Sub

ErrorHandler:
    ElementInProcesse.Remove pNewElement
    ShowStatus GetTranslation("AutoLengthsUpdateError") & Err.Description
End Sub

' Method to handle the selection logic for elements
Private Sub HandleElementSelection()
    If HasSingleLinkedElement() Then
        HandleSingleElement
    Else
        HandleMultipleElements
    End If
End Sub

' Method to check if there is only one linked element
Private Function HasSingleLinkedElement() As Boolean
    HasSingleLinkedElement = (UBound(pLinkedElements) - LBound(pLinkedElements) = 0)
End Function

' Method to handle the case of a single linked element
Private Sub HandleSingleElement()
    Dim Results() As String
    Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(UBound(pLinkedElements))), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
    ElementInProcesse.Remove pNewElement
End Sub

' Method to handle the case of multiple linked elements
Private Sub HandleMultipleElements()
    If CountNonZeroLengths() = 1 Then
        UpdateSingleNonZeroLength
    Else
        ShowElementSelectionForm
    End If
End Sub

' Method to count the number of non-zero lengths
Private Function CountNonZeroLengths() As Long
    Dim count As Long, j As Long
    count = 0
    For j = LBound(pLengths) To UBound(pLengths)
        If pLengths(j) <> 0 Then
            count = count + 1
        End If
    Next j
    CountNonZeroLengths = count
End Function

' Method to update the length for a single non-zero length
Private Sub UpdateSingleNonZeroLength()
    Dim Results() As String, j As Long
    For j = LBound(pLengths) To UBound(pLengths)
        If pLengths(j) <> 0 Then
            Results = StringsInEl.GetSetTextsInEl(pNewElement, CStr(pLengths(j)), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
            ElementInProcesse.Remove pNewElement
            Exit For
        End If
    Next j
End Sub

' Method to show a form to select an element from the linked elements
Private Sub ShowElementSelectionForm()
    On Error GoTo ErrorHandler

    Dim frm As New AutoLengths_GUI_SelectElements
    Dim i As Long

    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        If pLengths(i) <> 0 Then
            frm.ListBox1.AddItem CStr(pLengths(i))
            frm.ListBox1.List(frm.ListBox1.ListCount - 1, 1) = i
        End If
    Next i

    Set frm.AutoLengthsInstance = Me
    frm.SetLinkedElements pLinkedElements
    Set frm.SetMasterElement = ActiveModelReference.GetElementByID(pNewElement.id)
    frm.Show vbModeless

    Exit Sub

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsShowFormError") & Err.Description
End Sub

' Method to be called when an element is selected in the form
Public Sub OnElementSelected(ByVal selectedElement As Element, ByVal MasterElement As Element)
    On Error GoTo ErrorHandler

    Dim Results() As String
    Results = StringsInEl.GetSetTextsInEl(MasterElement, CStr(Length.GetLength(selectedElement, ARES_VAR.ARES_LENGTH_ROUND.Value)), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
    
    ElementInProcesse.Remove MasterElement
    Exit Sub

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsSelectionError") & Err.Description
End Sub

' Method to set the trigger list in the configuration
Public Function SetTrigger(ByVal trigger As String) As Boolean
    On Error GoTo ErrorHandler

    SetTrigger = False

    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = "" Then
        SetTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, trigger)
    Else
        SetTrigger = AddTrigger(trigger)
    End If

    Exit Function

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsSetTriggerError") & Err.Description
    SetTrigger = False
End Function

' Method to add a new trigger to the existing trigger list
Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    On Error GoTo ErrorHandler

    Dim currentTriggers As String
    currentTriggers = ARES_VAR.ARES_LENGTH_TRIGGER.Value
    AddTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, currentTriggers & ARES_VAR.ARES_VAR_DELIMITER & newTrigger)

    Exit Function

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsAddTriggerError") & Err.Description
    AddTrigger = False
End Function

' Method to reset the trigger list in the configuration
Public Function ResetTrigger() As Boolean
    On Error GoTo ErrorHandler

    ResetTrigger = False

    ARES_VAR.ResetMSVar ARES_VAR.ARES_LENGTH_TRIGGER

    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = ARES_VAR.ARES_LENGTH_TRIGGER.Default Then
        ResetTrigger = True
    End If

    Exit Function

ErrorHandler:
    ShowStatus GetTranslation("AutoLengthsResetTriggerError") & Err.Description
    ResetTrigger = False
End Function
