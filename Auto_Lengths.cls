' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.
' It handles the calculation of lengths for linked elements and updates the text in the main element accordingly.
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
' Sets up the linked elements and calculates their lengths
Public Sub Initialize(ByVal NewElement As Element)
    ' Set the new element
    Set pNewElement = NewElement

    ' Get the linked elements using the Link module
    pLinkedElements = Link.GetLink(NewElement)

    ' Redimension the lengths array to match the linked elements array
    ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))

    ' Calculate the lengths of the linked elements
    CalculateLengths
End Sub

' Method to calculate lengths for all linked elements
' Uses the Length module to get the length of each element with specified rounding
Private Sub CalculateLengths()
    Dim i As Long

    ' Loop through each linked element and calculate its length
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        pLengths(i) = Length.GetLength(pLinkedElements(i), ARES_VAR.ARES_LENGTH_ROUND.Value)
    Next i
End Sub

' Method to update lengths in the new element based on linked elements
' Handles different cases based on the number of linked elements and non-zero lengths
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
    ' Show an error message if updating lengths fails
    ShowStatus GetTranslation("AutoLengthsUpdateError")
End Sub

' Method to show a form to select an element from the linked elements
' Displays a form with non-zero lengths for user selection
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

    ' Set the linked elements and master element in the form
    frm.SetLinkedElements pLinkedElements

    ' Get the master element by ID to avoid unintended initialization issues
    Set frm.SetMasterElement = ActiveModelReference.GetElementByID(pNewElement.Id)

    ' Show the form modelessly
    frm.Show vbModeless
End Sub

' Method to be called when an element is selected in the form
' Updates the text in the master element with the length of the selected element
Public Sub OnElementSelected(ByVal selectedElement As Element, ByVal MasterElement As Element)
    Dim Results() As String

    ' Get and set the text in the master element with the length of the selected element
    Results = StringsInEl.GetSetTextsInEl(MasterElement, CStr(Length.GetLength(selectedElement, ARES_VAR.ARES_LENGTH_ROUND.Value)), ARES_VAR.ARES_LENGTH_TRIGGER.Value)
End Sub

' Method to set the trigger list in the configuration
' Updates or initializes the trigger list in the configuration variables
Public Function SetTrigger(ByVal trigger As String) As Boolean
    SetTrigger = False

    ' If the trigger list is empty, set the new trigger
    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = "" Then
        SetTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, trigger)
    Else
        ' Otherwise, add the new trigger to the existing list
        SetTrigger = AddTrigger(trigger)
    End If
End Function

' Method to add a new trigger to the existing trigger list
' Appends the new trigger to the current list using a delimiter
Private Function AddTrigger(ByVal newTrigger As String) As Boolean
    Dim currentTriggers As String

    ' Get the current triggers
    currentTriggers = ARES_VAR.ARES_LENGTH_TRIGGER.Value

    ' Append the new trigger to the current list
    AddTrigger = Config.SetVar(ARES_VAR.ARES_LENGTH_TRIGGER.key, currentTriggers & ARES_VAR.ARES_VAR_DELIMITER & newTrigger)
End Function

' Method to reset the trigger list in the configuration
' Resets the trigger list to its default value
Public Function ResetTrigger() As Boolean
    ResetTrigger = False

    ' Reset the trigger list to its default value
    ARES_VAR.ResetMSVar ARES_VAR.ARES_LENGTH_TRIGGER

    ' Check if the trigger list was reset successfully
    If ARES_VAR.ARES_LENGTH_TRIGGER.Value = ARES_VAR.ARES_LENGTH_TRIGGER.Default Then
        ResetTrigger = True
    Else
        ResetTrigger = False
    End If
End Function
