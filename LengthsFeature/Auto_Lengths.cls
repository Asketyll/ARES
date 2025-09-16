' Class Module: AutoLengths
' Description: This module provides functions to add length with rounding to a text if they are graphically linked and the trigger is present in the text.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, Length, ARESConfigClass, ARESConstants, AutoLengths_GUI_SelectElements, LangManager, ElementInProcesseClass, ElementChangeHandler, link, ErrorHandlerClass
Option Explicit

' Class to handle selection
Implements ILocateCommandEvents

' Variable to store the new element
Private pNewElement As element
' Array to store the linked elements
Private pLinkedElements() As element
' Array to store the lengths of the linked elements
Private pLengths() As Double

' Method to initialize the class with the new element
Public Sub Initialize(ByVal NewElement As element)
    On Error GoTo ErrorHandler

    Set pNewElement = NewElement
    pLinkedElements = GetLinkedElements()

    If Not UBound(pLinkedElements) = -1 Then
        ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))
        CalculateLengths
    Else
        ShowStatus "No linked elements found for element ID: " & DLongToString(pNewElement.ID)
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths.Initialize"
    ShowStatus GetTranslation("AutoLengthsInitError") & Err.Description
End Sub

' Method to get linked elements based on specific types
Private Function GetLinkedElements() As element()
    Dim ElType(5) As Long
    ElType(0) = 12 ' Complex Line
    ElType(1) = 4  ' Line String
    ElType(2) = 16 ' Arc
    ElType(3) = 3  ' Line
    ElType(4) = 6  ' Shape
    ElType(5) = 14 ' Complex Shape

    ' Use the GetLink method from the Link class with the specified element types
    GetLinkedElements = Link.GetLink(pNewElement, False, ElType)
End Function

' Method to calculate lengths for all linked elements
Private Sub CalculateLengths()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim linkedElement As element

    ' Check if the array pLinkedElements is initialized and not empty
    'If Not IsArrayInitialized(pLinkedElements) Then Exit Sub

    ReDim pLengths(LBound(pLinkedElements) To UBound(pLinkedElements))

    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        Set linkedElement = pLinkedElements(i)
        If Not linkedElement Is Nothing Then
            pLengths(i) = Length.GetLength(linkedElement, ARESConfig.ARES_LENGTH_ROUND.Value)
        Else
            pLengths(i) = 0
            ShowStatus "Linked element at index " & i & " is Nothing"
        End If
    Next i

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths.CalculateLengths"
    ShowStatus GetTranslation("AutoLengthsCalculationError") & Err.Description
End Sub

' Method to update lengths in the new element based on linked elements
Public Sub UpdateLengths()
    On Error GoTo ErrorHandler
    
    HandleElementSelection
    Exit Sub
    
ErrorHandler:
    ElementInProcesse.Remove pNewElement
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths.UpdateLengths"
    ShowStatus GetTranslation("AutoLengthsUpdateError") & Err.Description
End Sub

' Method to handle the selection logic for elements
Private Sub HandleElementSelection()
    If Not UBound(pLinkedElements) = -1 Then
        If HasSingleLinkedElement() Then
            HandleSingleElement
        Else
            HandleMultipleElements
        End If
    End If
End Sub

' Method to check if there is only one linked element
Private Function HasSingleLinkedElement() As Boolean
    HasSingleLinkedElement = (UBound(pLinkedElements) - LBound(pLinkedElements) = 0)
End Function

' Method to handle the case of a single linked element
Private Sub HandleSingleElement()
    Dim Results() As String
    Dim lengthValue As String
    Dim Color As Long
    
    ' Get the length of the first (and only) linked element
    lengthValue = CStr(pLengths(UBound(pLinkedElements)))
    
    If ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value Then
        Color = pLinkedElements(UBound(pLinkedElements)).Color
        Results = StringsInEl.GetSetTextsInEl(pNewElement, lengthValue, ARESConfig.ARES_LENGTH_TRIGGER.Value, Color)
    Else
        ' Get and update the texts in the element with the calculated length
        Results = StringsInEl.GetSetTextsInEl(pNewElement, lengthValue, ARESConfig.ARES_LENGTH_TRIGGER.Value)
    End If
    
    ' If the operation was successful, remove the element from the list of elements being processed
    If Not IsEmpty(Results) Then
        ElementInProcesse.Remove pNewElement
    Else
        ShowStatus "Failed to update texts in element with ID: " & DLongToString(pNewElement.ID)
    End If
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
    Dim Count As Long, j As Long
    Count = 0
    For j = LBound(pLengths) To UBound(pLengths)
        If pLengths(j) <> 0 Then
            Count = Count + 1
        End If
    Next j
    CountNonZeroLengths = Count
End Function

' Method to update the length for a single non-zero length
Private Sub UpdateSingleNonZeroLength()
    Dim Results() As String, j As Long
    Dim lengthValue As String
    Dim Success As Boolean

    Success = False

    For j = LBound(pLengths) To UBound(pLengths)
        If pLengths(j) <> 0 Then
            lengthValue = CStr(pLengths(j))
            ' Get and update the texts in the element with the calculated length
            Results = StringsInEl.GetSetTextsInEl(pNewElement, lengthValue, ARESConfig.ARES_LENGTH_TRIGGER.Value)
            If Not IsEmpty(Results) Then
                ElementInProcesse.Remove pNewElement
                Success = True
            Else
                ShowStatus "Failed to update texts in element with ID: " & DLongToString(pNewElement.ID)
            End If
            Exit For
        End If
    Next j

    If Not Success Then
        ErrorHandler.HandleError "No non-zero length found to update in element with ID: " & DLongToString(pNewElement.ID), 0, "AutoLengths.UpdateSingleNonZeroLength", "AutoLengths"
    End If
End Sub

' Method to show a form to select an element from the linked elements
Private Sub ShowElementSelectionForm()
    On Error GoTo ErrorHandler

    Dim frm As New AutoLengths_GUI_SelectElements
    Dim i As Long
    Dim lengthValue As String
    Dim lengthCount As Long

    lengthCount = 0

    ' Add only elements with non-zero lengths to the list
    For i = LBound(pLinkedElements) To UBound(pLinkedElements)
        If pLengths(i) <> 0 Then
            lengthValue = CStr(pLengths(i))
            frm.ListBox1.AddItem lengthValue
            frm.ListBox1.List(frm.ListBox1.ListCount - 1, 1) = i
            lengthCount = lengthCount + 1
        End If
    Next i

    ' Check if valid elements were added to the list
    If lengthCount = 0 Then
        ShowStatus GetTranslation("AutoLengthsNoValidElements")
        Exit Sub
    End If

    ' Configure and show the form
    Set frm.AutoLengthsInstance = Me
    frm.SetLinkedElements pLinkedElements
    Set frm.SetMasterElement = ActiveModelReference.GetElementById(pNewElement.ID)
    frm.Show vbModeless

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths.ShowElementSelectionForm"
    ShowStatus GetTranslation("AutoLengthsShowFormError") & Err.Description
End Sub

' Method to be called when an element is selected in the form
Public Sub OnElementSelected(ByVal selectedElement As element, ByVal MasterElement As element)
    On Error GoTo ErrorHandler

    Dim Results() As String
    Dim lengthValue As String
    
    ' Get the length of the selected element with the specified precision
    lengthValue = CStr(Length.GetLength(selectedElement, ARESConfig.ARES_LENGTH_ROUND.Value))
    
    If ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value Then
        ' Get and update the texts in the master element with the calculated length
        Results = StringsInEl.GetSetTextsInEl(MasterElement, lengthValue, ARESConfig.ARES_LENGTH_TRIGGER.Value, selectedElement.Color)
    Else
        Results = StringsInEl.GetSetTextsInEl(MasterElement, lengthValue, ARESConfig.ARES_LENGTH_TRIGGER.Value)
    End If
    
    ' If the operation was successful, remove the elements from the list of elements being processed
    If Not IsEmpty(Results) Then
        ElementInProcesse.Remove MasterElement
        Dim linkedElement As Variant
        For Each linkedElement In pLinkedElements
            ElementInProcesse.Remove linkedElement
        Next linkedElement
    Else
        ShowStatus "Failed to update texts in master element with ID: " & DLongToString(MasterElement.ID)
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths.OnElementSelected"
    ShowStatus GetTranslation("AutoLengthsSelectionError") & Err.Description
End Sub

Private Sub ILocateCommandEvents_Accept(ByVal element As element, Point As Point3d, ByVal View As View)
    ' Call the function with the selected object
End Sub

Private Sub ILocateCommandEvents_Cleanup()
    ' Cleanup after the command
End Sub

Private Sub ILocateCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    ' Handle dynamics
End Sub

Private Sub ILocateCommandEvents_LocateFailed()
    ' Handle selection failure
    ShowStatus GetTranslation("AutoLengthsNoValidElement")
End Sub

Private Sub ILocateCommandEvents_LocateFilter(ByVal element As element, Point As Point3d, Accepted As Boolean)
    ' Accept only specified elements
    Accepted = False
    
    Select Case True
        Case element.IsComplexStringElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsLineElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsArcElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsShapeElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsComplexShapeElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsTextElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsTextNodeElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
        Case element.IsCellElement And element.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID
            Accepted = True
    End Select
    If Accepted Then
        If ElementInProcesse.Add(element) Then
            ChangeHandler.HandleElementModified element
        End If
    End If
End Sub

Private Sub ILocateCommandEvents_LocateReset()
    ' Reset the command
End Sub

Private Sub ILocateCommandEvents_Start()
    Dim lc As LocateCriteria
    ' Configure selection criteria
    Set lc = CommandState.CreateLocateCriteria(False)
    CommandState.SetLocateCriteria lc
    ShowCommand GetTranslation("AutoLengthsSelectanelementC")
    ShowPrompt GetTranslation("AutoLengthsSelectanelementP")
End Sub