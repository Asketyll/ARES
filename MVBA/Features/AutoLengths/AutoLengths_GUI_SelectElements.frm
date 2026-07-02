VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoLengths_GUI_SelectElements 
   Caption         =   "Select :"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2535
   OleObjectBlob   =   "AutoLengths_GUI_SelectElements.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoLengths_GUI_SelectElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: AutoLengths_GUI_SelectElements
' Description: This UserForm is used for selecting an element in a list and returning it.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: MSGraphicalInteraction, AutoLengths, LangManager, ErrorHandlerClass, FormUXHelper
Option Explicit

' Array to store the linked elements
Private moLinkedElements() As element
' Variable to store the master element
Private moMasterElement As element
' Instance of the AutoLengths class to handle auto-length operations
Private moAutoLengths As AutoLengths

' Method to set the linked elements from an external array
Public Sub SetLinkedElements(Elements() As element)
    moLinkedElements = Elements
End Sub

' Property to set the master element
Public Property Set SetMasterElement(ByVal El As element)
    Set moMasterElement = El
End Property

' Property to set the instance of AutoLengths
Public Property Set AutoLengthsInstance(ByVal AutoLengths As AutoLengths)
    Set moAutoLengths = AutoLengths
End Property

' Event handler for clicking an item in the ListBox
' Zooms and highlights the selected element in the graphical interface
Private Sub ListBox_Lengths_Click()
    On Error GoTo ErrorHandler
    Dim idx As Long
    If ListBox_Lengths.ListIndex = -1 Then Exit Sub
    idx = SelectedElementIndex()
    If idx = -1 Then Exit Sub
    If Not moLinkedElements(idx) Is Nothing Then
        MSGraphicalInteraction.ZoomEl moLinkedElements(idx)
        MSGraphicalInteraction.HighlightEl moLinkedElements(idx)
    Else
        LangManager.ShowStatusT "AutoLengthsGUIInvalidSelectedElement"
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox_Lengths_Click"
End Sub

' Event handler for double-clicking an item in the ListBox: commit the selection.
Private Sub ListBox_Lengths_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    CommitSelection
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox_Lengths_DblClick"
End Sub

' Keyboard: Enter commits the selected row (same path as double-click), Esc closes.
Private Sub ListBox_Lengths_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            CommitSelection
        Case vbKeyEscape
            Unload Me
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox_Lengths_KeyUp"
End Sub

' OK button: commit the selected row, same path as double-click / Enter.
Private Sub OK_Command_Click()
    On Error GoTo ErrorHandler
    CommitSelection
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.OK_Command_Click"
End Sub

' Cancel button: close the picker without selecting.
Private Sub Cancel_Command_Click()
    On Error GoTo ErrorHandler
    Unload Me
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.Cancel_Command_Click"
End Sub

' Event handler for initializing the UserForm
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Me.Caption = GetTranslation("AutoLengthsGUISelectElementsCaption")
    ' Hide the index column: column 2 carries the moLinkedElements index
    ListBox_Lengths.ColumnCount = 2
    ListBox_Lengths.ColumnWidths = ";0"
    FormUXHelper.SetTip ListBox_Lengths, "AutoLengthsGUISelectElementsListTip"

    ' OK/Cancel buttons
    OK_Command.Caption = GetTranslation("AutoLengthsGUISelectElementsOK_CommandCaption")
    Cancel_Command.Caption = GetTranslation("AutoLengthsGUISelectElementsCancel_CommandCaption")
    OK_Command.Default = True
    Cancel_Command.Cancel = True
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.UserForm_Initialize"
End Sub

' Event handler for querying the close action of the UserForm
' Clears the transient element collection when the form is closed
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    ' Clear the transient element collection
    Set TEC = Nothing
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.UserForm_QueryClose"
End Sub

' Resolve the currently selected row to a validated moLinkedElements index, or -1 if invalid.
Private Function SelectedElementIndex() As Long
    On Error GoTo ErrorHandler
    SelectedElementIndex = -1
    If ListBox_Lengths.ListIndex = -1 Then Exit Function
    Dim idx As Long
    idx = CLng(ListBox_Lengths.List(ListBox_Lengths.ListIndex, 1))
    If idx < LBound(moLinkedElements) Or idx > UBound(moLinkedElements) Then
        LangManager.ShowStatusT "AutoLengthsGUIInvalidSelectedElement"
        Exit Function
    End If
    SelectedElementIndex = idx
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.SelectedElementIndex"
    SelectedElementIndex = -1
End Function

' Commit the currently selected element and close the modeless picker.
Private Sub CommitSelection()
    On Error GoTo ErrorHandler
    Dim idx As Long
    idx = SelectedElementIndex()
    If idx = -1 Then Exit Sub
    If moLinkedElements(idx) Is Nothing Then
        LangManager.ShowStatusT "AutoLengthsGUIInvalidSelectedElement"
        Exit Sub
    End If
    ' Clear the transient element collection (used by MSGraphicalInteraction), select, then close.
    Set TEC = Nothing
    OnElementSelected moLinkedElements(idx)
    Unload Me
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.CommitSelection"
End Sub

' Method to handle the selected element
' Calls the method in the AutoLengths instance to continue the execution with the selected element
Private Sub OnElementSelected(ByVal selectedElement As element)
    On Error GoTo ErrorHandler
    ' Call the method in the existing AutoLengths instance to continue the execution
    moAutoLengths.OnElementSelected selectedElement, moMasterElement
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.OnElementSelected"
End Sub
