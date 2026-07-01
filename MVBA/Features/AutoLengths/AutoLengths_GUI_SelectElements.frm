VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoLengths_GUI_SelectElements
   Caption         =   "Select :"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1815
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
Private Sub ListBox1_Click()
    On Error GoTo ErrorHandler
    Dim idx As Long
    If ListBox1.ListIndex = -1 Then Exit Sub
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox1_Click"
End Sub

' Event handler for double-clicking an item in the ListBox: commit the selection.
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    CommitSelection
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox1_DblClick"
End Sub

' Keyboard (AC-10): Enter commits the selected row (same path as double-click), Esc closes.
Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_SelectElements.ListBox1_KeyUp"
End Sub

' Event handler for initializing the UserForm
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Me.Caption = GetTranslation("AutoLengthsGUISelectElementsCaption")
    ' Hide the index column: column 2 carries the moLinkedElements index (AC-10)
    ListBox1.ColumnCount = 2
    ListBox1.ColumnWidths = ";0"
    FormUXHelper.SetTip ListBox1, "AutoLengthsGUISelectElementsListTip"
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
    If ListBox1.ListIndex = -1 Then Exit Function
    Dim idx As Long
    idx = CLng(ListBox1.List(ListBox1.ListIndex, 1))
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

' Commit the currently selected element and close the modeless picker (AC-10).
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
