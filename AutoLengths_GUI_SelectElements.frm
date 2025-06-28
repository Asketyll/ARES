' UserForm: AutoLengths_GUI_SelectElements
' Description: This UserForm is used for selecting an element in a list and returning it.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: MSGraphicalInteraction, AutoLengths, LangManager

Option Explicit

' Array to store the linked elements
Private pLinkedElements() As Element

' Variable to store the master element
Private pMasterElement As Element

' Instance of the AutoLengths class to handle auto-length operations
Private pAutoLengths As autoLengths

' Method to set the linked elements from an external array
Public Sub SetLinkedElements(elements() As Element)
    pLinkedElements = elements
End Sub

' Property to set the master element
Public Property Set SetMasterElement(ByVal El As Element)
    Set pMasterElement = El
End Property

' Property to set the instance of AutoLengths
Public Property Set AutoLengthsInstance(ByVal autoLengths As autoLengths)
    Set pAutoLengths = autoLengths
End Property

' Event handler for clicking an item in the ListBox
' Zooms and highlights the selected element in the graphical interface
Private Sub ListBox1_Click()
    Dim selectedIndex As Long

    ' Check if an item is selected
    If ListBox1.ListIndex <> -1 Then
        ' Get the index of the selected element
        selectedIndex = ListBox1.List(ListBox1.ListIndex, 1)
        ' Check if the selected element is valid
        If Not pLinkedElements(selectedIndex) Is Nothing Then
            ' Zoom and highlight the selected element
            MSGraphicalInteraction.ZoomEl pLinkedElements(selectedIndex)
            MSGraphicalInteraction.HighlightEl pLinkedElements(selectedIndex)
        Else
            ' Show an error message if the selected element is invalid
            MsgBox GetTranslation("AutoLengthsGUIInvalidSelectedElement"), vbExclamation
        End If
    End If
End Sub

' Event handler for double-clicking an item in the ListBox
' Hides the form and triggers the selection of the element
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Long

    ' Check if an item is selected
    If ListBox1.ListIndex <> -1 Then
        ' Get the index of the selected element
        selectedIndex = ListBox1.List(ListBox1.ListIndex, 1)
        ' Hide the form
        Me.Hide
        ' Clear the transient element collection
        Set TEC = Nothing   'Is public, Used in MSGraphicalInteraction for TransientElement
        ' Call the method to handle the selected element
        OnElementSelected pLinkedElements(selectedIndex)
    End If
End Sub

' Event handler for initializing the UserForm
' Sets the caption of the form using a translation key
Private Sub UserForm_Initialize()
    Me.Caption = GetTranslation("AutoLengthsGUICaption")
End Sub

' Event handler for querying the close action of the UserForm
' Clears the transient element collection when the form is closed
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If CloseMode = 0 Then MsgBox "Nop!"
    'Cancel = CloseMode = 0
    ' Clear the transient element collection
    Set TEC = Nothing   'Is public, Used in MSGraphicalInteraction for TransientElement
End Sub

' Method to handle the selected element
' Calls the method in the AutoLengths instance to continue the execution with the selected element
Private Sub OnElementSelected(ByVal selectedElement As Element)
    ' Call the method in the existing AutoLengths instance to continue the execution
    pAutoLengths.OnElementSelected selectedElement, pMasterElement
End Sub
