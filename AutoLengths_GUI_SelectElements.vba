' UserForm: AutoLengths_GUI_SelectElements
' Description: This UserForm is used for selecting an element in a list and returning it.

' Dependencies: MSGraphicalInteraction, AutoLengths, LangManager

Private pLinkedElements() As Element
Private pMasterElement As Element
Private pAutoLengths As autoLengths

Public Sub SetLinkedElements(elements() As Element)
    pLinkedElements = elements
End Sub

Public Property Set SetMasterElement(ByVal El As Element)
    Set pMasterElement = El
End Property

Public Property Set AutoLengthsInstance(ByVal autoLengths As autoLengths)
    Set pAutoLengths = autoLengths
End Property

Private Sub ListBox1_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim selectedIndex As Long
        selectedIndex = ListBox1.List(ListBox1.ListIndex, 1)
        If Not pLinkedElements(selectedIndex) Is Nothing Then
            MSGraphicalInteraction.ZoomEl pLinkedElements(selectedIndex)
            MSGraphicalInteraction.HighlightEl pLinkedElements(selectedIndex)
        Else
            MsgBox GetTranslation("AutoLengthsGUIInvalidSelectedElement"), vbExclamation
        End If
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox1.ListIndex <> -1 Then
        Dim selectedIndex As Long
        selectedIndex = ListBox1.List(ListBox1.ListIndex, 1)
        Me.Hide
        Set TEC = Nothing   'Is public, Used in MSGraphicalInteraction for TransientElement
        OnElementSelected pLinkedElements(selectedIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = GetTranslation("AutoLengthsGUICaption")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If CloseMode = 0 Then MsgBox "Vous devez faire un choix pour que la macro se termine sans encombre."
    'Cancel = CloseMode = 0
    Set TEC = Nothing   'Is public, Used in MSGraphicalInteraction for TransientElement
End Sub

Private Sub OnElementSelected(ByVal selectedElement As Element)
    ' Call the method in the existing AutoLengths instance to continue the execution
    pAutoLengths.OnElementSelected selectedElement, pMasterElement
End Sub
