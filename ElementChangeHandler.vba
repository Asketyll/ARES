' ClassModule: ElementChangeHandler
' Description: Handles element change events.

'Dependencies: Config, Auto_Lengths, ARES_VAR

Option Explicit

Implements IChangeTrackEvents

Private Const ARES_AUTO_LENGTHS_DEFAULT As Boolean = True

' Event handler for the beginning of an undo/redo action
Private Sub IChangeTrackEvents_BeginUndoRedo(ByVal AfterUndoRedo As Element, ByVal BeforeUndoRedo As Element, ByVal Action As MsdChangeTrackAction, ByVal IsUndo As Boolean)
    ' Add code to handle the beginning of an undo/redo action if needed
End Sub

' Event handler for when an element is changed
Private Sub IChangeTrackEvents_ElementChanged(ByVal AfterChange As Element, ByVal BeforeChange As Element, ByVal Action As MsdChangeTrackAction, CantBeUndone As Boolean)
    ' Example: when an element is added
    If Action = msdChangeTrackActionAdd Then
        ' Adding an element. BeforeChange is Nothing
        ' Call a sub or function when an element is added
        HandleElementAdded AfterChange
    End If
End Sub

' Event handler for the end of an undo/redo action
Private Sub IChangeTrackEvents_FinishUndoRedo(ByVal IsUndo As Boolean)
    ' Add code to handle the end of an undo/redo action if needed
End Sub

' Event handler for marking changes
Private Sub IChangeTrackEvents_Mark()
    ' Add code to handle marking changes if needed
End Sub

' Handle the addition of a new element
Private Sub HandleElementAdded(ByVal NewElement As Element)
    Dim AUTO_LENGTH As Boolean
    
    ' Logic to wright Length in text of element
    If Config.GetVar(ARES_VAR.ARES_AUTO_LENGTHS) = "" Then
        If Not Config.SetVar(ARES_VAR.ARES_AUTO_LENGTHS, ARES_AUTO_LENGTHS_DEFAULT) Then
            ShowStatus "Impossible de créer la variable " & ARES_VAR.ARES_AUTO_LENGTHS & " ou de la modifier."
        Else
            ShowStatus ARES_VAR.ARES_AUTO_LENGTHS & " défini à " & ARES_AUTO_LENGTHS_DEFAULT & " par défaut"
        End If
    End If
    
    AUTO_LENGTH = Config.GetVar(ARES_VAR.ARES_AUTO_LENGTHS)
    If AUTO_LENGTH And NewElement.GraphicGroup <> ARES_VAR.ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        If NewElement.IsTextElement Or NewElement.IsTextNodeElement Or NewElement.IsCellElement Then
            Dim autoLengths As New autoLengths
            autoLengths.Initialize NewElement
            autoLengths.UpdateLengths
        End If
    End If
End Sub
