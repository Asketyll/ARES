' ClassModule: ElementChangeHandler
' Description: Handles element change events.

Option Explicit

Implements IChangeTrackEvents

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
    ' Add your logic here to handle the newly added element
End Sub
