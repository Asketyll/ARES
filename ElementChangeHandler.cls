' Class Module: ElementChangeHandler
' Description: Handles element change events.
' This class module is responsible for managing events related to changes in elements within MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, Auto_Lengths, ARES_VAR, LangManager

Option Explicit
Implements IChangeTrackEvents

' Event handler for the beginning of an undo/redo action
Private Sub IChangeTrackEvents_BeginUndoRedo(ByVal AfterUndoRedo As Element, ByVal BeforeUndoRedo As Element, ByVal Action As MsdChangeTrackAction, ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler

    ' Add code to handle the beginning of an undo/redo action if needed

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Event handler for when an element is changed
Private Sub IChangeTrackEvents_ElementChanged(ByVal AfterChange As Element, ByVal BeforeChange As Element, ByVal Action As MsdChangeTrackAction, CantBeUndone As Boolean)
    On Error GoTo ErrorHandler

    ' Example: when an element is added
    If Action = msdChangeTrackActionAdd Then
        ' Adding an element. BeforeChange is Nothing
        ' Call a sub or function when an element is added
        HandleElementAdded AfterChange
    End If

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Event handler for the end of an undo/redo action
Private Sub IChangeTrackEvents_FinishUndoRedo(ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler

    ' Add code to handle the end of an undo/redo action if needed

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Event handler for marking changes
Private Sub IChangeTrackEvents_Mark()
    On Error GoTo ErrorHandler

    ' Add code to handle marking changes if needed

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Handle the addition of a new element
Private Sub HandleElementAdded(ByVal NewElement As Element)
    On Error GoTo ErrorHandler

    Dim AUTO_LENGTH As Boolean
    Dim TriggerFinded As Boolean
    Dim Txts() As String
    Dim i As Long

    TriggerFinded = False
    AUTO_LENGTH = ARES_VAR.ARES_AUTO_LENGTHS.Value

    If AUTO_LENGTH And NewElement.GraphicGroup <> ARES_VAR.ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        Txts = StringsInEl.GetSetTextsInEl(NewElement)
        For i = LBound(Txts) To UBound(Txts)
            If InStr(1, Txts(i), ARES_VAR.ARES_LENGTH_TRIGGER.Value, vbBinaryCompare) > 0 Then
                TriggerFinded = True
                Exit For
            End If
        Next i
        If TriggerFinded Then
            Dim autoLengths As New autoLengths
            autoLengths.Initialize NewElement
            autoLengths.UpdateLengths
        End If
    End If

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub
