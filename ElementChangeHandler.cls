' Class Module: ElementChangeHandler
' Description: Handles element change events.
' This class module is responsible for managing events related to changes in elements within MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, Auto_Lengths, ARES_VAR, LangManager, StringsInEl

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
            If ElementInProcesse.Add(AfterChange) Then
            ' Call a sub or function when an element is added
            HandleElementAdded AfterChange
            Exit Sub
        End If
    End If
    
    If Action = msdChangeTrackActionModify Then
        If ElementInProcesse.Add(AfterChange) Then
            ' Call a sub or function when an element is modified
            HandleElementModified AfterChange
            Exit Sub
        End If
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
Private Sub HandleElementAdded(ByVal NewElement As Element, Optional Recalculate As Boolean = False)
    On Error GoTo ErrorHandler

    Dim AUTO_LENGTH As Boolean
    Dim TriggerFinded As Boolean
    Dim Txts() As String
    Dim Triggers() As String
    Dim i As Long
    Dim j As Long
    Dim SplitedTrigger() As String
    Dim FirstPartIndex As Long
    Dim SecondPartIndex As Long
    Dim MiddleText As String
    Dim k As Long
    
    TriggerFinded = False
    AUTO_LENGTH = ARES_VAR.ARES_AUTO_LENGTHS.Value

    If AUTO_LENGTH And NewElement.GraphicGroup <> ARES_VAR.ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        Txts = StringsInEl.GetSetTextsInEl(NewElement)
        If Not IsArray(Txts) Then
            ElementInProcesse.Remove NewElement
            Exit Sub
        End If
        Triggers = Split(ARES_VAR.ARES_LENGTH_TRIGGER.Value, ARES_VAR.ARES_VAR_DELIMITER)
        If Recalculate Then
            For i = LBound(Txts) To UBound(Txts)
                For j = LBound(Triggers) To UBound(Triggers)
                    SplitedTrigger = Split(Triggers(j), ARES_VAR.ARES_LENGTH_TRIGGER_ID.Value)
                    If UBound(SplitedTrigger) = 1 Then ' Ensure it splits into exactly two parts
                        FirstPartIndex = InStr(1, Txts(i), SplitedTrigger(0))
                        SecondPartIndex = InStr(1, Txts(i), SplitedTrigger(1))
                        If FirstPartIndex > 0 And SecondPartIndex > FirstPartIndex Then
                            MiddleText = Mid(Txts(i), FirstPartIndex + Len(SplitedTrigger(0)), SecondPartIndex - FirstPartIndex - Len(SplitedTrigger(0)))
                            ' Check if MiddleText contains only numbers, spaces, commas, or points
                            For k = 1 To Len(MiddleText)
                                If Not (Mid(MiddleText, k, 1) Like "[0-9 ,.]" Or Mid(MiddleText, k, 1) = " ") Then
                                    Exit For
                                End If
                            Next k
                            If k > Len(MiddleText) Then
                                Txts(i) = Left(Txts(i), FirstPartIndex + Len(SplitedTrigger(0)) - 1) & Mid(Txts(i), SecondPartIndex)
                                StringsInEl.GetSetTextsInEl NewElement, Join(Txts, ARES_VAR.ARES_VAR_DELIMITER)
                                TriggerFinded = True
                            End If
                        End If
                    End If
                Next j
            Next i
        End If
        For i = LBound(Txts) To UBound(Txts)
            For j = LBound(Triggers) To UBound(Triggers)
                If InStr(1, Txts(i), StringsInEl.RemovePattern(Triggers(j), ARES_VAR.ARES_LENGTH_TRIGGER_ID.Value)) > 0 Then
                    TriggerFinded = True
                    Exit For
                End If
            Next j
        Next i
        If TriggerFinded Then
            Dim AutoLengths As New AutoLengths
            AutoLengths.Initialize NewElement
            AutoLengths.UpdateLengths
        Else
            ElementInProcesse.Remove NewElement
        End If
    End If

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Handle the addition of a new element
Private Sub HandleElementModified(ByVal ModifiedElement As Element)
    On Error GoTo ErrorHandler
    
    Dim UPDATE_LENGTH As Boolean
    Dim els() As Element
    Dim i As Byte
    
    UPDATE_LENGTH = ARES_VAR.ARES_UPDATE_LENGTHS.Value
    
    If UPDATE_LENGTH Then
        If ModifiedElement.IsCellElement Or ModifiedElement.IsTextElement Or ModifiedElement.IsTextNodeElement Then
            HandleElementAdded ModifiedElement, True
        Else
            els = Link.GetLink(ModifiedElement)
            For i = LBound(els) To UBound(els)
                If ElementInProcesse.Add(els(i)) Then
                    HandleElementAdded els(i), True
                End If
            Next i
            ElementInProcesse.Remove ModifiedElement
        End If
    End If
    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub
