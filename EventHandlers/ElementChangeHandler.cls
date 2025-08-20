' Class Module: ElementChangeHandler
' Description: Handles element change events.
' This class module is responsible for managing events related to changes in elements within MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, AutoLengths, ARESConfigClass, ARESConstants, LangManager, StringsInEl, ElementInProcesseClass, ErrorHandlerClass
Option Explicit
Implements IChangeTrackEvents

' Event handler for the beginning of an undo/redo action
Private Sub IChangeTrackEvents_BeginUndoRedo(ByVal AfterUndoRedo As element, ByVal BeforeUndoRedo As element, ByVal Action As MsdChangeTrackAction, ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    ' Add code to handle the beginning of an undo/redo action if needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_BeginUndoRedo"
    Err.Clear
End Sub

' Event handler for when an element is changed
Private Sub IChangeTrackEvents_ElementChanged(ByVal AfterChange As element, ByVal BeforeChange As element, ByVal Action As MsdChangeTrackAction, CantBeUndone As Boolean)
    On Error GoTo ErrorHandler

    Select Case Action
        Case msdChangeTrackActionAdd
            If AfterChange.IsGraphical Then
                If AfterChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If ElementInProcesse.Add(AfterChange) Then
                        HandleElementAdded AfterChange
                    End If
                End If
            End If
        Case msdChangeTrackActionModify
            If AfterChange.IsGraphical Then
                If AfterChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If ElementInProcesse.Add(AfterChange) Then
                        HandleElementModified AfterChange
                    End If
                End If
            End If
        Case msdChangeTrackActionDelete
            If BeforeChange.IsGraphical Then
                If BeforeChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If ElementInProcesse.Add(BeforeChange) Then
                        HandleElementDeleted BeforeChange
                    End If
                End If
            End If
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_ElementChanged"
    Err.Clear
End Sub

' Event handler for the end of an undo/redo action
Private Sub IChangeTrackEvents_FinishUndoRedo(ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    ' Add code to handle the end of an undo/redo action if needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_FinishUndoRedo"
    Err.Clear
End Sub

' Event handler for marking changes
Private Sub IChangeTrackEvents_Mark()
    On Error GoTo ErrorHandler
    ' Add code to handle marking changes if needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_Mark"
    Err.Clear
End Sub

' Handle the addition of a new element
Private Sub HandleElementAdded(ByVal NewElement As element, Optional Recalculate As Boolean = False)
    On Error GoTo ErrorHandler

    Dim AUTO_LENGTH As Boolean
    Dim TriggerFinded As Boolean
    Dim Txts() As String
    Dim Triggers() As String
    Dim i As Long, j As Long
    Dim SplitedTrigger() As String
    Dim FirstPartIndex As Long, SecondPartIndex As Long
    Dim MiddleText As String, k As Long
    Dim isValidElement As Boolean

    TriggerFinded = False
    AUTO_LENGTH = ARESConfig.ARES_AUTO_LENGTHS.Value

    ' Check if the element is valid for automatic length processing
    isValidElement = (AUTO_LENGTH And NewElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID)
    If Not isValidElement Then
        ElementInProcesse.Remove NewElement
        Exit Sub
    End If

    ' Get the texts in the element
    Txts = StringsInEl.GetSetTextsInEl(NewElement)
    If Not IsArray(Txts) Then
        ElementInProcesse.Remove NewElement
        Exit Sub
    End If
    
    On Error Resume Next
    If UBound(Txts) < LBound(Txts) Then
        ElementInProcesse.Remove NewElement
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get the triggers
    Triggers = Split(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARES_VAR_DELIMITER)

    ' If Recalculate is True, process the triggers in the texts
    If Recalculate Then
        For i = LBound(Txts) To UBound(Txts)
            For j = LBound(Triggers) To UBound(Triggers)
                SplitedTrigger = Split(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                If UBound(SplitedTrigger) = 1 Then ' Ensure it splits into exactly two parts
                    FirstPartIndex = InStr(1, Txts(i), SplitedTrigger(0))
                    SecondPartIndex = InStr(1, Txts(i), SplitedTrigger(1))
                    If FirstPartIndex > 0 And SecondPartIndex > FirstPartIndex Then
                        MiddleText = Mid(Txts(i), FirstPartIndex + Len(SplitedTrigger(0)), SecondPartIndex - FirstPartIndex - Len(SplitedTrigger(0)))
                        ' Check if MiddleText contains only numbers, spaces, commas, or points
                        If IsNumericText(MiddleText) Then
                            Txts(i) = Left(Txts(i), FirstPartIndex + Len(SplitedTrigger(0)) - 1) & Mid(Txts(i), SecondPartIndex)
                            StringsInEl.GetSetTextsInEl NewElement, Join(Txts, ARES_VAR_DELIMITER)
                            TriggerFinded = True
                        End If
                    End If
                End If
            Next j
        Next i
    End If

    ' Search for triggers in the texts
    For i = LBound(Txts) To UBound(Txts)
        For j = LBound(Triggers) To UBound(Triggers)
            If InStr(1, Txts(i), StringsInEl.RemovePattern(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)) > 0 Then
                TriggerFinded = True
                Exit For
            End If
        Next j
    Next i

    ' Update lengths if a trigger was found
    If TriggerFinded Then
        Dim AutoLengths As New AutoLengths
        AutoLengths.Initialize NewElement
        AutoLengths.UpdateLengths
    Else
        ElementInProcesse.Remove NewElement
    End If

    Exit Sub

ErrorHandler:
    ElementInProcesse.Remove NewElement
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.HandleElementAdded"
    Err.Clear
End Sub

' Handle the modification of an element
Sub HandleElementModified(ByVal ModifiedElement As element)
    On Error GoTo ErrorHandler

    Dim UPDATE_LENGTH As Boolean
    Dim els() As element
    Dim i As Long
    Dim linkedElement As element

    UPDATE_LENGTH = ARESConfig.ARES_UPDATE_LENGTHS.Value

    If UPDATE_LENGTH Then
        If IsTextOrCellElement(ModifiedElement) Then
            HandleElementAdded ModifiedElement, True
        Else
            els = Link.GetLink(ModifiedElement)
            For i = LBound(els) To UBound(els)
                Set linkedElement = els(i)
                If ElementInProcesse.Add(linkedElement) Then
                    HandleElementAdded linkedElement, True
                End If
            Next i
            ElementInProcesse.Remove ModifiedElement
        End If
    End If

    Exit Sub

ErrorHandler:
    ElementInProcesse.Remove ModifiedElement
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.HandleElementModified"
    Err.Clear
End Sub

Sub HandleElementDeleted(ByVal DeletedElement As element)
    On Error GoTo ErrorHandler
    
    Dim UPDATE_LENGTH As Boolean
    Dim els() As element
    Dim i As Integer
    
    UPDATE_LENGTH = ARESConfig.ARES_UPDATE_LENGTHS.Value
    
    If UPDATE_LENGTH Then
        els = Link.GetLink(DeletedElement)
        If UBound(els) > 1 Then
            For i = 0 To UBound(els)
                If ElementInProcesse.Add(els(1)) Then
                    HandleElementModified els(1)
                    ElementInProcesse.Remove els(1)
                End If
            Next i
        End If
    End If
    
    ElementInProcesse.Remove DeletedElement
    Exit Sub
    
ErrorHandler:
    ElementInProcesse.Remove DeletedElement
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.HandleElementDeleted"
    Err.Clear
End Sub

' Function to check if an element is a text or cell element
Private Function IsTextOrCellElement(ByVal element As element) As Boolean
    IsTextOrCellElement = (element.IsCellElement Or element.IsTextElement Or element.IsTextNodeElement)
End Function

' Function to check if a string contains only numbers, spaces, commas, or points
Private Function IsNumericText(ByVal text As String) As Boolean
    Dim k As Long
    For k = 1 To Len(text)
        If Not (Mid(text, k, 1) Like "[0-9 ,.]" Or Mid(text, k, 1) = " ") Then
            Exit For
        End If
    Next k
    IsNumericText = (k > Len(text))
End Function
