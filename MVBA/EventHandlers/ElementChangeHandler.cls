' Class Module: ElementChangeHandler
' Description: Handles element change events.
' This class module is responsible for managing events related to changes in elements within MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, AutoLengths, ARESConfigClass, ARESConstants, LangManager, StringsInEl, ElementInProcesseClass, ErrorHandlerClass
Option Explicit
Implements IChangeTrackEvents

' === PRIVATE MEMBERS ===
Private moIdleHandler As IdleEventHandler
Private mbIdleHandlerRegistered As Boolean

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
    Dim els() As element
    Dim i As Long
    
    Select Case Action
        Case msdChangeTrackActionAdd
            If AfterChange.IsGraphical Then
                If AfterChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If ElementInProcesse.Add(AfterChange) Then
                        EnsureIdleHandlerRegistered
                    End If
                End If
            End If
            
        Case msdChangeTrackActionModify
            If AfterChange.IsGraphical Then
                If AfterChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If ElementInProcesse.Add(AfterChange) Then
                        EnsureIdleHandlerRegistered
                    End If
                End If
            End If
            
        Case msdChangeTrackActionDelete
            If BeforeChange.IsGraphical Then
                If BeforeChange.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    els = Link.GetLink(BeforeChange)
                    If UBound(els) = -1 Then Exit Sub
                    For i = LBound(els) To UBound(els)
                        If ElementInProcesse.Add(els(i)) Then
                            EnsureIdleHandlerRegistered
                        End If
                    Next i
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

' Ensure IdleEventHandler is registered (created once per batch)
Private Sub EnsureIdleHandlerRegistered()
    On Error GoTo ErrorHandler
    
    ' Only register if not already registered
    If Not mbIdleHandlerRegistered Then
        Set moIdleHandler = New IdleEventHandler
        AddEnterIdleEventHandler moIdleHandler
        mbIdleHandlerRegistered = True
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.EnsureIdleHandlerRegistered"
End Sub

' Called by IdleEventHandler to reset the registration flag
Public Sub ResetIdleHandlerFlag()
    On Error GoTo ErrorHandler
    
    mbIdleHandlerRegistered = False
    Set moIdleHandler = Nothing
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ResetIdleHandlerFlag"
End Sub

' ========================================
' PUBLIC METHODS - Called by IdleHandler
' ========================================

' Unified handler for element processing (called by IdleEventHandler)
Public Sub ProcessElement(ByVal oElement As element)
    On Error GoTo ErrorHandler

    Dim AUTO_LENGTH As Boolean
    Dim UPDATE_LENGTH As Boolean
    Dim ONLY_COLOR As Boolean
    Dim TriggerFinded As Boolean
    Dim txts() As String
    Dim Triggers() As String
    Dim i As Long, j As Long
    Dim SplitedTrigger() As String
    Dim FirstPartIndex As Long, SecondPartIndex As Long
    Dim MiddleText As String
    Dim els() As element
    Dim linkedElement As element
    Dim subEl As element
    Dim ELEnum As ElementEnumerator
    Dim fillcolor As Long
    
    AUTO_LENGTH = ARESConfig.ARES_AUTO_LENGTHS.Value
    UPDATE_LENGTH = ARESConfig.ARES_UPDATE_LENGTHS.Value
    TriggerFinded = False
    ONLY_COLOR = ARESConfig.ARES_ONLY_COLOR.Value
    
    ' Check if the element is valid for automatic length processing
    If Not (oElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID) Then
        Exit Sub
    End If

    ' Branch logic based on element type
    If IsTextOrCellElement(oElement) Then
        ' Process text/cell elements directly
        If (AUTO_LENGTH Or UPDATE_LENGTH) Then
            ' Get the texts in the element
            txts = StringsInEl.GetSetTextsInEl(oElement)
            If Not IsArray(txts) Then Exit Sub
            On Error Resume Next
            If UBound(txts) < LBound(txts) Then Exit Sub
            On Error GoTo ErrorHandler
            ' Get the triggers
            Triggers = Split(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARES_VAR_DELIMITER)
            ' Clean up existing numeric values between triggers (recalculate mode)
            For i = LBound(txts) To UBound(txts)
                For j = LBound(Triggers) To UBound(Triggers)
                    SplitedTrigger = Split(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                    If UBound(SplitedTrigger) = 1 Then
                        FirstPartIndex = InStr(1, txts(i), SplitedTrigger(0))
                        SecondPartIndex = InStr(1, txts(i), SplitedTrigger(1))
                        If FirstPartIndex > 0 And SecondPartIndex > FirstPartIndex Then
                            MiddleText = Mid(txts(i), FirstPartIndex + Len(SplitedTrigger(0)), SecondPartIndex - FirstPartIndex - Len(SplitedTrigger(0)))
                            If IsNumericText(MiddleText) Then
                                txts(i) = Left(txts(i), FirstPartIndex + Len(SplitedTrigger(0)) - 1) & Mid(txts(i), SecondPartIndex)
                                StringsInEl.GetSetTextsInEl oElement, Join(txts, ARES_VAR_DELIMITER)
                                TriggerFinded = True
                            End If
                        End If
                    End If
                Next j
            Next i
    
            ' Search for triggers in the texts
            If Not TriggerFinded Then
                For i = LBound(txts) To UBound(txts)
                    For j = LBound(Triggers) To UBound(Triggers)
                        If InStr(1, txts(i), StringsInEl.RemovePattern(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)) > 0 Then
                            TriggerFinded = True
                            Exit For
                        End If
                    Next j
                    If TriggerFinded Then Exit For
                Next i
            End If
    
            ' Update lengths if a trigger was found
            If TriggerFinded Then
                Dim AutoLengths As New AutoLengths
                AutoLengths.Initialize oElement
                AutoLengths.UpdateLengths
                Exit Sub
            End If
        End If
        If ONLY_COLOR Then
            els = Link.GetLink(oElement)
            If UBound(els) = -1 Then Exit Sub
            Set linkedElement = els(1)
            If oElement.IsCellElement Then
                Set ELEnum = oElement.AsCellElement.GetSubElements
                Do While ELEnum.MoveNext
                    Set subEl = ELEnum.Current
                    If subEl.Color <> linkedElement.Color Then
                        If subEl.IsShapeElement Then
                            If subEl.AsShapeElement.FillMode = 2 Then
                                fillcolor = subEl.AsShapeElement.fillcolor
                                subEl.AsShapeElement.Color = linkedElement.Color
                                subEl.AsShapeElement.fillcolor = fillcolor
                            Else
                                subEl.AsShapeElement.Color = linkedElement.Color
                            End If
                        Else
                            subEl.Color = linkedElement.Color
                        End If
                        subEl.Rewrite
                    End If
                Loop
            ElseIf oElement.IsTextElement Then
                If oElement.AsTextElement.Color <> linkedElement.Color Then
                    oElement.AsTextElement.Color = linkedElement.Color
                    oElement.Rewrite
                End If
            ElseIf oElement.IsTextNodeElement Then
                If oElement.AsTextNodeElement.Color <> linkedElement.Color Then
                    oElement.AsTextNodeElement.Color = linkedElement.Color
                    oElement.Rewrite
                End If
            End If
            Exit Sub
        End If
        
    Else
        ' Process geometric elements (line, arc, shape, etc.) - update linked texts
        If Not UPDATE_LENGTH Then Exit Sub
        
        els = Link.GetLink(oElement)
        If UBound(els) = -1 Then Exit Sub
        For i = LBound(els) To UBound(els)
            Set linkedElement = els(i)
            If ElementInProcesse.Add(linkedElement) Then
                ' Recursive call for linked text elements
                ProcessElement linkedElement
            End If
        Next i
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ProcessElement"
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