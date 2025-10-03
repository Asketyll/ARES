' Module: StringsInEl
' Description: This module provides functions to get and set texts within elements in MicroStation.
' It handles text manipulation for TextElement, TextNodeElement, and CellElement types.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, CellRedreaw

                    'Color Property is eras if y use TextLine Write Property
                    'Using TextLine is not recommended. This feature has been buggy for 20 years and has numerous technical limitations not
                    'stated in Bentley's technical documentation. For example, if the TextNodeElement is in a cell, TextLine Property doesn't work.
                    'My workaround is to treat the TextNodeElement as a cell composed of TextElements.
                    'You need to create an ElementEnumerator and GetSubElements to interact directly with the sub-elements.
                    'or .color Property to get befor change and set color of TextNodeElement befor Rewrite

Option Explicit

' Public function to get and set texts within elements
Public Function GetSetTextsInEl(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -1) As String()
    On Error GoTo ErrorHandler
    Dim result() As String
    
    ' Determine the type of element and process accordingly
    If Not TextElement.IsLocked Then
        Select Case True
            Case TextElement.IsTextElement
                ' Process as a single text element
                result = ProcessTextElement(TextElement, txt, Triggers, Color)
            Case TextElement.IsTextNodeElement
                ' Process as a text node element
                result = ProcessTextNodeElement(TextElement, txt, Triggers, Color)
            Case TextElement.IsCellElement
                ' Process as a cell element, which may contain nested elements
                result = ProcessCellElement(TextElement, txt, Triggers, Color)
        End Select
    End If
    
    GetSetTextsInEl = result
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.GetSetTextsInEl"
    ' If an error occurs and Triggers is empty, return an array with an empty string
    If Triggers = "" Then
        GetSetTextsInEl = Array("")
    End If
End Function

' Private function to process a text element
Private Function ProcessTextElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -1) As String()
    On Error GoTo ErrorHandler

    If Triggers = "" And txt = "" Then
        ProcessTextElement = Split(TextElement.AsTextElement.text, "")
    ElseIf Triggers = "" Then
        TextElement.AsTextElement.text = txt
        TextElement.Rewrite
        ProcessTextElement = Split(txt, "")
    Else
        Dim OldTxt As String, NewTxt As String
        Dim trigger() As String, SplitedTriggers() As String
        Dim i As Long
        Dim oldcolor As Long
        oldcolor = TextElement.Color
        
        If Color <> -1 Then
            TextElement.Color = Color
        End If
        OldTxt = TextElement.AsTextElement.text
        NewTxt = OldTxt
        trigger = Split(Triggers, ARES_VAR_DELIMITER)
        
        For i = LBound(trigger) To UBound(trigger)
            SplitedTriggers = Split(trigger(i), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
            If UBound(SplitedTriggers) = 1 Then
                NewTxt = Replace(NewTxt, SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
            End If
        Next i
        
        If NewTxt <> OldTxt Then
            TextElement.AsTextElement.text = NewTxt
            TextElement.Rewrite
        End If
        
        ProcessTextElement = Split(NewTxt, "")
    End If

    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.ProcessTextElement"
    ProcessTextElement = Array("")
End Function

' Private function to process a text node element
Private Function ProcessTextNodeElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -1) As String()
    On Error GoTo ErrorHandler
    
    If Triggers = "" And txt = "" Then
        ProcessTextNodeElement = GetTextLines(TextElement)
    Else
        ProcessTextNodeElement = UpdateTextLines(TextElement, txt, Triggers, Color)
    End If

    Exit Function
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.ProcessTextNodeElement"
    ProcessTextNodeElement = Array("")
End Function

' Helper function to get text lines from a text node element
Private Function GetTextLines(ByVal TextElement As element) As String()
    Dim result() As String
    Dim i As Long
    
    ReDim result(TextElement.AsTextNodeElement.TextLinesCount - 1)
    For i = 0 To UBound(result)
        result(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
    Next i
    
    GetTextLines = result
End Function

' Helper function to update text lines in a text node element
Private Function UpdateTextLines(ByRef TextElement As element, ByVal txt As String, ByVal Triggers As String, Optional Color As Long = -1) As String()
    Dim i As Long, j As Long
    Dim OldTxts() As String, NewTxts() As String
    Dim SubTxtEnum As ElementEnumerator, SubTxt As TextElement
    Dim trigger() As String, SplitedTriggers() As String
    Dim oldcolor As Long
    oldcolor = TextElement.Color
        
    ReDim OldTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)
    ReDim NewTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)

    If Triggers = "" Then
        NewTxts = Split(txt, ARES_VAR_DELIMITER)
        If UBound(NewTxts) = UBound(OldTxts) Then
            Set SubTxtEnum = TextElement.AsTextNodeElement.GetSubElements
            For i = 0 To UBound(NewTxts)
                SubTxtEnum.MoveNext
                Set SubTxt = SubTxtEnum.Current
                If SubTxt.text <> NewTxts(i) Then
                    If Color <> -1 And oldcolor <> Color Then
                        TextElement.Color = Color
                        oldcolor = Color
                        TextElement.Rewrite
                        SubTxt.Color = Color
                    End If
                    SubTxt.text = NewTxts(i)
                    SubTxt.Rewrite
                    Set TextElement = ActiveModelReference.GetElementByID(TextElement.id)
                End If
            Next i
        End If
    Else
        trigger = Split(Triggers, ARES_VAR_DELIMITER)
        
        For i = 0 To UBound(OldTxts)
            OldTxts(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
            NewTxts(i) = OldTxts(i)
            
            For j = LBound(trigger) To UBound(trigger)
                SplitedTriggers = Split(trigger(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                If UBound(SplitedTriggers) = 1 Then
                    NewTxts(i) = Replace(NewTxts(i), SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
                End If
            Next j
        Next i

        Set SubTxtEnum = TextElement.AsTextNodeElement.GetSubElements
        For i = 0 To UBound(NewTxts)
            SubTxtEnum.MoveNext
            Set SubTxt = SubTxtEnum.Current
            If SubTxt.text <> NewTxts(i) Then
                If Color <> -1 And oldcolor <> Color Then
                    TextElement.Color = Color
                    oldcolor = Color
                    TextElement.Rewrite
                    SubTxt.Color = Color
                End If
                SubTxt.text = NewTxts(i)
                SubTxt.Rewrite
                Set TextElement = ActiveModelReference.GetElementByID(TextElement.id)
            End If
        Next i
    End If

    UpdateTextLines = NewTxts
End Function

' Private function to process a cell element, including nested cells
Private Function ProcessCellElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -1) As String()
    On Error GoTo ErrorHandler
    Dim ELEnum As ElementEnumerator
    Dim subel As element
    Dim result() As String
    Dim result2() As String
    Dim oldcolor As Long
    Dim fillcolor As Long
    Dim IsEdited As Boolean
    oldcolor = TextElement.Color
    Dim i As Long
    
    ' Get an enumerator for the sub-elements of the cell
    Set ELEnum = TextElement.AsCellElement.GetSubElements
    Do While ELEnum.MoveNext
        Set subel = ELEnum.Current
        ' Determine the type of sub-element and process accordingly
        Select Case True
            Case subel.IsTextElement
                result2 = GetSetTextsInEl(subel)
                result = ProcessTextElement(subel, txt, Triggers, Color)
                For i = 0 To UBound(result)
                    If result2(i) <> result(i) Then
                        IsEdited = True
                    End If
                Next i
            Case subel.IsTextNodeElement
                result2 = GetSetTextsInEl(subel)
                result = ProcessTextNodeElement(subel, txt, Triggers, Color)
                For i = 0 To UBound(result)
                    If result2(i) <> result(i) Then
                        IsEdited = True
                    End If
                Next i
            Case subel.IsCellElement
                ' Recursively process nested cells
                result2 = GetSetTextsInEl(subel)
                result = ProcessCellElement(subel, txt, Triggers, Color)
                For i = 0 To UBound(result)
                    If result2(i) <> result(i) Then
                        IsEdited = True
                    End If
                Next i
        End Select
        If subel.Color = oldcolor And Color <> -1 And Color <> oldcolor Then
            If subel.IsShapeElement Then
                If subel.AsShapeElement.FillMode = 2 Then
                    fillcolor = subel.AsShapeElement.fillcolor
                    subel.AsShapeElement.Color = Color
                    subel.AsShapeElement.fillcolor = fillcolor
                Else
                    subel.AsShapeElement.Color = Color
                End If
            Else
                subel.Color = Color
            End If
            subel.Rewrite
        End If
    Loop
    
    If IsEdited Then
        CellRedreaw.ATLASCellLabelUpdate TextElement
    End If
    ' Return the result of processing the cell element
    ProcessCellElement = result
    Exit Function
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.ProcessCellElement"
    ProcessCellElement = Array("")
End Function

' Function to remove a specific pattern from a string
Public Function RemovePattern(ByVal originalString As String, ByVal pattern As String) As String
    On Error GoTo ErrorHandler
    ' Use the Replace function to remove the pattern
    RemovePattern = Replace(originalString, pattern, "")
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.RemovePattern"
    RemovePattern = originalString
End Function
