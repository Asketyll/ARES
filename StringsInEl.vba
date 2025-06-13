' Module: StringsInEl
' Description: This module provides functions to get and set texts within elements in MicroStation.

' Dependencies: ARES_VAR

                    'Color Property is eras if y use TextLine Write Property
                    'Using TextLine is not recommended. This feature has been buggy for 20 years and has numerous technical limitations not
                    'stated in Bentley's technical documentation. For example, if the TextNodeElement is in a cell, TextLine Property doesn't work.
                    'My workaround is to treat the TextNodeElement as a cell composed of TextElements.
                    'You need to create an ElementEnumerator and GetSubElements to interact directly with the sub-elements.
                    'or .color Property to get befor change and set color of TextNodeElement befor Rewrite

Option Explicit

Public Function GetSetTextsInEl(ByVal TextElement As Element, txt As String, Optional Triggers As String) As String()
    On Error GoTo ErrorHandler

    Dim result() As String

    ' Determine the type of element and process accordingly
    Select Case True
        Case TextElement.IsTextElement
            ' Process as a single text element
            result = ProcessTextElement(TextElement, txt, Triggers)

        Case TextElement.IsTextNodeElement
            ' Process as a text node element
            result = ProcessTextNodeElement(TextElement, txt, Triggers)

        Case TextElement.IsCellElement
            ' Process as a cell element, which may contain nested elements
            result = ProcessCellElement(TextElement, txt, Triggers)
    End Select

    GetSetTextsInEl = result
    Exit Function

ErrorHandler:
    ' If an error occurs and Triggers is empty, return an array with an empty string
    If Triggers = "" Then
        GetSetTextsInEl = Array("")
    End If
End Function

Private Function ProcessTextElement(ByVal TextElement As Element, txt As String, Optional Triggers As String) As String()
    On Error GoTo ErrorHandler
    
    ' Process a single text element
    Dim OldTxt As String
    Dim NewTxt As String
    Dim trigger() As String
    Dim SplitedTriggers() As String
    Dim i As Long

    ' If no Triggers are provided, split the text element's text into an array
    If Triggers = "" Then
        ProcessTextElement = Split(TextElement.AsTextElement.Text, "")
    Else
        ' Retrieve the current text of the element
        OldTxt = TextElement.AsTextElement.Text
        ' Split the Triggers into an array using the delimiter
        trigger = Split(Triggers, ARES_VAR.ARES_VAR_DELIMITER)
        NewTxt = OldTxt

        ' Loop through each Trigger and process replacements
        For i = LBound(trigger) To UBound(trigger)
            ' Split the Trigger into parts using the trigger ID
            SplitedTriggers = Split(trigger(i), ARES_VAR.ARES_LENGTH_TRIGGER_ID.Value)
            ' If the Trigger is valid (contains the ID), perform the replacement
            If UBound(SplitedTriggers) = 1 Then
                NewTxt = Replace(NewTxt, SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
            End If
        Next i

        ' If the text has changed, update the element's text and rewrite it
        If NewTxt <> OldTxt Then
            TextElement.AsTextElement.Text = NewTxt
            TextElement.Rewrite
        End If
        ' Return the new text as an array
        ProcessTextElement = Split(NewTxt, "")
    End If
    Exit Function
    
ErrorHandler:
    ProcessTextElement = Array("")
End Function

Private Function ProcessTextNodeElement(ByVal TextElement As Element, txt As String, Optional Triggers As String) As String()
    On Error GoTo ErrorHandler
    
    ' Process a text node element
    Dim i As Long, j As Long
    Dim OldTxts() As String
    Dim NewTxts() As String
    Dim result() As String
    Dim SubTxtEnum As elementEnumerator
    Dim SubTxt As TextElement
    Dim trigger() As String
    Dim SplitedTriggers() As String

    ' If no Triggers are provided, retrieve all text lines from the text node element
    If Triggers = "" Then
        ReDim result(TextElement.AsTextNodeElement.TextLinesCount - 1)
        For i = 0 To UBound(result)
            result(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
        Next i
        ProcessTextNodeElement = result
    Else
        ' Initialize arrays to hold the old and new text lines
        ReDim OldTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)
        ReDim NewTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)
        ' Split the Triggers into an array using the delimiter
        trigger = Split(Triggers, ARES_VAR.ARES_VAR_DELIMITER)

        ' Loop through each text line and process replacements
        For i = 0 To UBound(OldTxts)
            OldTxts(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
            NewTxts(i) = OldTxts(i)

            ' Loop through each Trigger and process replacements
            For j = LBound(trigger) To UBound(trigger)
                ' Split the Trigger into parts using the trigger ID
                SplitedTriggers = Split(trigger(i), ARES_VAR.ARES_LENGTH_TRIGGER_ID.Value)
                ' If the Trigger is valid (contains the ID), perform the replacement
                If UBound(SplitedTriggers) = 1 Then
                    NewTxts(i) = Replace(NewTxts(i), SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
                End If
            Next j
        Next i

        ' Update the text of each sub-element if it has changed
        Set SubTxtEnum = TextElement.AsTextNodeElement.GetSubElements
        For i = 0 To UBound(NewTxts)
            SubTxtEnum.MoveNext
            Set SubTxt = SubTxtEnum.Current
            If SubTxt.Text <> NewTxts(i) Then
                SubTxt.Text = NewTxts(i)
                SubTxt.Rewrite
            End If
        Next i
        ' Return the new text lines as an array
        ProcessTextNodeElement = NewTxts
    End If
    Exit Function
    
ErrorHandler:
    ProcessTextNodeElement = Array("")
End Function

Private Function ProcessCellElement(ByVal TextElement As Element, txt As String, Optional Triggers As String) As String()
    On Error GoTo ErrorHandler
    
    ' Process a cell element, including nested cells
    Dim ElEnum As elementEnumerator
    Dim SubEl As Element
    Dim result() As String

    ' Get an enumerator for the sub-elements of the cell
    Set ElEnum = TextElement.AsCellElement.GetSubElements
    Do While ElEnum.MoveNext
        Set SubEl = ElEnum.Current
        ' Determine the type of sub-element and process accordingly
        Select Case True
            Case SubEl.IsTextElement
                result = ProcessTextElement(SubEl, txt, Triggers)
            Case SubEl.IsTextNodeElement
                result = ProcessTextNodeElement(SubEl, txt, Triggers)
            Case SubEl.IsCellElement
                ' Recursively process nested cells
                result = ProcessCellElement(SubEl, txt, Triggers)
        End Select
    Loop
    ' Return the result of processing the cell element
    ProcessCellElement = result
    Exit Function
    
ErrorHandler:
    ProcessCellElement = Array("")
End Function
