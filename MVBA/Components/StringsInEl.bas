' Module: StringsInEl
' Description: Provides functions to get and set texts within MicroStation elements.
' This module handles text manipulation for TextElement, TextNodeElement, and CellElement types.
' It supports trigger-based text replacement (for automatic length insertion) and color synchronization.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, CellRedreaw
'
' IMPORTANT NOTES ON TEXTLINE PROPERTY:
' - Color Property is erased if you use TextLine Write Property
' - Using TextLine is not recommended. This feature has been buggy for 20 years and has numerous
'   technical limitations not stated in Bentley's technical documentation.
' - For example, if the TextNodeElement is in a cell, TextLine Property doesn't work.
' - WORKAROUND: Treat the TextNodeElement as a cell composed of TextElements.
'   Create an ElementEnumerator and use GetSubElements to interact directly with the sub-elements.
'   Or use .Color Property to get the color before changes and set it on the TextNodeElement before Rewrite.

Option Explicit

' ========================================
' PUBLIC FUNCTIONS
' ========================================

' Main entry point for getting and setting texts within elements
' This function determines the element type and delegates to the appropriate processor
' Parameters:
'   TextElement - The element containing text to get or set (ByRef to allow updates)
'   txt         - Optional. The text value to insert (typically a length value)
'   Triggers    - Optional. Pipe-delimited trigger patterns (e.g., "(Xx_m)|(Xx_cm)")
'   Color       - Optional. The color to apply to the element (-2 = no change)
' Returns:
'   Array of strings containing the text content of the element
Public Function GetSetTextsInEl(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -2) As String()
    On Error GoTo ErrorHandler
    Dim Result() As String

    ' Only process unlocked elements
    If Not TextElement.IsLocked Then
        Select Case True
            Case TextElement.IsTextElement
                ' Process as a single text element (simple text string)
                Result = ProcessTextElement(TextElement, txt, Triggers, Color)

            Case TextElement.IsTextNodeElement
                ' Process as a text node element (multi-line text)
                Result = ProcessTextNodeElement(TextElement, txt, Triggers, Color)

            Case TextElement.IsCellElement
                ' Process as a cell element (container with nested elements)
                Result = ProcessCellElement(TextElement, txt, Triggers, Color)
        End Select
    End If

    GetSetTextsInEl = Result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.GetSetTextsInEl"
    ' If an error occurs and no triggers specified, return an array with an empty string
    If Triggers = "" Then
        GetSetTextsInEl = Array("")
    End If
End Function

' Removes a specific pattern from a string
' Used to extract the base trigger pattern without the ID placeholder
' Parameters:
'   originalString - The string to process
'   pattern        - The pattern to remove
' Returns:
'   The string with the pattern removed
Public Function RemovePattern(ByVal originalString As String, ByVal pattern As String) As String
    On Error GoTo ErrorHandler
    RemovePattern = Replace(originalString, pattern, "")
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.RemovePattern"
    RemovePattern = originalString
End Function

' ========================================
' PRIVATE FUNCTIONS - Element Processors
' ========================================

' Processes a single TextElement to get or set its text content
' Parameters:
'   TextElement - The TextElement to process
'   txt         - Optional. The text value to insert at trigger locations
'   Triggers    - Optional. Pipe-delimited trigger patterns
'   Color       - Optional. The color to apply (-2 = no change)
' Returns:
'   Array containing the text content (split by empty string for single elements)
Private Function ProcessTextElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -2) As String()
    On Error GoTo ErrorHandler

    ' GET MODE: Return current text content
    If Triggers = "" And txt = "" Then
        ProcessTextElement = Split(TextElement.AsTextElement.text, "")

    ' SET MODE (no triggers): Replace entire text content
    ElseIf Triggers = "" Then
        TextElement.AsTextElement.text = txt
        TextElement.Rewrite
        ProcessTextElement = Split(txt, "")

    ' TRIGGER MODE: Insert text at trigger locations
    Else
        Dim OldTxt As String, NewTxt As String
        Dim trigger() As String, SplitedTriggers() As String
        Dim i As Long
        Dim oldcolor As Long

        ' Save original color for comparison
        oldcolor = TextElement.Color

        ' Apply new color if specified
        If Color <> -2 Then
            TextElement.Color = Color
        End If

        ' Get current text and prepare for modification
        OldTxt = TextElement.AsTextElement.text
        NewTxt = OldTxt

        ' Parse trigger patterns (pipe-delimited)
        trigger = Split(Triggers, ARES_VAR_DELIMITER)

        ' Process each trigger pattern
        ' Trigger format: "prefix" + TRIGGER_ID + "suffix" (e.g., "(" + "Xx_" + "m)")
        ' We replace "prefix" + "suffix" with "prefix" + txt + "suffix"
        For i = LBound(trigger) To UBound(trigger)
            SplitedTriggers = Split(trigger(i), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
            If UBound(SplitedTriggers) = 1 Then
                NewTxt = Replace(NewTxt, SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
            End If
        Next i

        ' Only rewrite if text actually changed
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

' Processes a TextNodeElement (multi-line text) to get or set its content
' Uses sub-element iteration instead of TextLine property due to MicroStation bugs
' Parameters:
'   TextElement - The TextNodeElement to process
'   txt         - Optional. The text value to insert at trigger locations
'   Triggers    - Optional. Pipe-delimited trigger patterns
'   Color       - Optional. The color to apply (-2 = no change)
' Returns:
'   Array of strings, one per text line
Private Function ProcessTextNodeElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -2) As String()
    On Error GoTo ErrorHandler

    ' GET MODE: Return all text lines
    If Triggers = "" And txt = "" Then
        ProcessTextNodeElement = GetTextLines(TextElement)

    ' SET/TRIGGER MODE: Update text lines
    Else
        ProcessTextNodeElement = UpdateTextLines(TextElement, txt, Triggers, Color)
    End If

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.ProcessTextNodeElement"
    ProcessTextNodeElement = Array("")
End Function

' Processes a CellElement by iterating through its sub-elements
' Handles nested cells recursively and applies color changes to all sub-elements
' Parameters:
'   TextElement - The CellElement to process
'   txt         - Optional. The text value to insert at trigger locations
'   Triggers    - Optional. Pipe-delimited trigger patterns
'   Color       - Optional. The color to apply (-2 = no change)
' Returns:
'   Array of strings from the last processed text element
Private Function ProcessCellElement(ByRef TextElement As element, Optional txt As String, Optional Triggers As String, Optional Color As Long = -2) As String()
    On Error GoTo ErrorHandler

    ' === LOCAL VARIABLES ===
    Dim ELEnum As ElementEnumerator     ' Enumerator for iterating sub-elements
    Dim subEl As element                ' Current sub-element being processed
    Dim Result() As String              ' Result from processing current element
    Dim Result2() As String             ' Previous text content for comparison
    Dim oldcolor As Long                ' Original color of the cell
    Dim fillcolor As Long               ' Saved fill color for closed elements
    Dim IsEdited As Boolean             ' Flag: text content was modified
    Dim i As Long                       ' Loop counter

    ' Save original color for sub-element color matching
    oldcolor = TextElement.Color

    ' Get enumerator for sub-elements of the cell
    Set ELEnum = TextElement.AsCellElement.GetSubElements

    ' Process each sub-element
    Do While ELEnum.MoveNext
        Set subEl = ELEnum.Current

        ' Determine sub-element type and delegate to appropriate processor
        Select Case True
            Case subEl.IsTextElement
                ' Save current text for comparison, then process
                Result2 = GetSetTextsInEl(subEl)
                Result = ProcessTextElement(subEl, txt, Triggers, Color)
                ' Check if any text line changed
                For i = 0 To UBound(Result)
                    If Result2(i) <> Result(i) Then
                        IsEdited = True
                    End If
                Next i

            Case subEl.IsTextNodeElement
                ' Save current text for comparison, then process
                Result2 = GetSetTextsInEl(subEl)
                Result = ProcessTextNodeElement(subEl, txt, Triggers, Color)
                ' Check if any text line changed
                For i = 0 To UBound(Result)
                    If Result2(i) <> Result(i) Then
                        IsEdited = True
                    End If
                Next i

            Case subEl.IsCellElement
                ' Recursively process nested cells
                Result2 = GetSetTextsInEl(subEl)
                Result = ProcessCellElement(subEl, txt, Triggers, Color)
                ' Check if any text line changed
                For i = 0 To UBound(Result)
                    If Result2(i) <> Result(i) Then
                        IsEdited = True
                    End If
                Next i
        End Select

        ' Apply color change to sub-elements that match the original cell color
        ' This ensures consistent color across all elements in the cell
        If subEl.Color = oldcolor And Color <> -2 And Color <> oldcolor Then
            ' Handle closed elements (shapes, ellipses, etc.) specially to preserve fill color
            ' ClosedElement interface covers all fillable elements: ShapeElement, EllipseElement, etc.
            ' FillMode = 2 (msdFillModeOutlined) means the element has separate outline and fill colors
            If subEl.IsClosedElement Then
                If subEl.AsClosedElement.FillMode = 2 Then
                    ' Save fill color, update outline color, restore fill color
                    fillcolor = subEl.AsClosedElement.fillcolor
                    subEl.Color = Color
                    subEl.AsClosedElement.fillcolor = fillcolor
                Else
                    ' No fill or solid fill - just update the color
                    subEl.Color = Color
                End If
            Else
                ' Non-closed elements (lines, text, etc.) - simple color update
                subEl.Color = Color
            End If
            subEl.Rewrite
        End If
    Loop

    ' If text was edited, update ATLAS cell label (if applicable)
    If IsEdited Then
        CellRedreaw.ATLASCellLabelUpdate TextElement
    End If

    ProcessCellElement = Result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.ProcessCellElement"
    ProcessCellElement = Array("")
End Function

' ========================================
' PRIVATE HELPER FUNCTIONS
' ========================================

' Extracts all text lines from a TextNodeElement
' Parameters:
'   TextElement - The TextNodeElement to read
' Returns:
'   Array of strings, one per text line
Private Function GetTextLines(ByVal TextElement As element) As String()
    Dim Result() As String
    Dim i As Long

    ' Allocate array for all text lines
    ReDim Result(TextElement.AsTextNodeElement.TextLinesCount - 1)

    ' Extract each line (TextLine is 1-indexed in MicroStation)
    For i = 0 To UBound(Result)
        Result(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
    Next i

    GetTextLines = Result
End Function

' Updates text lines in a TextNodeElement using sub-element iteration
' This avoids MicroStation's buggy TextLine write property
' Parameters:
'   TextElement - The TextNodeElement to update
'   txt         - The text value to insert at trigger locations
'   Triggers    - Pipe-delimited trigger patterns (empty = direct replacement)
'   Color       - Optional. The color to apply (-2 = no change)
' Returns:
'   Array of the new text values
Private Function UpdateTextLines(ByRef TextElement As element, ByVal txt As String, ByVal Triggers As String, Optional Color As Long = -2) As String()
    Dim i As Long, j As Long
    Dim OldTxts() As String             ' Original text content
    Dim NewTxts() As String             ' New text content after modification
    Dim SubTxtEnum As ElementEnumerator ' Enumerator for text sub-elements
    Dim SubTxt As TextElement           ' Current text sub-element
    Dim trigger() As String             ' Array of trigger patterns
    Dim SplitedTriggers() As String     ' Trigger split by ID placeholder
    Dim oldcolor As Long                ' Original element color

    ' Save original color
    oldcolor = TextElement.Color

    ' Allocate arrays for text lines
    ReDim OldTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)
    ReDim NewTxts(TextElement.AsTextNodeElement.TextLinesCount - 1)

    ' DIRECT REPLACEMENT MODE: No triggers, replace lines directly
    If Triggers = "" Then
        ' Split input text by delimiter to get individual lines
        NewTxts = Split(txt, ARES_VAR_DELIMITER)

        ' Only proceed if line counts match
        If UBound(NewTxts) = UBound(OldTxts) Then
            Set SubTxtEnum = TextElement.AsTextNodeElement.GetSubElements
            For i = 0 To UBound(NewTxts)
                SubTxtEnum.MoveNext
                Set SubTxt = SubTxtEnum.Current

                ' Only update if text changed
                If SubTxt.text <> NewTxts(i) Then
                    ' Apply color change if specified
                    If Color <> -2 And oldcolor <> Color Then
                        TextElement.Color = Color
                        oldcolor = Color
                        TextElement.Rewrite
                        SubTxt.Color = Color
                    End If
                    SubTxt.text = NewTxts(i)
                    SubTxt.Rewrite
                    ' Refresh element reference after modification
                    Set TextElement = ActiveModelReference.GetElementById(TextElement.ID)
                End If
            Next i
        End If

    ' TRIGGER MODE: Insert text at trigger locations in each line
    Else
        ' Parse trigger patterns
        trigger = Split(Triggers, ARES_VAR_DELIMITER)

        ' Build new text content by processing each line
        For i = 0 To UBound(OldTxts)
            OldTxts(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
            NewTxts(i) = OldTxts(i)

            ' Apply each trigger pattern to this line
            For j = LBound(trigger) To UBound(trigger)
                SplitedTriggers = Split(trigger(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                If UBound(SplitedTriggers) = 1 Then
                    NewTxts(i) = Replace(NewTxts(i), SplitedTriggers(0) & SplitedTriggers(1), SplitedTriggers(0) & txt & SplitedTriggers(1))
                End If
            Next j
        Next i

        ' Apply changes to sub-elements
        Set SubTxtEnum = TextElement.AsTextNodeElement.GetSubElements
        For i = 0 To UBound(NewTxts)
            SubTxtEnum.MoveNext
            Set SubTxt = SubTxtEnum.Current

            ' Only update if text changed
            If SubTxt.text <> NewTxts(i) Then
                ' Apply color change if specified
                If Color <> -2 And oldcolor <> Color Then
                    TextElement.Color = Color
                    oldcolor = Color
                    TextElement.Rewrite
                    SubTxt.Color = Color
                End If
                SubTxt.text = NewTxts(i)
                SubTxt.Rewrite
                ' Refresh element reference after modification
                Set TextElement = ActiveModelReference.GetElementById(TextElement.ID)
            End If
        Next i
    End If

    UpdateTextLines = NewTxts
End Function
