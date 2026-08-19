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
    ReDim Result(0)

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
        Dim TriggerID As String

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

        ' Cache trigger ID to avoid repeated property access in loop
        TriggerID = ARESConfig.ARES_LENGTH_TRIGGER_ID.Value

        ' Process each trigger pattern
        ' Trigger format: "prefix" + TRIGGER_ID + "suffix" (e.g., "(" + "Xx_" + "m)")
        ' We replace "prefix" + "suffix" with "prefix" + txt + "suffix"
        For i = LBound(trigger) To UBound(trigger)
            SplitedTriggers = Split(trigger(i), TriggerID)
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

    If Not TextElement.IsTextNodeElement Then Exit Function

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
    ReDim Result(0)                     ' Default: empty result if no text sub-elements found
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
                ' Check if any text line changed (guard against mismatched array sizes)
                For i = 0 To UBound(Result)
                    If i <= UBound(Result2) Then
                        If Result2(i) <> Result(i) Then
                            IsEdited = True
                        End If
                    End If
                Next i

            Case subEl.IsTextNodeElement
                ' Save current text for comparison, then process
                Result2 = GetSetTextsInEl(subEl)
                Result = ProcessTextNodeElement(subEl, txt, Triggers, Color)
                ' Check if any text line changed (guard against mismatched array sizes)
                For i = 0 To UBound(Result)
                    If i <= UBound(Result2) Then
                        If Result2(i) <> Result(i) Then
                            IsEdited = True
                        End If
                    End If
                Next i

            Case subEl.IsCellElement
                ' Recursively process nested cells
                Result2 = GetSetTextsInEl(subEl)
                Result = ProcessCellElement(subEl, txt, Triggers, Color)
                ' Check if any text line changed (guard against mismatched array sizes)
                For i = 0 To UBound(Result)
                    If i <= UBound(Result2) Then
                        If Result2(i) <> Result(i) Then
                            IsEdited = True
                        End If
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
    On Error GoTo ErrorHandler

    Dim i As Long, j As Long
    Dim OldTxts() As String             ' Original text content
    Dim NewTxts() As String             ' New text content after modification
    Dim SubTxtEnum As ElementEnumerator ' Enumerator for text sub-elements
    Dim SubTxt As TextElement           ' Current text sub-element
    Dim trigger() As String             ' Array of trigger patterns
    Dim SplitedTriggers() As String     ' Trigger split by ID placeholder
    Dim oldcolor As Long                ' Original element color
    Dim TriggerID As String             ' Cached trigger ID to avoid repeated config access

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

        ' Cache trigger ID to avoid repeated property access in loop
        TriggerID = ARESConfig.ARES_LENGTH_TRIGGER_ID.Value

        ' Build new text content by processing each line
        For i = 0 To UBound(OldTxts)
            OldTxts(i) = TextElement.AsTextNodeElement.TextLine(i + 1)
            NewTxts(i) = OldTxts(i)

            ' Apply each trigger pattern to this line
            For j = LBound(trigger) To UBound(trigger)
                SplitedTriggers = Split(trigger(j), TriggerID)
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
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.UpdateTextLines"
    UpdateTextLines = NewTxts
End Function

' ========================================
' READ-ONLY TEXT AGGREGATION (PropertyCalculation) + TARGETED SUB-TEXT ACCESS (PropertyRendering)
' ========================================

' Concatenate ALL text an element contains, read-only, for PropertyCalculation. Depth-first over a
' cell's GetSubElements order: TextElement -> its whole .Text; TextNodeElement -> each .TextLine
' top-to-bottom; nested CellElement -> recurse. Each fragment is trimmed, empty fragments are
' dropped, kept fragments are joined by Separator (default a single space). Never writes, never
' touches color; returns "" on fault or when no text is present.
' NOTE: this deliberately does NOT reuse GetSetTextsInEl. That function, in GET mode, returns only
' the LAST text-bearing sub-element of a cell, wrapped in a ONE-element array (Split(s, "") on a
' zero-length delimiter yields the whole string in a single entry, NOT one entry per character) - it
' does not aggregate a cell's text. A fresh read-only extractor is used instead of changing that.
' (The constraint that froze GetSetTextsInEl's GET semantics was Auto Lengths, now removed - so changing
' it is no longer forbidden, merely unnecessary here. Do not read this note as a standing prohibition.)
' Thin delegator since epic 15: "exclude nothing" is the nIds = 0 case of GetConcatenatedTextExcluding.
Public Function GetConcatenatedText(ByRef El As element, Optional ByVal Separator As String = " ") As String
    On Error GoTo ErrorHandler
    Dim noIds() As Long             ' deliberately UNALLOCATED - nIds = 0 means it is never touched
    GetConcatenatedText = GetConcatenatedTextExcluding(El, noIds, 0, Separator)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.GetConcatenatedText"
    GetConcatenatedText = ""
End Function

' Same aggregation, SKIPPING every text-bearing sub-element whose SubId is in ExcludeIds(0 .. nIds-1).
' Exists for the renderer's containment rule: a sub-text the RENDERER writes must not feed the
' CellText[...] calc source that governs it, or the value would ratchet on its own output.
' nIds = 0 excludes nothing and is byte-identical to the pre-epic-15 GetConcatenatedText; ExcludeIds is
' then never read, so an UNALLOCATED array is a legal argument. The count is passed EXPLICITLY rather
' than inferred: VBA cannot detect the omission of an Optional typed array (IsMissing only works on
' Variants) and LBound on an unallocated array raises error 9.
Public Function GetConcatenatedTextExcluding(ByRef El As element, ByRef ExcludeIds() As Long, ByVal nIds As Long, Optional ByVal Separator As String = " ") As String
    On Error GoTo ErrorHandler

    Dim sResult As String
    Dim els() As element
    Dim texts() As String
    Dim nFound As Long
    Dim bFaulted As Boolean

    sResult = ""
    nFound = 0
    If Not El Is Nothing Then
        WalkTextBearers El, Separator, sResult, els, texts, nFound, bFaulted, ExcludeIds, nIds
    End If
    GetConcatenatedTextExcluding = sResult
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.GetConcatenatedTextExcluding"
    GetConcatenatedTextExcluding = ""
End Function

' Enumerate the text-bearing sub-elements of El. ids(0 .. n-1) receives the SubIds and texts(0 .. n-1)
' each bearer's WHOLE text (a TextNode's lines joined by vbLf). The SubIds are returned explicitly even
' though they are currently the plain ordinal, so the identity scheme can change without touching callers.
' Returns the count; 0 when the element bears no text; -1 when the walk FAULTED. A partial mapping is
' never handed out: a shifted ordinal would make a later SetTextAtSubId write into the wrong sub-text.
Public Function EnumerateTextSubIds(ByRef El As element, ByRef ids() As Long, ByRef texts() As String) As Long
    On Error GoTo ErrorHandler

    Dim els() As element
    Dim sSink As String
    Dim noIds() As Long
    Dim nFound As Long
    Dim bFaulted As Boolean
    Dim i As Long

    EnumerateTextSubIds = -1
    If El Is Nothing Then Exit Function

    nFound = 0
    WalkTextBearers El, " ", sSink, els, texts, nFound, bFaulted, noIds, 0
    If bFaulted Then Exit Function

    If nFound > 0 Then
        ReDim ids(0 To nFound - 1)
        For i = 0 To nFound - 1
            ids(i) = i
        Next i
    End If

    EnumerateTextSubIds = nFound
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.EnumerateTextSubIds"
    EnumerateTextSubIds = -1
End Function

' Read the WHOLE string of ONE text-bearing sub-element, addressed by SubId. A TextNode's lines come back
' joined by vbLf. Returns "" when the SubId does not exist, when the walk faulted, or when that sub-text is
' genuinely empty - the caller must treat all three the same way (never as "the user wiped the text").
' Deliberately NOT the shape of the existing GET path, which returns a one-entry array and, on a cell,
' only the LAST text-bearing sub-element.
Public Function GetTextAtSubId(ByRef El As element, ByVal SubId As Long) As String
    On Error GoTo ErrorHandler

    Dim els() As element
    Dim texts() As String
    Dim sSink As String
    Dim noIds() As Long
    Dim nFound As Long
    Dim bFaulted As Boolean

    GetTextAtSubId = ""
    If El Is Nothing Then Exit Function
    If SubId < 0 Then Exit Function

    nFound = 0
    WalkTextBearers El, " ", sSink, els, texts, nFound, bFaulted, noIds, 0
    If bFaulted Then Exit Function
    If SubId > nFound - 1 Then Exit Function

    GetTextAtSubId = texts(SubId)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.GetTextAtSubId"
    GetTextAtSubId = ""
End Function

' Write NewText into ONE text-bearing sub-element, addressed by SubId, and touch NOTHING else.
' This is the renderer's only text write. It exists because GetSetTextsInEl's SET path on a cell
' overwrites EVERY text-bearing sub-element, which would destroy a label cell's authentic source text.
' Guarantees: the parent AND the sub-element are checked for IsLocked; identical text is a strict no-op
' (no Rewrite, no ATLAS rebuild); only the addressed sub-element is rewritten; the ATLAS leader-label
' rebuild fires only when something really changed; a TextNode write whose vbLf line count differs from
' the node's TextLinesCount is REFUSED rather than silently partially applied.
' Returns True when the text now reads NewText (including the no-op case), False on any refusal or fault.
Public Function SetTextAtSubId(ByRef El As element, ByVal SubId As Long, ByVal NewText As String) As Boolean
    On Error GoTo ErrorHandler

    Dim els() As element
    Dim texts() As String
    Dim sSink As String
    Dim noIds() As Long
    Dim nFound As Long
    Dim bFaulted As Boolean
    Dim oTarget As element
    Dim bChanged As Boolean

    SetTextAtSubId = False
    If El Is Nothing Then Exit Function
    If SubId < 0 Then Exit Function
    If El.IsLocked Then Exit Function

    nFound = 0
    WalkTextBearers El, " ", sSink, els, texts, nFound, bFaulted, noIds, 0
    If bFaulted Then Exit Function
    If SubId > nFound - 1 Then Exit Function

    Set oTarget = els(SubId)
    If oTarget Is Nothing Then Exit Function
    ' The existing cell SET path only tests the top-level element (:38); a locked sub-element must be
    ' honoured too, so the asymmetry is not reproduced here.
    If oTarget.IsLocked Then Exit Function

    If texts(SubId) = NewText Then
        SetTextAtSubId = True
        Exit Function
    End If

    Select Case True
        Case oTarget.IsTextElement
            oTarget.AsTextElement.text = NewText
            oTarget.Rewrite
            bChanged = True

        Case oTarget.IsTextNodeElement
            bChanged = WriteTextNodeLines(oTarget, NewText)
    End Select

    If Not bChanged Then Exit Function

    ' Rebuild the ATLAS leader-label geometry - only on a real change, mirroring ProcessCellElement.
    If El.IsCellElement Then
        CellRedreaw.ATLASCellLabelUpdate El
    End If

    ' The caller's handle on El is STALE after a sub-element Rewrite, and it keeps being used: on a cell
    ' carrying two rendered sub-texts, the SECOND write's walk and the metadata write that closes the pass
    ' both run off it. The pre-existing UpdateTextLines re-fetches after every sub-write for exactly this
    ' reason. El is ByRef, so refreshing it here reaches the caller's variable.
    RefreshElementHandle El

    SetTextAtSubId = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.SetTextAtSubId"
    SetTextAtSubId = False
End Function

' Re-fetch a model element after one of its sub-elements was rewritten, so the caller stops working off a
' stale handle. Best effort by design and deliberately silent: a cell COMPONENT has no retrievable model
' id, GetElementById then faults or yields Nothing, and the old handle - which is all that ever existed
' for such an element - is simply kept.
Private Sub RefreshElementHandle(ByRef El As element)
    On Error Resume Next
    Dim oFresh As element
    Set oFresh = ActiveModelReference.GetElementById(El.ID)
    If Not oFresh Is Nothing Then Set El = oFresh
End Sub

' Validates if a string contains only numeric characters
' Used to identify length values between trigger patterns
' Allowed characters: digits (0-9), spaces, commas, and decimal points. An EMPTY string passes.
' Moved here from ElementChangeHandler (epic 15) with the rest of the string helpers. It was long
' documented as an Auto Lengths-owned helper; that was wrong even then, and Auto Lengths is now gone.
' This is the NUMERIC GATE of PropertyRendering's D6/D7/D8 forge protection - the first condition of every
' one of those guards. Sole callers: PropertyRendering.SuffixIsSafeAddition and PrefixIsSafeAddition.
' Never reintroduce a local copy: the "no forged number" guarantee rests on ONE definition of numeric.
' Parameters:
'   text - The string to validate
' Returns:
'   True if the string contains only numeric characters
Public Function IsNumericText(ByVal text As String) As Boolean
    On Error GoTo ErrorHandler
    Dim k As Long
    For k = 1 To Len(text)
        If Not (Mid(text, k, 1) Like "[0-9 ,.]") Then
            Exit For
        End If
    Next k
    IsNumericText = (k > Len(text))
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.IsNumericText"
    IsNumericText = False
End Function

' ========================================
' PRIVATE HELPER FUNCTIONS - Shared text-bearer walk
' ========================================

' THE single DFS walk over an element's TEXT-BEARING nodes, and the SOLE source of SubIds. The read API
' (GetConcatenatedTextExcluding) and the targeted write API (EnumerateTextSubIds / GetTextAtSubId /
' SetTextAtSubId) both run off this one procedure ON PURPOSE: two separate walks would drift the moment
' either side changed, and a drifted ordinal makes the renderer write into the WRONG sub-text.
'
' SubId = the DFS ordinal over text-bearing nodes only, counted from 0 at the FIRST one encountered.
' A TextNodeElement counts as ONE bearer (never one per line); a cell root is not itself numbered, so a
' standalone Text/TextNode is SubId 0. Nested cells are recursed into, in GetSubElements order. The
' Select Case order (Text -> TextNode -> Cell) matters: a TextNode also answers IsCellElement.
'
' Two outputs with two DIFFERENT fault contracts, deliberately:
'   - sResult (read): fragments appended exactly as the pre-epic-15 aggregation did - a TextElement
'     contributes its whole .Text, a TextNode ONE FRAGMENT PER LINE, each trimmed, empties dropped,
'     joined by Separator. A fault is logged and that branch abandoned, but the walk CONTINUES, which is
'     the historical behaviour and is preserved byte-for-byte.
'   - els()/texts()/nFound (write): the bearer elements and their WHOLE text (a TextNode's lines joined
'     by vbLf). A partial mapping is dangerous here, so bFaulted is raised and every write-API caller
'     refuses outright rather than address a possibly-shifted ordinal.
Private Sub WalkTextBearers(ByRef El As element, ByVal Separator As String, ByRef sResult As String, ByRef els() As element, ByRef texts() As String, ByRef nFound As Long, ByRef bFaulted As Boolean, ByRef ExcludeIds() As Long, ByVal nIds As Long)
    On Error GoTo ErrorHandler
    Dim ELEnum As ElementEnumerator
    Dim subEl As element
    Dim i As Long
    Dim SubId As Long
    Dim bExcluded As Boolean
    Dim sWhole As String
    Dim sLine As String

    Select Case True
        Case El.IsTextElement
            SubId = nFound
            nFound = nFound + 1
            sWhole = El.AsTextElement.text
            RecordTextBearer El, sWhole, SubId, els, texts
            If Not IsExcludedSubId(SubId, ExcludeIds, nIds) Then
                AppendFragment sWhole, Separator, sResult
            End If

        Case El.IsTextNodeElement
            SubId = nFound
            nFound = nFound + 1
            bExcluded = IsExcludedSubId(SubId, ExcludeIds, nIds)
            sWhole = ""
            For i = 1 To El.AsTextNodeElement.TextLinesCount
                ' TextLine READ is safe; only the TextLine WRITE property is unusable (module header).
                sLine = El.AsTextNodeElement.TextLine(i)
                If i > 1 Then sWhole = sWhole & vbLf
                sWhole = sWhole & sLine
                If Not bExcluded Then
                    AppendFragment sLine, Separator, sResult
                End If
            Next i
            RecordTextBearer El, sWhole, SubId, els, texts

        Case El.IsCellElement
            Set ELEnum = El.AsCellElement.GetSubElements
            Do While ELEnum.MoveNext
                Set subEl = ELEnum.Current
                WalkTextBearers subEl, Separator, sResult, els, texts, nFound, bFaulted, ExcludeIds, nIds
            Loop
    End Select
    Exit Sub

ErrorHandler:
    bFaulted = True
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.WalkTextBearers"
End Sub

' Record one text bearer at its SubId. No handler on purpose: a ReDim fault must propagate to
' WalkTextBearers, which is what raises bFaulted and invalidates the whole mapping.
Private Sub RecordTextBearer(ByRef El As element, ByVal sWhole As String, ByVal SubId As Long, ByRef els() As element, ByRef texts() As String)
    ReDim Preserve els(0 To SubId)
    ReDim Preserve texts(0 To SubId)
    Set els(SubId) = El
    texts(SubId) = sWhole
End Sub

' Membership test over the exclusion set. ExcludeIds may be UNALLOCATED when nIds = 0, so the count is
' tested FIRST and the array is not touched at all in that case (VBA has no short-circuit, hence the
' nested exit rather than an And chain).
Private Function IsExcludedSubId(ByVal SubId As Long, ByRef ExcludeIds() As Long, ByVal nIds As Long) As Boolean
    Dim i As Long
    IsExcludedSubId = False
    If nIds <= 0 Then Exit Function
    For i = 0 To nIds - 1
        If ExcludeIds(i) = SubId Then
            IsExcludedSubId = True
            Exit Function
        End If
    Next i
End Function

' Write a multi-line string into a TextNodeElement through the sub-element walk (the TextLine WRITE
' property is unusable - see the module header). REFUSES when the vbLf line count differs from
' TextLinesCount: the existing path silently skips that case (:360), which would leave the renderer
' believing it had written. Only the lines that actually differ are rewritten.
'
' ALL OR NOTHING, in two phases. TextLinesCount is not a promise about what GetSubElements yields: the
' enumerator can hand back FEWER nodes (the write loop would then end early having changed only some
' lines, and still report success, so the renderer would store LastValues for a text it only partly
' wrote) or a node that is not a TextElement (error 13 raised MID-write, with the earlier lines already
' changed and nothing left to roll them back). Phase 1 therefore inspects and refuses before a single
' line is touched, and phase 2 restores what it wrote if it faults anyway.
Private Function WriteTextNodeLines(ByRef oNode As element, ByVal NewText As String) As Boolean
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim SubTxtEnum As ElementEnumerator
    Dim subEls() As element
    Dim oldTxts() As String
    Dim nSub As Long
    Dim nLines As Long
    Dim i As Long
    Dim nWritten As Long
    Dim bAny As Boolean
    Dim sErrDesc As String, lErrNum As Long, sErrSrc As String

    WriteTextNodeLines = False
    parts = Split(NewText, vbLf)
    nLines = UBound(parts) - LBound(parts) + 1
    If nLines <> oNode.AsTextNodeElement.TextLinesCount Then Exit Function

    ' PHASE 1 - inspect only, write nothing.
    nSub = 0
    Set SubTxtEnum = oNode.AsTextNodeElement.GetSubElements
    Do While SubTxtEnum.MoveNext
        ReDim Preserve subEls(0 To nSub)
        ReDim Preserve oldTxts(0 To nSub)
        Set subEls(nSub) = SubTxtEnum.Current
        If subEls(nSub) Is Nothing Then Exit Function
        If Not subEls(nSub).IsTextElement Then Exit Function
        oldTxts(nSub) = subEls(nSub).AsTextElement.text
        nSub = nSub + 1
    Loop
    If nSub <> nLines Then Exit Function

    ' PHASE 2 - write. A fault here puts back every line already written, so the node is never left
    ' half-rendered while the caller is being told the write failed.
    nWritten = 0
    On Error GoTo RollbackHandler
    For i = 0 To nSub - 1
        ' Counted BEFORE the write, not after: a fault between the assignment and the Rewrite would
        ' otherwise leave line i outside the rollback range. Restoring a line that was never touched is a
        ' no-op anyway - the rollback compares before it writes.
        nWritten = i + 1
        If subEls(i).AsTextElement.text <> parts(LBound(parts) + i) Then
            subEls(i).AsTextElement.text = parts(LBound(parts) + i)
            subEls(i).Rewrite
            bAny = True
        End If
    Next i

    WriteTextNodeLines = bAny
    Exit Function

RollbackHandler:
    ' Capture the error BEFORE any On Error statement resets the Err object.
    sErrDesc = Err.Description
    lErrNum = Err.Number
    sErrSrc = Err.Source
    On Error Resume Next
    For i = 0 To nWritten - 1
        If subEls(i).AsTextElement.text <> oldTxts(i) Then
            subEls(i).AsTextElement.text = oldTxts(i)
            subEls(i).Rewrite
        End If
    Next i
    On Error GoTo 0
    ErrorHandler.HandleError sErrDesc, lErrNum, sErrSrc, "StringsInEl.WriteTextNodeLines"
    WriteTextNodeLines = False
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.WriteTextNodeLines"
    WriteTextNodeLines = False
End Function

' Trim a text fragment; drop it when empty, otherwise append it to sResult separated by Separator.
Private Sub AppendFragment(ByVal sFragment As String, ByVal Separator As String, ByRef sResult As String)
    Dim sTrim As String
    sTrim = Trim(sFragment)
    If Len(sTrim) = 0 Then Exit Sub
    If Len(sResult) = 0 Then
        sResult = sTrim
    Else
        sResult = sResult & Separator & sTrim
    End If
End Sub
