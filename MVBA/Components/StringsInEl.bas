' Module: StringsInEl
' Description: Provides functions to get and set texts within MicroStation elements.
' This module handles text manipulation for TextElement, TextNodeElement, and CellElement types.
' Two services: read-only aggregation of everything an element contains (PropertyCalculation), and
' targeted per-sub-text read/write addressed by SubId (PropertyRendering). The trigger-based
' replacement that served Auto Lengths was deleted with it (2026-09-10).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ErrorHandlerClass, CellRedreaw, CallStackClass
'
' IMPORTANT NOTES ON TEXTLINE PROPERTY - the two halves do NOT share a reputation:
' - WRITE is the broken half. TextLine(i) = s erases the element's Color, and does not work at all on a
'   TextNodeElement nested in a cell. Buggy for 20 years, with limitations Bentley documents nowhere.
' - READ is fine, cell-nested nodes included (verified 2026-09-10 - see mvba-cheatsheet.md for the
'   measurement). WalkTextBearers reads through it deliberately.
' - WORKAROUND for WRITING: treat the TextNodeElement as what it structurally is, a cell composed of
'   TextElements - walk GetSubElements and write each sub-element (see WriteTextNodeLines).
'   Or read .Color first and set it back on the TextNodeElement before Rewrite.

Option Explicit

' ========================================
' PUBLIC FUNCTIONS
' ========================================

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
' READ-ONLY TEXT AGGREGATION (PropertyCalculation) + TARGETED SUB-TEXT ACCESS (PropertyRendering)
' ========================================

' Concatenate ALL text an element contains, read-only, for PropertyCalculation. Depth-first over a
' cell's GetSubElements order: TextElement -> its whole .Text; TextNodeElement -> each .TextLine
' top-to-bottom; nested CellElement -> recurse. Each fragment is trimmed, empty fragments are
' dropped, kept fragments are joined by Separator (default a single space). Never writes, never
' touches color; returns "" on fault or when no text is present.
' Written as a fresh read-only extractor rather than on top of the legacy GetSetTextsInEl, which
' returned only the LAST text-bearing sub-element of a cell and never aggregated. That function went
' with Auto Lengths (deleted 2026-09-10, its last callers being its own tests).
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
' GroupCellText[...] calc source that governs it, or the value would ratchet on its own output.
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
' This is the renderer's only text write. It exists because the legacy SET path (GetSetTextsInEl,
' deleted 2026-09-10) overwrote EVERY text-bearing sub-element of a cell, which would destroy a label
' cell's authentic source text.
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
    Dim bStackPushed As Boolean      ' Guards ErrorHandler's Pop against a fault raised by El.IsLocked
                                      ' below, i.e. BEFORE the Push a few lines down ever runs

    SetTextAtSubId = False
    If El Is Nothing Then Exit Function
    If SubId < 0 Then Exit Function
    If El.IsLocked Then Exit Function

    ' Pushed only past the cheap validity guards above. Every Exit Function below this point (7 of
    ' them) plus ErrorHandler pops this frame.
    CallStack.Push "StringsInEl.SetTextAtSubId", El
    bStackPushed = True

    nFound = 0
    WalkTextBearers El, " ", sSink, els, texts, nFound, bFaulted, noIds, 0
    If bFaulted Then
        CallStack.Pop
        Exit Function
    End If
    If SubId > nFound - 1 Then
        CallStack.Pop
        Exit Function
    End If

    Set oTarget = els(SubId)
    If oTarget Is Nothing Then
        CallStack.Pop
        Exit Function
    End If
    ' The existing cell SET path only tests the top-level element (:38); a locked sub-element must be
    ' honoured too, so the asymmetry is not reproduced here.
    If oTarget.IsLocked Then
        CallStack.Pop
        Exit Function
    End If

    If texts(SubId) = NewText Then
        SetTextAtSubId = True
        CallStack.Pop
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

    If Not bChanged Then
        CallStack.Pop
        Exit Function
    End If

    ' Rebuild the ATLAS leader-label geometry - only on a real change.
    If El.IsCellElement Then
        CellRedreaw.ATLASCellLabelUpdate El
    End If

    ' The caller's handle on El is STALE after a sub-element Rewrite, and it keeps being used: on a cell
    ' carrying two rendered sub-texts, the SECOND write's walk and the metadata write that closes the pass
    ' both run off it. El is ByRef, so refreshing it here reaches the caller's variable.
    RefreshElementHandle El

    SetTextAtSubId = True
    CallStack.Pop
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "StringsInEl.SetTextAtSubId"
    If bStackPushed Then CallStack.Pop
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

' Where a text-bearing element sits. A TextElement and a TextNodeElement both carry an Origin, but under
' different interfaces, and that two-way dispatch is the ONLY thing the callers of this ever shared -
' CellRedreaw asks it of a cell's first text sub-element, PropertyCalculation of the bearing element
' itself, as two branches of a much wider anchor cascade. Neither question belongs here; the dispatch
' does. Returns False when El is neither flavour, or when the read faults - a caller must never
' fabricate a point. Silent by design: each caller keeps its own reporting contract.
Public Function GetTextAnchor(ByRef El As element, ByRef ptOut As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetTextAnchor = False
    If El Is Nothing Then Exit Function

    If El.IsTextElement Then
        ptOut = El.AsTextElement.Origin
    ElseIf El.IsTextNodeElement Then
        ptOut = El.AsTextNodeElement.Origin
    Else
        Exit Function
    End If

    GetTextAnchor = True
    Exit Function

ErrorHandler:
    GetTextAnchor = False
End Function

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
