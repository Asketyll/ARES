' Module: RuleEditorUX
' Description: Reusable rule-editor UX: the one-rule editable-ComboBox mechanics (sentinel, clean-pick
'              index, rebuild) and the read-only coloured syntax preview (runtime Labels, metachar colours,
'              keyword bold, contradiction red, MeasureCharAdvance calibration, per-label StdFont). Grammar-
'              agnostic - the caller passes the render string + validity + red segments + bold-keyword list.
'              Used by both options forms (Property Tagging + Property Calculation, epic 14).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass (+ MSForms, stdole)

Option Explicit

' Rule separator for the editable-ComboBox split/join (mirrors the grammar modules' RULE_SEPARATOR).
Private Const RULE_SEP As String = ";"

' --- Coloured preview: runtime Labels rendered inside the caller's Frame ---
' Each label's name starts with this prefix so ClearPreview can remove exactly the preview labels from the
' passed Frame (stateless, per-Frame - the preview Frame contains only these runtime labels).
Private Const PREVIEW_PREFIX As String = "lblPreview"
' Monotonic sequence for unique Label names (never reset, so a failed Remove can never cause a name clash).
Private mlPreviewSeq As Long

' Preview colours (OLE_COLOR = RGB Long; tunable live). Render priority: contradiction RED > metachar > mono.
Private Const PREVIEW_PINK   As Long = 220 + 60 * 256& + 160 * 65536     ' @   group modifier    RGB(220,60,160)
Private Const PREVIEW_BLUE   As Long = 0 + 0 * 256& + 160 * 65536        ' &   AND               RGB(0,0,160)
Private Const PREVIEW_ORANGE As Long = 255 + 140 * 256& + 0 * 65536      ' !   negation          RGB(255,140,0)
Private Const PREVIEW_GREEN  As Long = 0 + 140 * 256& + 0 * 65536        ' */? wildcards         RGB(0,140,0)
Private Const PREVIEW_RED    As Long = 200 + 0 * 256& + 0 * 65536        ' contradiction segment RGB(200,0,0)
Private Const PREVIEW_MONO   As Long = 0                                 ' everything else       RGB(0,0,0)
Private Const PREVIEW_FONT   As String = "Consolas"                     ' fixed-pitch -> runs stay aligned
Private Const PREVIEW_SIZE   As Single = 9
Private Const PREVIEW_X0     As Single = 2                              ' left margin inside the preview Frame
Private Const PREVIEW_Y0     As Single = 2                              ' top margin inside the preview Frame
Private Const PREVIEW_CHARW  As Single = PREVIEW_SIZE * 0.55            ' FALLBACK advance/char if MeasureCharAdvance fails (Consolas ~0.55 em); tunable
Private Const PREVIEW_INSET  As Single = 4                             ' a Label's fixed left/right text inset
Private Const PREVIEW_CHARH  As Single = PREVIEW_SIZE * 1.7            ' line height (incl. margins)

'######################################################################################################################
'                                          EDITOR MECHANICS (grammar-agnostic)
'######################################################################################################################

' Re-seed the editable ComboBox from a raw ";"-joined value: clear, split on ";" (trimmed, empties dropped),
' add each item, append the trailing empty "new rule" sentinel (UI only, never stored), reset the edit text.
' The caller keeps its own edit-index (reset to -1 after this) and triggers the preview.
Public Sub SeedRulesCombo(ByVal combo As MSForms.ComboBox, ByVal sRawValue As String)
    On Error GoTo ErrorHandler
    combo.Clear
    Dim vRules As Variant, i As Long
    vRules = Split(sRawValue, RULE_SEP)
    For i = LBound(vRules) To UBound(vRules)
        If Len(Trim(vRules(i))) > 0 Then combo.AddItem Trim(vRules(i))
    Next i
    combo.AddItem ""                             ' trailing "new rule" sentinel (UI only, never stored)
    combo.text = ""
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RuleEditorUX.SeedRulesCombo"
End Sub

' Return the edit index the caller should hold. On a CLEAN pick (Text = List(ListIndex)) it becomes that
' index, or -1 when the pick is the trailing sentinel (ADD mode). Otherwise (the user is typing, ListIndex
' reset) the current index is PRESERVED - so the captured pick survives editing. The caller assigns:
'   mRuleEditIndex = RuleEditorUX.CaptureEditIndex(combo, mRuleEditIndex)
' (lCurrentIndex is required to preserve the capture across keystrokes - a bare, index-less form would
' reset it every _Change and break replace-vs-append.)
Public Function CaptureEditIndex(ByVal combo As MSForms.ComboBox, ByVal lCurrentIndex As Long) As Long
    On Error GoTo ErrorHandler

    CaptureEditIndex = lCurrentIndex             ' preserve unless a clean pick updates it
    If combo.ListIndex >= 0 Then
        If combo.text = combo.List(combo.ListIndex) Then
            If combo.ListIndex = combo.ListCount - 1 Then
                CaptureEditIndex = -1
            Else
                CaptureEditIndex = combo.ListIndex
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    CaptureEditIndex = lCurrentIndex
End Function

' Rebuild the ";"-joined rules value from the current combo items with the edit applied, using the edit
' index and the (already-validated) canonical form (or "" to remove):
'   index >= 0, canonical non-empty -> REPLACE that item     index >= 0, empty -> REMOVE that item
'   index <  0, canonical non-empty -> APPEND                index <  0, empty -> no-op
' The empty "new rule" sentinel is skipped so it is never stored. Validation stays in the caller; this only
' rebuilds/joins. Returns the new stored value.
Public Function RebuildRules(ByVal combo As MSForms.ComboBox, ByVal lEditIndex As Long, ByVal sCanonical As String) As String
    On Error GoTo ErrorHandler

    Dim bHasIndex As Boolean
    bHasIndex = (lEditIndex >= 0)

    Dim rebuilt() As String
    Dim nCount As Long, i As Long, w As Long
    Dim bIsTarget As Boolean
    Dim sItem As String
    nCount = combo.ListCount
    ReDim rebuilt(0 To nCount)                    ' room for every item + one possible append
    w = 0
    For i = 0 To nCount - 1
        bIsTarget = False
        If bHasIndex Then
            If i = lEditIndex Then bIsTarget = True
        End If
        If bIsTarget Then
            If Len(sCanonical) > 0 Then
                rebuilt(w) = sCanonical           ' replace with the canonical form
                w = w + 1
            End If
            ' empty -> skip (remove)
        Else
            sItem = Trim(combo.List(i))
            If Len(sItem) > 0 Then                ' skip the empty "new rule" sentinel (never stored)
                rebuilt(w) = sItem
                w = w + 1
            End If
        End If
    Next i
    If Not bHasIndex Then                         ' free-typed rule -> append (canonical form)
        If Len(sCanonical) > 0 Then
            rebuilt(w) = sCanonical
            w = w + 1
        End If
    End If

    Dim sJoined As String
    If w = 0 Then
        sJoined = ""
    Else
        ReDim Preserve rebuilt(0 To w - 1)
        sJoined = Join(rebuilt, RULE_SEP)
    End If
    RebuildRules = sJoined
    Exit Function

ErrorHandler:
    ' Propagate to the caller's commit handler so it does NOT write on a fault (byte-preserving: the
    ' original inline rebuild shared CommitRuleEdit's handler, which writes nothing on error).
    Err.Raise Err.Number, "RuleEditorUX.RebuildRules", Err.Description
End Function

' Match a ComboBox' font to the coloured preview (fixed-pitch), so the combo text and the preview below it
' share the same character widths. A FRESH StdFont is assigned via Set (never combo.Font.* - the shared
' Font-object trap); the dropdown list also switches to the preview font (intended). The 9pt font can be a
' touch taller than a combo drawn for 8pt, so bump the height at runtime ONLY if it would be too short.
Public Sub MatchComboFont(ByVal combo As MSForms.ComboBox)
    On Error GoTo ErrorHandler
    Dim fCombo As stdole.StdFont
    Set fCombo = New stdole.StdFont
    fCombo.Name = PREVIEW_FONT
    fCombo.Size = PREVIEW_SIZE
    Set combo.Font = fCombo
    If combo.Height < PREVIEW_SIZE * 2 Then combo.Height = PREVIEW_SIZE * 2
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RuleEditorUX.MatchComboFont"
End Sub

'######################################################################################################################
'                                          COLOURED PREVIEW (grammar-agnostic, data in)
'######################################################################################################################

' Render sRender into previewFrame as coloured (optionally bold/underlined) runtime Labels: @ pink, & dark
' blue, ! orange, */? green, and RED over each conflicting redSegments range (valid rules only); the tokens
' in boldKeywords render bold (whether followed by "[" or standing alone); an invalid rule (bValid = False)
' is underlined and prefixed with a red-bold ballot-X marker. Read-only + fail-safe (any fault clears the
' preview and logs - a broken cosmetic preview must never break the editor). The caller passes the render
' string + validity + red segments (from ITS grammar); this module runs its grammar-specific validation.
Public Sub RenderPreview(ByVal previewFrame As MSForms.Frame, ByVal sRender As String, ByVal bValid As Boolean, ByRef redSegments() As String, ByRef boldKeywords() As String)
    On Error GoTo ErrorHandler

    ClearPreview previewFrame
    If Len(Trim(sRender)) = 0 Then Exit Sub       ' empty -> cleared (no Labels)

    ' Per-character colour + bold maps. MarkSegmentsRed touches colours ONLY, so bold never shifts the red
    ' character positions.
    Dim colours() As Long
    Dim bolds() As Boolean
    colours = BuildColourMap(sRender)
    bolds = BuildBoldMap(sRender, boldKeywords)
    If bValid Then
        MarkSegmentsRed sRender, redSegments, colours
    End If

    ' Stop laying runs out once past the Frame's inner width - a long rule is clipped on the right (no wrap).
    Dim dMaxX As Single
    dMaxX = 0
    On Error Resume Next
    dMaxX = previewFrame.InsideWidth
    On Error GoTo ErrorHandler

    ' Coalesce neighbours sharing the SAME (colour, bold) into runs, laid out left-to-right. Each run
    ' advances x by its EXACT text width so runs sit seamlessly (fixed pitch; transparent labels hide the
    ' generous right padding that overlaps the next run).
    Dim n As Long, i As Long, runStart As Long
    Dim x As Single
    Dim bBreak As Boolean
    Dim bUnderline As Boolean
    Dim dCharW As Single
    n = Len(sRender)
    x = PREVIEW_X0
    runStart = 1
    bUnderline = Not bValid
    dCharW = MeasureCharAdvance(previewFrame)      ' real rendered advance (GDI pixel/DPI), not the theoretical const

    ' Invalid rule: an error cue - a red-bold ballot-X (U+2717) marker + space at the head, and every text
    ' run underlined (spell-checker style). The glyph is built with ChrW (NEVER a literal - the .bas is ANSI
    ' and would eat it); if MSForms/Consolas does not render it, replace ChrW(&H2717) with "X". No red
    ' analysis runs on an invalid rule, so the marker's leading x offset affects no colour/segment mapping.
    If Not bValid Then
        x = EmitRun(previewFrame, ChrW(&H2717) & " ", PREVIEW_RED, True, False, dCharW, x)
    End If

    For i = 1 To n
        bBreak = False
        If i = n Then
            bBreak = True
        Else
            If colours(i + 1) <> colours(runStart) Then bBreak = True
            If bolds(i + 1) <> bolds(runStart) Then bBreak = True
        End If
        If bBreak Then
            x = EmitRun(previewFrame, Mid(sRender, runStart, i - runStart + 1), colours(runStart), bolds(runStart), bUnderline, dCharW, x)
            runStart = i + 1
        End If
        If dMaxX > 0 Then
            If x >= dMaxX Then Exit For
        End If
    Next i
    Exit Sub

ErrorHandler:
    ClearPreview previewFrame
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RuleEditorUX.RenderPreview"
End Sub

' Per-character colour map (1-based) for the render string: metachars coloured, everything else monochrome.
' Colouring metachars wherever they appear (not only at bracket depth 0) is acceptable for a preview - a
' literal & / @ inside [...] is rare and harmless. Contradiction red is overlaid separately (valid only).
Private Function BuildColourMap(ByVal s As String) As Long()
    On Error GoTo ErrorHandler

    Dim colours() As Long
    Dim n As Long, i As Long
    Dim ch As String
    n = Len(s)
    If n < 1 Then n = 1                           ' never ReDim(1 To 0)
    ReDim colours(1 To n)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "@"
                colours(i) = PREVIEW_PINK
            Case "&"
                colours(i) = PREVIEW_BLUE
            Case "!"
                colours(i) = PREVIEW_ORANGE
            Case "*", "?"
                colours(i) = PREVIEW_GREEN
            Case Else
                colours(i) = PREVIEW_MONO
        End Select
    Next i
    BuildColourMap = colours
    Exit Function

ErrorHandler:
    ' Fail-safe: an all-monochrome map of the right size so the caller can still render.
    ReDim colours(1 To IIf(Len(s) < 1, 1, Len(s)))
    For i = LBound(colours) To UBound(colours)
        colours(i) = PREVIEW_MONO
    Next i
    BuildColourMap = colours
End Function

' Overlay PREVIEW_RED on the characters of each conflicting segment (an exact substring of the canonical
' text). Search from a running position so repeated segments map to distinct ranges; red overrides any
' metachar colour in range (a dead condition reads as unmistakably red).
Private Sub MarkSegmentsRed(ByVal sCanonical As String, ByRef segments() As String, ByRef colours() As Long)
    On Error GoTo ErrorHandler

    Dim si As Long, p As Long, k As Long, pos As Long
    pos = 1
    For si = LBound(segments) To UBound(segments)
        If Len(segments(si)) > 0 Then
            p = InStr(pos, sCanonical, segments(si))
            If p > 0 Then
                For k = p To p + Len(segments(si)) - 1
                    If k >= LBound(colours) Then
                        If k <= UBound(colours) Then colours(k) = PREVIEW_RED
                    End If
                Next k
                pos = p + Len(segments(si))
            End If
        End If
    Next si
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RuleEditorUX.MarkSegmentsRed"
End Sub

' Per-character BOLD map (1-based): each depth-0 run of letters that (case-insensitively) equals a member
' of boldKeywords renders bold - a keyword before a "[" (Lvl[..]/Cell[..]/Type[..]) AND a bare keyword
' standing alone (so a calc source like Coord/Id also bolds). Names inside [...] are at bracket depth > 0
' and are never bolded.
Private Function BuildBoldMap(ByVal s As String, ByRef boldKeywords() As String) As Boolean()
    On Error GoTo ErrorHandler

    Dim bold() As Boolean
    Dim n As Long, i As Long, depth As Long, ch As String, runStart As Long
    n = Len(s)
    If n < 1 Then n = 1
    ReDim bold(1 To n)                           ' defaults all False
    depth = 0
    runStart = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "[" Then
            If runStart > 0 Then MarkTokenBold s, runStart, i - 1, boldKeywords, bold
            runStart = 0
            depth = depth + 1
        ElseIf ch = "]" Then
            depth = depth - 1
        ElseIf depth > 0 Then
            ' inside [...] - not a depth-0 keyword token
        ElseIf IsLetter(ch) Then
            If runStart = 0 Then runStart = i
        Else
            If runStart > 0 Then MarkTokenBold s, runStart, i - 1, boldKeywords, bold
            runStart = 0
        End If
    Next i
    If runStart > 0 Then MarkTokenBold s, runStart, Len(s), boldKeywords, bold
    BuildBoldMap = bold
    Exit Function

ErrorHandler:
    ' Fail-safe: an all-regular map of the right size.
    ReDim bold(1 To IIf(Len(s) < 1, 1, Len(s)))
    BuildBoldMap = bold
End Function

' If the token s[startPos..endPos] (case-insensitively) matches a member of boldKeywords, mark that range bold.
Private Sub MarkTokenBold(ByVal s As String, ByVal startPos As Long, ByVal endPos As Long, ByRef boldKeywords() As String, ByRef bold() As Boolean)
    On Error GoTo ErrorHandler

    If startPos < 1 Then Exit Sub
    If endPos < startPos Then Exit Sub

    Dim tok As String
    tok = Mid(s, startPos, endPos - startPos + 1)

    Dim kw As Long
    Dim bMatch As Boolean
    bMatch = False
    For kw = LBound(boldKeywords) To UBound(boldKeywords)
        If Len(boldKeywords(kw)) > 0 Then
            If StrComp(tok, boldKeywords(kw), vbTextCompare) = 0 Then
                bMatch = True
                Exit For
            End If
        End If
    Next kw

    If bMatch Then
        Dim k As Long
        For k = startPos To endPos
            If k >= LBound(bold) Then
                If k <= UBound(bold) Then bold(k) = True
            End If
        Next k
    End If
    Exit Sub

ErrorHandler:
End Sub

' True when ch is an ASCII letter A-Z / a-z (nested Ifs, no And; module is Option Compare Binary).
Private Function IsLetter(ByVal ch As String) As Boolean
    IsLetter = False
    If Len(ch) = 0 Then Exit Function
    Dim u As String
    u = UCase(ch)
    If u >= "A" Then
        If u <= "Z" Then IsLetter = True
    End If
End Function

' Create one Label for a coloured (optionally bold) run, size it EXPLICITLY (deterministic fixed-pitch
' cell, no AutoSize), and return the x for the NEXT run advanced by the run's EXACT text width so runs sit
' seamlessly (the generous right padding overlaps the next run but is invisible - transparent BackStyle).
Private Function EmitRun(ByVal previewFrame As MSForms.Frame, ByVal sRun As String, ByVal lColour As Long, ByVal bBold As Boolean, ByVal bUnderline As Boolean, ByVal dCharW As Single, ByVal x As Single) As Single
    On Error GoTo ErrorHandler

    EmitRun = x
    If Len(sRun) = 0 Then Exit Function

    Dim oLbl As MSForms.Label
    Set oLbl = AddPreviewLabel(previewFrame)
    If oLbl Is Nothing Then Exit Function

    ' Order: font (name/size/bold/underline) FIRST, then caption, then EXPLICIT width/height (AutoSize off,
    ' no wrap) so the label never depends on a per-label AutoSize recalculation. dCharW is the RUNTIME-measured
    ' advance (MeasureCharAdvance) - Width = Len*dCharW + INSET is generous so the text never clips; x below
    ' advances by the EXACT Len*dCharW so successive runs join seamlessly with no cumulative pixel drift.
    ' Consolas keeps the same advance in bold/underlined, so a regular-weight calibration suffices.
    oLbl.Visible = False
    ' A fresh per-label StdFont: MSForms controls added via Controls.Add SHARE their container's Font object,
    ' so mutating oLbl.Font.* would contaminate every other (and future) label - a valid rule would then
    ' inherit the bold/underline left by a previous invalid render. Assigning a new StdFont isolates this
    ' label's font. (No oLbl.Font.* mutation anywhere after this.)
    Dim f As stdole.StdFont
    Set f = New stdole.StdFont
    f.Name = PREVIEW_FONT
    f.Size = PREVIEW_SIZE
    f.bold = bBold
    f.Underline = bUnderline
    Set oLbl.Font = f
    oLbl.AutoSize = False
    oLbl.WordWrap = False
    oLbl.Caption = sRun
    oLbl.ForeColor = lColour
    oLbl.Width = Len(sRun) * dCharW + PREVIEW_INSET
    oLbl.Height = PREVIEW_CHARH
    oLbl.Left = x
    oLbl.Top = PREVIEW_Y0
    oLbl.Visible = True
    EmitRun = x + Len(sRun) * dCharW
    Exit Function

ErrorHandler:
    ' Silent fail-safe (per-run): a failed run just does not advance x. RenderPreview is the single logger
    ' of the render path; per-run logging would spam if Controls.Add is systemically unavailable.
    EmitRun = x
End Function

' Measure the REAL rendered character advance once per render: a hidden calibration label (never shown)
' gets a fresh StdFont (regular weight - Consolas' bold advance is identical) and a known 64-char etalon,
' then AutoSize = True is set AFTER the caption (toggled off->on to force the recompute - the inverse of the
' round-1 order bug). dCharW = .Width / N absorbs the fixed label margin over N=64 (negligible). SANITY
' CHECK: an implausible result (AutoSize did not recompute) falls back to the theoretical PREVIEW_CHARW
' constant (no worse than before). This removes the GDI pixel/DPI drift a theoretical advance caused.
Private Function MeasureCharAdvance(ByVal previewFrame As MSForms.Frame) As Single
    On Error GoTo ErrorHandler

    MeasureCharAdvance = PREVIEW_CHARW            ' fallback = theoretical constant

    Const CAL_N As Long = 64
    Dim oCal As MSForms.Label
    Set oCal = AddPreviewLabel(previewFrame)      ' hidden, cleared next render - no leak
    If oCal Is Nothing Then Exit Function

    Dim f As stdole.StdFont
    Set f = New stdole.StdFont
    f.Name = PREVIEW_FONT
    f.Size = PREVIEW_SIZE
    f.bold = False
    f.Underline = False
    Set oCal.Font = f

    oCal.WordWrap = False
    oCal.Caption = String(CAL_N, "M")            ' known N-char etalon
    oCal.AutoSize = False
    oCal.AutoSize = True                          ' AFTER the caption (round-1 order inverted) -> forces the recompute

    If oCal.Width > 0 Then
        Dim dMeasured As Single
        dMeasured = oCal.Width / CAL_N
        ' Plausible fixed-pitch advance? else AutoSize did not recompute -> keep the fallback constant.
        If dMeasured >= PREVIEW_SIZE * 0.3 Then
            If dMeasured <= PREVIEW_SIZE * 1.2 Then
                MeasureCharAdvance = dMeasured
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    MeasureCharAdvance = PREVIEW_CHARW
End Function

' Create a runtime Label inside the preview Frame with a unique PREVIEW_PREFIX name, return it (Nothing on
' fault). Created hidden; shown by EmitRun once sized. Cleared by ClearPreview (by name-prefix) next render.
Private Function AddPreviewLabel(ByVal previewFrame As MSForms.Frame) As MSForms.Label
    On Error GoTo ErrorHandler

    mlPreviewSeq = mlPreviewSeq + 1
    Dim sName As String
    sName = PREVIEW_PREFIX & mlPreviewSeq

    Dim oLbl As MSForms.Label
    Set oLbl = previewFrame.Controls.Add("Forms.Label.1", sName, False)   ' created hidden; shown once sized
    oLbl.BackStyle = fmBackStyleTransparent
    Set AddPreviewLabel = oLbl
    Exit Function

ErrorHandler:
    ' Silent fail-safe: a failed Controls.Add just yields no Label (the run is skipped). See EmitRun.
    Set AddPreviewLabel = Nothing
End Function

' Remove exactly the runtime preview Labels (PREVIEW_PREFIX names) from the passed Frame - stateless and
' per-Frame (bounded control count, no leak). Collect names first, then remove (never modify the collection
' mid-iteration). Silent (On Error Resume Next): a stale name is skipped; cleanup faults must never surface.
Private Sub ClearPreview(ByVal previewFrame As MSForms.Frame)
    On Error Resume Next
    Dim names() As String
    Dim nn As Long
    Dim oCtrl As MSForms.control
    ReDim names(0 To 0)
    nn = 0
    For Each oCtrl In previewFrame.Controls
        If Left(oCtrl.Name, Len(PREVIEW_PREFIX)) = PREVIEW_PREFIX Then
            If nn > UBound(names) Then ReDim Preserve names(0 To nn)
            names(nn) = oCtrl.Name
            nn = nn + 1
        End If
    Next oCtrl
    Dim i As Long
    For i = 0 To nn - 1
        previewFrame.Controls.Remove names(i)
    Next i
    On Error GoTo 0
End Sub
