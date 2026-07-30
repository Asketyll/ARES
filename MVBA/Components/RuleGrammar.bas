' Module: RuleGrammar
' Description: Shared rule CONDITION grammar (Lvl/Cell/Type conditions, & AND, ! negation, */? wildcards via
'              Like, contradiction reasoning, bracket-depth-aware split). Used by PropertyTagging (tag rules)
'              and PropertyCalculation (calc rules, epic 14). Pure/stateless - no config, no cache, no model write.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, MicroStationDefinition, ErrorHandlerClass

Option Explicit

Private Const NEG_MARKER As String = "!"
Private Const BRK_OPEN As String = "["
Private Const BRK_CLOSE As String = "]"
Private Const RULE_SEPARATOR As String = ";"
' The name separator inside [...] is ARESConstants.ARES_VAR_DELIMITER ("|").

' Rule keyword vocabulary (canonicalised, case-insensitive on input).
Public Enum RuleKeyword
    rkLvl
    rkCell
    rkType
End Enum

' One parsed condition: [!] Keyword[name|name|...]. Names are kept VERBATIM (trimmed) for Like matching.
' For rkType, Names resolve to Types() (MsdElementType values) and/or MatchesAnyCell (the special "Cell"
' token = any cell, since there is no single MsdElementType for a cell and StringToMsdElementType("Cell")
' does not resolve to one).
Public Type RuleCondition
    Keyword As RuleKeyword
    Negated As Boolean
    Names() As String
    Types() As Long
    MatchesAnyCell As Boolean
End Type

'######################################################################################################################
'                                          CONDITION PARSE
'######################################################################################################################

' Parse ONE condition segment "[!] Keyword[name|name|...]" into c. Returns "" on success or a reason.
Public Function ParseCondition(ByVal segInput As String, ByRef c As RuleCondition) As String
    On Error GoTo ErrorHandler

    ParseCondition = ""

    ' Reset so a previous condition cannot leak in.
    c.Negated = False
    c.MatchesAnyCell = False
    Erase c.Names
    Erase c.Types

    Dim seg As String
    seg = Trim(segInput)
    If Len(seg) = 0 Then
        ParseCondition = "empty condition"
        Exit Function
    End If

    If Left(seg, 1) = NEG_MARKER Then
        c.Negated = True
        seg = Trim(Mid(seg, 2))
        If Len(seg) = 0 Then
            ParseCondition = "empty condition after '!'"
            Exit Function
        End If
    End If

    Dim nOpen As Long, nClose As Long
    nOpen = InStr(seg, BRK_OPEN)
    If nOpen = 0 Then
        ParseCondition = "condition '" & seg & "' has no keyword (expected Lvl[..]/Cell[..]/Type[..])"
        Exit Function
    End If
    nClose = InStr(seg, BRK_CLOSE)
    If nClose <= nOpen Then
        ParseCondition = "malformed [...] in '" & seg & "'"
        Exit Function
    End If
    If nClose <> Len(seg) Then
        ParseCondition = "unexpected text after ']' in '" & seg & "'"
        Exit Function
    End If

    Dim sKw As String, body As String
    sKw = Trim(Left(seg, nOpen - 1))
    body = Mid(seg, nOpen + 1, nClose - nOpen - 1)

    Select Case UCase(sKw)
        Case "LVL"
            c.Keyword = rkLvl
        Case "CELL"
            c.Keyword = rkCell
        Case "TYPE"
            c.Keyword = rkType
        Case Else
            If Len(sKw) = 0 Then
                ParseCondition = "condition has no keyword (expected Lvl/Cell/Type)"
            Else
                ParseCondition = "unknown keyword '" & sKw & "' (expected Lvl/Cell/Type)"
            End If
            Exit Function
    End Select

    ' Forbidden characters inside [...] (";" and "[" - "]" cannot be here since nClose is the first "]").
    If InStr(body, RULE_SEPARATOR) > 0 Then
        ParseCondition = "';' not allowed inside [...]"
        Exit Function
    End If
    If InStr(body, BRK_OPEN) > 0 Then
        ParseCondition = "'[' not allowed inside [...]"
        Exit Function
    End If

    ' Split the body on "|" into trimmed, NON-EMPTY names (kept verbatim for Like).
    Dim vNames As Variant, nm As String, nCount As Long
    Dim namesOut() As String
    vNames = Split(body, ARESConstants.ARES_VAR_DELIMITER)
    ReDim namesOut(0 To UBound(vNames))
    nCount = 0
    Dim i As Long
    For i = LBound(vNames) To UBound(vNames)
        nm = Trim(vNames(i))
        If Len(nm) = 0 Then
            ParseCondition = "empty name in " & KeywordName(c.Keyword) & "[...]"
            Exit Function
        End If
        namesOut(nCount) = nm
        nCount = nCount + 1
    Next i
    ReDim Preserve namesOut(0 To nCount - 1)
    c.Names = namesOut

    ' Type resolution: each name resolves to an MsdElementType, or is the special "Cell" token (any cell).
    ' Wildcards are not meaningful for a type (an enum, not a name) -> rejected.
    If c.Keyword = rkType Then
        Dim typesOut() As Long
        Dim nt As Long
        ReDim typesOut(0 To nCount - 1)
        nt = 0
        For i = 0 To nCount - 1
            Dim bWild As Boolean
            bWild = False
            If InStr(c.Names(i), "*") > 0 Then bWild = True
            If InStr(c.Names(i), "?") > 0 Then bWild = True
            If bWild Then
                ParseCondition = "wildcards not allowed in Type[...]"
                Exit Function
            End If
            If UCase(c.Names(i)) = "CELL" Then
                c.MatchesAnyCell = True
            Else
                Dim t As Long
                t = MicroStationDefinition.StringToMsdElementType(c.Names(i))
                If t = 0 Then
                    ParseCondition = "unknown element type '" & c.Names(i) & "'"
                    Exit Function
                End If
                typesOut(nt) = t
                nt = nt + 1
            End If
        Next i
        If nt > 0 Then
            ReDim Preserve typesOut(0 To nt - 1)
            c.Types = typesOut
        End If
    End If
    Exit Function

ErrorHandler:
    ParseCondition = "invalid condition"
End Function

'######################################################################################################################
'                                          MATCHER
'######################################################################################################################

' Evaluate ONE condition with strict negation. The positive result is computed with per-keyword guards
' (never And-chained across a possibly-raising read); a negated condition returns Not(positive). On a
' level-less element a positive Lvl is False (a negated !Lvl is True); on a non-cell a positive Cell is
' False (!Cell is True) - so "Type[Cell]&!Cell[A]" means exactly "a cell, but not the one named A".
Public Function ConditionMatches(ByRef c As RuleCondition, ByVal oElement As element, ByVal sLevel As String, ByVal bHasLevel As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim bPos As Boolean
    bPos = False

    Select Case c.Keyword
        Case rkLvl
            If bHasLevel Then
                bPos = LikeAnyCI(sLevel, c.Names)
            End If
        Case rkCell
            If oElement.IsCellElement Then
                bPos = LikeAnyCI(oElement.AsCellElement.Name, c.Names)
            End If
        Case rkType
            If c.MatchesAnyCell Then
                If oElement.IsCellElement Then bPos = True
            End If
            If Not bPos Then
                If HasLongs(c.Types) Then
                    Dim ti As Long
                    For ti = LBound(c.Types) To UBound(c.Types)
                        If oElement.Type = c.Types(ti) Then
                            bPos = True
                            Exit For
                        End If
                    Next ti
                End If
            End If
    End Select

    If c.Negated Then
        ConditionMatches = Not bPos
    Else
        ConditionMatches = bPos
    End If
    Exit Function

ErrorHandler:
    ' Fail-closed: an unexpected fault counts as no match (never an errant attach).
    ConditionMatches = False
End Function

' Case-insensitive Like match of value against any of names. VBA Like metacharacters that could appear
' literally in a name are neutralised: only "#" can occur ("[" / "]" are forbidden inside a name), so
' escape "#" -> "[#]"; "*"/"?" stay wildcards. Case-insensitivity via UCase on both sides (the module has
' no Option Compare Text). Nested guards, no short-circuit.
Private Function LikeAnyCI(ByVal value As String, ByRef names() As String) As Boolean
    On Error GoTo ErrorHandler

    LikeAnyCI = False
    Dim uv As String
    uv = UCase(value)

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            If uv Like UCase(EscapeLikePattern(names(i))) Then
                LikeAnyCI = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    LikeAnyCI = False
End Function

' Escape the only Like metacharacter that can appear literally in a name: "#" -> "[#]". "*"/"?" are kept
' as wildcards; "[" / "]" cannot occur in a name (grammar-forbidden), so nothing else needs escaping.
Private Function EscapeLikePattern(ByVal name As String) As String
    EscapeLikePattern = Replace(name, "#", "[#]")
End Function

' Case-insensitive Like match of a single value against a single pattern (wildcards */?, "#"-escaped) -
' the scalar counterpart of the internal LikeAnyCI, exposed for PropertyCalculation's CellText[pattern]
' cell-name matching (epic 14). Reuses EscapeLikePattern so the wildcard semantics stay identical. An
' empty pattern never matches. Nested guards, no short-circuit.
Public Function LikeCI(ByVal value As String, ByVal pattern As String) As Boolean
    On Error GoTo ErrorHandler

    LikeCI = False
    If Len(pattern) = 0 Then Exit Function
    If UCase(value) Like UCase(EscapeLikePattern(pattern)) Then LikeCI = True
    Exit Function

ErrorHandler:
    LikeCI = False
End Function

'######################################################################################################################
'                                          CANONICALISATION
'######################################################################################################################

' Canonical text of one condition: [!] Keyword[name|name|...].
Public Function ConditionToCanonical(ByRef c As RuleCondition) As String
    Dim s As String
    s = ""
    If c.Negated Then s = NEG_MARKER
    s = s & KeywordName(c.Keyword)
    s = s & BRK_OPEN & Join(c.Names, ARESConstants.ARES_VAR_DELIMITER) & BRK_CLOSE
    ConditionToCanonical = s
End Function

' Canonical keyword casing.
Public Function KeywordName(ByVal kw As RuleKeyword) As String
    Select Case kw
        Case rkLvl
            KeywordName = "Lvl"
        Case rkCell
            KeywordName = "Cell"
        Case rkType
            KeywordName = "Type"
        Case Else
            KeywordName = ""
    End Select
End Function

'######################################################################################################################
'                                          CONTRADICTION DETECTOR
'######################################################################################################################

' Rule-agnostic contradiction over a conditions array: True (with the two conflicting condition segments,
' canonical text) when the conditions can never be satisfied together. Callable by both the tag
' RuleHasNoEffect and the coming calc detector. Needs >= 2 conditions.
Public Function ConditionsHaveContradiction(ByRef conds() As RuleCondition, ByVal nCond As Long, ByRef segments() As String) As Boolean
    On Error GoTo ErrorHandler

    ConditionsHaveContradiction = False
    ReDim segments(0 To 0)
    segments(0) = ""
    If nCond < 2 Then Exit Function

    Dim i As Long, j As Long
    For i = 0 To nCond - 2
        For j = i + 1 To nCond - 1
            If PairContradicts(conds(i), conds(j)) Then
                ReDim segments(0 To 1)
                segments(0) = ConditionToCanonical(conds(i))
                segments(1) = ConditionToCanonical(conds(j))
                ConditionsHaveContradiction = True
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    ConditionsHaveContradiction = False
    ReDim segments(0 To 0)
    segments(0) = ""
End Function

' True when two conditions can never be satisfied together (see ConditionsHaveContradiction for the cases).
Private Function PairContradicts(ByRef a As RuleCondition, ByRef b As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    PairContradicts = False

    ' (a) Same keyword, both positive, disjoint (no wildcard for Lvl/Cell; Type has none).
    If Not a.Negated Then
        If Not b.Negated Then
            If a.Keyword = b.Keyword Then
                Select Case a.Keyword
                    Case rkLvl, rkCell
                        If Not HasWildcard(a) Then
                            If Not HasWildcard(b) Then
                                If NamesDisjoint(a.Names, b.Names) Then
                                    PairContradicts = True
                                    Exit Function
                                End If
                            End If
                        End If
                    Case rkType
                        If TypeCondsDisjoint(a, b) Then
                            PairContradicts = True
                            Exit Function
                        End If
                End Select
            End If
        End If
    End If

    ' (b) Cell[...] (requires a cell) vs a Type condition that forbids cells - either order.
    If CellTypeContradict(a, b) Then
        PairContradicts = True
        Exit Function
    End If
    If CellTypeContradict(b, a) Then
        PairContradicts = True
        Exit Function
    End If
    Exit Function

ErrorHandler:
    PairContradicts = False
End Function

' c must be a positive Cell[...]; t a Type[...] that forbids cells: a positive Type whose resolved set
' contains no cell, or a negated Type[...] that covers all cells (!Type[Cell]). This contradiction is
' STRUCTURAL - "is the element a cell?" does not depend on the cell NAME - so there is NO wildcard guard
' here (unlike the same-keyword disjoint-list check, where a wildcard makes disjointness undecidable).
Private Function CellTypeContradict(ByRef c As RuleCondition, ByRef t As RuleCondition) As Boolean
    CellTypeContradict = False

    If c.Keyword <> rkCell Then Exit Function
    If c.Negated Then Exit Function
    If t.Keyword <> rkType Then Exit Function

    If Not t.Negated Then
        ' A positive Type with no cell in its match-set forbids the cell that Cell[...] requires.
        If Not t.MatchesAnyCell Then
            If Not HasCellType(t) Then
                CellTypeContradict = True
            End If
        End If
    Else
        ' A negated Type covering all cells (!Type[Cell]) forbids the cell Cell[...] requires.
        If t.MatchesAnyCell Then
            CellTypeContradict = True
        End If
    End If
End Function

' True when a condition carries a "*" or "?" in any of its names.
Private Function HasWildcard(ByRef c As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    HasWildcard = False
    Dim i As Long
    For i = LBound(c.Names) To UBound(c.Names)
        If InStr(c.Names(i), "*") > 0 Then
            HasWildcard = True
            Exit Function
        End If
        If InStr(c.Names(i), "?") > 0 Then
            HasWildcard = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasWildcard = True                           ' fail-safe: treat as "wildcard present" -> no verdict
End Function

' True when name-lists a and b share no name (case-insensitive, trimmed) -> disjoint.
Private Function NamesDisjoint(ByRef a() As String, ByRef b() As String) As Boolean
    On Error GoTo ErrorHandler

    NamesDisjoint = True
    Dim i As Long, j As Long
    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If StrComp(Trim(a(i)), Trim(b(j)), vbTextCompare) = 0 Then
                NamesDisjoint = False
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    NamesDisjoint = False                        ' fail-safe: assume they overlap -> no verdict
End Function

' True when two positive Type conditions can never match the same element (their match-sets are disjoint,
' accounting for the "any cell" token).
Private Function TypeCondsDisjoint(ByRef a As RuleCondition, ByRef b As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    TypeCondsDisjoint = False

    If a.MatchesAnyCell Then
        If b.MatchesAnyCell Then Exit Function       ' both any-cell -> overlap
        If HasCellType(b) Then Exit Function         ' a any-cell, b lists a cell type -> overlap
    End If
    If b.MatchesAnyCell Then
        If HasCellType(a) Then Exit Function         ' b any-cell, a lists a cell type -> overlap
    End If
    If TypesIntersect(a.Types, b.Types) Then Exit Function

    TypeCondsDisjoint = True
    Exit Function

ErrorHandler:
    TypeCondsDisjoint = False
End Function

' True when a resolved type list contains a cell type (CellHeader or SharedCell).
Private Function HasCellType(ByRef c As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    HasCellType = False
    If Not HasLongs(c.Types) Then Exit Function
    Dim i As Long
    For i = LBound(c.Types) To UBound(c.Types)
        If c.Types(i) = msdElementTypeCellHeader Then
            HasCellType = True
            Exit Function
        End If
        If c.Types(i) = msdElementTypeSharedCell Then
            HasCellType = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasCellType = False
End Function

' True when two resolved type lists share a value.
Private Function TypesIntersect(ByRef a() As Long, ByRef b() As Long) As Boolean
    On Error GoTo ErrorHandler

    TypesIntersect = False
    If Not HasLongs(a) Then Exit Function
    If Not HasLongs(b) Then Exit Function
    Dim i As Long, j As Long
    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If a(i) = b(j) Then
                TypesIntersect = True
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    TypesIntersect = False
End Function

'######################################################################################################################
'                                          BRACKET-DEPTH-AWARE LOW-LEVEL HELPERS
'######################################################################################################################

' Position (1-based) of the first occurrence of the single character ch at bracket depth 0, or 0 if none.
Public Function FindTopLevelChar(ByVal s As String, ByVal ch As String) As Long
    Dim depth As Long, i As Long, c As String
    depth = 0
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = BRK_OPEN Then
            depth = depth + 1
        ElseIf c = BRK_CLOSE Then
            depth = depth - 1
        ElseIf c = ch Then
            If depth = 0 Then
                FindTopLevelChar = i
                Exit Function
            End If
        End If
    Next i
    FindTopLevelChar = 0
End Function

' Split s on the single character ch at bracket depth 0 (a ch inside [...] is literal). Returns a 0-based
' array of the raw (untrimmed) segments, including empties.
Public Function SplitTopLevel(ByVal s As String, ByVal ch As String) As String()
    Dim out() As String
    Dim n As Long, depth As Long, i As Long, c As String, seg As String
    ReDim out(0 To 0)
    n = 0
    depth = 0
    seg = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = BRK_OPEN Then
            depth = depth + 1
            seg = seg & c
        ElseIf c = BRK_CLOSE Then
            depth = depth - 1
            seg = seg & c
        ElseIf c = ch Then
            If depth = 0 Then
                If n > UBound(out) Then ReDim Preserve out(0 To n)
                out(n) = seg
                n = n + 1
                seg = ""
            Else
                seg = seg & c
            End If
        Else
            seg = seg & c
        End If
    Next i
    If n > UBound(out) Then ReDim Preserve out(0 To n)
    out(n) = seg
    SplitTopLevel = out
End Function

' Split a string and trim each entry, dropping empties. Returns a 0-based array (a single "" when none).
Public Function SplitTrim(ByVal s As String, ByVal Delim As String) As String()
    Dim vParts As Variant, i As Long, n As Long
    Dim out() As String

    vParts = Split(s, Delim)
    ReDim out(0 To UBound(vParts))
    n = 0
    For i = LBound(vParts) To UBound(vParts)
        If Len(Trim(vParts(i))) > 0 Then
            out(n) = Trim(vParts(i))
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim out(0 To 0)
        out(0) = ""
    Else
        ReDim Preserve out(0 To n - 1)
    End If
    SplitTrim = out
End Function

' Safe "Long array is allocated and non-empty" check.
Private Function HasLongs(ByRef arr() As Long) As Boolean
    On Error Resume Next
    HasLongs = False
    If UBound(arr) <> -1 Then HasLongs = True
    On Error GoTo 0
End Function
