' Module: PropertyRendering_TemplateModel
' Description: The template model - ExpandTemplate/AlignVisible/AlignByValues/TemplateIsWellFormed (the
'              Public test seams external code still qualifies as PropertyRendering_TemplateModel.Xxx),
'              the tokenizer/canonicalizer (ParseTemplate), and the D6-D8 forge-protection guards they all
'              reach into directly. Full mechanism: see _bmad/docs/property-rendering-mechanics.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: StringsInEl, ErrorHandlerClass (global ErrorHandler), PropertyRendering (IsKnownProperty),
'               PropertyRendering_Types, PropertyRendering_Serialization (ValueHasIllegalChar),
'               PropertyRendering_Reporting

Option Explicit

' Token grammar - deliberately the exact form the calc grammar already uses on its left-hand side
' (Prop[name]=Source), so the user learns ONE spelling. The "]" terminator makes prefix collisions
' impossible, so no longest-match rule is needed; the tag/calc grammar already forbids "[", "]" and ";"
' inside a property name.
Private Const TOKEN_OPEN As String = "Prop["
Private Const TOKEN_CLOSE As String = "]"

' D6 - WHITELIST (not a blacklist) of symbols an addition may border a value with: a blacklist could
' silently miss a forging character (e.g. a locale thousands separator) - this list is closed by
' construction instead. Full admission criterion and named exceptions: see
' property-rendering-mechanics.md.
Private Const SAFE_BOUNDARY_SYMBOLS As String = "%()[]{}\*#~<>=:;!?&@_|"""

' D8 - characters that can appear INSIDE a number literal (any base), bounding NumericTailIsPossible's
' scan only - never granting safety by itself (too broad is harmless, too narrow just stops the scan
' early). Full rationale: see property-rendering-mechanics.md.
Private Const NUMBER_CAPABLE_CHARS As String = "0123456789abcdefABCDEFxXbBoO"

' Expand a Template (L0 T0 L1 T1 ... Ln) against a set of values; an EMPTY/ABSENT value renders the token's
' OWN LITERAL TEXT. bOk reports whether the expansion is TRUSTWORTHY (fails OPEN on a fault) - see
' "ExpandTemplate" in property-rendering-mechanics.md.
Public Function ExpandTemplate(ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, Optional ByVal bValidateNames As Boolean = True, Optional ByRef bOk As Boolean) As String
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim sOut As String
    Dim sVal As String

    bOk = False
    ExpandTemplate = sTemplate
    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    ' The parse held, so sTemplate IS its own expansion when it carries no token.
    bOk = True
    If nTok = 0 Then Exit Function

    sOut = ""
    For i = 0 To nTok - 1
        sOut = sOut & lits(i)
        sVal = LookupValue(toks(i), ValNames, ValValues, nVals)
        If Len(sVal) = 0 Then
            sOut = sOut & TokenLiteral(toks(i))
        Else
            sOut = sOut & sVal
        End If
    Next i
    sOut = sOut & lits(nTok)

    ExpandTemplate = sOut
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_TemplateModel.ExpandTemplate"
    ExpandTemplate = sTemplate
    bOk = False
End Function

' Align a user-edited visible string against the Template that produced the last rendering, deriving the
' new Template plus the surviving LastValues (the per-token release). Full mechanism (literal walk, span
' survival, why anchor-first): see "AlignVisible (per-token release, literal walk)" in
' property-rendering-mechanics.md.
Public Function AlignVisible(ByVal sVisible As String, ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, ByRef sNewTemplate As String, ByRef NewNames() As String, ByRef NewValues() As String, ByRef nNew As Long, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim cursor As Long
    Dim pos As Long
    Dim sSpan As String
    Dim sKeep As String
    Dim sKeepAfter As String
    Dim sLit As String
    Dim sAnchor As String
    Dim sEntry As String
    Dim bHasEntry As Boolean
    Dim bSurvived As Boolean
    Dim bSpanFound As Boolean
    Dim lCmp As Long
    Dim sOut As String

    AlignVisible = False
    sNewTemplate = ""
    nNew = 0

    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then Exit Function

    ' L0 must sit at the very start (empty when the Template opens with a token), matched case-insensitively
    ' like the tokeniser. A surviving span is re-materialised CANONICALLY (TokenLiteral), an edited one keeps
    ' the user's own casing verbatim - see "AlignVisible" in property-rendering-mechanics.md.
    If Len(lits(0)) > 0 Then
        If StrComp(Left(sVisible, Len(lits(0))), lits(0), vbTextCompare) <> 0 Then Exit Function
    End If
    cursor = Len(lits(0)) + 1
    sOut = Left(sVisible, Len(lits(0)))

    For i = 0 To nTok - 1
        bHasEntry = HasValueEntry(toks(i), ValNames, nVals)
        sEntry = LookupValue(toks(i), ValNames, ValValues, nVals)

        ' What the LAST rendering put in this span is KNOWN: the value itself, or - in the unset state -
        ' the token's own literal text.
        sAnchor = ""
        lCmp = vbBinaryCompare
        If bHasEntry Then
            If Len(sEntry) > 0 Then
                sAnchor = sEntry
            Else
                sAnchor = TokenLiteral(toks(i))
                lCmp = vbTextCompare
            End If
        End If

        ' ANCHOR FIRST, literal search only as fallback: locating the closing literal blindly breaks the
        ' moment a VALUE contains it (values are copied verbatim). See property-rendering-mechanics.md.
        bSpanFound = False
        If Len(sAnchor) > 0 Then
            If SpanAnchorsAt(sVisible, cursor, sAnchor, lits(i + 1), lCmp) Then
                sSpan = Mid(sVisible, cursor, Len(sAnchor))
                sLit = Mid(sVisible, cursor + Len(sAnchor), Len(lits(i + 1)))
                cursor = cursor + Len(sAnchor) + Len(lits(i + 1))
                bSpanFound = True
            End If
        End If

        If Not bSpanFound Then
            If Len(lits(i + 1)) = 0 Then
                ' Trailing empty literal: the span runs to end-of-string.
                sSpan = Mid(sVisible, cursor)
                sLit = ""
                cursor = Len(sVisible) + 1
            Else
                pos = InStr(cursor, sVisible, lits(i + 1), vbTextCompare)
                If pos = 0 Then Exit Function          ' literal lost -> conservative fallback
                sSpan = Mid(sVisible, cursor, pos - cursor)
                sLit = Mid(sVisible, pos, Len(lits(i + 1)))
                cursor = pos + Len(lits(i + 1))
            End If
        End If

        bSurvived = False
        If bHasEntry Then
            If Len(sEntry) > 0 Then
                ' A VALUE is compared BINARY: values are copied verbatim, so their casing is data.
                bSurvived = (sSpan = sEntry)
            Else
                ' Unset state: the literal token IS the last rendering, so finding it intact means the
                ' user left the token alone (a static-text edit must not release it). Case-insensitive,
                ' like the tokeniser: TokenLiteral re-materialises the canonical "Prop[" prefix, which is
                ' not necessarily the casing the user typed and the Template stored.
                bSurvived = (StrComp(sSpan, TokenLiteral(toks(i)), vbTextCompare) = 0)
            End If
        End If

        ' D7 - the span still CONTAINS the value intact, with the user's text beside it: may that addition be
        ' kept as static content while the token stays live? Each test is handed the addition concatenated
        ' with its context (the neighbouring literal), never the addition alone. Full rationale (the two
        ' forges a context-free test would miss): see "D7" in property-rendering-mechanics.md.
        sKeep = ""
        sKeepAfter = ""
        If Not bSurvived And bHasEntry Then
            If Len(sEntry) > 0 And Len(sSpan) > Len(sEntry) Then
                If Right(sSpan, Len(sEntry)) = sEntry Then
                    sKeep = Left(sSpan, Len(sSpan) - Len(sEntry))
                    If PrefixIsSafeAddition(lits(i) & sKeep, sEntry) Then
                        bSurvived = True
                    Else
                        sKeep = ""
                    End If
                ElseIf Left(sSpan, Len(sEntry)) = sEntry Then
                    sKeepAfter = Mid(sSpan, Len(sEntry) + 1)
                    If SuffixIsSafeAddition(sEntry, sKeepAfter & lits(i + 1)) Then
                        bSurvived = True
                    Else
                        sKeepAfter = ""
                    End If
                End If
            End If
        End If

        If bSurvived Then
            ' Both are empty on every path except an accepted addition, where they carry the user's own
            ' text back into the Template as static content - on the correct side of the token, which is
            ' what makes the re-authored Template render the next value in the place the user chose.
            sOut = sOut & sKeep & TokenLiteral(toks(i)) & sKeepAfter
            AppendValue toks(i), sEntry, NewNames, NewValues, nNew
        Else
            sOut = sOut & sSpan
        End If

        sOut = sOut & sLit
    Next i

    ' Anything the user appended past the final literal is kept as static text.
    If cursor <= Len(sVisible) Then sOut = sOut & Mid(sVisible, cursor)

    sNewTemplate = sOut
    AlignVisible = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_TemplateModel.AlignVisible"
    AlignVisible = False
    nNew = 0
End Function

' D8 - alignment by VALUE RECOGNITION, the fallback for everything the literal walk cannot follow. Runs
' ONLY after AlignVisible has declined. Public test seam (AC12). See "D8 - AlignByValues" in
' property-rendering-mechanics.md.
Public Function AlignByValues(ByVal sVisible As String, ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, ByRef sNewTemplate As String, ByRef NewNames() As String, ByRef NewValues() As String, ByRef nNew As Long, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim pos As Long
    Dim cursor As Long
    Dim sEntry As String
    Dim sPrevEntry As String
    Dim sGap As String
    Dim sOut As String

    AlignByValues = False
    sNewTemplate = ""
    nNew = 0

    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then Exit Function

    cursor = 1
    sOut = ""
    sPrevEntry = ""

    For i = 0 To nTok - 1
        If Not HasValueEntry(toks(i), ValNames, nVals) Then Exit Function
        sEntry = LookupValue(toks(i), ValNames, ValValues, nVals)
        If Len(sEntry) = 0 Then Exit Function

        ' A VALUE is matched BINARY, exactly as AlignVisible does: values are copied verbatim, so casing
        ' is data and a case-insensitive hit would align the wrong text.
        pos = InStr(cursor, sVisible, sEntry, vbBinaryCompare)
        If pos = 0 Then Exit Function
        If InStr(pos + 1, sVisible, sEntry, vbBinaryCompare) > 0 Then Exit Function

        sGap = Mid(sVisible, cursor, pos - cursor)

        ' The gap touches TWO values - the previous one on its left edge, this one on its right - and both
        ' edges must be judged, against the authored literal that used to sit there.
        If i > 0 Then
            If Not RightContextIsSafe(lits(i), sGap, sPrevEntry) Then Exit Function
        End If
        If Not LeftContextIsSafe(lits(i), sGap, sEntry) Then Exit Function

        sOut = sOut & sGap & TokenLiteral(toks(i))
        AppendValue toks(i), sEntry, NewNames, NewValues, nNew
        cursor = pos + Len(sEntry)
        sPrevEntry = sEntry
    Next i

    sGap = Mid(sVisible, cursor)
    If Not RightContextIsSafe(lits(nTok), sGap, sPrevEntry) Then Exit Function
    sOut = sOut & sGap

    If Not ReauthoredTemplateIsSound(sOut, nTok, NewNames, NewValues, nNew, sVisible, bValidateNames) Then Exit Function

    sNewTemplate = sOut
    AlignByValues = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_TemplateModel.AlignByValues"
    AlignByValues = False
    nNew = 0
End Function

' A Template is well formed when no property is tokenised twice (v1: ONE token per property per text) and
' no two tokens are adjacent - adjacency produces an interior EMPTY literal, leaving the span boundary
' undefined for the per-token release. Public test seam, like ExpandTemplate/AlignVisible.
Public Function TemplateIsWellFormed(ByVal sTemplate As String, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long, j As Long

    TemplateIsWellFormed = False
    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then
        TemplateIsWellFormed = True
        Exit Function
    End If

    For i = 0 To nTok - 1
        For j = i + 1 To nTok - 1
            If StrComp(toks(i), toks(j), vbTextCompare) = 0 Then Exit Function
        Next j
    Next i

    For i = 1 To nTok - 1
        If Len(lits(i)) = 0 Then Exit Function
    Next i

    TemplateIsWellFormed = True
    Exit Function

ErrorHandler:
    TemplateIsWellFormed = False
End Function

' Split a Template into its alternating literals and tokens: lits(0..nTok) and toks(0..nTok-1).
' bValidateNames = True (the runtime) accepts only tokens naming a property the DGNLib actually declares;
' anything else stays part of the surrounding literal, fail-closed. False gives the pure grammar, which
' is what the unit tests assert without MicroStation.
Public Function ParseTemplate(ByVal sTemplate As String, ByRef lits() As String, ByRef toks() As String, ByRef nTok As Long, ByVal bValidateNames As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim cursor As Long
    Dim posOpen As Long
    Dim posClose As Long
    Dim sName As String
    Dim sLit As String

    ParseTemplate = False
    nTok = 0
    ReDim lits(0 To 0)
    lits(0) = ""
    cursor = 1
    sLit = ""

    Do
        posOpen = InStr(cursor, sTemplate, TOKEN_OPEN, vbTextCompare)
        If posOpen = 0 Then Exit Do

        posClose = InStr(posOpen + Len(TOKEN_OPEN), sTemplate, TOKEN_CLOSE)
        If posClose = 0 Then Exit Do

        sName = Mid(sTemplate, posOpen + Len(TOKEN_OPEN), posClose - posOpen - Len(TOKEN_OPEN))

        If IsAcceptableTokenName(sName, bValidateNames) Then
            sLit = sLit & Mid(sTemplate, cursor, posOpen - cursor)
            lits(nTok) = sLit
            ReDim Preserve toks(0 To nTok)
            toks(nTok) = sName
            nTok = nTok + 1
            ReDim Preserve lits(0 To nTok)
            lits(nTok) = ""
            sLit = ""
            cursor = posClose + Len(TOKEN_CLOSE)
        Else
            ' Unknown / malformed name: the whole "Prop[...]" run stays literal text.
            sLit = sLit & Mid(sTemplate, cursor, posClose + Len(TOKEN_CLOSE) - cursor)
            cursor = posClose + Len(TOKEN_CLOSE)
            If bValidateNames Then PropertyRendering_Reporting.ReportTokenUnknown
        End If
    Loop

    lits(nTok) = sLit & Mid(sTemplate, cursor)
    ParseTemplate = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_TemplateModel.ParseTemplate"
    ParseTemplate = False
    nTok = 0
End Function

' A token name is acceptable when it is non-empty, carries no grammar metacharacter, and - at runtime -
' names a property the DGNLib really declares. A "Base:Member" name (the split-coordinate field syntax,
' Prop[Name:X]/Prop[Name:Y] - see plan-xy-split-coordinate-properties.md) validates the BASE against the
' DGNLib as before, and the MEMBER strictly against the fixed X/Y convention (no ItemType introspection
' here by design - the convention is fixed, not discovered) - regardless of bValidateNames, since this is a
' grammar-level constraint like the "[" / ";" checks above, not a DGNLib-existence check like
' IsKnownProperty is.
Private Function IsAcceptableTokenName(ByVal sName As String, ByVal bValidateNames As Boolean) As Boolean
    IsAcceptableTokenName = False
    If Len(Trim(sName)) = 0 Then Exit Function
    If InStr(1, sName, "[") > 0 Then Exit Function
    If InStr(1, sName, ";") > 0 Then Exit Function
    If PropertyRendering_Serialization.ValueHasIllegalChar(sName) Then Exit Function

    Dim sBase As String, sMember As String
    If Not SplitTokenMember(sName, sBase, sMember) Then Exit Function     ' malformed "A:B:C" / ":X" / "Name:"
    If Len(Trim(sBase)) = 0 Then Exit Function

    If Len(sMember) > 0 Then
        If StrComp(sMember, "X", vbTextCompare) <> 0 Then
            If StrComp(sMember, "Y", vbTextCompare) <> 0 Then Exit Function
        End If
    End If

    If Not bValidateNames Then
        IsAcceptableTokenName = True
        Exit Function
    End If
    IsAcceptableTokenName = PropertyRendering.IsKnownProperty(sBase)
End Function

' Splits a token name into its BASE property name and, if present, a ":Member" suffix - the split-
' coordinate field syntax, Prop[Name:X]/Prop[Name:Y]. sMember = "" when sTokenName carries no ":" (the
' ordinary, unsplit case - every token before this feature). Pure syntax splitting only, no DGNLib/ItemType
' validation here (see IsAcceptableTokenName for that) - False only on a MALFORMED colon shape (more than
' one ":", or nothing on one side of it), never on an unrecognised member name (that is
' IsAcceptableTokenName's job). Public: also used by PropertyRendering_Authoring to route the three
' different name resolutions a token can need (base-only for the calc-rule/attachment lookups, base+member
' split for the value read) - see the plan's §5.3 for why these must not be conflated.
Public Function SplitTokenMember(ByVal sTokenName As String, ByRef sBase As String, ByRef sMember As String) As Boolean
    Dim nPos As Long

    SplitTokenMember = False
    sBase = sTokenName
    sMember = ""

    nPos = InStr(1, sTokenName, ":")
    If nPos = 0 Then
        SplitTokenMember = True
        Exit Function
    End If

    If InStr(nPos + 1, sTokenName, ":") > 0 Then Exit Function   ' more than one ":" - malformed
    If nPos = 1 Then Exit Function                                ' ":X" - nothing before the ":"
    If nPos = Len(sTokenName) Then Exit Function                  ' "Name:" - nothing after the ":"

    sBase = Left(sTokenName, nPos - 1)
    sMember = Mid(sTokenName, nPos + 1)
    SplitTokenMember = True
End Function

' The exact substring a token occupies in the text. This is the "unset" cue Expand writes, and the thing
' the conservative fallback keeps. Private: every caller (ExpandTemplate, AlignVisible, AlignByValues) is
' now inside this same module.
Private Function TokenLiteral(ByVal sName As String) As String
    TokenLiteral = TOKEN_OPEN & sName & TOKEN_CLOSE
End Function

' True when sAnchor - what the LAST rendering left in this span - still sits exactly at the cursor AND is
' immediately followed by the closing literal. D6 relaxes the trailing-empty-literal case in one direction
' (safe appended text). Full rationale: see "SpanAnchorsAt" in property-rendering-mechanics.md.
' Private: its only caller (AlignVisible) is now inside this same module.
Private Function SpanAnchorsAt(ByVal sVisible As String, ByVal cursor As Long, ByVal sAnchor As String, ByVal sNextLit As String, ByVal lAnchorCompare As Long) As Boolean
    SpanAnchorsAt = False
    If Len(sAnchor) = 0 Then Exit Function
    If cursor < 1 Then Exit Function
    If cursor + Len(sAnchor) - 1 > Len(sVisible) Then Exit Function
    If StrComp(Mid(sVisible, cursor, Len(sAnchor)), sAnchor, lAnchorCompare) <> 0 Then Exit Function

    If Len(sNextLit) = 0 Then
        If cursor + Len(sAnchor) = Len(sVisible) + 1 Then
            SpanAnchorsAt = True
        Else
            SpanAnchorsAt = SuffixIsSafeAddition(sAnchor, Mid(sVisible, cursor + Len(sAnchor)))
        End If
    Else
        SpanAnchorsAt = (StrComp(Mid(sVisible, cursor + Len(sAnchor), Len(sNextLit)), sNextLit, vbTextCompare) = 0)
    End If
End Function

' D6 - ASCII letter / ASCII digit, by CODE POINT (AscW, not Asc - locale-independent). An accented letter,
' a Unicode digit or a fullwidth digit is none of these, and is always the REFUSING side below.
Private Function IsAsciiLetter(ByVal sChar As String) As Boolean
    Dim n As Long
    IsAsciiLetter = False
    If Len(sChar) <> 1 Then Exit Function
    n = AscW(sChar)
    IsAsciiLetter = (n >= 65 And n <= 90) Or (n >= 97 And n <= 122)
End Function

Private Function IsAsciiDigit(ByVal sChar As String) As Boolean
    Dim n As Long
    IsAsciiDigit = False
    If Len(sChar) <> 1 Then Exit Function
    n = AscW(sChar)
    IsAsciiDigit = (n >= 48 And n <= 57)
End Function

' D6 - the whitelist test: True only for a character explicitly admitted (ASCII letter, or a
' SAFE_BOUNDARY_SYMBOLS member). Refusal is the DEFAULT - see property-rendering-mechanics.md.
Private Function IsSafeBoundaryChar(ByVal sChar As String) As Boolean
    On Error GoTo ErrorHandler

    IsSafeBoundaryChar = False
    If Len(sChar) <> 1 Then Exit Function

    If IsAsciiLetter(sChar) Then
        IsSafeBoundaryChar = True
    Else
        IsSafeBoundaryChar = (InStr(1, SAFE_BOUNDARY_SYMBOLS, sChar, vbBinaryCompare) > 0)
    End If
    Exit Function

ErrorHandler:
    IsSafeBoundaryChar = False
End Function

' D6 - may an addition AFTER the value be kept while the token keeps updating? Three load-bearing
' conditions (numeric anchor, whitelisted boundary char, exponent guard) that bound the CONSEQUENCE of a
' wrong call rather than guess intent. sSuffix is EVERYTHING to the right of the value, not just what the
' user typed. Full rationale: see "D6 - SuffixIsSafeAddition" in property-rendering-mechanics.md.
' Private: every caller (SpanAnchorsAt, AlignVisible, RightContextIsSafe) is now inside this same module.
Private Function SuffixIsSafeAddition(ByVal sAnchor As String, ByVal sSuffix As String) As Boolean
    On Error GoTo ErrorHandler
    Dim s As String
    Dim k As Long

    SuffixIsSafeAddition = False
    If Len(sSuffix) = 0 Then Exit Function
    If Not StringsInEl.IsNumericText(sAnchor) Then Exit Function
    If Len(sAnchor) = 0 Then Exit Function          ' IsNumericText("") is True - never let it through

    s = LTrim(sSuffix)
    If Len(s) = 0 Then Exit Function                ' whitespace only: fuses with the next value
    If Not IsSafeBoundaryChar(Left(s, 1)) Then Exit Function

    If IsAsciiLetter(Left(s, 1)) Then
        k = 2
        If Mid(s, k, 1) = "+" Or Mid(s, k, 1) = "-" Then k = k + 1
        If IsAsciiDigit(Mid(s, k, 1)) Then Exit Function
    End If

    SuffixIsSafeAddition = True
    Exit Function

ErrorHandler:
    SuffixIsSafeAddition = False
End Function

' D6 mirror - may an addition BEFORE the value be kept? Same three conditions, opposite end. sPrefix is
' EVERYTHING to the left of the value, not just what the user typed. Full rationale: see "D6 mirror -
' PrefixIsSafeAddition" in property-rendering-mechanics.md.
' Private: every caller (AlignVisible, LeftContextIsSafe) is now inside this same module.
Private Function PrefixIsSafeAddition(ByVal sPrefix As String, ByVal sAnchor As String) As Boolean
    On Error GoTo ErrorHandler
    Dim s As String

    PrefixIsSafeAddition = False
    If Len(sPrefix) = 0 Then Exit Function
    If Not StringsInEl.IsNumericText(sAnchor) Then Exit Function
    If Len(sAnchor) = 0 Then Exit Function

    s = RTrim(sPrefix)
    If Len(s) = 0 Then Exit Function
    If Not IsSafeBoundaryChar(Right(s, 1)) Then Exit Function

    If IsAsciiLetter(Right(s, 1)) And Len(s) >= 2 Then
        If IsAsciiDigit(Mid(s, Len(s) - 1, 1)) Then Exit Function
    End If

    PrefixIsSafeAddition = True
    Exit Function

ErrorHandler:
    PrefixIsSafeAddition = False
End Function

' D8 - the RAW frontier of a literal (first/last two characters, BEFORE any trimming): did the text
' touching the value change? Full rationale: see "D8" in property-rendering-mechanics.md.
Private Function RawHead(ByVal s As String) As String
    RawHead = Left(s, 2)
End Function

Private Function RawTail(ByVal s As String) As String
    RawTail = Right(s, 2)
End Function

' D8 - could the END of this literal be the beginning of a number? Answered on the trailing run of
' number-capable characters taken WHOLE, bounded by the data never a constant: a number starts with a
' digit. Closes the non-local base-literal forge D6's fixed-distance guards cannot see. Full rationale:
' see "D8" in property-rendering-mechanics.md.
Private Function NumericTailIsPossible(ByVal sLit As String) As Boolean
    On Error GoTo ErrorHandler
    Dim i As Long

    NumericTailIsPossible = False
    i = Len(sLit)
    Do While i > 0
        If InStr(1, NUMBER_CAPABLE_CHARS, Mid(sLit, i, 1), vbBinaryCompare) = 0 Then Exit Do
        i = i - 1
    Loop

    If i >= Len(sLit) Then Exit Function           ' no number-capable run at all
    NumericTailIsPossible = IsAsciiDigit(Mid(sLit, i + 1, 1))
    Exit Function

ErrorHandler:
    NumericTailIsPossible = False
End Function

' D8 - is the text now sitting to the LEFT of a value acceptable? Three ordered questions (new base
' literal possible? did the RAW frontier change at all? then D6's guard). Full rationale: see "D8" in
' property-rendering-mechanics.md. Private: its only caller (AlignByValues) is now inside this same module.
Private Function LeftContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    LeftContextIsSafe = True

    ' EMPTY current literal provably cannot weld (nothing to weld with) - handled HERE, never in the shared
    ' D6 predicates, where empty means the opposite ("no addition was made").
    If Len(sCurrent) = 0 Then Exit Function

    LeftContextIsSafe = False
    If NumericTailIsPossible(sCurrent) And Not NumericTailIsPossible(sAuthored) Then Exit Function

    LeftContextIsSafe = True
    If RawTail(sAuthored) = RawTail(sCurrent) Then Exit Function

    LeftContextIsSafe = PrefixIsSafeAddition(sCurrent, sValue)
End Function

' D8 mirror, for the text now sitting to the RIGHT of a value - no base-literal guard needed (that family
' can only form to the LEFT). See property-rendering-mechanics.md. Private: its only caller (AlignByValues)
' is now inside this same module.
Private Function RightContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    RightContextIsSafe = True

    ' Emptied literal (nothing to weld) handled here, same as the left half. An emptied gap BETWEEN two
    ' values is caught by ReauthoredTemplateIsSound's well-formedness check instead.
    If Len(sCurrent) = 0 Then Exit Function
    If RawHead(sAuthored) = RawHead(sCurrent) Then Exit Function

    RightContextIsSafe = SuffixIsSafeAddition(sValue, sCurrent)
End Function

' D8 - the ONE check that makes a re-authored Template safe to store: well-formed, carries exactly the
' expected tokens, and expanding it with the SAME LastValues reproduces the visible text byte-for-byte
' (the engine's fixed point, verified not inherited). Any failure falls back conservatively. Full
' rationale (the round-8/10 wedge case): see property-rendering-mechanics.md. Private: its only caller
' (AlignByValues) is now inside this same module.
Private Function ReauthoredTemplateIsSound(ByVal sNewTemplate As String, ByVal nExpectedTok As Long, ByRef names() As String, ByRef values() As String, ByVal n As Long, ByVal sVisible As String, ByVal bValidateNames As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim bOk As Boolean

    ReauthoredTemplateIsSound = False

    If Not TemplateIsWellFormed(sNewTemplate, bValidateNames) Then Exit Function
    If Not ParseTemplate(sNewTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok <> nExpectedTok Then Exit Function

    bOk = False
    If ExpandTemplate(sNewTemplate, names, values, n, bValidateNames, bOk) <> sVisible Then Exit Function
    If Not bOk Then Exit Function

    ReauthoredTemplateIsSound = True
    Exit Function

ErrorHandler:
    ReauthoredTemplateIsSound = False
End Function
