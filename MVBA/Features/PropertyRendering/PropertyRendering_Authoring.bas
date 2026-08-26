' Module: PropertyRendering_Authoring
' Description: Bind / first-author / top-up - the authoring rule shared by the first-author scan and the
'              top-up scan on an already-bound bearer. Full mechanism (hybrid policy, source preservation):
'              see _bmad/docs/property-rendering-mechanics.md "Authoring".
' License: This project is licensed under the AGPL-3.0.
' Dependencies: CustomPropertyHandler, PropertyTagging, PropertyCalculation, StringsInEl, LangManager,
'               ErrorHandlerClass (global ErrorHandler), PropertyRendering (Core),
'               PropertyRendering_Types, PropertyRendering_TemplateModel, PropertyRendering_Serialization,
'               PropertyRendering_Reporting

Option Explicit

' Author ONE bearer, if it qualifies, appending an entry to the list. THE unit of binding is the SUB-TEXT,
' not the element. Both the first-author scan and the top-up scan run this one function, so the rules
' cannot drift. Full rationale (hybrid policy §6.2, source-preservation via bKeepSource/nFree): see
' "Authoring (TryAuthorBearer / AuthorUnclaimedBearers / TryFirstAuthor)" in property-rendering-mechanics.md.
Private Function TryAuthorBearer(ByRef oEl As element, ByVal sText As String, ByVal SubId As Long, ByRef ents() As RenderEntry, ByRef nEnts As Long, ByVal bKeepSource As Boolean, ByRef nFree As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim names() As String
    Dim values() As String
    Dim n As Long
    Dim j As Long
    Dim bAllAttached As Boolean

    TryAuthorBearer = False

    If Not PropertyRendering_TemplateModel.ParseTemplate(sText, lits, toks, nTok, True) Then Exit Function
    If nTok = 0 Then Exit Function

    ' SOURCE PRESERVATION, after the token test on purpose - see property-rendering-mechanics.md.
    If bKeepSource Then
        If nFree <= 1 Then
            PropertyRendering_Reporting.ReportCycleWarning
            Exit Function
        End If
    End If

    If Not PropertyRendering_TemplateModel.TemplateIsWellFormed(sText, True) Then
        ' Duplicate or adjacent tokens: refused at author time, no metadata written.
        PropertyRendering_Reporting.ReportStructuralRefusal sText
        Exit Function
    End If
    If Not CanBearTokens(oEl, sText) Then
        PropertyRendering_Reporting.ReportTextNodeRefused
        Exit Function
    End If

    bAllAttached = True
    n = 0
    For j = 0 To nTok - 1
        If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, toks(j)) Then
            bAllAttached = False
        Else
            AppendValue toks(j), "", names, values, n
            WarnStaticCellTextCycle oEl, toks(j)
        End If
    Next j

    If Not bAllAttached Then
        PropertyRendering_Reporting.ReportNotBound
        Exit Function
    End If

    ReDim Preserve ents(0 To nEnts)
    ents(nEnts).SubId = SubId
    ents(nEnts).Template = sText
    ents(nEnts).Dropped = False
    SetEntryValues ents, nEnts, names, values, n
    nEnts = nEnts + 1
    nFree = nFree - 1
    TryAuthorBearer = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Authoring.TryAuthorBearer"
    TryAuthorBearer = False
End Function

' Does this element FEED a Cell* calc source? If so it must keep at least ONE sub-text out of the exclusion
' set (TryAuthorBearer enforces it) - binding the last one would let the containment destroy the data it
' protects. Full rationale: see "Authoring" in property-rendering-mechanics.md.
Private Function FeedsCellSource(ByRef oEl As element) As Boolean
    On Error GoTo ErrorHandler

    FeedsCellSource = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    FeedsCellSource = PropertyCalculation.IsTriggerCell(oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: if we cannot tell, assume the cell feeds a source and protect it.
    FeedsCellSource = True
End Function

' TOP-UP AUTHORING on a bearer that is ALREADY bound: every sub-text no stored entry claims is offered to
' the same authoring rule. Returns how many entries were added. Without this a second sub-text whose
' property arrives later stays inert forever. Full rationale: see "Authoring" in
' property-rendering-mechanics.md.
Public Function AuthorUnclaimedBearers(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long, ByRef ents() As RenderEntry, ByRef nEnts As Long, ByRef subIds() As Long) As Long
    On Error GoTo ErrorHandler

    Dim claimed() As Boolean
    Dim i As Long
    Dim nBefore As Long
    Dim nFree As Long
    Dim bKeepSource As Boolean

    AuthorUnclaimedBearers = 0
    If nBearers <= 0 Then Exit Function
    If Not IsAcceptableBearerElement(oEl) Then Exit Function

    ' An UNRESOLVED live entry (subIds = -1) bails the whole scan this pass, rather than risk authoring a
    ' duplicate for the text it drives. See "Authoring" in property-rendering-mechanics.md.
    ReDim claimed(0 To nBearers - 1)
    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then
            If subIds(i) < 0 Then Exit Function
            If subIds(i) <= nBearers - 1 Then claimed(subIds(i)) = True
        End If
    Next i

    nFree = 0
    For i = 0 To nBearers - 1
        If Not claimed(i) Then nFree = nFree + 1
    Next i
    If nFree = 0 Then Exit Function

    ' Evaluated once, and only once there is something to author, so a fully-bound cell pays nothing.
    ' TryAuthorBearer owns the rule itself and maintains nFree.
    bKeepSource = FeedsCellSource(oEl)

    nBefore = nEnts
    For i = 0 To nBearers - 1
        If Not claimed(i) Then TryAuthorBearer oEl, texts(i), i, ents, nEnts, bKeepSource, nFree
    Next i

    AuthorUnclaimedBearers = nEnts - nBefore
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Authoring.AuthorUnclaimedBearers"
    AuthorUnclaimedBearers = 0
End Function

' Branch 4 - the first author, hybrid policy. bManual = the BindPropertyRender key-in (reports refusals);
' the automatic path uses identical rules. A bearer binds only when every token names a property ALREADY
' ATTACHED to the element - attachment is the intent signal, so nothing is ever bound by accident.
Public Function TryFirstAuthor(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long, ByVal bManual As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim ents() As RenderEntry
    Dim nEnts As Long
    Dim i As Long
    Dim nFree As Long
    Dim bKeepSource As Boolean
    Dim sExp As String
    Dim bOkExp As Boolean
    Dim curNames() As String
    Dim curValues() As String
    Dim nCur As Long
    Dim wrIds() As Long
    Dim wrPrev() As String
    Dim nWr As Long

    TryFirstAuthor = False
    nEnts = 0

    ' Element-level bearer guard, evaluated once: never author on anything but a TOP-LEVEL model element.
    If Not IsAcceptableBearerElement(oEl) Then Exit Function

    ' One decision per BEARER, via the shared authoring rule (includes SOURCE PRESERVATION - see
    ' "Authoring" in property-rendering-mechanics.md). Nothing is bound yet, so every bearer starts free.
    bKeepSource = FeedsCellSource(oEl)
    nFree = nBearers
    For i = 0 To nBearers - 1
        TryAuthorBearer oEl, texts(i), i, ents, nEnts, bKeepSource, nFree
    Next i

    If nEnts = 0 Then Exit Function

    ' Render every new entry BEFORE any metadata is touched (text first, then metadata).
    nWr = 0
    For i = 0 To nEnts - 1
        If ReadCurrentValues(oEl, ents(i), curNames, curValues, nCur) Then
            ' A faulted expansion returns the Template unchanged; binding on it would store LastValues
            ' that do not match what is actually visible, and the next pass would read a user edit that
            ' never happened. Drop the entry instead.
            sExp = PropertyRendering_TemplateModel.ExpandTemplate(ents(i).Template, curNames, curValues, nCur, True, bOkExp)
            If Not bOkExp Then
                ents(i).Dropped = True
            ElseIf WriteRenderedText(oEl, ents(i).SubId, sExp) Then
                ReDim Preserve wrIds(0 To nWr)
                ReDim Preserve wrPrev(0 To nWr)
                wrIds(nWr) = ents(i).SubId
                wrPrev(nWr) = texts(ents(i).SubId)
                nWr = nWr + 1
                SetEntryValues ents, i, curNames, curValues, nCur
            Else
                ents(i).Dropped = True
            End If
        Else
            ents(i).Dropped = True
        End If
    Next i

    If PropertyRendering_Types.CountLiveEntries(ents, nEnts) = 0 Then Exit Function

    If Not PropertyTagging.AttachRenderMetadata(oEl) Then
        RestoreWrittenTexts oEl, wrIds, wrPrev, nWr
        PropertyRendering_Reporting.ReportLibraryMissing
        Exit Function
    End If

    If Not PropertyRendering_Serialization.WriteRenderMetadata(oEl, ents, nEnts) Then
        ' The binding never landed: restore the text and undo our own attach, so no half-bound element
        ' is left behind and the next pass sees exactly the state it saw this time.
        RestoreWrittenTexts oEl, wrIds, wrPrev, nWr
        PropertyTagging.DetachRenderMetadata oEl
        PropertyRendering_Reporting.ReportMetadataUnreadable
        PropertyRendering.DisableAfterWriteFailure
        Exit Function
    End If

    If bManual Then LangManager.ShowStatusT "RenderBindDone"
    TryFirstAuthor = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Authoring.TryFirstAuthor"
    TryFirstAuthor = False
End Function

' Put back the texts written earlier in this pass, after the metadata write failed. Best effort by
' definition - a failure here leaves the visible ahead of the stored state, which the state machine then
' reads as a user edit rather than as a corrupt binding.
' oEl is ByRef so SetTextAtSubId's post-Rewrite handle refresh reaches the caller (see WriteRenderedText).
Private Sub RestoreWrittenTexts(ByRef oEl As element, ByRef ids() As Long, ByRef prev() As String, ByVal n As Long)
    On Error Resume Next
    Dim i As Long
    For i = 0 To n - 1
        StringsInEl.SetTextAtSubId oEl, ids(i), prev(i)
    Next i
End Sub

' The ONE place visible text is written. Refuses a reserved serialisation delimiter (and a stray CR) before
' it reaches the file. Returns True only when the sub-text now reads sNew. oEl is ByRef down the whole call
' chain on purpose - see "WriteRenderedText and the ByRef element chain" in property-rendering-mechanics.md.
' Public: called from Module D (StateMachine)'s RenderEntryOnElement (branch 2), not just internally.
Public Function WriteRenderedText(ByRef oEl As element, ByVal SubId As Long, ByVal sNew As String) As Boolean
    On Error GoTo ErrorHandler

    WriteRenderedText = False

    ' vbLf is legal here (a multi-line TextNode's rendering carries it); only the delimiters and a stray
    ' CR are refused. VALUES are held to the stricter rule in ReadCurrentValues.
    If PropertyRendering_Serialization.ContainsSerialisationDelimiter(sNew) Then
        PropertyRendering_Reporting.ReportValueIllegal
        Exit Function
    End If
    If InStr(1, sNew, vbCr) > 0 Then
        PropertyRendering_Reporting.ReportValueIllegal
        Exit Function
    End If

    WriteRenderedText = StringsInEl.SetTextAtSubId(oEl, SubId, sNew)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Authoring.WriteRenderedText"
    WriteRenderedText = False
End Function

' Read the CURRENT value of every token of an entry, off the bearing element's OWN attached properties.
' Values are copied VERBATIM (never CStr/Format/rounding). Full rationale: see "ReadCurrentValues" in
' property-rendering-mechanics.md.
' Public: called from Module D (StateMachine)'s RenderEntryOnElement, not just internally.
Public Function ReadCurrentValues(ByVal oEl As element, ByRef ent As RenderEntry, ByRef names() As String, ByRef values() As String, ByRef n As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim vVal As Variant
    Dim sVal As String

    ReadCurrentValues = False
    n = 0
    If Not PropertyRendering_TemplateModel.ParseTemplate(ent.Template, lits, toks, nTok, True) Then Exit Function

    For i = 0 To nTok - 1
        sVal = ""
        vVal = CustomPropertyHandler.GetPropertyValueFromElement(oEl, toks(i), toks(i))
        ' Nested tests, never an And chain: VBA evaluates both operands, and VarType/CStr on an array
        ' would raise.
        If IsNull(vVal) Then
            sVal = ""                                   ' unset -> the literal token, per Expand
        ElseIf IsArray(vVal) Then
            PropertyRendering_Reporting.ReportValueUnsupported
            sVal = ""
        ElseIf VarType(vVal) = vbString Then
            sVal = vVal                                 ' VERBATIM - no CStr, no Format, no rounding
        Else
            PropertyRendering_Reporting.ReportValueUnsupported
            sVal = ""
        End If

        ' A VALUE is held to the strict rule: no delimiter, no CR, no LF. Refusing here is what makes it
        ' impossible for a value to render the metadata unparseable.
        If Len(sVal) > 0 Then
            If PropertyRendering_Serialization.ValueHasIllegalChar(sVal) Then
                PropertyRendering_Reporting.ReportValueIllegal
                Exit Function
            End If
        End If

        AppendValue toks(i), sVal, names, values, n
    Next i

    ReadCurrentValues = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Authoring.ReadCurrentValues"
    ReadCurrentValues = False
End Function

' Defensive bearer guard (the Ouroboros exclusion): the bearer is always a TOP-LEVEL model element, never a
' cell component. Refuses SILENTLY on purpose. Full rationale: see "Bearer guards" in
' property-rendering-mechanics.md.
Private Function IsAcceptableBearerElement(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsAcceptableBearerElement = False
    If oEl Is Nothing Then Exit Function
    If oEl.IsComponentElement Then Exit Function
    IsAcceptableBearerElement = True
    Exit Function

ErrorHandler:
    IsAcceptableBearerElement = False
End Function

' v1 refuses to author a token inside a TEXTNODE belonging to a cell fed by an active group source: the
' exclusion granularity is one whole bearer, and a TextNode IS one bearer. Full rationale: see "Bearer
' guards" in property-rendering-mechanics.md.
Private Function CanBearTokens(ByVal oEl As element, ByVal sBearerText As String) As Boolean
    On Error GoTo ErrorHandler

    CanBearTokens = True
    If Not oEl.IsCellElement Then Exit Function
    If Not PropertyCalculation.IsTriggerCell(oEl) Then Exit Function

    ' A multi-line bearer is a TextNode: the whole-text form joins its lines with vbLf.
    If InStr(1, sBearerText, vbLf) > 0 Then CanBearTokens = False
    Exit Function

ErrorHandler:
    CanBearTokens = False
End Function

' Bind-time discoverability: warn about the one static cycle v1 can detect - a token rendered inside the
' very cell whose text feeds the property through a CellText rule. This is a structural trap in the
' binding choice itself, independent of any later edit, so it stays a bind-time check - unlike the general
' "this value is governed by a calc rule" notice, which moved to StateMachine's WarnGovernedTokensLost:
' at bind time every token here is freshly attached and unedited, so a "won't survive an edit" message
' would be misleading before any edit ever happened. See "WarnGovernedTokensLost" in
' property-rendering-mechanics.md, PropertyRendering_StateMachine.bas.
Private Sub WarnStaticCellTextCycle(ByVal oEl As element, ByVal P As String)
    On Error Resume Next

    Dim kind As CalcSource
    Dim sArg As String
    Dim sCanonical As String

    If Not PropertyCalculation.GetCalcRuleForProperty(P, oEl, kind, sArg, sCanonical) Then Exit Sub

    If kind = csCellText Then
        If oEl.IsCellElement Then PropertyRendering_Reporting.ReportCycleWarning
    End If
End Sub
