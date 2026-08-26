' Module: PropertyRendering_StateMachine
' Description: The render branches (1-3b) for an already-bound element. Full mechanism (the 4-branch
'              release state machine): see _bmad/docs/property-rendering-mechanics.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: PropertyTagging, StringsInEl, PropertyCalculation, ErrorHandlerClass (global ErrorHandler),
'               PropertyRendering (Core), PropertyRendering_Types, PropertyRendering_TemplateModel,
'               PropertyRendering_Authoring, PropertyRendering_Serialization, PropertyRendering_Reporting

Option Explicit

' Branches 1-3, per stored entry. texts() holds the whole text of every bearer, indexed by SubId. Write
' ORDER is load-bearing: every TEXT write happens first, the ONE metadata write last. Full rationale: see
' "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
Public Sub RenderBoundElement(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long)
    On Error GoTo ErrorHandler

    Dim ents() As RenderEntry
    Dim nEnts As Long
    Dim i As Long
    Dim outcome As Long
    Dim bDirty As Boolean
    Dim nLive As Long
    Dim wrIds() As Long
    Dim wrPrev() As String
    Dim nWr As Long
    Dim wSubId As Long
    Dim wPrev As String
    Dim subIds() As Long
    Dim nAdded As Long
    Dim bRepaint As Boolean

    ' ReadRenderMetadata surfaces the right status itself (a schema mismatch has a dedicated one, and the
    ' generic "unreadable" must not overwrite it).
    If Not PropertyRendering_Serialization.ReadRenderMetadata(oEl, ents, nEnts) Then Exit Sub

    ' Every entry is located up front, in two passes over the whole list, so that one sub-text is driven
    ' by ONE entry only and the result does not depend on the order the entries happen to sit in.
    If nEnts > 0 Then PropertyRendering.ResolveAllSubIds ents, nEnts, texts, nBearers, subIds

    ' THEN top up: a sub-text that no entry claims has never been authored, and being on an element that
    ' is already bound is not a reason to ignore it. Binding is per SUB-TEXT; the item on the cell header
    ' is just where the whole list lives. This also covers the legal "attached but Entries empty" state.
    nAdded = PropertyRendering_Authoring.AuthorUnclaimedBearers(oEl, texts, nBearers, ents, nEnts, subIds)
    If nEnts = 0 Then Exit Sub
    If nAdded > 0 Then PropertyRendering.ResolveAllSubIds ents, nEnts, texts, nBearers, subIds

    ' A newly authored entry must be PERSISTED even when its own render is a no-op (an unset value renders
    ' the literal token, which is what the bearer already shows - branch 1). Without this the new entry
    ' would be re-authored, and thrown away, on every single pass.
    bDirty = (nAdded > 0)
    nWr = 0
    For i = 0 To nEnts - 1
        wSubId = -1
        wPrev = ""
        outcome = RenderEntryOnElement(oEl, ents, i, texts, nBearers, subIds(i), wSubId, wPrev, bRepaint)
        If wSubId >= 0 Then
            ReDim Preserve wrIds(0 To nWr)
            ReDim Preserve wrPrev(0 To nWr)
            wrIds(nWr) = wSubId
            wrPrev(nWr) = wPrev
            nWr = nWr + 1
        End If
        If outcome = ENTRY_DROP Then
            ents(i).Dropped = True
            bDirty = True
        ElseIf outcome = ENTRY_UPDATED Then
            bDirty = True
        End If
    Next i

    If Not bDirty Then Exit Sub

    ' The whole binding was released: take the metadata off rather than leave an empty shell behind.
    nLive = PropertyRendering_Types.CountLiveEntries(ents, nEnts)
    If nLive = 0 Then
        PropertyTagging.DetachRenderMetadata oEl
        Exit Sub
    End If

    If Not PropertyRendering_Serialization.WriteRenderMetadata(oEl, ents, nEnts) Then
        For i = 0 To nWr - 1
            StringsInEl.SetTextAtSubId oEl, wrIds(i), wrPrev(i)
        Next i
        PropertyRendering_Reporting.ReportMetadataUnreadable
        PropertyRendering.DisableAfterWriteFailure
    ElseIf bRepaint Then
        ' Branch 3 transitioned an entry WITHOUT writing text. Noted only HERE, after the metadata write
        ' succeeded - see "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
        PropertyRendering.NoteDirtyGroup oEl
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_StateMachine.RenderBoundElement"
End Sub

' The three branches for ONE entry. Returns ENTRY_UNCHANGED / ENTRY_UPDATED / ENTRY_DROP. SubId is the
' ordinal ResolveAllSubIds settled (-1 = refuse; a drift relocation may differ from the stored one). oEl is
' ByRef down to StringsInEl.SetTextAtSubId (see WriteRenderedText). Full rationale for each branch: see
' "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
Private Function RenderEntryOnElement(ByRef oEl As element, ByRef ents() As RenderEntry, ByVal idx As Long, ByRef texts() As String, ByVal nBearers As Long, ByVal SubId As Long, ByRef wSubId As Long, ByRef wPrev As String, ByRef bRepaint As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim sVisible As String
    Dim sExpCur As String
    Dim sExpLast As String
    Dim bOkCur As Boolean
    Dim bOkLast As Boolean
    Dim curNames() As String
    Dim curValues() As String
    Dim nCur As Long
    Dim newNames() As String
    Dim newValues() As String
    Dim nNew As Long
    Dim sNewTemplate As String
    Dim newLits() As String
    Dim newToks() As String
    Dim nNewTok As Long
    Dim k As Long
    Dim bRelocated As Boolean

    RenderEntryOnElement = ENTRY_UNCHANGED
    wSubId = -1
    wPrev = ""

    ' The sub-text could not be identified (its ordinal drifted and its text is nowhere to be found, or
    ' another entry owns it): refuse rather than write blind into a sub-text that is not ours.
    If SubId < 0 Then
        PropertyRendering_Reporting.ReportDrift
        Exit Function
    End If
    If SubId > nBearers - 1 Then
        PropertyRendering_Reporting.ReportDrift
        Exit Function
    End If
    ' Relocation held in a LOCAL and committed only once the entry proves usable - see
    ' property-rendering-mechanics.md.
    bRelocated = (SubId <> ents(idx).SubId)

    sVisible = texts(SubId)
    ' An empty or faulted read is NEVER "the user wiped everything": skip, no transition.
    If Len(sVisible) = 0 Then
        PropertyRendering_Reporting.ReportMetadataUnreadable
        Exit Function
    End If

    ' Validation on read: the Template's token set must match the LastValues key set exactly. Metadata
    ' vandalised through the native Properties pane is never rendered as if it had been intended.
    If Not EntryIsConsistent(ents(idx)) Then
        PropertyRendering_Reporting.ReportMetadataInvalid
        Exit Function
    End If

    If Not PropertyRendering_Authoring.ReadCurrentValues(oEl, ents(idx), curNames, curValues, nCur) Then Exit Function

    ' ExpandTemplate fails OPEN and reaches COM - skip the entry on a fault, exactly as an empty read does.
    ' See "ExpandTemplate" in property-rendering-mechanics.md.
    sExpCur = PropertyRendering_TemplateModel.ExpandTemplate(ents(idx).Template, curNames, curValues, nCur, True, bOkCur)
    If Not bOkCur Then Exit Function
    sExpLast = PropertyRendering_TemplateModel.ExpandTemplate(ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, True, bOkLast)
    If Not bOkLast Then Exit Function

    ' Relocation is persisted only once the entry is proven readable/consistent/legal (edge #17). An
    ' undrifted entry still returns ENTRY_UNCHANGED, so branch 1 stays the strict no-op AC2 requires.
    If bRelocated Then
        ents(idx).SubId = SubId
        RenderEntryOnElement = ENTRY_UPDATED
    End If

    ' --- BRANCH 1: up to date. STRICT no-op - no text write, no Rewrite, no metadata write. The loop
    ' terminator, and the reason a re-queued unchanged element costs nothing.
    If sVisible = sExpCur Then Exit Function

    ' --- BRANCH 2: the visible still matches the LAST rendering, so the VALUES moved. Re-render. Second
    ' test also catches an entry authored but never rendered (branch 3 trap otherwise) - full rationale:
    ' see "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
    If sVisible = sExpLast Or StrComp(sVisible, ents(idx).Template, vbTextCompare) = 0 Then
        If PropertyRendering_Authoring.WriteRenderedText(oEl, SubId, sExpCur) Then
            wSubId = SubId
            wPrev = sVisible
            SetEntryValues ents, idx, curNames, curValues, nCur
            RenderEntryOnElement = ENTRY_UPDATED
        End If
        Exit Function
    End If

    ' --- BRANCH 3: a NON-EMPTY read differing from BOTH expansions - positive proof of a user edit.
    ' Per-token release, no text written. Full rationale: see "RenderBoundElement / RenderEntryOnElement"
    ' in property-rendering-mechanics.md.
    If PropertyRendering_TemplateModel.AlignVisible(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ' The re-authored Template is re-validated exactly like a first author.
        If PropertyRendering_TemplateModel.TemplateIsWellFormed(sNewTemplate, True) Then
            If PropertyRendering_TemplateModel.ParseTemplate(sNewTemplate, newLits, newToks, nNewTok, True) Then
                If nNewTok > 0 Then
                    For k = 0 To nNewTok - 1
                        If Not HasValueEntry(newToks(k), newNames, nNew) Then
                            AppendValue newToks(k), "", newNames, newValues, nNew
                        End If
                    Next k
                    ' The user's edit is about to release every token in ents(idx) that has no surviving
                    ' entry in newNames - warn for the ones a calc rule governs, BEFORE ents(idx) is
                    ' overwritten below. See "WarnGovernedTokensLost" in property-rendering-mechanics.md.
                    WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, newNames, nNew
                    ents(idx).Template = sNewTemplate
                    SetEntryValues ents, idx, newNames, newValues, nNew
                    bRepaint = True
                    RenderEntryOnElement = ENTRY_UPDATED
                Else
                    ' Alignment SUCCEEDED and concluded nothing survived - a deliberate release, not a
                    ' failure to understand the text; gets its own status rather than ReportAmbiguous below.
                    ' Every old token is lost here (newNames is empty) - warn for the governed ones.
                    WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, newNames, nNew
                    PropertyRendering_Reporting.ReportBindingReleased
                    RenderEntryOnElement = ENTRY_DROP
                End If
                Exit Function
            End If
        End If
    End If

    ' --- BRANCH 3b (D8): the literal walk could not follow the text (a position moved anywhere). Try to
    ' recognise the values themselves instead. Runs ONLY after AlignVisible has declined.
    If PropertyRendering_TemplateModel.AlignByValues(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ents(idx).Template = sNewTemplate
        SetEntryValues ents, idx, newNames, newValues, nNew
        bRepaint = True
        RenderEntryOnElement = ENTRY_UPDATED
        Exit Function
    End If

    ' Ambiguous alignment (or an invalid re-authored Template): keep only the literal tokens the visible
    ' still carries, drop everything else, and say so.
    PropertyRendering_Reporting.ReportAmbiguous
    RenderEntryOnElement = ApplyConservativeFallback(oEl, ents, idx, sVisible)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_StateMachine.RenderEntryOnElement"
    RenderEntryOnElement = ENTRY_UNCHANGED
End Function

' Conservative outcome: the visible text BECOMES the Template. Converges next pass (branch 2 or 1); an
' ill-formed or token-free visible RELEASES the entry outright instead. Full rationale: see
' "ApplyConservativeFallback" in property-rendering-mechanics.md. oEl is needed only to warn for a
' released token a calc rule governs (WarnGovernedTokensLost) - this is still "the user edited the text",
' the same release event as Branch 3's, just reached via the ambiguous fallback instead of AlignVisible.
Private Function ApplyConservativeFallback(ByVal oEl As element, ByRef ents() As RenderEntry, ByVal idx As Long, ByVal sVisible As String) As Long
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim names() As String
    Dim values() As String
    Dim n As Long

    ApplyConservativeFallback = ENTRY_DROP
    n = 0

    If Not PropertyRendering_TemplateModel.TemplateIsWellFormed(sVisible, True) Then
        WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, names, n
        Exit Function
    End If
    If Not PropertyRendering_TemplateModel.ParseTemplate(sVisible, lits, toks, nTok, True) Then
        WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, names, n
        Exit Function
    End If
    If nTok = 0 Then
        WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, names, n
        Exit Function
    End If

    For i = 0 To nTok - 1
        AppendValue toks(i), "", names, values, n
    Next i
    WarnGovernedTokensLost oEl, ents(idx).ValNames, ents(idx).nVals, names, n

    ents(idx).Template = sVisible
    SetEntryValues ents, idx, names, values, n
    ApplyConservativeFallback = ENTRY_UPDATED
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_StateMachine.ApplyConservativeFallback"
    ApplyConservativeFallback = ENTRY_DROP
End Function

' Validation on read: the Template's token set and the LastValues key set must be the SAME set. Anything
' else is metadata that was hand-edited through the native Properties pane, and it is never rendered as
' if it had been intended. Relocated here (from the file's tail "value-map helpers" section) - its only
' caller is RenderEntryOnElement, so it travels with its one consumer.
Private Function EntryIsConsistent(ByRef ent As RenderEntry) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long

    EntryIsConsistent = False
    If Not PropertyRendering_TemplateModel.ParseTemplate(ent.Template, lits, toks, nTok, True) Then Exit Function
    If nTok <> ent.nVals Then Exit Function

    For i = 0 To nTok - 1
        If Not HasValueEntry(toks(i), ent.ValNames, ent.nVals) Then Exit Function
    Next i

    EntryIsConsistent = PropertyRendering_TemplateModel.TemplateIsWellFormed(ent.Template, True)
    Exit Function

ErrorHandler:
    EntryIsConsistent = False
End Function

' A user edit just released oldNames(0..oldCount-1)'s tokens that have no surviving entry in
' newNames(0..newCount-1) - warn for each released token whose property is governed by a calc rule
' (PropertyCalculation.GetCalcRuleForProperty). One-shot per element pass via ReportGovernedValue, so
' losing several governed tokens in the same pass still surfaces a single status line, consistent with
' every other one-shot report in this module.
' Deliberately NOT called at bind time (TryAuthorBearer): every token is freshly attached and unedited
' there, so "editing this won't change it" would be true before any edit ever happened - it only means
' something once an edit has actually cost the binding a governed token. See "WarnGovernedTokensLost" in
' property-rendering-mechanics.md.
Private Sub WarnGovernedTokensLost(ByVal oEl As element, ByRef oldNames() As String, ByVal oldCount As Long, ByRef newNames() As String, ByVal newCount As Long)
    On Error Resume Next

    Dim i As Long
    Dim kind As CalcSource
    Dim sArg As String
    Dim sCanonical As String
    Dim sBase As String, sMember As String

    For i = 0 To oldCount - 1
        If Not HasValueEntry(oldNames(i), newNames, newCount) Then
            ' The calc rule targets the ITEM as a whole - a lost "Base:X"/"Base:Y" split-coordinate token
            ' resolves the rule lookup against the BASE name, never the full token. See
            ' plan-xy-split-coordinate-properties.md §5.3.
            PropertyRendering_TemplateModel.SplitTokenMember oldNames(i), sBase, sMember
            If PropertyCalculation.GetCalcRuleForProperty(sBase, oEl, kind, sArg, sCanonical) Then
                PropertyRendering_Reporting.ReportGovernedValue oldNames(i)
            End If
        End If
    Next i
End Sub
