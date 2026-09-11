' Module: GroupSplitGuard
' Description: A cut that leaves several cable geometries in one graphic group gives every text reading its
'              values off that group (GroupLength, GroupProp, GroupColor) several sources of truth. No hook can
'              veto a command before it runs, so the change-track stream is recorded instead: a MARK closes a
'              command's block, and a block that deletes a geometry of group G and adds two or more geometries
'              into that same G is a cut. The user is then asked, at the next idle, to separate the group, undo
'              the cut, or keep it. Measured sequences and undo scope:
'              _bmad-output/implementation-artifacts/group-split-guard.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass (global ARESConfig), ARESConstants, Length, Link, CustomPropertyHandler,
'               PropertyCalculation (HasElements), PropertyCalculation_SourceEval (GetElementAnchorPoint),
'               Zoning_Buffer (ChainSeg, Flatten, DistanceToChain), ElementChangeHandler (global ChangeHandler),
'               LangManager, ErrorHandlerClass (global ErrorHandler), CallStackClass (global CallStack),
'               GroupSplitGuardIdle

Option Explicit

' A block larger than this is a bulk operation, not a cut: recording stops until the next MARK.
Private Const MAX_BLOCK_EVENTS As Long = 2000

' The block being recorded - every geometry delete and add since the last MARK.
Private mDelGroups() As Long
Private mnDel As Long
Private mAddGroups() As Long
Private mAddIds() As DLong
Private mnAdd As Long
Private mbBlockOverflow As Boolean

' The cut waiting for its question. mlChangesSince counts what happened after its MARK: once anything has,
' UNDO would no longer reach the cut.
Private mbPending As Boolean
Private mlPendingGroup As Long
Private mPendingIds() As DLong
Private mnPendingIds As Long
Private mlChangesSince As Long

' Held here so VBA cannot collect the one-shot idle handler before it fires.
Private moIdle As GroupSplitGuardIdle

'######################################################################################################################
'                                          SWITCH
'######################################################################################################################

' Master switch (ARES_Group_Split_Guard, True by default). Fail-closed False on any nil. Read inside
' change-track callbacks, so it never initialises the configuration itself.
Public Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then Exit Function
    If ARESConfig.ARES_GROUP_SPLIT_GUARD Is Nothing Then Exit Function
    IsEnabled = CBool(ARESConfig.ARES_GROUP_SPLIT_GUARD.Value)
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

'######################################################################################################################
'                                          RECORDING (change-track callbacks)
'######################################################################################################################

' Called for every ElementChanged. Records the geometry deletes and adds of grouped elements; anything at all
' that happens while a question is pending makes UNDO unsafe.
Public Sub NoteChange(ByVal AfterChange As element, ByVal BeforeChange As element, ByVal Action As Long)
    On Error GoTo ErrorHandler

    If Not IsEnabled Then Exit Sub
    If mbPending Then mlChangesSince = mlChangesSince + 1

    Select Case Action
        Case msdChangeTrackActionDelete
            If IsGroupedGeometry(BeforeChange) Then RecordGeometryDelete BeforeChange.GraphicGroup
        Case msdChangeTrackActionAdd
            If IsGroupedGeometry(AfterChange) Then RecordGeometryAdd AfterChange.GraphicGroup, AfterChange.ID
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.NoteChange"
End Sub

' Called for every MARK: closes the block and, when it holds a cut, schedules the question for the next idle.
' One question at a time - a second cut before that idle is not asked about.
Public Sub NoteMark()
    On Error GoTo ErrorHandler

    Dim lGroup As Long
    Dim ids() As DLong
    Dim nIds As Long

    If Not IsEnabled Then Exit Sub
    If Not CloseBlock(lGroup, ids, nIds) Then Exit Sub
    If mbPending Then Exit Sub

    mbPending = True
    mlPendingGroup = lGroup
    mPendingIds = ids
    mnPendingIds = nIds
    mlChangesSince = 0
    ScheduleQuestion
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.NoteMark"
End Sub

' An undo or redo already settles the question - never ask about a cut that has been undone or replayed.
Public Sub NoteUndoRedo()
    On Error GoTo ErrorHandler

    ClearBlock
    mbPending = False
    mnPendingIds = 0
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.NoteUndoRedo"
End Sub

'######################################################################################################################
'                                          BLOCK SIGNATURE (pure - test seam)
'######################################################################################################################

Public Sub RecordGeometryDelete(ByVal lGroup As Long)
    On Error GoTo ErrorHandler

    If BlockIsFull() Then Exit Sub
    ReDim Preserve mDelGroups(0 To mnDel)
    mDelGroups(mnDel) = lGroup
    mnDel = mnDel + 1
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.RecordGeometryDelete"
End Sub

Public Sub RecordGeometryAdd(ByVal lGroup As Long, ByRef id As DLong)
    On Error GoTo ErrorHandler

    If BlockIsFull() Then Exit Sub
    ReDim Preserve mAddGroups(0 To mnAdd)
    ReDim Preserve mAddIds(0 To mnAdd)
    mAddGroups(mnAdd) = lGroup
    mAddIds(mnAdd) = id
    mnAdd = mnAdd + 1
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.RecordGeometryAdd"
End Sub

' Closes the recorded block and says whether it is a cut: a geometry of group G deleted and at least two
' geometries added into that same G. Returns that group and the IDs added into it. The block is emptied either
' way, and an overflowed block is never a cut.
Public Function CloseBlock(ByRef lGroup As Long, ByRef ids() As DLong, ByRef nIds As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim d As Long
    Dim a As Long

    CloseBlock = False
    lGroup = 0
    nIds = 0

    If Not mbBlockOverflow Then
        For d = 0 To mnDel - 1
            nIds = 0
            For a = 0 To mnAdd - 1
                If mAddGroups(a) = mDelGroups(d) Then
                    ReDim Preserve ids(0 To nIds)
                    ids(nIds) = mAddIds(a)
                    nIds = nIds + 1
                End If
            Next a
            If nIds >= 2 Then
                lGroup = mDelGroups(d)
                CloseBlock = True
                Exit For
            End If
        Next d
        If Not CloseBlock Then nIds = 0
    End If

    ClearBlock
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.CloseBlock"
    ClearBlock
    CloseBlock = False
    nIds = 0
End Function

' Clears the block and any pending question, leaving the idle handler alone. For the unit tests.
Public Sub ResetForTests()
    On Error GoTo ErrorHandler

    ClearBlock
    mbPending = False
    mnPendingIds = 0
    mlChangesSince = 0
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.ResetForTests"
End Sub

Private Sub ClearBlock()
    mnDel = 0
    mnAdd = 0
    mbBlockOverflow = False
End Sub

Private Function BlockIsFull() As Boolean
    If mnDel + mnAdd >= MAX_BLOCK_EVENTS Then mbBlockOverflow = True
    BlockIsFull = mbBlockOverflow
End Function

'######################################################################################################################
'                                          THE QUESTION (idle)
'######################################################################################################################

Private Sub ScheduleQuestion()
    On Error GoTo ErrorHandler

    If Not moIdle Is Nothing Then Exit Sub
    Set moIdle = New GroupSplitGuardIdle
    AddEnterIdleEventHandler moIdle
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.ScheduleQuestion"
    Set moIdle = Nothing
End Sub

' Called by GroupSplitGuardIdle once it has left the idle list.
Public Sub ReleaseIdleHandler()
    Set moIdle = Nothing
End Sub

' The pending cut is taken and cleared first: whatever the answer, it is never asked about twice, and the
' writes a separation makes do not count as changes since the cut. Cancel undoes the cut, and Escape and the
' close box answer Cancel too - acceptable, an undo being one Redo away.
Public Sub AskPending()
    On Error GoTo ErrorHandler

    Dim bStackPushed As Boolean
    Dim lGroup As Long
    Dim ids() As DLong
    Dim nIds As Long
    Dim lChanges As Long
    Dim pieces() As element
    Dim nPieces As Long
    Dim sMsg As String
    Dim lAnswer As VbMsgBoxResult

    If Not mbPending Then Exit Sub
    lGroup = mlPendingGroup
    ids = mPendingIds
    nIds = mnPendingIds
    lChanges = mlChangesSince
    mbPending = False
    mnPendingIds = 0
    mlChangesSince = 0

    CallStack.Push "GroupSplitGuard.AskPending"
    bStackPushed = True

    If Not CollectPieces(lGroup, ids, nIds, pieces, nPieces) Then GoTo Done
    If Not GroupCarriesAresItems(pieces(0)) Then GoTo Done

    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    sMsg = GetTranslation("GroupSplitGuardIntro", nPieces, lGroup) & vbCrLf & vbCrLf & _
           GetTranslation("GroupSplitGuardYesSeparate") & vbCrLf
    If lChanges = 0 Then
        sMsg = sMsg & GetTranslation("GroupSplitGuardNoKeep") & vbCrLf & GetTranslation("GroupSplitGuardCancelUndo")
        lAnswer = MsgBox(sMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1, "ARES")
    Else
        sMsg = sMsg & GetTranslation("GroupSplitGuardNoKeep") & vbCrLf & vbCrLf & GetTranslation("GroupSplitGuardUndoUnavailable")
        lAnswer = MsgBox(sMsg, vbYesNo + vbQuestion + vbDefaultButton1, "ARES")
    End If

    Select Case lAnswer
        Case vbYes
            SeparateGroup pieces, nPieces, lGroup
        Case vbCancel
            CadInputQueue.SendKeyin "UNDO"
        Case Else
    End Select

Done:
    If bStackPushed Then CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.AskPending"
    If bStackPushed Then CallStack.Pop
End Sub

'######################################################################################################################
'                                          SEPARATION
'######################################################################################################################

' Piece 0 keeps lGroup, every other piece gets a fresh group, and every other member joins the piece it lies
' nearest to. The members are listed before anything moves - once a piece leaves lGroup no scan of lGroup
' sees it again - and the pieces move before the members, so a text never lands in a group that has no
' geometry yet. A member that stays is re-processed: while the pieces shared the group, its values may have
' come off a piece that just left.
Private Sub SeparateGroup(ByRef pieces() As element, ByVal nPieces As Long, ByVal lGroup As Long)
    On Error GoTo ErrorHandler

    Dim members() As element
    Dim newGroups() As Long
    Dim bMoved() As Boolean
    Dim i As Long
    Dim k As Long
    Dim bJoin As Boolean
    Dim nMoved As Long
    Dim nLocked As Long

    members = Link.GetLink(pieces(0), True)

    ReDim newGroups(0 To nPieces - 1)
    ReDim bMoved(0 To nPieces - 1)
    newGroups(0) = lGroup
    For k = 1 To nPieces - 1
        If pieces(k).IsLocked Then
            nLocked = nLocked + 1
        Else
            newGroups(k) = Application.UpdateGraphicGroupNumber
            pieces(k).GraphicGroup = newGroups(k)
            pieces(k).Rewrite
            bMoved(k) = True
        End If
    Next k

    If PropertyCalculation.HasElements(members) Then
        For i = LBound(members) To UBound(members)
            If Not IsPiece(members(i), pieces, nPieces) Then
                k = NearestPiece(members(i), pieces, nPieces)
                bJoin = False
                If k > 0 Then
                    If bMoved(k) Then bJoin = True
                End If
                If Not bJoin Then
                    RefreshStayingMember members(i)
                ElseIf members(i).IsLocked Then
                    nLocked = nLocked + 1
                Else
                    members(i).GraphicGroup = newGroups(k)
                    members(i).Rewrite
                    nMoved = nMoved + 1
                End If
            End If
        Next i
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    ShowStatus GetTranslation("GroupSplitGuardSeparated", nPieces, nMoved, nLocked)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.SeparateGroup"
End Sub

' A member left in the original group goes back through ARES's normal pass, as if it had just changed.
Private Sub RefreshStayingMember(ByVal oMember As element)
    On Error GoTo ErrorHandler

    If ChangeHandler Is Nothing Then Exit Sub
    ChangeHandler.ProcessElement oMember
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.RefreshStayingMember"
End Sub

' Index of the piece oMember lies nearest to, from its anchor to each piece's own geometry; -1 when no
' distance can be measured.
Private Function NearestPiece(ByVal oMember As element, ByRef pieces() As element, ByVal nPieces As Long) As Long
    On Error GoTo ErrorHandler

    Dim pt As Point3d
    Dim segs() As ChainSeg
    Dim nSeg As Long
    Dim k As Long
    Dim d As Double
    Dim best As Double

    NearestPiece = -1
    If Not PropertyCalculation_SourceEval.GetElementAnchorPoint(oMember, pt) Then Exit Function

    best = 1E+30
    For k = 0 To nPieces - 1
        If Zoning_Buffer.Flatten(pieces(k), segs, nSeg) Then
            d = Zoning_Buffer.DistanceToChain(pt, segs, nSeg)
            If d < best Then
                best = d
                NearestPiece = k
            End If
        End If
    Next k
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.NearestPiece"
    NearestPiece = -1
End Function

'######################################################################################################################
'                                          HELPERS
'######################################################################################################################

' Graphical, grouped, and a geometry GroupLength would read. Nested Ifs, never an And chain: VBA does not
' short-circuit, and .GraphicGroup raises on a non-graphical element.
Private Function IsGroupedGeometry(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsGroupedGeometry = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsGraphical Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function
    IsGroupedGeometry = Length.IsLengthCapable(oEl)
    Exit Function

ErrorHandler:
    IsGroupedGeometry = False
End Function

' The recorded pieces, re-fetched after the commit, that are still geometries of lGroup. False when fewer than
' two remain - the user or a tool has already sorted the group out.
Private Function CollectPieces(ByVal lGroup As Long, ByRef ids() As DLong, ByVal nIds As Long, ByRef pieces() As element, ByRef nPieces As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim oEl As element

    CollectPieces = False
    nPieces = 0
    For i = 0 To nIds - 1
        Set oEl = FetchById(ids(i))
        If IsGroupedGeometry(oEl) Then
            If oEl.GraphicGroup = lGroup Then
                ReDim Preserve pieces(0 To nPieces)
                Set pieces(nPieces) = oEl
                nPieces = nPieces + 1
            End If
        End If
    Next i

    CollectPieces = (nPieces >= 2)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.CollectPieces"
    CollectPieces = False
End Function

' The trigger condition chosen for the question: some member of the group carries an ARES property. A group
' nothing reads from has no source of truth to confuse.
Private Function GroupCarriesAresItems(ByVal oAnyMember As element) As Boolean
    On Error GoTo ErrorHandler

    Dim members() As element
    Dim i As Long

    GroupCarriesAresItems = False
    members = Link.GetLink(oAnyMember, True)
    If Not PropertyCalculation.HasElements(members) Then Exit Function

    For i = LBound(members) To UBound(members)
        If CustomPropertyHandler.IsAnyItemAttachedToElement(members(i)) Then
            GroupCarriesAresItems = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuard.GroupCarriesAresItems"
    GroupCarriesAresItems = False
End Function

Private Function IsPiece(ByVal oEl As element, ByRef pieces() As element, ByVal nPieces As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim k As Long

    IsPiece = False
    For k = 0 To nPieces - 1
        If DLongComp(oEl.ID, pieces(k).ID) = 0 Then
            IsPiece = True
            Exit Function
        End If
    Next k
    Exit Function

ErrorHandler:
    IsPiece = False
End Function

' Nothing when the element no longer exists. Silent on purpose: a piece that is gone is an answer, not a fault.
Private Function FetchById(ByRef id As DLong) As element
    On Error Resume Next
    Set FetchById = Nothing
    Set FetchById = ActiveModelReference.GetElementByID(id)
End Function
