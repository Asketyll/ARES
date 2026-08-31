' Module: RegionMerge
' Description: Geometry engine for the MergeRegion command. Fuses two closed regions
'              (Shape / ComplexShape) into a single region via GetRegionUnion. The merged
'              result inherits the FIRST region's level + symbology; both originals are
'              deleted (default) or kept (ARES_RegionMerge_Keep_Originals).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ErrorHandlerClass, LangManager
'
' Mechanic: GetRegionUnion, evaluated near the origin (Zoning / RegionSplit precision
' workaround). GetRegionUnion can return more than one element when the two regions do
' not overlap - that case is REJECTED (no model change, both originals kept): a disjoint
' result is not a real visual fusion, and applying oFirst's symbology to an untouched
' second shape was confusing in practice (decided 2026-08-27, after live testing).
Option Explicit

' MergeElements
' ---------------------------------------------------------------------------
' Sole public engine entry. Validates inputs, computes the union near the origin, writes
' the resulting element(s) with oFirst's symbology, then disposes of both originals per
' ARES_RegionMerge_Keep_Originals.
'
' Ordering guarantee: build + validate the merged result(s) FIRST, add them, THEN delete
' both originals. Any error before completion leaves both originals intact (no destructive
' partial edit).
'
' Parameters:
'   oFirst  - the first region clicked; its level + symbology is what the merged result inherits
'   oSecond - the second region clicked
Public Sub MergeElements(ByVal oFirst As Element, ByVal oSecond As Element)
    On Error GoTo ErrorHandler

    Dim bKeepOriginals As Boolean
    Dim merged()       As Element
    Dim nMerged        As Long

    ' --- Read fate from config (config var, not a literal) ---
    If Not ReadConfig(bKeepOriginals) Then
        ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion: configuration unavailable"
        Exit Sub
    End If

    ' --- Validate both regions + active model ---
    If Not IsMergeableRegion(oFirst) Or Not IsMergeableRegion(oSecond) Then
        ShowMergeStatus "MergeRegionNoRegion", "MergeRegion: not a supported closed region"
        ErrorHandler.HandleError "oFirst/oSecond is Nothing / not a supported closed region", 0, "", "RegionMerge.MergeElements"
        Exit Sub
    End If
    If DLongComp(oFirst.ID, oSecond.ID) = 0 Then
        ShowMergeStatus "MergeRegionSameZone", "MergeRegion: same element clicked twice"
        Exit Sub
    End If
    If Not Application.HasActiveModelReference Then
        ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion: no active model"
        ErrorHandler.HandleError "No active model reference", 0, "", "RegionMerge.MergeElements"
        Exit Sub
    End If

    ' --- Boolean union near the origin -> one region (or several, when disjoint) ---
    nMerged = 0
    If Not UnionNearOrigin(oFirst, oSecond, merged, nMerged) Then
        ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion: boolean union failed"
        ErrorHandler.HandleError "GetRegionUnion failed during merge", 0, "", "RegionMerge.MergeElements"
        Exit Sub
    End If

    ' Must yield >= 1 non-empty region, else abort with no model change.
    If nMerged < 1 Then
        ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion: union produced no region"
        ErrorHandler.HandleError "Boolean union produced no region", 0, "", "RegionMerge.MergeElements"
        Exit Sub
    End If

    ' Disjoint regions (they do not overlap/touch): GetRegionUnion returns > 1 element, i.e. no
    ' real visual fusion happened. Reject outright rather than silently recolour a second,
    ' untouched shape with oFirst's symbology (decided 2026-08-27, after live testing) - no model
    ' change, both originals left intact.
    If nMerged > 1 Then
        ShowMergeStatus "MergeRegionDisjoint", "MergeRegion: regions do not overlap, refused"
        Exit Sub
    End If

    ' --- Write the merged result with oFirst's symbology, THEN delete both originals ---
    ' WriteMerged is a Function returning True ONLY if every produced element was added and
    ' styled. Both originals are deleted ONLY on real success AND Not bKeepOriginals, so a
    ' partial write failure leaves both originals intact (anti-destructive ordering holds on
    ' the error path too).
    If Not WriteMerged(oFirst, merged, nMerged) Then
        ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion: failed to write the merged region"
        ErrorHandler.HandleError "WriteMerged failed; originals left intact", 0, "", "RegionMerge.MergeElements"
        Exit Sub
    End If
    If Not bKeepOriginals Then
        ActiveModelReference.RemoveElement oFirst
        ActiveModelReference.RemoveElement oSecond
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMerge.MergeElements"
    ShowMergeStatus "MergeRegionCannotMerge", "MergeRegion failed: " & Err.Description
End Sub

' ============================================================
'  CONFIG / VALIDATION
' ============================================================

' ReadConfig
' Reads the single RegionMerge config var. Boolean parsed with the UCase(Trim(...)) =
' "TRUE" idiom (as RegionSplit.ReadConfig / Command.bas do).
Private Function ReadConfig(ByRef bKeepOriginals As Boolean) As Boolean
    On Error GoTo ErrorHandler

    ReadConfig = False
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then Exit Function

    bKeepOriginals = (UCase(Trim(ARESConfig.ARES_REGIONMERGE_KEEP_ORIGINALS.Value)) = "TRUE")

    ReadConfig = True
    Exit Function

ErrorHandler:
    ReadConfig = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMerge.ReadConfig"
End Function

' IsMergeableRegion
' True only for a supported closed region: Shape or ComplexShape. Mirrors
' RegionSplit.IsSplittableRegion. Defence-in-depth even though the locator already filtered.
Private Function IsMergeableRegion(ByVal oRegion As Element) As Boolean
    On Error GoTo ErrorHandler
    IsMergeableRegion = False
    If oRegion Is Nothing Then Exit Function
    If Not oRegion.IsGraphical Then Exit Function
    IsMergeableRegion = (oRegion.IsShapeElement Or oRegion.IsComplexShapeElement)
    Exit Function

ErrorHandler:
    IsMergeableRegion = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMerge.IsMergeableRegion"
End Function

' ============================================================
'  BOOLEAN UNION
' ============================================================

' UnionNearOrigin
' Performs merged = oFirst + oSecond via GetRegionUnion, evaluated near the origin
' (Zoning / RegionSplit precision workaround): clone both regions, translate them by
' -oFirst.Range.High, run the boolean, then translate each result back. Accumulates the
' resulting region(s) into outMerged() -- GetRegionUnion can return more than one element
' when the two regions do not overlap (see Zoning.FuseRegions, same idiom). Returns False
' only on a hard failure (GetRegionUnion itself returning Nothing); an empty result still
' returns True with nMerged = 0, letting the caller report it as a clean "no region" abort.
Private Function UnionNearOrigin(ByVal oFirst As Element, _
                                 ByVal oSecond As Element, _
                                 ByRef outMerged() As Element, _
                                 ByRef nMerged As Long) As Boolean
    On Error GoTo ErrorHandler

    UnionNearOrigin = False
    nMerged = 0

    Dim toOrigin   As Point3d
    Dim fromOrigin As Point3d
    toOrigin = Point3dNegate(oFirst.Range.High)
    fromOrigin = Point3dNegate(toOrigin)

    Dim firstClone  As Element
    Dim secondClone As Element
    Set firstClone = oFirst.Clone
    Set secondClone = oSecond.Clone
    firstClone.Move toOrigin
    secondClone.Move toOrigin

    Dim region1(0 To 0) As Element
    Dim region2(0 To 0) As Element
    Set region1(0) = firstClone
    Set region2(0) = secondClone

    Dim oEnum As ElementEnumerator
    Set oEnum = GetRegionUnion(region1, region2, Nothing, msdFillModeNotFilled)
    If oEnum Is Nothing Then Exit Function

    Dim oRes As Element
    Do While oEnum.MoveNext
        Set oRes = oEnum.Current
        If Not oRes Is Nothing Then
            oRes.Move fromOrigin                  ' restore to the original location
            ReDim Preserve outMerged(0 To nMerged)
            Set outMerged(nMerged) = oRes
            nMerged = nMerged + 1
        End If
    Loop

    UnionNearOrigin = True
    Exit Function

ErrorHandler:
    UnionNearOrigin = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMerge.UnionNearOrigin"
End Function

' ============================================================
'  OUTPUT
' ============================================================

' WriteMerged
' Applies oFirst's level + symbology (the first region clicked) to every produced element
' and adds it to the active model. Called only after >= 1 result is validated; both
' originals are deleted by the caller AFTER this returns True (add-both-then-delete
' ordering).
'
' This is a Function returning True ONLY if every produced element was added and styled
' successfully. Any AddElement / property-set / Rewrite failure routes to ErrorHandler and
' returns False, so the caller does NOT delete the originals (anti-destructive on the error
' path too). Mirrors RegionSplit.WriteHalves: Level/LineStyle are object-valued since
' MicroStation 8.1 (read into locals with Set; written onto the element by reference, no
' Set); Level can only be set once the element is a model member, so each result is
' AddElement'd FIRST, then its level + symbology applied.
Private Function WriteMerged(ByVal oFirst As Element, _
                             ByRef merged() As Element, _
                             ByVal nMerged As Long) As Boolean
    On Error GoTo ErrorHandler
    WriteMerged = False

    Dim srcLevel  As Level
    Dim srcStyle  As LineStyle
    Dim srcColor  As Long
    Dim srcWeight As Long
    Set srcLevel = oFirst.Level
    Set srcStyle = oFirst.LineStyle
    srcColor = oFirst.Color
    srcWeight = oFirst.LineWeight

    Dim i        As Long
    Dim nWritten As Long
    nWritten = 0
    For i = 0 To nMerged - 1
        If Not merged(i) Is Nothing Then
            ' Add first: Level cannot be assigned to a non-member element (see doc).
            ActiveModelReference.AddElement merged(i)
            merged(i).Level = srcLevel
            merged(i).Color = srcColor
            merged(i).LineStyle = srcStyle
            merged(i).LineWeight = srcWeight
            merged(i).Rewrite
            nWritten = nWritten + 1
        End If
    Next i

    ' True only if every produced result really made it into the model.
    WriteMerged = (nWritten >= 1) And (nWritten = nMerged)
    Exit Function

ErrorHandler:
    WriteMerged = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMerge.WriteMerged"
End Function

' ShowMergeStatus
' Shows a translated, user-facing status for a merge outcome through the shared
' LangManager.ShowStatusT (the single channel for all user status; it self-initialises the
' translation system). ReasonEN is the inline English reason for that abort branch, kept as
' code documentation -- genuine faults are logged separately by the ErrorHandler blocks.
Private Sub ShowMergeStatus(ByVal Key As String, ByVal ReasonEN As String)
    ShowStatusT Key
End Sub
