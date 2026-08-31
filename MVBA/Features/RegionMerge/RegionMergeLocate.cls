' Class Module: RegionMergeLocate
' Description: Two-click driver for the MergeRegion command. Implements
'              IPrimitiveCommandEvents. The FIRST datapoint locates a region under the
'              click (CommandState.LocateElement) and remembers it in moFirstZone.
'              StartPrimitive is deliberately NOT called here (it would clear the highlight
'              and re-run Start, see the IPrimitiveCommandEvents doc) -- but MicroStation's
'              own LocateElement highlight does NOT reliably persist across a second,
'              unrelated DataPoint either (confirmed live, 2026-08-27), so the first region
'              is explicitly re-highlighted via MSGraphicalInteraction.HighlightEl (a
'              transient-element overlay, independent of the native locate highlight) while
'              the second click is awaited. The SECOND datapoint locates the other region
'              and passes both, in click order, to RegionMerge.MergeElements; only THEN is
'              StartPrimitive re-issued (clears the native highlight) and the transient
'              overlay released, ready for the next pair. Reset ends the command and
'              releases the pending first pick the same way.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: RegionMerge, CommandState, LangManager, ErrorHandlerClass, MSGraphicalInteraction
Option Explicit

Implements IPrimitiveCommandEvents

' The region acquired on the first click of the current pair; Nothing while awaiting it.
Private moFirstZone As Element

' _Start: show the locate cursor, reset any pending first pick, and prompt for the first click.
Private Sub IPrimitiveCommandEvents_Start()
    On Error GoTo ErrorHandler
    Set moFirstZone = Nothing
    CommandState.SetLocateCursor
    ShowCommand GetTranslation("MergeRegionSelectFirstC")
    ShowPrompt GetTranslation("MergeRegionSelectFirstP")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMergeLocate.Start"
End Sub

' _DataPoint: routes on whether a first region is already acquired (moFirstZone).
Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    On Error GoTo ErrorHandler

    Dim oEle As Element
    Set oEle = LocateMergeableRegion(Point, View)

    If moFirstZone Is Nothing Then
        ' --- Awaiting the FIRST click ---
        If oEle Is Nothing Then
            ShowStatus GetTranslation("MergeRegionNoRegion")
            Exit Sub                                ' no state to lose; redo the first click
        End If
        Set moFirstZone = oEle
        ' Deliberately no StartPrimitive here: it would clear the highlight and re-run Start.
        ' The native LocateElement highlight alone does not reliably survive to the second
        ' DataPoint (confirmed live), so make the pick visible explicitly instead.
        MSGraphicalInteraction.HighlightEl moFirstZone
        ShowPrompt GetTranslation("MergeRegionSelectSecondP")
        Exit Sub
    End If

    ' --- Awaiting the SECOND click ---
    If oEle Is Nothing Then
        ShowStatus GetTranslation("MergeRegionNoRegion")
        Exit Sub                                    ' keep moFirstZone; redo the second click only
    End If
    If DLongComp(moFirstZone.ID, oEle.ID) = 0 Then
        ShowStatus GetTranslation("MergeRegionSameZone")
        Exit Sub                                    ' keep moFirstZone; redo the second click only
    End If

    RegionMerge.MergeElements moFirstZone, oEle
    Set moFirstZone = Nothing
    Set MSGraphicalInteraction.TEC = Nothing         ' release the first-click highlight overlay
    CommandState.StartPrimitive Me                  ' clear highlight + ready for the next pair
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMergeLocate.DataPoint"
    On Error Resume Next
    Set moFirstZone = Nothing
    Set MSGraphicalInteraction.TEC = Nothing
    CommandState.StartPrimitive Me
End Sub

' LocateMergeableRegion
' Steps through the hit list at Point/View until a supported closed region
' (Shape / ComplexShape) is found. Returns Nothing when none is found under the click.
Private Function LocateMergeableRegion(ByRef Point As Point3d, ByVal View As View) As Element
    On Error GoTo ErrorHandler
    Set LocateMergeableRegion = Nothing

    Dim oEle As Element
    Set oEle = CommandState.LocateElement(Point, View, True)
    Do While Not oEle Is Nothing
        If oEle.IsShapeElement Or oEle.IsComplexShapeElement Then
            Set LocateMergeableRegion = oEle
            Exit Function
        End If
        Set oEle = CommandState.LocateElement(Point, View, False)   ' next element under the click
    Loop
    Exit Function

ErrorHandler:
    Set LocateMergeableRegion = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionMergeLocate.LocateMergeableRegion"
End Function

' _Reset: release the pending first pick (and its highlight overlay) and end the command.
Private Sub IPrimitiveCommandEvents_Reset()
    On Error Resume Next
    Set moFirstZone = Nothing
    Set MSGraphicalInteraction.TEC = Nothing
    CommandState.StartDefaultCommand
End Sub

' _Keyin: not used (StartPrimitive was not asked to forward key-ins).
Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
    ' Intentionally empty.
End Sub

' _Dynamics: no rubber-band preview of the merge (out of scope).
Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    ' Intentionally empty: the merge is computed and applied on the confirming click only.
End Sub

' _Cleanup: nothing to release beyond Class_Terminate.
Private Sub IPrimitiveCommandEvents_Cleanup()
    ' Intentionally empty.
End Sub

' Explicit cleanup to avoid errors during implicit destruction, and to release the pending
' first pick if the instance is destroyed mid-sequence. Standard idiom for the
' ILocateCommandEvents classes; no dependency on any other feature.
Private Sub Class_Terminate()
    On Error Resume Next
    Set moFirstZone = Nothing
    Set MSGraphicalInteraction.TEC = Nothing
End Sub
