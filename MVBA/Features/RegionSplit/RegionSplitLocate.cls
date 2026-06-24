' Class Module: RegionSplitLocate
' Description: Single-click driver for the SplitRegion command. Implements
'              IPrimitiveCommandEvents so ONE datapoint does everything: DataPoint locates the
'              element under the click (CommandState.LocateElement, a single call) and passes
'              it, with the clicked point, to RegionSplit.SplitElementAt. After each cut it
'              re-issues StartPrimitive to clear the locate highlight (so no element is left
'              selected) and ready the next cut; Reset ends the command.
'              (Replaces the earlier ILocateCommandEvents model, which required TWO datapoints
'              -- identify then accept -- and left the located element highlighted.)
' License: This project is licensed under the AGPL-3.0.
' Dependencies: RegionSplit, CommandState, LangManager, ErrorHandlerClass
Option Explicit

Implements IPrimitiveCommandEvents

' _Start: show the locate cursor and the command/prompt strings.
Private Sub IPrimitiveCommandEvents_Start()
    On Error GoTo ErrorHandler
    CommandState.SetLocateCursor
    ShowCommand GetTranslation("RegionSplitSelectRegionC")
    ShowPrompt GetTranslation("RegionSplitSelectRegionP")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplitLocate.Start"
End Sub

' _DataPoint: THE single click. LocateElement finds the element under the point in one
' datapoint; step through the hit list until a supported closed region (Shape / ComplexShape)
' is found, cut it, then restart the primitive to clear the highlight and accept the next cut.
Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    On Error GoTo ErrorHandler

    Dim oEle As Element
    Set oEle = CommandState.LocateElement(Point, View, True)
    Do While Not oEle Is Nothing
        If oEle.IsShapeElement Or oEle.IsComplexShapeElement Then
            RegionSplit.SplitElementAt oEle, Point
            CommandState.StartPrimitive Me      ' clear highlight + ready for the next region
            Exit Sub
        End If
        Set oEle = CommandState.LocateElement(Point, View, False)   ' next element under the click
    Loop

    ' Nothing splittable under the click.
    ShowStatus GetTranslation("RegionSplitNoRegion")
    CommandState.StartPrimitive Me
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplitLocate.DataPoint"
    On Error Resume Next
    CommandState.StartPrimitive Me
End Sub

' _Reset: end the command and return MicroStation to its default tool.
Private Sub IPrimitiveCommandEvents_Reset()
    On Error Resume Next
    CommandState.StartDefaultCommand
End Sub

' _Keyin: not used (StartPrimitive was not asked to forward key-ins).
Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
    ' Intentionally empty.
End Sub

' _Dynamics: no rubber-band preview of the cut (out of scope).
Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    ' Intentionally empty: the cut is computed and applied on the confirming click only.
End Sub

' _Cleanup: nothing to release beyond Class_Terminate.
Private Sub IPrimitiveCommandEvents_Cleanup()
    ' Intentionally empty.
End Sub

' Explicit cleanup to avoid errors during implicit destruction (Auto_Lengths idiom).
Private Sub Class_Terminate()
    On Error Resume Next
End Sub
