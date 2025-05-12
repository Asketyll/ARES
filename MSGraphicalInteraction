' Module: MSGraphicalInteraction
' Description: This Module provides functions to interact with views, zoom and Highlight.

Option Explicit

Public TEC As TransientElementContainer

Public Function ZoomEl(ByVal el As Element, Optional Factor As Single = 1.3, Optional intView As Integer = 1) As Boolean
    On Error GoTo ErrorHandler
    
    Dim Rng As Range3d
    Dim PntZoom As Point3d
    Dim oView As View
    Dim pntCenter As Point3d
    Dim Pnt As Point3d
    ZoomEl = False
    
    If el.IsGraphical Then
    
        Set oView = ActiveDesignFile.Views(intView)
        
        Rng = el.Range
        
        With Rng
            PntZoom.X = .High.X - .Low.X
            PntZoom.Y = .High.Y - .Low.Y
            PntZoom.Z = .High.Z - .Low.Z
        End With
        With Pnt
            .X = PntZoom.X * Factor
            .Y = PntZoom.Y * Factor
            .Z = PntZoom.Z * Factor
        End With
        oView.SetArea Rng.Low, Pnt, oView.Rotation, Rng.High.Z
        oView.ZoomAboutPoint Point3dAddScaled(Rng.Low, PntZoom, 0.5), 1
        oView.Redraw
        ZoomEl = True
    End If
    Exit Function
    
ErrorHandler:
    ZoomEl = False
End Function
Public Function HighlightEl(ByVal el As Element) As Boolean
    On Error GoTo ErrorHandler
    
    Dim Flags As MsdTransientFlags
    HighlightEl = False
    Set TEC = Nothing
    Flags = msdTransientFlagsOverlay + msdTransientFlagsSnappable
    el.IsHighlighted = True
    Set TEC = CreateTransientElementContainer1(el, Flags, msdViewAll, msdDrawingModeHilite)
    el.IsHighlighted = False
    HighlightEl = True
    Exit Function
    
ErrorHandler:
    HighlightEl = False
End Function
