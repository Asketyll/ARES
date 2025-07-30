' Module: MSGraphicalInteraction
' Description: This module provides functions to interact with views, zoom, and highlight elements in MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None

Option Explicit

' Container for transient elements
Public TEC As TransientElementContainer

' Function to zoom on an element in a specified view
Public Function ZoomEl(ByVal el As Element, Optional Factor As Single = 1.3) As Boolean
    On Error GoTo ErrorHandler

    Dim Rng As Range3d
    Dim PntZoom As Point3d
    Dim oView As View
    Dim pntCenter As Point3d
    Dim Pnt As Point3d

    ZoomEl = False

    ' Check if the element is graphical
    If el.IsGraphical Then
        ' Get the Last View
        Set oView = CommandState.LastView
        ' Get the range of the element
        Rng = el.Range

        ' Calculate the zoom point based on the range of the element
        With Rng
            PntZoom.X = .High.X - .Low.X
            PntZoom.Y = .High.Y - .Low.Y
            PntZoom.Z = .High.Z - .Low.Z
        End With

        ' Set the point for zooming
        With Pnt
            .X = PntZoom.X * Factor
            .Y = PntZoom.Y * Factor
            .Z = PntZoom.Z * Factor
        End With

        ' Set the view area and zoom
        oView.SetArea Rng.Low, Pnt, oView.Rotation, Rng.High.Z
        oView.ZoomAboutPoint Point3dAddScaled(Rng.Low, PntZoom, 0.5), 1
        oView.Redraw

        ZoomEl = True
    End If

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    ZoomEl = False
End Function

' Function to highlight an element
Public Function HighlightEl(ByVal el As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim Flags As MsdTransientFlags

    HighlightEl = False

    ' Clear the transient element container
    Set TEC = Nothing

    ' Set the flags for the transient element
    Flags = msdTransientFlagsOverlay + msdTransientFlagsSnappable

    ' Highlight the element
    el.IsHighlighted = True
    Set TEC = CreateTransientElementContainer1(el, Flags, msdViewAll, msdDrawingModeHilite)
    el.IsHighlighted = False

    HighlightEl = True

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    HighlightEl = False
End Function
