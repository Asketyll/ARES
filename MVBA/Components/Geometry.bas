' Module: Geometry
' Description: Pure geometry helpers shared across features. These operate ONLY on Point3d /
'              vectors / angles using native MVBA point math — no document state, no config,
'              no model I/O — so they are safe to call from any feature.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass
Option Explicit

' Perp2D
' ---------------------------------------------------------------------------
' Returns the left-hand perpendicular vector for segment A→B, scaled to Dist.
' "Left" = 90° counter-clockwise from the direction of travel.
'
'   A ──────────────────► B
'            ↑
'         result (this function, length = Dist)
'
' Returns a zero Point3d if A and B are coincident (zero-length segment).
' Callers should check Point3dMagnitudeSquared(result) < 1E-24 to detect this.
'
' Built entirely from native MVBA point math: the direction A→B is rotated 90° CCW
' (Point3dRotateXY), normalised (Point3dNormalize), then scaled to Dist (Point3dScale).
' Point3dNormalize returns a zero vector for a zero-length input, so a coincident A/B
' naturally yields a zero result.
' ---------------------------------------------------------------------------
Public Function Perp2D(ByRef A As Point3d, ByRef B As Point3d, ByVal Dist As Double) As Point3d
    On Error GoTo ErrorHandler
    Perp2D = Point3dScale(Point3dNormalize(Point3dRotateXY(Point3dSubtract(B, A), Application.PI / 2)), Dist)
    Exit Function

ErrorHandler:
    ' Default-initialised Point3d (zero vector) tells the caller "degenerate input".
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Geometry.Perp2D"
End Function

' NormalizeAngle
' ---------------------------------------------------------------------------
' Adjusts a sweep angle (delta) to lie in the correct half-open interval
' for CreateArcElement2 based on the intended sweep direction.
'
'   direction > 0  → result in (0,  2π]   (counter-clockwise sweep)
'   direction < 0  → result in [-2π, 0)   (clockwise sweep)
'
' WHY: When computing the angular difference between two points that cross the
' ±π boundary, raw subtraction can produce a value with the wrong sign or
' magnitude. This function corrects it by adding/subtracting 2π as needed.
' ---------------------------------------------------------------------------
Public Function NormalizeAngle(ByVal delta As Double, ByVal direction As Double) As Double
    On Error GoTo ErrorHandler
    If direction > 0 Then
        Do While delta <= 0                       : delta = delta + 2# * Application.PI : Loop
        Do While delta > 2# * Application.PI     : delta = delta - 2# * Application.PI : Loop
    Else
        Do While delta >= 0                       : delta = delta - 2# * Application.PI : Loop
        Do While delta < -2# * Application.PI    : delta = delta + 2# * Application.PI : Loop
    End If
    NormalizeAngle = delta
    Exit Function

ErrorHandler:
    NormalizeAngle = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Geometry.NormalizeAngle"
End Function
