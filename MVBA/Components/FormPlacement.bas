' Module: FormPlacement
' Description: Remembers each ARES option form's on-screen position across sessions. The position is
'              stored as a percentage of the virtual desktop in the ARES_Form_Layout config var (one
'              entry per form), restored on open, kept grabbable (the caption stays reachable) and
'              multi-monitor aware via MonitorFromRect. When nothing is stored the form is centered on
'              the primary monitor. DPI is read live so 125/150% users are handled (mixed-DPI is
'              approximate but the grabbable guard still keeps the form reachable).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass (ARES_FORM_LAYOUT), ARESConstants, ErrorHandlerClass
Option Explicit

' === Win32 ===
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpmi As MONITORINFO) As Long
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79
Private Const SPI_GETWORKAREA As Long = 48
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const POINTS_PER_INCH As Double = 72
Private Const MIN_VISIBLE_PX As Long = 90        ' min horizontal strip of the form kept on-screen
Private Const MIN_TITLE_PX As Long = 30          ' min height keeping the caption/title bar grabbable
Private Const ENTRY_DELIM As String = "="        ' one entry is <FormName>=<xPct>:<yPct>
Private Const COORD_DELIM As String = ":"

' Restore a form to its stored position, or center it on the primary monitor if none is stored.
' Call from UserForm_Initialize: RestoreFormPosition Me, Me.Name
Public Sub RestoreFormPosition(ByVal oForm As Object, ByVal sKey As String)
    On Error GoTo ErrorHandler

    ' The form must be free-positioned for Left/Top to take effect (belt-and-suspenders; the .frm
    ' header is also set to Manual). Ignore the error if the property is read-only at this point.
    On Error Resume Next
    oForm.StartUpPosition = 0
    On Error GoTo ErrorHandler

    Dim xPct As Double, yPct As Double
    If Not LayoutGet(sKey, xPct, yPct) Then
        CenterForm oForm
        Exit Sub
    End If

    Dim vsX As Long, vsY As Long, vsCx As Long, vsCy As Long
    vsX = GetSystemMetrics(SM_XVIRTUALSCREEN)
    vsY = GetSystemMetrics(SM_YVIRTUALSCREEN)
    vsCx = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vsCy = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    If vsCx <= 0 Or vsCy <= 0 Then
        CenterForm oForm
        Exit Sub
    End If

    Dim dpiX As Long, dpiY As Long
    ScreenDPI dpiX, dpiY

    Dim formW_px As Long, formH_px As Long
    formW_px = PtToPx(oForm.Width, dpiX)
    formH_px = PtToPx(oForm.Height, dpiY)

    Dim L_px As Long, T_px As Long
    L_px = vsX + CLng(xPct * vsCx)
    T_px = vsY + CLng(yPct * vsCy)

    ' Work area of the monitor the form would land on (nearest if that spot no longer exists).
    Dim rcWork As RECT
    If Not WorkAreaForRect(L_px, T_px, formW_px, formH_px, rcWork) Then
        If Not PrimaryWorkArea(rcWork) Then
            CenterForm oForm
            Exit Sub
        End If
    End If

    ' Keep the caption grabbable; allow the rest of the form to overhang the screen.
    L_px = ClampRange(L_px, rcWork.Left + MIN_VISIBLE_PX - formW_px, rcWork.Right - MIN_VISIBLE_PX)
    T_px = ClampRange(T_px, rcWork.Top, rcWork.Bottom - MIN_TITLE_PX)

    oForm.Left = PxToPt(L_px, dpiX)
    oForm.Top = PxToPt(T_px, dpiY)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.RestoreFormPosition"
End Sub

' Persist a form's current position as a percentage of the virtual desktop.
' Call from UserForm_QueryClose (real-close path): SaveFormPosition Me, Me.Name
Public Sub SaveFormPosition(ByVal oForm As Object, ByVal sKey As String)
    On Error GoTo ErrorHandler

    Dim vsX As Long, vsY As Long, vsCx As Long, vsCy As Long
    vsX = GetSystemMetrics(SM_XVIRTUALSCREEN)
    vsY = GetSystemMetrics(SM_YVIRTUALSCREEN)
    vsCx = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vsCy = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    If vsCx <= 0 Or vsCy <= 0 Then Exit Sub

    Dim dpiX As Long, dpiY As Long
    ScreenDPI dpiX, dpiY

    Dim xPct As Double, yPct As Double
    xPct = (PtToPx(oForm.Left, dpiX) - vsX) / vsCx
    yPct = (PtToPx(oForm.Top, dpiY) - vsY) / vsCy

    LayoutSet sKey, xPct, yPct
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.SaveFormPosition"
End Sub

' Center a form on the primary monitor's work area.
Public Sub CenterForm(ByVal oForm As Object)
    On Error GoTo ErrorHandler

    On Error Resume Next
    oForm.StartUpPosition = 0
    On Error GoTo ErrorHandler

    Dim rcWork As RECT
    If Not PrimaryWorkArea(rcWork) Then Exit Sub

    Dim dpiX As Long, dpiY As Long
    ScreenDPI dpiX, dpiY

    Dim formW_px As Long, formH_px As Long
    formW_px = PtToPx(oForm.Width, dpiX)
    formH_px = PtToPx(oForm.Height, dpiY)

    Dim L_px As Long, T_px As Long
    L_px = rcWork.Left + ((rcWork.Right - rcWork.Left) - formW_px) \ 2
    T_px = rcWork.Top + ((rcWork.Bottom - rcWork.Top) - formH_px) \ 2
    If L_px < rcWork.Left Then L_px = rcWork.Left
    If T_px < rcWork.Top Then T_px = rcWork.Top

    oForm.Left = PxToPt(L_px, dpiX)
    oForm.Top = PxToPt(T_px, dpiY)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.CenterForm"
End Sub

' Forget every stored form position (next open of each form will center it).
Public Sub ClearFormPositions()
    On Error GoTo ErrorHandler
    If Not (ARESConfig.ARES_FORM_LAYOUT Is Nothing) Then ARESConfig.ARES_FORM_LAYOUT.Value = ""
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.ClearFormPositions"
End Sub

' === PRIVATE HELPERS ===

' Read the stored xPct/yPct for sKey. Returns False when there is no valid entry.
Private Function LayoutGet(ByVal sKey As String, ByRef xPct As Double, ByRef yPct As Double) As Boolean
    On Error GoTo ErrorHandler
    LayoutGet = False
    If ARESConfig.ARES_FORM_LAYOUT Is Nothing Then Exit Function

    Dim sAll As String
    sAll = ARESConfig.ARES_FORM_LAYOUT.Value
    If Len(sAll) = 0 Then Exit Function

    Dim entries() As String, i As Long, kv() As String, xy() As String
    entries = Split(sAll, ARESConstants.ARES_VAR_DELIMITER)
    For i = LBound(entries) To UBound(entries)
        kv = Split(entries(i), ENTRY_DELIM)
        If UBound(kv) = 1 Then
            If kv(0) = sKey Then
                xy = Split(kv(1), COORD_DELIM)
                If UBound(xy) = 1 Then
                    xPct = Val(Replace(xy(0), ",", "."))
                    yPct = Val(Replace(xy(1), ",", "."))
                    LayoutGet = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.LayoutGet"
    LayoutGet = False
End Function

' Upsert the entry for sKey, preserving the other forms' entries.
Private Sub LayoutSet(ByVal sKey As String, ByVal xPct As Double, ByVal yPct As Double)
    On Error GoTo ErrorHandler
    If ARESConfig.ARES_FORM_LAYOUT Is Nothing Then Exit Sub

    Dim sAll As String, sOut As String
    sAll = ARESConfig.ARES_FORM_LAYOUT.Value

    Dim entries() As String, i As Long, kv() As String
    If Len(sAll) > 0 Then
        entries = Split(sAll, ARESConstants.ARES_VAR_DELIMITER)
        For i = LBound(entries) To UBound(entries)
            If Len(entries(i)) > 0 Then
                kv = Split(entries(i), ENTRY_DELIM)
                If kv(0) <> sKey Then
                    If Len(sOut) > 0 Then sOut = sOut & ARESConstants.ARES_VAR_DELIMITER
                    sOut = sOut & entries(i)
                End If
            End If
        Next i
    End If

    If Len(sOut) > 0 Then sOut = sOut & ARESConstants.ARES_VAR_DELIMITER
    sOut = sOut & sKey & ENTRY_DELIM & FmtPct(xPct) & COORD_DELIM & FmtPct(yPct)

    ARESConfig.ARES_FORM_LAYOUT.Value = sOut
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormPlacement.LayoutSet"
End Sub

' Format a percentage with a dot decimal separator, locale-independent.
Private Function FmtPct(ByVal d As Double) As String
    FmtPct = Replace(Format(d, "0.#####"), ",", ".")
End Function

' Fill rcWork with the work area of the monitor nearest the given rectangle (pixels). False on failure.
Private Function WorkAreaForRect(ByVal L As Long, ByVal T As Long, ByVal W As Long, ByVal H As Long, ByRef rcWork As RECT) As Boolean
    On Error GoTo ErrorHandler
    WorkAreaForRect = False

    Dim rc As RECT
    rc.Left = L
    rc.Top = T
    rc.Right = L + W
    rc.Bottom = T + H

    Dim hMon As LongPtr
    hMon = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    If hMon = 0 Then Exit Function

    Dim mi As MONITORINFO
    mi.cbSize = Len(mi)
    If GetMonitorInfo(hMon, mi) = 0 Then Exit Function

    rcWork = mi.rcWork
    WorkAreaForRect = True
    Exit Function

ErrorHandler:
    WorkAreaForRect = False
End Function

' Fill rcWork with the primary monitor work area (pixels). False on failure.
Private Function PrimaryWorkArea(ByRef rcWork As RECT) As Boolean
    On Error GoTo ErrorHandler
    PrimaryWorkArea = False

    Dim rc As RECT
    If SystemParametersInfo(SPI_GETWORKAREA, 0, rc, 0) <> 0 Then
        rcWork = rc
        PrimaryWorkArea = True
    Else
        rc.Left = 0
        rc.Top = 0
        rc.Right = GetSystemMetrics(SM_CXSCREEN)
        rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
        rcWork = rc
        PrimaryWorkArea = (rc.Right > 0)
    End If
    Exit Function

ErrorHandler:
    PrimaryWorkArea = False
End Function

' Read the screen DPI (defaults to 96 on any failure).
Private Sub ScreenDPI(ByRef dpiX As Long, ByRef dpiY As Long)
    On Error GoTo ErrorHandler
    dpiX = 96
    dpiY = 96

    Dim hDC As LongPtr
    hDC = GetDC(0)
    If hDC <> 0 Then
        dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
        dpiY = GetDeviceCaps(hDC, LOGPIXELSY)
        ReleaseDC 0, hDC
    End If
    If dpiX <= 0 Then dpiX = 96
    If dpiY <= 0 Then dpiY = 96
    Exit Sub

ErrorHandler:
    dpiX = 96
    dpiY = 96
End Sub

Private Function PtToPx(ByVal pts As Double, ByVal dpi As Long) As Long
    PtToPx = CLng(pts * dpi / POINTS_PER_INCH)
End Function

Private Function PxToPt(ByVal px As Long, ByVal dpi As Long) As Double
    PxToPt = px * POINTS_PER_INCH / dpi
End Function

Private Function ClampRange(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If lo > hi Then
        ClampRange = lo
    ElseIf v < lo Then
        ClampRange = lo
    ElseIf v > hi Then
        ClampRange = hi
    Else
        ClampRange = v
    End If
End Function
