' Module: ColorDialog
' Description: Windows Color Dialog wrapper with MicroStation color index support.
'              Exposes PickMsColorIndex and ApplyColorToButton for reuse across ARES UserForms.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: comdlg32.dll, MicroStation MVBA object model (ColorTable, ModelReference)
Option Explicit

#If VBA7 Then
' 64-bit — explicit padding required to match native x64 struct alignment
Private Type CHOOSECOLOR
    lStructSize    As Long      ' 4 bytes
    pad1           As Long      ' 4 bytes padding (align next pointer to 8)
    hwndOwner      As LongPtr   ' 8 bytes
    hInstance      As LongPtr   ' 8 bytes
    rgbResult      As Long      ' 4 bytes
    pad2           As Long      ' 4 bytes padding
    lpCustColors   As LongPtr   ' 8 bytes (pointer to array of 16 LONGs)
    Flags          As Long      ' 4 bytes
    pad3           As Long      ' 4 bytes padding
    lCustData      As LongPtr   ' 8 bytes
    lpfnHook       As LongPtr   ' 8 bytes (null — no hook)
    lpTemplateName As LongPtr   ' 8 bytes (null — no template)
End Type

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (pChoosecolor As CHOOSECOLOR) As Long
#Else
Private Type CHOOSECOLOR
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (pChoosecolor As CHOOSECOLOR) As Long
#End If

Private Const CC_RGBINIT  As Long = &H1
Private Const CC_FULLOPEN As Long = &H2

' Persists custom colors across dialog calls within the same session.
Private mCustColors(0 To 15) As Long

' ============================================================
' PUBLIC API
' ============================================================

' Opens the Windows Color Dialog initialised with the current MicroStation color index.
' Returns the nearest MicroStation color index for the selected color,
' or -1 if the user cancelled.
Public Function PickMsColorIndex(ByVal msColorIndex As Long) As Long
    On Error GoTo ErrorHandler

    Dim selectedRgb As Long
    selectedRgb = SelectColorDialog(GetRgbForIndex(msColorIndex))

    If selectedRgb = -1 Then
        PickMsColorIndex = -1
        Exit Function
    End If

    PickMsColorIndex = RgbToColorIndex(selectedRgb)
    Exit Function

ErrorHandler:
    PickMsColorIndex = -1
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ColorDialog.PickMsColorIndex"
End Function

' Sets a TextBox's BackColor, ForeColor and Text to reflect a MicroStation color index.
' The TextBox should be Locked=True so the user cannot type in it directly.
Public Sub ApplyColorToTextBox(ByRef txt As MSForms.TextBox, ByVal colorIdx As Long)
    On Error GoTo ErrorHandler

    Dim rgb As Long
    rgb = GetRgbForIndex(colorIdx)
    txt.Text = CStr(colorIdx)
    If rgb >= 0 Then
        txt.BackColor = rgb
        txt.ForeColor = ContrastColor(rgb)
    End If
    Exit Sub

ErrorHandler:
    txt.Text = CStr(colorIdx)
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ColorDialog.ApplyColorToTextBox"
End Sub

' Sets a CommandButton's BackColor, ForeColor and Caption to reflect a MicroStation color index.
Public Sub ApplyColorToButton(ByRef btn As MSForms.CommandButton, ByVal colorIdx As Long)
    On Error GoTo ErrorHandler

    Dim rgb As Long
    rgb = GetRgbForIndex(colorIdx)
    btn.Caption = CStr(colorIdx)
    If rgb >= 0 Then
        btn.BackColor = rgb
        btn.ForeColor = ContrastColor(rgb)
    End If
    Exit Sub

ErrorHandler:
    btn.Caption = CStr(colorIdx)
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ColorDialog.ApplyColorToButton"
End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================

' Returns white or black depending on which contrasts better against the given RGB background.
Private Function ContrastColor(ByVal rgb As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = rgb And &HFF
    g = (rgb \ &H100) And &HFF
    b = (rgb \ &H10000) And &HFF
    Dim luminance As Double
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    ContrastColor = IIf(luminance < 128, VBA.RGB(255, 255, 255), VBA.RGB(0, 0, 0))
End Function

' Returns the RGB Long for a MicroStation color index using the active DGN color table.
' The returned Long packs colors as R + G*256 + B*65536 (same as VBA.RGB and OLE BackColor).
' Returns -1 on failure.
Private Function GetRgbForIndex(ByVal colorIdx As Long) As Long
    On Error GoTo Fallback
    Dim ct As ColorTable
    Set ct = Application.ActiveDesignFile.ExtractColorTable
    GetRgbForIndex = ct.GetColorAtIndex(colorIdx)
    Exit Function

Fallback:
    GetRgbForIndex = -1
End Function

' Returns the nearest MicroStation color index (0-254) for a given RGB Long.
' Uses InternalColorFromRGBColor: the low byte of the returned internal color is the index.
Private Function RgbToColorIndex(ByVal rgbValue As Long) As Long
    On Error GoTo Fallback
    RgbToColorIndex = ActiveModelReference.InternalColorFromRGBColor(rgbValue) And &HFF
    Exit Function

Fallback:
    RgbToColorIndex = 0
End Function

' Opens the Windows Color Dialog initialised with initialRgb.
' Returns the selected COLORREF Long, or -1 if the user cancelled.
Private Function SelectColorDialog(ByVal initialRgb As Long) As Long
    On Error GoTo ErrorHandler

    If initialRgb < 0 Then initialRgb = 0

    Dim cc As CHOOSECOLOR
    cc.lStructSize  = LenB(cc)
    cc.hwndOwner    = 0
    cc.rgbResult    = initialRgb
    cc.lpCustColors = VarPtr(mCustColors(0))
    cc.Flags        = CC_RGBINIT Or CC_FULLOPEN

    If ChooseColor(cc) <> 0 Then
        SelectColorDialog = cc.rgbResult
    Else
        SelectColorDialog = -1
    End If
    Exit Function

ErrorHandler:
    SelectColorDialog = -1
End Function
