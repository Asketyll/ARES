' Module: Length
' Description: This module provides functions to calculate lengths of elements in MicroStation with silent error handling.
' The module includes functions to determine the length of various element types, handle rounding logic,
' and manage configuration variables for rounding.
' NEVER USE Rnd = 255 in GetLength function! It is reserved for errors.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_VAR, LangManager

Option Explicit

' Public function to get the length of an element
Public Function GetLength(ByVal el As Element, Optional RND As Variant, Optional RndLength As Boolean = True, Optional ErasRnd As Boolean = False) As Double
    On Error GoTo ErrorHandler

    ' Determine the length based on the element type
    GetLength = GetElementLength(el)

    ' Handle rounding if required
    If RndLength Then
        RND = HandleRounding(RND, ErasRnd)
        If RND = ARES_VAR.ARES_RND_ERROR_VALUE Then
            ShowStatus GetTranslation("LengthRoundError") & ARES_VAR.ARES_RND_ERROR_VALUE
            GetLength = 0
            Exit Function
        End If
        GetLength = RoundedLength(GetLength, CByte(RND))
    ElseIf ErasRnd Then
        If Not HandleRoundingForErase(RND) Then
            GetLength = 0
            Exit Function
        End If
    End If

    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    GetLength = 0
End Function

' Private function to get the length of an element based on its type
Private Function GetElementLength(ByVal el As Element) As Double
    On Error GoTo ErrorHandler

    ' Determine the length based on the element type
    Select Case True
        Case el.IsComplexStringElement
            GetElementLength = el.AsComplexStringElement.Length
        Case el.IsComplexShapeElement
            GetElementLength = LengthComplexShape(el)
        Case el.IsLineElement
            GetElementLength = el.AsLineElement.Length
        Case el.IsArcElement
            GetElementLength = el.AsArcElement.Length
        Case el.IsShapeElement
            GetElementLength = LengthShape(el)
        Case Else
            GetElementLength = 0
            ShowStatus GetTranslation("LengthElementTypeNotSupportedByInterface", DLongToString(el.ID), el.Type)
    End Select
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    GetElementLength = 0
End Function

' Private function to handle rounding logic
Private Function HandleRounding(Optional RND As Variant, Optional ErasRnd As Boolean) As Variant
    On Error GoTo ErrorHandler

    ' Handle missing rounding value
    If IsMissing(RND) Then
        RND = GetRoundValue()
    ElseIf ErasRnd And (VarType(RND) = vbByte Or VarType(RND) = vbInteger) Then
        ' Set rounding value if erase rounding is true
        If Not SetRound(CByte(RND)) Then
            HandleRounding = ARES_VAR.ARES_RND_ERROR_VALUE
            Exit Function
        End If
    End If

    HandleRounding = RND

    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    HandleRounding = ARES_VAR.ARES_RND_ERROR_VALUE
End Function

' Private function to handle rounding logic for erase
Private Function HandleRoundingForErase(Optional RND As Variant) As Boolean
    On Error GoTo ErrorHandler

    HandleRoundingForErase = False

    ' Set default rounding if Rnd is missing
    If IsMissing(RND) Then
        If Not ResetRound() Then Exit Function
    ElseIf VarType(RND) = vbByte Or VarType(RND) = vbInteger Then
        ' Set rounding value if Rnd is provided
        If Not SetRound(CByte(RND)) Then Exit Function
    End If

    HandleRoundingForErase = True

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    HandleRoundingForErase = False
End Function

' Private function to calculate the length of a complex shape element
Private Function LengthComplexShape(ByVal el As ComplexShapeElement) As Double
    On Error GoTo ErrorHandler

    Dim ElEnum As ElementEnumerator
    Dim SubEl As Element
    Dim i As Long

    LengthComplexShape = el.Perimeter
    Set ElEnum = el.GetSubElements

    ' Iterate through sub-elements
    For i = 0 To UBound(ElEnum.BuildArrayFromContents)
        ElEnum.MoveNext
    Next

    Set SubEl = ElEnum.Current

    If SubEl.IsLineElement Then
        LengthComplexShape = (LengthComplexShape - (SubEl.AsLineElement.Length * 2)) / 2
    Else
        LengthComplexShape = 0
    End If

    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthComplexShape = 0
End Function
' Private function to calculate the length of a shape element
Private Function LengthShape(ByVal el As ShapeElement) As Double
    On Error GoTo ErrorHandler

    Dim ElEnum As ElementEnumerator
    Dim SubEl As Element
    Dim i As Long
    
    LengthShape = el.Perimeter
    Set ElEnum = el.GetSubElements

    ' Iterate through sub-elements
    For i = 0 To UBound(ElEnum.BuildArrayFromContents)
        ElEnum.MoveNext
    Next

    Set SubEl = ElEnum.Current

    If SubEl.IsLineElement Then
        LengthShape = (LengthShape - (SubEl.AsLineElement.Length * 2)) / 2
    Else
        LengthShape = 0
    End If

    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthShape = 0
End Function

' Private function to round the length to a specified number of decimal places
Private Function RoundedLength(Length As Double, RND As Byte) As Double
    On Error GoTo ErrorHandler

    RoundedLength = Round(Length, RND)

    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    RoundedLength = 0
End Function

' Private function to get the rounding value
Private Function GetRoundValue() As Variant
    On Error GoTo ErrorHandler

    Dim roundValue As String
    roundValue = ARES_VAR.ARES_ROUNDS.Value

    ' Handle empty rounding value
    If roundValue = "" Then
        ResetRound
    End If

    GetRoundValue = CByte(roundValue)

    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    GetRoundValue = ARES_VAR.ARES_RND_ERROR_VALUE
End Function

' Public function to set the rounding configuration variable
Public Function SetRound(RND As Byte) As Boolean
    On Error GoTo ErrorHandler

    If RND <> ARES_VAR.ARES_RND_ERROR_VALUE Then
        SetRound = Config.SetVar(ARES_VAR.ARES_ROUNDS.key, RND)
    Else
        ShowStatus GetTranslation("LengthRoundError") & ARES_VAR.ARES_RND_ERROR_VALUE
    End If

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    SetRound = False
End Function

' Public function to reset the rounding configuration variable
Public Function ResetRound() As Boolean
    On Error GoTo ErrorHandler

    ARES_VAR.ResetMSVar ARES_VAR.ARES_ROUNDS
    ResetRound = (ARES_VAR.ARES_ROUNDS.Value = ARES_VAR.ARES_ROUNDS.Default)

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    ResetRound = False
End Function
