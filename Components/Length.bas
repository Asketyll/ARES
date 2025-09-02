' Module: Length
' Description: This module provides functions to calculate lengths of elements in MicroStation with silent error handling.
' The module includes functions to determine the length of various element types, handle rounding logic,
' NEVER USE Rnd = 255 in GetLength function! It is reserved for errors.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARESConfigClass, ARESConstants, LangManager, ErrorHandlerClass
Option Explicit

' Public function to get the length of an element
Public Function GetLength(ByVal El As element, Optional RND As Variant, Optional RndLength As Boolean = True, Optional ErasRnd As Boolean = False) As Double
    On Error GoTo ErrorHandler
    ' Determine the length based on the element type
    GetLength = GetElementLength(El)
    ' Handle rounding if required
    If RndLength Then
        RND = HandleRounding(RND, ErasRnd)
        If RND = ARES_RND_ERROR_VALUE Then
            ShowStatus GetTranslation("LengthRoundError") & ARES_RND_ERROR_VALUE
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLength"
End Function

' Private function to get the length of an element based on its type
Private Function GetElementLength(ByVal El As element) As Double
    On Error GoTo ErrorHandler
    ' Determine the length based on the element type
    Select Case True
        Case El.IsComplexStringElement
            GetElementLength = El.AsComplexStringElement.Length
        Case El.IsComplexShapeElement
            GetElementLength = LengthComplexShape(El)
        Case El.IsLineElement
            GetElementLength = El.AsLineElement.Length
        Case El.IsArcElement
            GetElementLength = El.AsArcElement.Length
        Case El.IsShapeElement
            GetElementLength = LengthShape(El)
        Case Else
            GetElementLength = 0
            ShowStatus GetTranslation("LengthElementTypeNotSupportedByInterface", DLongToString(El.ID), El.Type)
    End Select
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    GetElementLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetElementLength"
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
            HandleRounding = ARES_RND_ERROR_VALUE
            Exit Function
        End If
    End If
    HandleRounding = RND
    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    HandleRounding = ARES_RND_ERROR_VALUE
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.HandleRounding"
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.HandleRoundingForErase"
End Function

' Private function to calculate the length of a complex shape element
Private Function LengthComplexShape(ByVal El As ComplexShapeElement) As Double
    On Error GoTo ErrorHandler
    Dim ELEnum As ElementEnumerator
    Dim subel As element
    Dim i As Long
    LengthComplexShape = El.Perimeter
    Set ELEnum = El.GetSubElements
    ' Iterate through sub-elements
    For i = 0 To UBound(ELEnum.BuildArrayFromContents)
        ELEnum.MoveNext
    Next
    Set subel = ELEnum.Current
    If subel.IsLineElement Then
        LengthComplexShape = (LengthComplexShape - (subel.AsLineElement.Length * 2)) / 2
    Else
        LengthComplexShape = 0
    End If
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthComplexShape"
End Function
' Private function to calculate the length of a shape element
Private Function LengthShape(ByVal El As ShapeElement) As Double
    On Error GoTo ErrorHandler
    Dim ELEnum As ElementEnumerator
    Dim subel As element
    Dim i As Long
    
    LengthShape = El.Perimeter
    Set ELEnum = El.GetSubElements
    ' Iterate through sub-elements
    For i = 0 To UBound(ELEnum.BuildArrayFromContents)
        ELEnum.MoveNext
    Next
    Set subel = ELEnum.Current
    If subel.IsLineElement Then
        LengthShape = (LengthShape - (subel.AsLineElement.Length * 2)) / 2
    Else
        LengthShape = 0
    End If
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthShape"
End Function

' Private function to round the length to a specified number of decimal places
Private Function RoundedLength(Length As Double, RND As Byte) As Double
    On Error GoTo ErrorHandler
    RoundedLength = Round(Length, RND)
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    RoundedLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.RoundedLength"
End Function

' Private function to get the rounding value
Private Function GetRoundValue() As Variant
    On Error GoTo ErrorHandler
    Dim roundValue As String
    roundValue = ARESConfig.ARES_ROUNDS.Value
    ' Handle empty rounding value
    If roundValue = "" Then
        ResetRound
    End If
    GetRoundValue = CByte(roundValue)
    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    GetRoundValue = ARES_RND_ERROR_VALUE
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetRoundValue"
End Function

' Public function to set the rounding configuration variable
Public Function SetRound(RND As Byte) As Boolean
    SetRound = False
    On Error GoTo ErrorHandler

    If RND <> ARES_VAR.ARES_RND_ERROR_VALUE Then
        SetRound = Config.SetVar(ARES_VAR.ARES_ROUNDS.Key, RND)
    Else
        ShowStatus GetTranslation("LengthRoundError") & ARES_VAR.ARES_RND_ERROR_VALUE
    End If
End Function