' Module: Length
' Description: This module provides functions to calculate lengths of elements in MicroStation with silent error handling.
' The module includes functions to determine the length of various element types, handle rounding logic,
' and manage configuration variables for rounding.
' NEVER USE Rnd = 255 in GetLength function ! it reserved for error

' Dependencies: Config module

Const ARES_RND_VAR As String = "ARES_RND"
Const ARES_RND_DEFAULT As Byte = 1
Const ARES_RND_ERROR_VALUE As Byte = 255

Option Explicit

' Public function to get the length of an element
Public Function GetLength(ByVal El As Element, Optional Rnd As Variant, Optional RndLength As Boolean = True, Optional ErasRnd As Boolean = False) As Double
    On Error GoTo ErrorHandler

    ' Determine the length based on the element type
    GetLength = GetElementLength(El)

    ' Handle rounding if required
    If RndLength Then
        Rnd = HandleRounding(Rnd, ErasRnd)
        If Rnd = ARES_RND_ERROR_VALUE Then
            ShowStatus "Valeur d'arrondi interdit: " & ARES_RND_ERROR_VALUE
            GetLength = 0
            Exit Function
        End If
        GetLength = RoundedLength(GetLength, CByte(Rnd))
    ElseIf ErasRnd Then
        If Not HandleRoundingForErase(Rnd) Then
            GetLength = 0
            Exit Function
        End If
    End If

    Exit Function

ErrorHandler:
    GetLength = 0
End Function

' Private function to get the length of an element based on its type
Private Function GetElementLength(ByVal El As Element) As Double
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
        Case Else
            GetElementLength = 0
    End Select

    Exit Function

ErrorHandler:
    GetElementLength = 0
End Function

' Private function to handle rounding logic
Private Function HandleRounding(Optional Rnd As Variant, Optional ErasRnd As Boolean) As Variant
    On Error GoTo ErrorHandler
    
    ' Handle missing rounding value
    If IsMissing(Rnd) Then
        Rnd = GetRoundValue()
    ElseIf ErasRnd And (VarType(Rnd) = vbByte Or VarType(Rnd) = vbInteger) Then
        ' Set rounding value if erase rounding is true
        If Not SetRound(CByte(Rnd)) Then
            HandleRounding = ARES_RND_ERROR_VALUE
            Exit Function
        End If
    End If
    HandleRounding = Rnd
    
    Exit Function

ErrorHandler:
    HandleRounding = ARES_RND_ERROR_VALUE
End Function

' Private function to handle rounding logic for erase
Private Function HandleRoundingForErase(Optional Rnd As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    HandleRoundingForErase = False
    
    ' Set default rounding if Rnd is missing
    If IsMissing(Rnd) Then
        If Not SetRound(ARES_RND_DEFAULT) Then Exit Function
        ShowStatus ARES_RND_VAR & " défini à " & ARES_RND_DEFAULT & " par défaut"
    ElseIf VarType(Rnd) = vbByte Or VarType(Rnd) = vbInteger Then
        ' Set rounding value if Rnd is provided
        If Not SetRound(CByte(Rnd)) Then Exit Function
    End If
    
    HandleRoundingForErase = True
    Exit Function

ErrorHandler:
    HandleRoundingForErase = False
End Function

' Private function to get the rounding value
Private Function GetRoundValue() As Variant
    On Error GoTo ErrorHandler

    Dim roundValue As String
    roundValue = GetRound()
    
    ' Handle empty rounding value
    If roundValue = "" Then
        If Not SetRound(ARES_RND_DEFAULT) Then Exit Function
        roundValue = GetRound()
        If roundValue = "" Then Exit Function
        ShowStatus ARES_RND_VAR & " défini à " & ARES_RND_DEFAULT & " par défaut"
    End If
    GetRoundValue = CByte(roundValue)

    Exit Function

ErrorHandler:
    GetRoundValue = ARES_RND_ERROR_VALUE
End Function

' Private function to get the rounding configuration variable
Private Function GetRound() As String
    On Error GoTo ErrorHandler
    
    GetRound = Config.GetVar(ARES_RND_VAR)
    Exit Function
    
ErrorHandler:
    GetRound = ""
End Function

' Private function to set the rounding configuration variable
Private Function SetRound(Rnd As Byte) As Boolean
    On Error GoTo ErrorHandler
    
    SetRound = Config.SetVar(ARES_RND_VAR, Rnd)
    Exit Function
    
ErrorHandler:
    SetRound = False
End Function

' Private function to calculate the length of a complex shape element
Private Function LengthComplexShape(ByVal El As ComplexShapeElement) As Double
    On Error GoTo ErrorHandler

    Dim ELEnum As ElementEnumerator
    Dim SubEl As Element
    Dim i As Long

    LengthComplexShape = El.Perimeter
    Set ELEnum = El.GetSubElements

    ' Iterate through sub-elements
    For i = 0 To UBound(ELEnum.BuildArrayFromContents)
        ELEnum.MoveNext
    Next

    Set SubEl = ELEnum.Current
    If SubEl.IsLineElement Then
        LengthComplexShape = (LengthComplexShape - (SubEl.AsLineElement.Length * 2)) / 2
    Else
        LengthComplexShape = 0
    End If

    Exit Function

ErrorHandler:
    LengthComplexShape = 0
End Function

' Private function to round the length to a specified number of decimal places
Private Function RoundedLength(Length As Double, Rnd As Byte) As Double
    On Error GoTo ErrorHandler
    
    RoundedLength = Round(Length, Rnd)
    Exit Function
    
ErrorHandler:
    RoundedLength = 0
End Function
