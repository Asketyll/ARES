' Module: Length
' Description: This module provides functions to calculate lengths of elements in MicroStation with silent error handling.
' The module includes functions to determine the length of various element types, handle rounding logic,
' and manage configuration variables for rounding.
' NEVER USE Rnd = 255 in GetLength function ! it reserved for error

' Dependencies: Config , ARES_VAR

Option Explicit

' Public function to get the length of an element
Public Function GetLength(ByVal El As Element, Optional RND As Variant, Optional RndLength As Boolean = True, Optional ErasRnd As Boolean = False) As Double
    On Error GoTo ErrorHandler

    ' Determine the length based on the element type
    GetLength = GetElementLength(El)

    ' Handle rounding if required
    If RndLength Then
        RND = HandleRounding(RND, ErasRnd)
        If RND = ARES_VAR.ARES_RND_ERROR_VALUE Then
            ShowStatus "Valeur d'arrondi interdit: " & ARES_VAR.ARES_RND_ERROR_VALUE
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
    HandleRounding = ARES_RND_ERROR_VALUE
End Function

' Private function to handle rounding logic for erase
Private Function HandleRoundingForErase(Optional RND As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    HandleRoundingForErase = False
    
    ' Set default rounding if Rnd is missing
    If IsMissing(RND) Then
        If Not SetRound(ARES_RND_DEFAULT) Then Exit Function
        ShowStatus ARES_VAR.ARES_ROUNDS & " défini à " & ARES_RND_DEFAULT & " par défaut"
    ElseIf VarType(RND) = vbByte Or VarType(RND) = vbInteger Then
        ' Set rounding value if Rnd is provided
        If Not SetRound(CByte(RND)) Then Exit Function
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
        ShowStatus ARES_VAR.ARES_ROUNDS & " défini à " & ARES_RND_DEFAULT & " par défaut"
    End If
    GetRoundValue = CByte(roundValue)

    Exit Function

ErrorHandler:
    GetRoundValue = ARES_RND_ERROR_VALUE
End Function

' Private function to get the rounding configuration variable
Private Function GetRound() As String
    On Error GoTo ErrorHandler
    
    GetRound = Config.GetVar(ARES_VAR.Round)
    Exit Function
    
ErrorHandler:
    GetRound = ""
End Function

' Private function to set the rounding configuration variable
Private Function SetRound(RND As Byte) As Boolean
    On Error GoTo ErrorHandler
    
    SetRound = Config.SetVar(ARES_VAR.ARES_ROUNDS, RND)
    Exit Function
    
ErrorHandler:
    SetRound = False
End Function

' Private function to calculate the length of a complex shape element
Private Function LengthComplexShape(ByVal El As ComplexShapeElement) As Double
    On Error GoTo ErrorHandler

    Dim ElEnum As elementEnumerator
    Dim SubEl As Element
    Dim i As Long

    LengthComplexShape = El.Perimeter
    Set ElEnum = El.GetSubElements

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
    LengthComplexShape = 0
End Function

' Private function to round the length to a specified number of decimal places
Private Function RoundedLength(Length As Double, RND As Byte) As Double
    On Error GoTo ErrorHandler
    
    RoundedLength = Round(Length, RND)
    Exit Function
    
ErrorHandler:
    RoundedLength = 0
End Function
