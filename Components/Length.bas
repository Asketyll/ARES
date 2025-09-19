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
            GetElementLength = LengthComplexShape(El, True)
        Case El.IsLineElement
            GetElementLength = El.AsLineElement.Length
        Case El.IsArcElement
            GetElementLength = El.AsArcElement.Length
        Case El.IsShapeElement
            GetElementLength = LengthShape(El, True)
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
Private Function LengthComplexShape(ByVal El As ComplexShapeElement, Optional ByVal LongestSideOnly As Boolean = False) As Double
    On Error GoTo ErrorHandler
    
    If LongestSideOnly Then
        ' Find the longest sub-element
        LengthComplexShape = GetLongestSideFromComplexShape(El)
    Else
        ' Return the perimeter (default behavior)
        LengthComplexShape = El.Perimeter
    End If
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthComplexShape"
End Function

' Private function to calculate the length of a shape element
Private Function LengthShape(ByVal El As ShapeElement, Optional ByVal LongestSideOnly As Boolean = False) As Double
    On Error GoTo ErrorHandler
    
    If LongestSideOnly Then
        ' Calculate the length of the longest side
        LengthShape = GetLongestSideFromShape(El)
    Else
        ' Return the perimeter (default behavior)
        LengthShape = El.Perimeter
    End If
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthShape"
End Function

' Helper function to get the longest side from a simple shape
Private Function GetLongestSideFromShape(ByVal El As ShapeElement) As Double
    On Error GoTo ErrorHandler
    
    Dim Vertices() As Point3d
    Dim i As Long
    Dim SideLength As Double
    Dim LongestSide As Double
    
    LongestSide = 0
    Vertices = El.GetVertices()
    
    ' Calculate length of each side and find the longest
    For i = 0 To UBound(Vertices) - 1
        SideLength = Point3dDistance(Vertices(i), Vertices(i + 1))
        If SideLength > LongestSide Then
            LongestSide = SideLength
        End If
    Next i
    
    GetLongestSideFromShape = LongestSide
    Exit Function

ErrorHandler:
    GetLongestSideFromShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLongestSideFromShape"
End Function

' Helper function to get the longest side from a complex shape
Private Function GetLongestSideFromComplexShape(ByVal El As ComplexShapeElement) As Double
    On Error GoTo ErrorHandler
    
    Dim ELEnum As ElementEnumerator
    Dim subel As element
    Dim ElementLength As Double
    Dim LongestSide As Double
    
    LongestSide = 0
    Set ELEnum = El.GetSubElements
    
    ' Iterate through sub-elements and find the longest one
    Do While ELEnum.MoveNext
        Set subel = ELEnum.Current
        ElementLength = 0
        
        Select Case True
            Case subel.IsLineElement
                ElementLength = subel.AsLineElement.Length
            Case subel.IsArcElement
                ElementLength = subel.AsArcElement.Length
            Case subel.IsShapeElement
                ' For nested shapes, get their longest side recursively
                ElementLength = GetLongestSideFromShape(subel.AsShapeElement)
            Case subel.IsComplexShapeElement
                ' For nested complex shapes, get their longest side recursively
                ElementLength = GetLongestSideFromComplexShape(subel.AsComplexShapeElement)
        End Select
        
        If ElementLength > LongestSide Then
            LongestSide = ElementLength
        End If
    Loop
    
    GetLongestSideFromComplexShape = LongestSide
    Exit Function

ErrorHandler:
    GetLongestSideFromComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLongestSideFromComplexShape"
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
    On Error GoTo ErrorHandler
    If RND <> ARES_RND_ERROR_VALUE Then
        SetRound = Config.SetVar(ARESConfig.ARES_ROUNDS.key, RND)
    Else
        ShowStatus GetTranslation("LengthRoundError") & ARES_RND_ERROR_VALUE
    End If
    Exit Function

ErrorHandler:
    ' Return False in case of an error
    SetRound = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.SetRound"
End Function

' Public function to reset the rounding configuration variable
Public Function ResetRound() As Boolean
    On Error GoTo ErrorHandler
    ARESConfig.ResetConfigVar ARESConfig.ARES_ROUNDS.key
    ResetRound = (ARESConfig.ARES_ROUNDS.Value = ARESConfig.ARES_ROUNDS.defaultValue)
    Exit Function

ErrorHandler:
    ' Return False in case of an error
    ResetRound = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.ResetRound"
End Function