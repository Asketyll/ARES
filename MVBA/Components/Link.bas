' Module: Link
' Description: This module provides functions to retrieve linked elements in MicroStation.
' It includes functions to scan for elements in a specified graphic group and filter them by type.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: MicroStationDefinition, ARESConfigClass, ARESConstants, ErrorHandlerClass
Option Explicit

' Public function to get the linked elements of a graphical element
Public Function GetLink(ByRef El As element, _
                        Optional ReturnMe As Boolean = False, _
                        Optional FilterByTypes As Variant, _
                        Optional MaxCount As Byte = 255) As Variant
    On Error GoTo ErrorHandler
    Dim LinkedElements() As element
    Dim MSDEType() As MsdElementType
    Dim Esc As ElementScanCriteria
    Dim i As Long
    Dim ee As ElementEnumerator
    Dim Count As Byte
    Dim subel As element
    
    ' Check if there is an active model reference
    If Not Application.HasActiveModelReference Then Exit Function

    ' Check if the element is graphical and has a valid graphic group ID
    If El.IsGraphical Then
        If El.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
            ' Initialize the element scan criteria
            Set Esc = New ElementScanCriteria
            ' Include only the specified graphic group
            Esc.IncludeOnlyGraphicGroup El.GraphicGroup
    
            ' Include the specified element types if provided
            If Not IsMissing(FilterByTypes) Then
                MSDEType = EnsureArray(FilterByTypes)
                Esc.ExcludeAllTypes
                For i = LBound(MSDEType) To UBound(MSDEType)
                    If IsValidElementType(MSDEType(i)) And MSDEType(i) <> ARES_MSDETYPE_ERROR Then
                        Esc.IncludeType MSDEType(i)
                    End If
                Next i
            End If
    
            ' Scan for elements in the specified graphic group
            Set ee = ActiveModelReference.Scan(Esc)
    
            If ReturnMe Then
                ' Return all elements in the enumerator as an array
                GetLink = ee.BuildArrayFromContents
                Exit Function
            Else
                ' Count the number of valid elements
                Count = 0
                Do While ee.MoveNext
                    Set subel = ee.Current
                    If DLongComp(El.id, subel.id) <> 0 Then
                        Count = Count + 1
                        If Count = MaxCount Then Exit Do
                    End If
                Loop
    
                ' If valid elements were found, collect them
                If Count > 0 Then
                    ReDim LinkedElements(1 To Count)
                    Count = 0
                    ' Reset the enumerator and populate the array
                    ee.Reset
                    Do While ee.MoveNext And Count < MaxCount
                        Set subel = ee.Current
                        If DLongComp(El.id, subel.id) <> 0 Then
                            Count = Count + 1
                            Set LinkedElements(Count) = subel
                        End If
                    Loop
                End If
            End If
        End If
    End If

    ' Return the array of linked elements
    GetLink = LinkedElements
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Link.GetLink"
    ' Handle errors by returning an empty array of Element type
    ReDim LinkedElements(0) As element
    GetLink = LinkedElements
End Function

' Private function to ensure a variant is an array of MsdElementType
Private Function EnsureArray(ByVal Value As Variant) As Variant
    On Error GoTo ErrorHandler
    Dim tempArray() As MsdElementType
    Dim i As Long

    If IsArray(Value) Then
        ' Check each element in the array and convert if necessary
        ReDim tempArray(LBound(Value) To UBound(Value))
        For i = LBound(Value) To UBound(Value)
            Select Case VarType(Value(i))
                Case vbString
                    tempArray(i) = StringToMsdElementType(Value(i))
                Case vbLong
                    tempArray(i) = Value(i)
                Case vbInteger
                    ' Convert Integer to Long
                    tempArray(i) = CLng(Value(i))
                Case Else
                    ' Return error value if type is not recognized
                    ReDim tempArray(0)
                    tempArray(0) = ARES_MSDETYPE_ERROR
                    EnsureArray = tempArray
                    Exit Function
            End Select
        Next i
        EnsureArray = tempArray
    ElseIf Not IsMissing(Value) And Not IsEmpty(Value) Then
        ' Create a single-element array containing the value
        ReDim tempArray(0)
        Select Case VarType(Value)
            Case vbString
                tempArray(0) = StringToMsdElementType(Value)
            Case vbLong
                tempArray(0) = Value
            Case vbInteger
                ' Convert Integer to Long
                tempArray(0) = CLng(Value)
            Case Else
                ' Return error value if type is not recognized
                ReDim tempArray(0)
                tempArray(0) = ARES_MSDETYPE_ERROR
                EnsureArray = tempArray
                Exit Function
        End Select
        EnsureArray = tempArray
    Else
        EnsureArray = Array(ARES_MSDETYPE_ERROR)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Link.EnsureArray"
    ' Handle errors by returning an array with one MSDETYPE_ERROR
    ReDim tempArray(0)
    tempArray(0) = ARES_MSDETYPE_ERROR
    EnsureArray = tempArray
End Function

' Private function to collect linked elements excluding the original element
Private Function CollectLinkedElements(ByRef ee As ElementEnumerator, _
                                      ByRef El As element, _
                                      ByVal MaxCount As Byte) As Variant
    On Error GoTo ErrorHandler

    Dim LinkedElements() As element
    Dim Count As Byte
    Dim subel As element

    ' Count the number of elements to size the array
    Count = CountValidElements(ee, El, MaxCount)

    ' Initialize the array with the correct size if count is greater than 0
    If Count > 0 Then
        ReDim LinkedElements(1 To Count)
        Count = 0
        ' Reset the enumerator and populate the array
        ee.Reset
        Do While ee.MoveNext
            Set subel = ee.Current
            If isValidElement(El, subel) Then
                Count = Count + 1
                Set LinkedElements(Count) = subel
                ' Stop if max count is reached
                If Count = MaxCount Then Exit Do
            End If
        Loop
    End If

    CollectLinkedElements = LinkedElements
    Exit Function

ErrorHandler:
    ' Return an empty array in case of any error
    CollectLinkedElements = LinkedElements
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Link.CollectLinkedElements"
End Function

' Private function to count valid elements excluding the original element
Private Function CountValidElements(ByRef ee As ElementEnumerator, _
                                    ByRef El As element, _
                                    ByVal MaxCount As Byte) As Byte
    On Error GoTo ErrorHandler

    Dim Count As Byte
    Dim subel As element

    Do While ee.MoveNext
        Set subel = ee.Current
        If isValidElement(El, subel) Then
            Count = Count + 1
            ' Stop if max count is reached
            If Count = MaxCount Then Exit Do
        End If
    Loop

    CountValidElements = Count
    Exit Function

ErrorHandler:
    CountValidElements = 0
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Link.CountValidElements"
End Function

' Private function to check if an element is valid (not the original element)
Private Function isValidElement(ByRef El As element, ByRef subel As element) As Boolean
    isValidElement = (DLongComp(El.id, subel.id) <> 0)
End Function
