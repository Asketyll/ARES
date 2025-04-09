' Module: Link
' Description: This module provides functions to retrieve linked elements in MicroStation.

' Dependencies: MicroStationDefinition

Option Explicit

' Public function to get the linked elements of a graphical element
Public Function GetLink(ByRef el As Element, _
                        Optional ReturnMe As Boolean = False, _
                        Optional FilterByTypes As Variant, _
                        Optional MaxCount As Byte = 255) As Variant
    On Error GoTo ErrorHandler

    Dim linkedElements() As Element
    Dim count As Byte
    Dim MSDEType() As MsdElementType

    ' Check if there is an active model reference
    If Not Application.HasActiveModelReference Then Exit Function

    ' Check if the element is graphical and has a valid graphic group ID
    If el.IsGraphical And el.GraphicGroup <> ARES_VAR.DEFAULT_GRAPHIC_GROUP_ID Then
        ' Initialize the element scan criteria
        Dim Esc As ElementScanCriteria
        Set Esc = New ElementScanCriteria

        ' Include only the specified graphic group
        Esc.IncludeOnlyGraphicGroup el.GraphicGroup

        ' Include the specified element types if provided
        If Not IsMissing(FilterByTypes) Then
            ' Ensure FilterByTypes is an array of MsdElementType
            MSDEType = EnsureArray(FilterByTypes)
            Esc.ExcludeAllTypes
            Dim i As Long
            For i = LBound(MSDEType) To UBound(MSDEType)
                If IsValidElementType(MSDEType(i)) And MSDEType(i) <> ARES_VAR.MSDETYPE_ERROR Then
                    Esc.IncludeType MSDEType(i)
                End If
            Next i
        End If

        ' Scan for elements in the specified graphic group
        Dim EE As ElementEnumerator
        Set EE = ActiveModelReference.Scan(Esc)

        If ReturnMe Then
            ' Return all elements in the enumerator as an array
            GetLink = EE.BuildArrayFromContents
            Exit Function
        Else
            ' Collect linked elements excluding the original element
            linkedElements = CollectLinkedElements(EE, el, MaxCount)
        End If
    End If

    ' Return the array of linked elements
    GetLink = linkedElements
    Exit Function

ErrorHandler:
    ' Handle errors by returning an empty array of Element type
    ReDim linkedElements(0) As Element
    GetLink = linkedElements
End Function

' Private function to ensure a variant is an array of MsdElementType
Private Function EnsureArray(ByVal value As Variant) As Variant
    On Error GoTo ErrorHandler

    Dim tempArray() As MsdElementType
    
    If IsArray(value) Then
        ' Check each element in the array and convert if necessary
        Dim i As Long
        ReDim tempArray(LBound(value) To UBound(value))
        For i = LBound(value) To UBound(value)
            Select Case VarType(value(i))
                Case vbString
                    tempArray(i) = StringToMsdElementType(value(i))
                Case vbLong
                    tempArray(i) = value(i)
                Case vbInteger
                    ' Convert Integer to Long
                    tempArray(i) = CLng(value(i))
                Case Else
                    ' Return error value if type is not recognized
                    ReDim tempArray(0)
                    tempArray(0) = ARES_VAR.MSDETYPE_ERROR
                    EnsureArray = tempArray
                    Exit Function
            End Select
        Next i
        EnsureArray = tempArray
    ElseIf Not IsMissing(value) And Not IsEmpty(value) Then
        ' Create a single-element array containing the value
        ReDim tempArray(0)
        Select Case VarType(value)
            Case vbString
                tempArray(0) = StringToMsdElementType(value)
            Case vbLong
                tempArray(0) = value
            Case vbInteger
                ' Convert Integer to Long
                tempArray(0) = CLng(value)
            Case Else
                ' Return error value if type is not recognized
                ReDim tempArray(0)
                tempArray(0) = ARES_VAR.MSDETYPE_ERROR
                EnsureArray = tempArray
                Exit Function
        End Select
        EnsureArray = tempArray
    Else
        EnsureArray = Array(ARES_VAR.MSDETYPE_ERROR)
    End If

    Exit Function

ErrorHandler:
    ' Return error value in case of any error
    ReDim tempArray(0)
    tempArray(0) = ARES_VAR.MSDETYPE_ERROR
    EnsureArray = tempArray
End Function


' Private function to collect linked elements excluding the original element
Private Function CollectLinkedElements(ByRef EE As ElementEnumerator, _
                                      ByRef el As Element, _
                                      ByVal MaxCount As Byte) As Variant
    On Error GoTo ErrorHandler
    
    Dim linkedElements() As Element
    Dim count As Byte
    Dim SubEl As Element

    ' Count the number of elements to size the array
    count = CountValidElements(EE, el, MaxCount)

    ' Initialize the array with the correct size if count is greater than 0
    If count > 0 Then
        ReDim linkedElements(1 To count)
        count = 0

        ' Reset the enumerator and populate the array
        EE.Reset
        Do While EE.MoveNext
            Set SubEl = EE.Current
            If IsValidElement(el, SubEl) Then
                count = count + 1
                Set linkedElements(count) = SubEl
                ' Stop if max count is reached
                If count = MaxCount Then Exit Do
            End If
        Loop
    End If

    CollectLinkedElements = linkedElements
    Exit Function

ErrorHandler:
    ' Return an empty array in case of any error
    CollectLinkedElements = linkedElements
End Function

' Private function to count valid elements excluding the original element
Private Function CountValidElements(ByRef EE As ElementEnumerator, _
                                    ByRef el As Element, _
                                    ByVal MaxCount As Byte) As Byte
    Dim count As Byte
    Dim SubEl As Element

    Do While EE.MoveNext
        Set SubEl = EE.Current
        If IsValidElement(el, SubEl) Then
            count = count + 1
            ' Stop if max count is reached
            If count = MaxCount Then Exit Do
        End If
    Loop

    CountValidElements = count
End Function

' Private function to check if an element is valid (not the original element)
Private Function IsValidElement(ByRef el As Element, ByRef SubEl As Element) As Boolean
    IsValidElement = (DLongComp(el.id, SubEl.id) <> 0)
End Function
