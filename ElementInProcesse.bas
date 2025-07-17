' Module: ElementInProcesse
' Description: This module provides functions to ensure a element is (not) in a ARES processe.
' License: This project is licensed under the AGPL-3.0.

Option Explicit

'store all id of element in processe
Dim ElementInPro() As DLong

'Add a element to ElementInProcesse
Public Function Add(ByVal el As Element) As Boolean
    Add = False
    If Not IsIn(el) Then
        If UBound(ElementInPro) <> -1 Then
            If DLongToLong(ElementInPro(UBound(ElementInPro))) <> 0 Then
                ReDim Preserve ElementInPro(LBound(ElementInPro) To UBound(ElementInPro) + 1)
            End If
            ElementInPro(UBound(ElementInPro)) = el.id
            Add = True
            Exit Function
        Else
            ReDim ElementInPro(0)
            ElementInPro(0) = el.id
            Add = True
            Exit Function
        End If
    End If
End Function

'Remove a element from ElementInProcesse
Public Function Remove(ByVal el As Element) As Boolean
    Dim i As Integer
    Dim indexToRemove As Integer
    Dim arrayLength As Integer
    Remove = False

    If Not IsArray(ElementInPro) Then Exit Function

    ' Find the index of the element to remove
    For i = LBound(ElementInPro) To UBound(ElementInPro)
        If DLongComp(ElementInPro(i), el.id) = 0 Then
            indexToRemove = i
            Remove = True
            Exit For
        End If
    Next i

    ' If the element was found, remove it by shifting elements
    If Remove Then
        arrayLength = UBound(ElementInPro) - LBound(ElementInPro) + 1
        ' Shift elements to the left of the index to remove
        For i = indexToRemove To arrayLength - 2
            ElementInPro(i) = ElementInPro(i + 1)
        Next i

        ' Resize the array
        If arrayLength - 1 > 0 Then
            ReDim Preserve ElementInPro(LBound(ElementInPro) To arrayLength - 2)
        Else
            ' If the array has only one element, erase the array
            Erase ElementInPro
        End If
    End If
End Function

'Check if a element is in ElementInProcesse
Public Function IsIn(ByVal el As Element) As Boolean
    Dim i As Integer
    IsIn = False
    
    If Not IsArray(ElementInPro) Then Exit Function
    For i = LBound(ElementInPro) To UBound(ElementInPro)
        If DLongComp(ElementInPro(i), el.id) = 0 Then
            IsIn = True
            Exit For
        End If
    Next i
End Function
' Reset the ElementInPro array
Public Sub Reset()
    Erase ElementInPro
End Sub
