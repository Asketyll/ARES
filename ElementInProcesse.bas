' Class Module: ElementInProcesseClass
' Description: Manages a collection of elements being processed.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

Private pElements As Collection

Private Sub Class_Initialize()
    Set pElements = New Collection
End Sub

' Method to add an element to the collection
Public Function Add(ByVal element As element) As Boolean
    On Error Resume Next
    Dim elementId As String
    elementId = DLongToString(element.id)
    pElements.Add element, elementId
    If Err.Number = 0 Then
        Add = True
    Else
        Add = False
        Err.Clear
    End If
End Function

' Method to remove an element from the collection
Public Sub Remove(ByVal element As element)
    On Error Resume Next
    Dim elementId As String
    elementId = DLongToString(element.id)
    pElements.Remove elementId
End Sub

' Method to check if an element is in the collection
Public Function IsIn(ByVal element As element) As Boolean
    On Error Resume Next
    Dim elementId As String
    elementId = DLongToString(element.id)
    Dim temp As element
    temp = pElements(elementId)
    If Err.Number = 0 Then
        IsIn = True
    Else
        IsIn = False
        Err.Clear
    End If
End Function

' Method to reset the collection
Public Sub Reset()
    Set pElements = New Collection
End Sub

' Method to get the count of elements in the collection
Public Function Count() As Long
    Count = pElements.Count
End Function
