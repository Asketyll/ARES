' Class Module: ElementInProcesseClass
' Description: Thread-safe collection manager for MicroStation elements currently being processed
' Prevents duplicate processing and provides efficient element tracking
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass (via BootLoader), ARESConstants
Option Explicit

' === PRIVATE MEMBERS ===
Private moElements As Collection    ' Collection of elements keyed by ID string
Private mlElementCount As Long      ' Cached count for performance
Private mbInitialized As Boolean    ' Initialization flag

' Initialize the collection when class is created
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set moElements = New Collection
    mlElementCount = 0
    mbInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Class_Initialize"
    mbInitialized = False
End Sub

' Add an element to the processing collection
' Returns True if added successfully, False if already exists or error
Public Function Add(ByVal oElement As Element) As Boolean
    On Error GoTo ErrorHandler
    
    Add = False
    
    ' Validate inputs
    If Not mbInitialized Then
        ErrorHandler.HandleError "ElementInProcesse not initialized", 0, "ElementInProcesseClass.Add", "ERROR"
        Exit Function
    End If
    
    If oElement Is Nothing Then
        ErrorHandler.HandleError "Cannot add null element", 0, "ElementInProcesseClass.Add", "WARNING"
        Exit Function
    End If
    
    Dim strElementId As String
    strElementId = ConvertElementIdToString(oElement.ID)
    
    ' Check if element already exists
    If ContainsElementId(strElementId) Then
        ' Element already in collection - this is normal, not an error
        Exit Function
    End If
    
    ' Add element to collection
    moElements.Add oElement, strElementId
    mlElementCount = mlElementCount + 1
    Add = True
    
    ' Log for debugging if needed
    ' ErrorHandler.HandleError "Added element ID: " & strElementId, 0, "ElementInProcesseClass.Add", "DEBUG"
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Add"
    Add = False
End Function

' Remove an element from the processing collection
Public Function Remove(ByVal oElement As Element) As Boolean
    On Error GoTo ErrorHandler
    
    Remove = False
    
    ' Validate inputs
    If Not mbInitialized Then
        ErrorHandler.HandleError "ElementInProcesse not initialized", 0, "ElementInProcesseClass.Remove", "ERROR"
        Exit Function
    End If
    
    If oElement Is Nothing Then
        ErrorHandler.HandleError "Cannot remove null element", 0, "ElementInProcesseClass.Remove", "WARNING"
        Exit Function
    End If
    
    Dim strElementId As String
    strElementId = ConvertElementIdToString(oElement.ID)
    
    ' Remove element if it exists
    If ContainsElementId(strElementId) Then
        moElements.Remove strElementId
        mlElementCount = mlElementCount - 1
        Remove = True
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Remove"
    Remove = False
End Function

' Remove element by ID string (more efficient)
Public Function RemoveById(ByVal strElementId As String) As Boolean
    On Error GoTo ErrorHandler
    
    RemoveById = False
    
    If Not mbInitialized Or Len(strElementId) = 0 Then Exit Function
    
    If ContainsElementId(strElementId) Then
        moElements.Remove strElementId
        mlElementCount = mlElementCount - 1
        RemoveById = True
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.RemoveById"
    RemoveById = False
End Function

' Check if an element is currently being processed
Public Function Contains(ByVal oElement As Element) As Boolean
    On Error GoTo ErrorHandler
    
    Contains = False
    
    ' Validate inputs
    If Not mbInitialized Or oElement Is Nothing Then Exit Function
    
    Dim strElementId As String
    strElementId = ConvertElementIdToString(oElement.ID)
    
    Contains = ContainsElementId(strElementId)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Contains"
    Contains = False
End Function

' Check if element ID is in collection (more efficient)
Public Function ContainsId(ByVal strElementId As String) As Boolean
    On Error GoTo ErrorHandler
    
    ContainsId = False
    
    If Not mbInitialized Or Len(strElementId) = 0 Then Exit Function
    
    ContainsId = ContainsElementId(strElementId)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.ContainsId"
    ContainsId = False
End Function

' Internal method to check if element ID exists
Private Function ContainsElementId(ByVal strElementId As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim oTempElement As Element
    Set oTempElement = moElements(strElementId)
    ContainsElementId = True
    Exit Function
    
ErrorHandler:
    ' Element not found - this is normal
    ContainsElementId = False
End Function

' Clear all elements from the collection
Public Sub Clear()
    On Error GoTo ErrorHandler
    
    If mbInitialized Then
        Set moElements = New Collection
        mlElementCount = 0
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Clear"
End Sub

' Legacy method for backward compatibility
Public Sub Reset()
    Clear
End Sub

' Get the count of elements currently being processed
Public Property Get Count() As Long
    On Error GoTo ErrorHandler
    
    If mbInitialized Then
        Count = mlElementCount
    Else
        Count = 0
    End If
    Exit Property
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Count"
    Count = 0
End Property

' Check if collection is empty
Public Property Get IsEmpty() As Boolean
    IsEmpty = (Count = 0)
End Property

' Get element by ID string
Public Function GetElementById(ByVal strElementId As String) As Element
    On Error GoTo ErrorHandler
    
    Set GetElementById = Nothing
    
    If mbInitialized And Len(strElementId) > 0 Then
        Set GetElementById = moElements(strElementId)
    End If
    
    Exit Function
    
ErrorHandler:
    ' Element not found
    Set GetElementById = Nothing
End Function

' Convert Element ID to string representation
Private Function ConvertElementIdToString(ByVal lngElementId As Long) As String
    On Error GoTo ErrorHandler
    
    ConvertElementIdToString = CStr(lngElementId)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.ConvertElementIdToString"
    ConvertElementIdToString = "0"
End Function

' Get all element IDs as an array (for debugging)
Public Function GetAllElementIds() As String()
    On Error GoTo ErrorHandler
    
    Dim strIds() As String
    Dim oElement As Element
    Dim intIndex As Integer
    
    If Count = 0 Then
        ReDim strIds(0 To 0)
        strIds(0) = ""
        GetAllElementIds = strIds
        Exit Function
    End If
    
    ReDim strIds(0 To Count - 1)
    intIndex = 0
    
    For Each oElement In moElements
        strIds(intIndex) = ConvertElementIdToString(oElement.ID)
        intIndex = intIndex + 1
    Next oElement
    
    GetAllElementIds = strIds
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.GetAllElementIds"
    ReDim strIds(0 To 0)
    GetAllElementIds = strIds
End Function

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    
    If mbInitialized And mlElementCount > 0 Then
        ErrorHandler.HandleError "ElementInProcesse terminated with " & mlElementCount & " elements still in queue", 0, "ElementInProcesseClass.Class_Terminate", "WARNING"
    End If
    
    Set moElements = Nothing
    mlElementCount = 0
    mbInitialized = False
End Sub
