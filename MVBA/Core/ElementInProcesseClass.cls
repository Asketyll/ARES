' Class Module: ElementInProcesseClass
' Description: Thread-safe collection manager for MicroStation element IDs currently being processed.
'              Prevents duplicate processing and provides efficient element ID tracking.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass (via BootLoader), ARESConstants

Option Explicit

' === PRIVATE MEMBERS ===
Private moElementIds As Collection   ' Collection of element ID strings
Private mbInitialized As Boolean     ' Initialization flag

' Initialize the collection when class is created
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set moElementIds = New Collection
    mbInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Class_Initialize"
    mbInitialized = False
End Sub

' Add an element ID to the processing collection
' Returns True if added successfully, False if already exists or error
Public Function Add(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler
    
    Add = False
    
    If Not mbInitialized Then Exit Function
    If oElement Is Nothing Then Exit Function
    
    Dim strElementId As String
    strElementId = DLongToString(oElement.ID)
    
    ' Check if element ID already exists
    If ContainsId(strElementId) Then Exit Function
    
    ' Add element ID to collection
    moElementIds.Add strElementId, strElementId
    Add = True
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Add"
    Add = False
End Function

' Remove an element ID from the collection
Public Function Remove(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler
    
    Remove = False
    
    If Not mbInitialized Then Exit Function
    If oElement Is Nothing Then Exit Function
    
    Dim strElementId As String
    strElementId = DLongToString(oElement.ID)
    
    Remove = RemoveById(strElementId)
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Remove"
    Remove = False
End Function

' Remove element by ID string
Private Function RemoveById(ByVal strElementId As String) As Boolean
    On Error GoTo ErrorHandler
    
    RemoveById = False
    
    If Not mbInitialized Then Exit Function
    If Len(strElementId) = 0 Then Exit Function
    
    If ContainsId(strElementId) Then
        moElementIds.Remove strElementId
        RemoveById = True
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.RemoveById"
    RemoveById = False
End Function

' Check if an element is in the collection
Public Function Contains(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler
    
    Contains = False
    
    If Not mbInitialized Then Exit Function
    If oElement Is Nothing Then Exit Function
    
    Dim strElementId As String
    strElementId = DLongToString(oElement.ID)
    
    Contains = ContainsId(strElementId)
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Contains"
    Contains = False
End Function

' Check if element ID is in collection
Private Function ContainsId(ByVal strElementId As String) As Boolean
    On Error Resume Next

    Dim strTest As String
    strTest = moElementIds(strElementId)
    ContainsId = (Err.Number = 0)
    Err.Clear  ' Important: clear error to prevent leaking to caller
End Function

' Clear all element IDs from the collection
Public Sub Clear()
    On Error GoTo ErrorHandler
    
    If mbInitialized Then
        Set moElementIds = New Collection
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.Clear"
End Sub

' Legacy method for backward compatibility
Public Sub Reset()
    Clear
End Sub

' Get the count of element IDs
Public Property Get Count() As Long
    On Error GoTo ErrorHandler
    
    If mbInitialized Then
        Count = moElementIds.Count
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

' Check if collection has elements
Public Property Get HasElements() As Boolean
    HasElements = (Count > 0)
End Property

' Get all elements as a Collection (retrieved fresh from MicroStation)
' Returns a Collection of valid Element objects, skipping deleted/invalid elements
Public Function GetAllElements() As Collection
    On Error GoTo ErrorHandler
    
    Dim colElements As Collection
    Dim strElementId As Variant
    Dim oElement As element
    
    Set colElements = New Collection
    
    ' Check if initialized and has IDs
    If Not mbInitialized Then
        Set GetAllElements = colElements
        Exit Function
    End If
    
    If moElementIds.Count = 0 Then
        Set GetAllElements = colElements
        Exit Function
    End If
    
    ' Iterate through all stored IDs and retrieve fresh elements
    For Each strElementId In moElementIds
        On Error Resume Next
        
        ' Retrieve fresh element from MicroStation
        Set oElement = Nothing
        Set oElement = ActiveModelReference.GetElementById(DLongFromString(CStr(strElementId)))
        
        ' Only add if element was successfully retrieved
        If Err.Number = 0 And Not oElement Is Nothing Then
            colElements.Add oElement
        End If
        
        On Error GoTo ErrorHandler
    Next strElementId
    
    Set GetAllElements = colElements
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.GetAllElements"
    
    ' Return empty collection on error
    If colElements Is Nothing Then
        Set colElements = New Collection
    End If
    Set GetAllElements = colElements
End Function

' Get all element IDs as a Collection (raw IDs, no validation)
Public Function GetAllElementIds() As Collection
    On Error GoTo ErrorHandler
    
    Set GetAllElementIds = moElementIds
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementInProcesseClass.GetAllElementIds"
    Set GetAllElementIds = Nothing
End Function

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    
    If mbInitialized And Count > 0 Then
        ErrorHandler.HandleError "ElementInProcesse terminated with " & Count & " element IDs still in queue", 0, "", "ElementInProcesseClass.Class_Terminate"
    End If
    
    Set moElementIds = Nothing
    mbInitialized = False
End Sub