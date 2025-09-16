' Module: CustomPropertyHandler
' Description: This module provides functions to manipulate Custom Property (EC) in MicroStation with silent error handling.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass

Option Explicit

' Function to get or create an ItemTypeLibrary by name
Public Function GetItemTypeLibrary(Optional LibraryName As String = ARESConfig.ARES_NAME_LIBRARY_TYPE.Value, Optional ItemName As String = ARESConfig.ARES_NAME_ITEM_TYPE.Value) As ItemTypeLibrary
    On Error GoTo ErrorHandler

    Dim ItemLibs As ItemTypeLibraries
    Dim ITL As ItemTypeLibrary
    
    ' Instantiate the ItemTypeLibraries collection
    Set ItemLibs = New ItemTypeLibraries
    
    ' Try to find the library by name
    Set ITL = ItemLibs.FindByName(LibraryName)
    
    ' If the library does not exist, create it
    If ITL Is Nothing Then
        Set GetItemTypeLibrary = CreateItemTypeLibrary(LibraryName, ItemName)
    Else
        Set GetItemTypeLibrary = ITL
    End If
    
    Exit Function

ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypeLibrary"
    Set GetItemTypeLibrary = Nothing
End Function

' Function to create a new ItemTypeLibrary and its associated ItemType
Private Function CreateItemTypeLibrary(Optional LibraryName As String = ARESConfig.ARES_NAME_LIBRARY_TYPE.Value, Optional ItemName As String = ARESConfig.ARES_NAME_ITEM_TYPE.Value) As ItemTypeLibrary
    On Error GoTo ErrorHandler

    Dim item As ItemType
    Dim ItemProp As ItemTypeProperty
    Dim ItemLibs As ItemTypeLibraries
    Dim ITL As ItemTypeLibrary
    
    ' Instantiate the ItemTypeLibraries collection
    Set ItemLibs = New ItemTypeLibraries
    
    ' Try to find the library by name
    Set ITL = ItemLibs.FindByName(LibraryName)
    
    ' If the library does not exist, create it
    If ITL Is Nothing Then
        'Create ItemType Library
        Set CreateItemTypeLibrary = ItemLibs.CreateLib(LibraryName, False)
        
        ' Create the ItemType within the library
        Set item = CreateItemTypeLibrary.AddItemType(ItemName)
        
        ' Add properties to the ItemType
        Set ItemProp = item.AddProperty("EditedBy" & LibraryName, ItemPropertyTypeBoolean)
        Set ItemProp = item.AddProperty("UpdatedString", ItemPropertyTypeString)
        Set ItemProp = item.AddProperty("DateOfEdit", ItemPropertyTypeDateTime)
        
        ' Write the library to the DGN file
        CreateItemTypeLibrary.Write
    End If
    
    Exit Function

ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.CreateItemTypeLibrary"
    Set CreateItemTypeLibrary = Nothing
End Function

' Function to delete an ItemTypeLibrary by name
Public Function DeleteItemTypeLibrary(Optional LibraryName As String = ARESConfig.ARES_NAME_LIBRARY_TYPE.Value) As Boolean
    On Error GoTo ErrorHandler

    Dim oItemLibs As ItemTypeLibraries
    Dim ITL As ItemTypeLibrary
    
    ' Initialize the return value to False
    DeleteItemTypeLibrary = False
    
    ' Instantiate the ItemTypeLibraries collection
    Set oItemLibs = New ItemTypeLibraries
    
    ' Try to find the library by name
    Set ITL = oItemLibs.FindByName(LibraryName)
    
    ' If the library exists, delete it
    If Not ITL Is Nothing Then
        ITL.DeleteLib
        DeleteItemTypeLibrary = True
    End If
    
    Exit Function

ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.DeleteItemTypeLibrary"
    DeleteItemTypeLibrary = False
End Function

' Function to attach an ItemType to an Element
Public Function AttachItemToElement(ByVal El As element, Optional LibraryName As String = ARESConfig.ARES_NAME_LIBRARY_TYPE.Value, Optional ItemName As String = ARESConfig.ARES_NAME_ITEM_TYPE.Value) As Boolean
    On Error GoTo ErrorHandler

    Dim ITL As ItemTypeLibrary
    Dim item As ItemType
    Dim itemPropHandler As ItemTypePropertyHandler
    
    ' Initialize the return value to False
    AttachItemToElement = False
    
    ' Get the ItemTypeLibrary
    Set ITL = GetItemTypeLibrary(LibraryName, ItemName)
    
    ' If the library exists, proceed
    If Not ITL Is Nothing Then
        ' Get the ItemType by name
        Set item = ITL.GetItemTypeByName(ItemName)
        
        ' If the ItemType exists, attach it to the Element
        If Not item Is Nothing Then
            If Not El.Items.HasItems(LibraryName, ItemName) Then
                Set itemPropHandler = item.AttachItem(El)
            End If
            ' Set the return value to True if successful
            AttachItemToElement = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.AttachItemToElement"
    AttachItemToElement = False
End Function

' Function to remove an ItemType to an Element
Public Function RemoveItemToElement(ByVal El As element, Optional LibraryName As String = ARESConfig.ARES_NAME_LIBRARY_TYPE.Value, Optional ItemName As String = ARESConfig.ARES_NAME_ITEM_TYPE.Value) As Boolean
    On Error GoTo ErrorHandler

    Dim ITL As ItemTypeLibrary
    Dim item As ItemType
    
    ' Initialize the return value to False
    RemoveItemToElement = False
    
    ' Get the ItemTypeLibrary
    Set ITL = GetItemTypeLibrary(LibraryName, ItemName)
    
    ' If the library exists, proceed
    If Not ITL Is Nothing Then
        ' Get the ItemType by name
        Set item = ITL.GetItemTypeByName(ItemName)
        
        ' If the ItemType exists, remove it from the Element
        If Not item Is Nothing Then
            If El.Items.HasItems(LibraryName, ItemName) Then
                El.Items.RemoveItem item
                ' Set the return value to True if successful
                RemoveItemFromElement = True
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.RemoveItemToElement"
    RemoveItemToElement = False
End Function
