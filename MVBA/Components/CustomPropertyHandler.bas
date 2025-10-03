' Module: CustomPropertyHandler
' Description: This module provides functions to manipulate Custom Property (EC) in MicroStation with silent error handling.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass

Option Explicit

' Function to get or create an ItemTypeLibrary by name
Public Function GetItemTypeLibrary(Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = ARESConstants.ARES_NAME_ITEM_TYPE) As ItemTypeLibrary
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
Private Function CreateItemTypeLibrary(Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = ARESConstants.ARES_NAME_ITEM_TYPE) As ItemTypeLibrary
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
Public Function DeleteItemTypeLibrary(Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
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
Public Function AttachItemToElement(ByVal El As element, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = ARESConstants.ARES_NAME_ITEM_TYPE) As Boolean
    On Error GoTo ErrorHandler

    Dim ITL As ItemTypeLibrary
    Dim item As ItemType
    Dim ItemPropHandler As ItemTypePropertyHandler
    
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
                Set ItemPropHandler = item.AttachItem(El)
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

' Function to remove an ItemType from an Element
Public Function RemoveItemToElement(ByVal El As element, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = ARESConstants.ARES_NAME_ITEM_TYPE) As Boolean
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
            ' Check if element has this item type attached
            If El.Items.HasItems(LibraryName, ItemName) Then
                ' Use the correct DetachItem method from ItemType
                item.DetachItem El
                ' Set the return value to True if successful
                RemoveItemToElement = True
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.RemoveItemToElement"
    RemoveItemToElement = False
End Function

' Function to get ItemTypeLibrary from an element (if it has items from that library)
Public Function GetItemTypeLibraryFromElement(ByVal El As element, Optional LibraryName As String = "") As ItemTypeLibrary
    On Error GoTo ErrorHandler
    
    Set GetItemTypeLibraryFromElement = Nothing
    
    Dim oItems As Items
    Dim oItemPropHandler As ItemTypePropertyHandler
    Dim oItemLibs As ItemTypeLibraries
    Dim oItemLib As ItemTypeLibrary
    
    Set oItems = El.Items
    
    ' If specific library name provided, look for that one
    If Len(LibraryName) > 0 Then
        oItems.Refresh LibraryName
        Set oItemPropHandler = oItems.Find(LibraryName, "*", Nothing)
        If Not oItemPropHandler Is Nothing Then
            Set oItemLibs = New ItemTypeLibraries
            Set GetItemTypeLibraryFromElement = oItemLibs.FindByName(LibraryName)
        End If
    Else
        ' Return first available library (for backward compatibility)
        Set oItemLibs = New ItemTypeLibraries
        Set GetItemTypeLibraryFromElement = oItemLibs.FindByName(ARESConstants.ARES_NAME_LIBRARY_TYPE)
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypeLibraryFromElement"
    Set GetItemTypeLibraryFromElement = Nothing
End Function

' Function to get ItemType from an element
Public Function GetItemTypeFromElement(ByVal El As element, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = "") As ItemType
    On Error GoTo ErrorHandler
    
    Set GetItemTypeFromElement = Nothing
    
    Dim oItems As Items
    Dim oItemPropHandler As ItemTypePropertyHandler
    Dim oItemLibs As ItemTypeLibraries
    Dim oItemLib As ItemTypeLibrary
    
    Set oItems = El.Items
    oItems.Refresh LibraryName
    
    ' If specific item name provided, look for that one
    If Len(ItemName) > 0 Then
        Set oItemLibs = New ItemTypeLibraries
        Set oItemLib = oItemLibs.FindByName(LibraryName)
        If Not oItemLib Is Nothing Then
            Set GetItemTypeFromElement = oItemLib.GetItemTypeByName(ItemName)
            
            ' Verify the element actually has this item type
            Set oItemPropHandler = oItems.FindForItemType(GetItemTypeFromElement)
            If oItemPropHandler Is Nothing Then
                Set GetItemTypeFromElement = Nothing
            End If
        End If
    Else
        ' Return first available item type from the library
        Set oItemPropHandler = oItems.Find(LibraryName, "*", Nothing)
        If Not oItemPropHandler Is Nothing Then
            Set oItemLibs = New ItemTypeLibraries
            Set oItemLib = oItemLibs.FindByName(LibraryName)
            If Not oItemLib Is Nothing Then
                Set GetItemTypeFromElement = oItemLib.GetItemTypeByName(oItemPropHandler.ItemTypeName)
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypeFromElement"
    Set GetItemTypeFromElement = Nothing
End Function

' Function to get ItemTypePropertyHandler from an element
Public Function GetItemTypePropertyHandlerFromElement(ByVal El As element, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = "") As ItemTypePropertyHandler
    On Error GoTo ErrorHandler
    
    Set GetItemTypePropertyHandlerFromElement = Nothing
    
    Dim oItems As Items
    Dim oItemPropHandler As ItemTypePropertyHandler
    Dim oItemType As ItemType
    
    Set oItems = El.Items
    oItems.Refresh LibraryName
    
    ' If specific item name provided, find by item type
    If Len(ItemName) > 0 Then
        Set oItemType = GetItemTypeFromElement(El, LibraryName, ItemName)
        If Not oItemType Is Nothing Then
            Set GetItemTypePropertyHandlerFromElement = oItems.FindForItemType(oItemType)
        End If
    Else
        ' Return first available handler from the library
        Set GetItemTypePropertyHandlerFromElement = oItems.Find(LibraryName, "*", Nothing)
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypePropertyHandlerFromElement"
    Set GetItemTypePropertyHandlerFromElement = Nothing
End Function

' Function to get property value from an element
Public Function GetPropertyValueFromElement(ByVal El As element, ByVal PropertyName As String, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = "") As Variant
    On Error GoTo ErrorHandler
    
    GetPropertyValueFromElement = Null
    
    Dim oItemPropHandler As ItemTypePropertyHandler
    
    Set oItemPropHandler = GetItemTypePropertyHandlerFromElement(El, LibraryName, ItemName)
    If Not oItemPropHandler Is Nothing Then
        GetPropertyValueFromElement = oItemPropHandler.GetPropertyValue(PropertyName)
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetPropertyValueFromElement"
    GetPropertyValueFromElement = Null
End Function

' Function to set property value to an element
Public Function SetPropertyValueToElement(ByVal El As element, ByVal PropertyName As String, ByVal PropertyValue As Variant, Optional LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ItemName As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    SetPropertyValueToElement = False
    
    Dim oItemPropHandler As ItemTypePropertyHandler
    
    Set oItemPropHandler = GetItemTypePropertyHandlerFromElement(El, LibraryName, ItemName)
    If Not oItemPropHandler Is Nothing Then
        SetPropertyValueToElement = oItemPropHandler.SetPropertyValue(PropertyName, PropertyValue)
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.SetPropertyValueToElement"
    SetPropertyValueToElement = False
End Function