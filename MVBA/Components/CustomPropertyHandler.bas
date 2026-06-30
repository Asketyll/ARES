' Module: CustomPropertyHandler
' Description: Attaches, reads and writes ARES custom properties (MicroStation Item Types) on
'              elements, with silent error handling.
'
'              The item-type DEFINITIONS and their value lists live in a DGNLib (the "ARES"
'              ItemTypeLibrary), authored once through the Item Types dialog and deployed via
'              MS_DGNLIBLIST - they are NOT created from VBA (the MVBA Item Type API cannot author
'              a native value list / picklist). This module only ATTACHES the types to elements and
'              reads/writes their values (the value stored on the element is a plain string).
'              The library is resolved with FindForDesignFile(..., includeDgnLibs:=True), so the
'              definitions are found whether they live in the active file or in a referenced DGNLib.
'
'              The managed property names are user-editable via the ARES_Custom_Property_List config
'              var (default "Commune|Coupe_Type") - each name is BOTH the ItemType name and the
'              property name. A user adds a custom property by authoring it in the DGNLib (ItemType +
'              value list) and adding its name to that list; no code change needed.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), ErrorHandlerClass (global ErrorHandler)

Option Explicit

' Default managed property names when ARES_Custom_Property_List is unset (name = ItemType = property).
Private Const DEFAULT_CUSTOM_PROPERTIES As String = "Commune|Coupe_Type"

'######################################################################################################################
'                              CONFIGURED PROPERTY NAMES (user-editable list)
'######################################################################################################################

' The ARES custom-property names ARES manages, from the ARES_Custom_Property_List config var
' (| -delimited). Each entry is both the ItemType name and the property name. A user can add their
' own after authoring the matching ItemType + value list in the "ARES" DGNLib. 0-based array; use
' the standard safe bounds-check before reading it.
Public Function GetCustomPropertyNames() As String()
    On Error GoTo ErrorHandler

    GetCustomPropertyNames = Split(GetCustomPropertyListRaw(), ARESConstants.ARES_VAR_DELIMITER)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetCustomPropertyNames"
    GetCustomPropertyNames = Split(DEFAULT_CUSTOM_PROPERTIES, ARESConstants.ARES_VAR_DELIMITER)
End Function

' Raw | -delimited list from configuration; falls back to the default when config is unavailable or
' the variable is empty. Lazily initialises ARESConfig like the other modules.
Private Function GetCustomPropertyListRaw() As String
    On Error GoTo ErrorHandler

    GetCustomPropertyListRaw = DEFAULT_CUSTOM_PROPERTIES

    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_CUSTOM_PROPERTY_LIST Is Nothing Then Exit Function

    Dim s As String
    s = ARESConfig.ARES_CUSTOM_PROPERTY_LIST.Value
    If Len(Trim(s)) > 0 Then GetCustomPropertyListRaw = s
    Exit Function

ErrorHandler:
    GetCustomPropertyListRaw = DEFAULT_CUSTOM_PROPERTIES
End Function

'######################################################################################################################
'                              GENERIC LIBRARY HELPERS (reusable, schema-agnostic)
'######################################################################################################################

' Resolve the ARES ItemTypeLibrary, searching the active design file AND any referenced DGNLibs
' (the definitions normally live in a DGNLib declared in MS_DGNLIBLIST). Returns Nothing if absent.
Public Function FindItemTypeLibrary(Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As ItemTypeLibrary
    On Error GoTo ErrorHandler

    Dim ItemLibs As ItemTypeLibraries
    Set ItemLibs = New ItemTypeLibraries
    Set FindItemTypeLibrary = ItemLibs.FindForDesignFile(LibraryName, ActiveDesignFile, True)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.FindItemTypeLibrary"
    Set FindItemTypeLibrary = Nothing
End Function

'######################################################################################################################
'                              GENERIC ELEMENT HELPERS (attach / read / write)
'######################################################################################################################

' Attach an ItemType (by name) to an element. Idempotent: returns True if already attached.
Public Function AttachItemToElement(ByVal El As element, ByVal ItemName As String, Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
    On Error GoTo ErrorHandler

    AttachItemToElement = False
    If El Is Nothing Then Exit Function
    If Len(ItemName) = 0 Then Exit Function

    Dim ITL As ItemTypeLibrary
    Dim oItem As ItemType
    Dim oHandler As ItemTypePropertyHandler

    Set ITL = FindItemTypeLibrary(LibraryName)
    If ITL Is Nothing Then Exit Function

    Set oItem = ITL.GetItemTypeByName(ItemName)
    If oItem Is Nothing Then Exit Function

    If Not El.Items.HasItems(LibraryName, ItemName) Then
        Set oHandler = oItem.AttachItem(El)
        If oHandler Is Nothing Then Exit Function
    End If

    AttachItemToElement = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.AttachItemToElement"
    AttachItemToElement = False
End Function

' Detach an ItemType (by name) from an element. Returns True only when an attached item was removed.
Public Function RemoveItemFromElement(ByVal El As element, ByVal ItemName As String, Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
    On Error GoTo ErrorHandler

    RemoveItemFromElement = False
    If El Is Nothing Then Exit Function
    If Len(ItemName) = 0 Then Exit Function

    Dim ITL As ItemTypeLibrary
    Dim oItem As ItemType

    Set ITL = FindItemTypeLibrary(LibraryName)
    If ITL Is Nothing Then Exit Function

    Set oItem = ITL.GetItemTypeByName(ItemName)
    If oItem Is Nothing Then Exit Function

    If El.Items.HasItems(LibraryName, ItemName) Then
        oItem.DetachItem El
        RemoveItemFromElement = True
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.RemoveItemFromElement"
    RemoveItemFromElement = False
End Function

' Get the ItemType actually attached to an element. With ItemName empty, returns the first ARES item
' type found on the element; otherwise returns the named item type only if the element carries it.
Public Function GetItemTypeFromElement(ByVal El As element, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As ItemType
    On Error GoTo ErrorHandler

    Set GetItemTypeFromElement = Nothing
    If El Is Nothing Then Exit Function

    Dim oItems As Items
    Dim oHandler As ItemTypePropertyHandler
    Dim ITL As ItemTypeLibrary

    Set oItems = El.Items
    oItems.Refresh LibraryName

    Set ITL = FindItemTypeLibrary(LibraryName)
    If ITL Is Nothing Then Exit Function

    If Len(ItemName) > 0 Then
        Set GetItemTypeFromElement = ITL.GetItemTypeByName(ItemName)
        ' Verify the element actually carries this item type
        If Not GetItemTypeFromElement Is Nothing Then
            Set oHandler = oItems.FindForItemType(GetItemTypeFromElement)
            If oHandler Is Nothing Then Set GetItemTypeFromElement = Nothing
        End If
    Else
        Set oHandler = oItems.Find(LibraryName, "*", Nothing)
        If Not oHandler Is Nothing Then
            Set GetItemTypeFromElement = ITL.GetItemTypeByName(oHandler.ItemTypeName)
        End If
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypeFromElement"
    Set GetItemTypeFromElement = Nothing
End Function

' Get the ItemTypeLibrary an element references items from (Nothing if the element has none).
Public Function GetItemTypeLibraryFromElement(ByVal El As element, Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As ItemTypeLibrary
    On Error GoTo ErrorHandler

    Set GetItemTypeLibraryFromElement = Nothing
    If El Is Nothing Then Exit Function

    Dim oItems As Items
    Dim oHandler As ItemTypePropertyHandler

    Set oItems = El.Items
    oItems.Refresh LibraryName

    Set oHandler = oItems.Find(LibraryName, "*", Nothing)
    If Not oHandler Is Nothing Then
        Set GetItemTypeLibraryFromElement = FindItemTypeLibrary(LibraryName)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypeLibraryFromElement"
    Set GetItemTypeLibraryFromElement = Nothing
End Function

' Get the property handler for an element's item. With ItemName empty, returns the first ARES handler.
Public Function GetItemTypePropertyHandlerFromElement(ByVal El As element, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As ItemTypePropertyHandler
    On Error GoTo ErrorHandler

    Set GetItemTypePropertyHandlerFromElement = Nothing
    If El Is Nothing Then Exit Function

    Dim oItems As Items
    Dim oItem As ItemType

    Set oItems = El.Items
    oItems.Refresh LibraryName

    If Len(ItemName) > 0 Then
        Set oItem = GetItemTypeFromElement(El, ItemName, LibraryName)
        If Not oItem Is Nothing Then
            Set GetItemTypePropertyHandlerFromElement = oItems.FindForItemType(oItem)
        End If
    Else
        Set GetItemTypePropertyHandlerFromElement = oItems.Find(LibraryName, "*", Nothing)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetItemTypePropertyHandlerFromElement"
    Set GetItemTypePropertyHandlerFromElement = Nothing
End Function

' Read a property value from an element. Returns Null when the item/property is not present.
Public Function GetPropertyValueFromElement(ByVal El As element, ByVal PropertyName As String, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Variant
    On Error GoTo ErrorHandler

    GetPropertyValueFromElement = Null

    Dim oHandler As ItemTypePropertyHandler
    Set oHandler = GetItemTypePropertyHandlerFromElement(El, ItemName, LibraryName)
    If Not oHandler Is Nothing Then
        GetPropertyValueFromElement = oHandler.GetPropertyValue(PropertyName)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetPropertyValueFromElement"
    GetPropertyValueFromElement = Null
End Function

' Write a property value to an element. Returns True on success.
Public Function SetPropertyValueToElement(ByVal El As element, ByVal PropertyName As String, ByVal PropertyValue As Variant, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
    On Error GoTo ErrorHandler

    SetPropertyValueToElement = False

    Dim oHandler As ItemTypePropertyHandler
    Set oHandler = GetItemTypePropertyHandlerFromElement(El, ItemName, LibraryName)
    If Not oHandler Is Nothing Then
        SetPropertyValueToElement = oHandler.SetPropertyValue(PropertyName, PropertyValue)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.SetPropertyValueToElement"
    SetPropertyValueToElement = False
End Function
