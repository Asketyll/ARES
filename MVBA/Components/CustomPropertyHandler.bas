' Module: CustomPropertyHandler
' Description: Creates and manipulates ARES custom properties (MicroStation Item Types / EC properties)
'              with silent error handling.
'
'              ARES exposes ONE ItemType per custom property, all inside a single "ARES"
'              ItemTypeLibrary, so each property can be attached independently to different elements:
'                - "Commune"    : free editable text          (ItemPropertyTypeString)
'                - "Coupe Type" : value taken from a list      (ItemPropertyTypeString)
'
'              ---------------------------------------------------------------------------------------
'              DROPDOWN FEASIBILITY (Coupe Type)
'              ---------------------------------------------------------------------------------------
'              The MicroStation VBA Item Type API (ItemTypeLibrary / ItemType / ItemTypeProperty)
'              exposes NO member to define a value list / picklist / enumeration. A property can only
'              be Boolean/DateTime/Double/Integer/Point/String (see ItemTypePropertyType enum),
'              optionally with a default value or a calculated expression. There is therefore no way to
'              author a NATIVE dropdown (a value-list-bound property shown in the Properties dialog)
'              from VBA - that requires a .dgnlib authored through the Item Types dialog / ECSchema.
'
'              ARES handles this at application level instead:
'                * "Coupe Type" is stored as a plain String property.
'                * The list of allowed values is NOT hard-coded; it lives in the ARES_Coupe_Type_List
'                  configuration variable (see ARESConfigClass). GetCoupeTypeValues() returns it.
'                * The dropdown itself will be presented later by an ARES UserForm (ComboBox) at
'                  apply-time; the value finally written on the element is the chosen string.
'              The first configured value is used here to seed the property default value.
'
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), ErrorHandlerClass (global ErrorHandler)

Option Explicit

'######################################################################################################################
'                                        CREATION (the deliverable)
'######################################################################################################################

' Create / ensure the ARES item types and their properties in the active design file.
' Idempotent: missing item types are added, existing ones are left untouched, and the library is
' written to the DGN only when something actually changed.
Public Function EnsureARESItemTypes() As Boolean
    On Error GoTo ErrorHandler

    EnsureARESItemTypes = False

    Dim ItemLibs As ItemTypeLibraries
    Dim ITL As ItemTypeLibrary
    Dim bChanged As Boolean

    Set ItemLibs = New ItemTypeLibraries
    Set ITL = ItemLibs.FindByName(ARESConstants.ARES_NAME_LIBRARY_TYPE)

    ' Create the library if it does not exist yet
    If ITL Is Nothing Then
        Set ITL = ItemLibs.CreateLib(ARESConstants.ARES_NAME_LIBRARY_TYPE, False)
        If ITL Is Nothing Then Exit Function
        bChanged = True
    End If

    ' One ItemType per property so they remain distinct and attachable to different elements.
    If EnsureStringItemType(ITL, ARESConstants.ARES_ITEM_COMMUNE, ARESConstants.ARES_PROP_COMMUNE, "") Then
        bChanged = True
    End If
    If EnsureStringItemType(ITL, ARESConstants.ARES_ITEM_COUPE_TYPE, ARESConstants.ARES_PROP_COUPE_TYPE, GetFirstCoupeTypeValue()) Then
        bChanged = True
    End If

    ' Persist to the DGN only when something was actually added
    If bChanged Then ITL.Write

    EnsureARESItemTypes = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.EnsureARESItemTypes"
    EnsureARESItemTypes = False
End Function

' Ensure an ItemType holding a single string property exists in the given library.
' Returns True when it had to create the ItemType (caller must Write the library afterwards),
' False when the ItemType was already present (nothing to persist).
Private Function EnsureStringItemType(ByVal ITL As ItemTypeLibrary, ByVal ItemName As String, ByVal PropertyName As String, ByVal DefaultValue As String) As Boolean
    On Error GoTo ErrorHandler

    EnsureStringItemType = False

    Dim oItem As ItemType
    Dim oProp As ItemTypeProperty

    ' Already there? leave it untouched (idempotent)
    Set oItem = ITL.GetItemTypeByName(ItemName)
    If Not oItem Is Nothing Then Exit Function

    Set oItem = ITL.AddItemType(ItemName)
    If oItem Is Nothing Then Exit Function

    Set oProp = oItem.AddProperty(PropertyName, ItemPropertyTypeString)
    If Not oProp Is Nothing Then
        ' Seed the default value when one is provided (e.g. first Coupe Type choice)
        If Len(DefaultValue) > 0 Then oProp.SetDefaultValue DefaultValue
    End If

    EnsureStringItemType = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.EnsureStringItemType"
    EnsureStringItemType = False
End Function

'######################################################################################################################
'                                   COUPE TYPE VALUE LIST (from configuration)
'######################################################################################################################

' Allowed values for the "Coupe Type" property, read from the ARES_Coupe_Type_List configuration
' variable so the list is configurable and never hard-coded here. Intended for the apply-time
' dropdown UI. Returns a 0-based array; use the standard safe bounds-check before reading it.
Public Function GetCoupeTypeValues() As String()
    On Error GoTo ErrorHandler

    GetCoupeTypeValues = Split(GetCoupeTypeListRaw(), ARESConstants.ARES_VAR_DELIMITER)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetCoupeTypeValues"
    Dim emptyArr() As String
    GetCoupeTypeValues = emptyArr
End Function

' Raw pipe-delimited list string from configuration. Lazily initialises ARESConfig like the other
' modules (Command, FileDialogs, LangManager) so the value list is available even on a standalone call.
Private Function GetCoupeTypeListRaw() As String
    On Error GoTo ErrorHandler

    GetCoupeTypeListRaw = ""

    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_COUPE_TYPE_LIST Is Nothing Then Exit Function

    GetCoupeTypeListRaw = ARESConfig.ARES_COUPE_TYPE_LIST.Value
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetCoupeTypeListRaw"
    GetCoupeTypeListRaw = ""
End Function

' First configured Coupe Type value, used to seed the property default value; "" when none configured.
Private Function GetFirstCoupeTypeValue() As String
    On Error GoTo ErrorHandler

    GetFirstCoupeTypeValue = ""

    Dim vals() As String
    vals = Split(GetCoupeTypeListRaw(), ARESConstants.ARES_VAR_DELIMITER)
    If UBound(vals) >= LBound(vals) Then GetFirstCoupeTypeValue = Trim(vals(LBound(vals)))
    Exit Function

ErrorHandler:
    GetFirstCoupeTypeValue = ""
End Function

'######################################################################################################################
'                              GENERIC LIBRARY HELPERS (reusable, schema-agnostic)
'######################################################################################################################

' Find an ItemTypeLibrary by name (no creation). Returns Nothing if absent.
Public Function FindItemTypeLibrary(Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As ItemTypeLibrary
    On Error GoTo ErrorHandler

    Dim ItemLibs As ItemTypeLibraries
    Set ItemLibs = New ItemTypeLibraries
    Set FindItemTypeLibrary = ItemLibs.FindByName(LibraryName)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.FindItemTypeLibrary"
    Set FindItemTypeLibrary = Nothing
End Function

' Delete an ItemTypeLibrary by name. Returns True when a library was found and deleted.
Public Function DeleteItemTypeLibrary(Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
    On Error GoTo ErrorHandler

    DeleteItemTypeLibrary = False

    Dim ITL As ItemTypeLibrary
    Set ITL = FindItemTypeLibrary(LibraryName)
    If Not ITL Is Nothing Then
        ITL.DeleteLib
        DeleteItemTypeLibrary = True
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.DeleteItemTypeLibrary"
    DeleteItemTypeLibrary = False
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
