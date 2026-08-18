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
'              The managed property names are ENUMERATED FROM THE DGNLIB ITSELF (GetCustomPropertyNames)
'              - each ItemType in the "ARES" library is one custom property, its name being BOTH the
'              ItemType name and the property name. The library IS the list: authoring an ItemType (+ its
'              value list) in the DGNLib is all it takes for ARES to know about it - no config var to
'              keep in sync, no code change.
'
'              It also owns the MicroStation-side Item Type STATE refresh (RefreshItemTypes):
'              MicroStation reads MS_DGNLIBLIST only at boot, so a DGNLib deployed or edited
'              afterwards needs an explicit refresh to become visible - and the round trip to the
'              library itself (OpenCustomPropertyLibrary), which opens the DGNLib FILE and raises the
'              Item Types dialog on it so the definitions can be edited.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, Config, ErrorHandlerClass (global ErrorHandler)

Option Explicit

' The DGNLib FILE that ships the ARES ItemTypes. NOT the same thing as the ItemTypeLibrary NAME
' (ARESConstants.ARES_NAME_LIBRARY_TYPE = "ARES"): this is the file on disk, that is the library
' authored inside it. Change this const if the file is ever renamed.
Private Const DGNLIB_FILE_NAME As String = "ARES_Custom_Properties.dgnlib"
' Where the ARES installer deploys it (it also writes "MS_DGNLIBLIST > c:/ares/rsc/*.dgnlib"). Used only
' as a last resort, when MS_DGNLIBLIST is undefined or points somewhere the file is not.
Private Const DGNLIB_FALLBACK_DIR As String = "C:\ARES\Rsc"
' MicroStation's own search list for DGN libraries: ";"-separated entries, each a file, a folder or a
' wildcard pattern.
Private Const MS_DGNLIBLIST_VAR As String = "MS_DGNLIBLIST"
Private Const DGNLIBLIST_SEPARATOR As String = ";"

'######################################################################################################################
'                              MANAGED PROPERTY NAMES (enumerated from the DGNLib)
'######################################################################################################################

' The ARES custom-property names ARES manages: every ItemType of the "ARES" ItemTypeLibrary, read from
' the library itself (one ItemType per property, so the ItemType name IS the property name). The DGNLib
' is the single source of truth - authoring an ItemType there is enough; nothing to declare elsewhere.
'
' Enumerated with the documented successive-Find idiom - ItemTypeLibrary.Find(NamePattern[, previous])
' returns the next ItemType, and "*" walks them all (mvba-docs/03-methods/Find_Method.md: "You can
' successively find all ItemTypes by giving the name pattern, '*'"); each name comes from
' ItemType.ItemTypeName (mvba-docs/04-properties/ItemTypeName_Property.md).
'
' ORDER is the library's own enumeration order (authoring order in the DGNLib) - NOT sorted, and not
' guaranteed stable across edits. No consumer depends on it: Zone Export uses the array as a membership
' set and as combo content.
'
' 0-based array, ALWAYS allocated. Library missing (no DGNLib deployed / not yet refreshed) or holding no
' ItemType -> the one-empty-entry array [""], the same shape an empty list used to produce, which both
' consumers already absorb (combo skips empty names; membership test simply fails -> "Zone <n>" labels).
Public Function GetCustomPropertyNames() As String()
    On Error GoTo ErrorHandler

    Dim names() As String
    Dim ITL As ItemTypeLibrary
    Dim oItem As ItemType
    Dim n As Long

    ReDim names(0 To 0)
    names(0) = ""
    n = 0

    ' LOAD-BEARING: FindItemTypeLibrary is called with NO argument, so it resolves the USER-facing library
    ' (ARES_NAME_LIBRARY_TYPE) only. This is what keeps the internal ARES_SYS library out of every
    ' user-facing enumeration (Zone Export's property picker, the tag/calc editors, the render token
    ' validation). Do NOT parameterise this call.
    Set ITL = FindItemTypeLibrary()
    If ITL Is Nothing Then
        GetCustomPropertyNames = names
        Exit Function
    End If

    ' Pick up ItemTypes authored since this cache was built (Refresh takes no argument on an
    ' ItemTypeLibrary - mvba-docs/03-methods/Refresh_Method.md). Best-effort: a refresh fault must not
    ' cost us the enumeration that follows.
    On Error Resume Next
    ITL.Refresh
    On Error GoTo ErrorHandler

    Do
        Set oItem = ITL.Find("*", oItem)
        If oItem Is Nothing Then Exit Do
        If Len(Trim(oItem.ItemTypeName)) > 0 Then
            ReDim Preserve names(0 To n)
            names(n) = oItem.ItemTypeName
            n = n + 1
        End If
    Loop

    GetCustomPropertyNames = names
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetCustomPropertyNames"
    ReDim names(0 To 0)
    names(0) = ""
    GetCustomPropertyNames = names
End Function

'######################################################################################################################
'                              ITEM TYPE STATE REFRESH (MicroStation side)
'######################################################################################################################

' Force MicroStation to re-read its Item Type state, so ItemTypes deployed or edited in a DGNLib become
' visible without restarting the session. MicroStation only scans MS_DGNLIBLIST at boot; this key-in is
' the supported way to refresh that state afterwards. Note: SendKeyin is synchronous (it returns once
' MicroStation has processed the key-in - mvba-docs/03-methods/SendKeyin_Method.md).
'
' Idempotent and deliberately silent: refreshing an already-current state is a no-op, and this is a
' background consistency step, so there is no status message and no translation key.
Public Sub RefreshItemTypes()
    On Error GoTo ErrorHandler

    ' The UPDATEALL key-in only takes effect while the Item Types dialog is OPEN (live-established by
    ' Asketyll, 2026-08-10) - hence the open / update / close sandwich. Works inline from the DGN-open
    ' event too (the earlier on-open failure was the missing dialog, not timing).
    CadInputQueue.SendKeyin "DIALOG ITEMTYPE OPEN"
    CadInputQueue.SendKeyin "ITEMTYPE DIALOG UPDATEALL"
    CadInputQueue.SendKeyin "DIALOG ITEMTYPE CLOSE"

    ' Restore the default command state after the key-ins (the documented SendKeyin pattern - see the
    ' CadInputQueue example in mvba-docs; same call as RegionSplitLocate).
    CommandState.StartDefaultCommand
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.RefreshItemTypes"
End Sub

'######################################################################################################################
'                              DGNLIB ROUND TRIP (edit the definitions)
'######################################################################################################################

' Open the DGNLib that holds the ARES ItemTypes, then raise the Item Types dialog on it, so the user can
' edit the custom-property definitions (add an ItemType, extend a value list) without navigating there by
' hand. Returns False when the library file cannot be located - the caller owns the user message.
'
' MicroStation supports ONE open design file, so OpenDesignFile closes the working file first
' (mvba-docs/03-methods/OpenDesignFile_Method.md; on error the original file is left open). Read-write,
' since editing is the whole point. When the user re-opens the working file afterwards, DGNOpenClose's
' OnDesignFileOpened already calls RefreshItemTypes - so the edits are picked up without a restart and
' the edit loop closes itself.
'
' Unlike RefreshItemTypes' open/update/close sandwich, the dialog is left OPEN here: the user edits in it.
Public Function OpenCustomPropertyLibrary() As Boolean
    On Error GoTo ErrorHandler

    OpenCustomPropertyLibrary = False

    Dim sPath As String
    sPath = FindCustomPropertyLibraryPath()
    If Len(sPath) = 0 Then Exit Function

    ' Already sitting in the library: re-opening it would be a pointless round trip - just raise the dialog.
    If Not IsActiveDesignFile(sPath) Then OpenDesignFile sPath, False

    CadInputQueue.SendKeyin "DIALOG ITEMTYPE OPEN"
    ' Restore the default command state after the key-in (the documented SendKeyin pattern, as in
    ' RefreshItemTypes - it does not close the dialog).
    CommandState.StartDefaultCommand

    OpenCustomPropertyLibrary = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.OpenCustomPropertyLibrary"
    OpenCustomPropertyLibrary = False
End Function

' Full path of the ARES DGNLib, or "" when it cannot be found.
'
' Resolved through MS_DGNLIBLIST - the list MicroStation itself scans - rather than from the resolved
' ItemTypeLibrary: the MVBA ItemTypeLibrary object exposes no source file at all (LibName / Write /
' AddItemType / Find / GetItemTypeByName / GetSchemaAccessString / RemoveItemType / DeleteLib / Refresh -
' mvba-docs/02-objects/ItemTypeLibrary_Object.md), so there is no path to read back from it.
' Each MS_DGNLIBLIST entry may be a file, a folder or a wildcard pattern, so every entry is probed both as
' a folder and as a path whose parent folder holds the library. Falls back to the installer's own
' deployment folder when the list yields nothing.
Public Function FindCustomPropertyLibraryPath() As String
    On Error GoTo ErrorHandler

    FindCustomPropertyLibraryPath = ""

    Dim sList As String
    Dim entries() As String
    Dim sHit As String
    Dim i As Long

    sList = Config.GetVar(MS_DGNLIBLIST_VAR)          ' expanded value; ARES_NAVD when undefined

    If sList <> ARESConstants.ARES_NAVD Then
        entries = Split(sList, DGNLIBLIST_SEPARATOR)
        For i = LBound(entries) To UBound(entries)
            sHit = ResolveDgnLibEntry(entries(i))
            If Len(sHit) > 0 Then
                FindCustomPropertyLibraryPath = sHit
                Exit Function
            End If
        Next i
    End If

    FindCustomPropertyLibraryPath = LibraryFileIn(DGNLIB_FALLBACK_DIR)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.FindCustomPropertyLibraryPath"
    FindCustomPropertyLibraryPath = ""
End Function

' Probe ONE MS_DGNLIBLIST entry for the ARES DGNLib. The entry is tried as a folder first, then as a file
' or wildcard pattern (in which case its parent folder is the one to look in) - which also covers an entry
' naming the library file itself. Returns the full path or "".
Private Function ResolveDgnLibEntry(ByVal sEntry As String) As String
    On Error GoTo ErrorHandler

    ResolveDgnLibEntry = ""

    ' MS_DGNLIBLIST entries commonly use forward slashes (the ARES installer writes "c:/ares/rsc/*.dgnlib").
    Dim s As String
    s = Trim(Replace(sEntry, "/", "\"))
    Do While Len(s) > 0
        If Right(s, 1) <> "\" Then Exit Do
        s = Left(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then Exit Function

    ResolveDgnLibEntry = LibraryFileIn(s)
    If Len(ResolveDgnLibEntry) > 0 Then Exit Function

    Dim nSep As Long
    nSep = InStrRev(s, "\")
    If nSep > 1 Then ResolveDgnLibEntry = LibraryFileIn(Left(s, nSep - 1))
    Exit Function

ErrorHandler:
    ' Silent fail-closed: an unreachable path (offline network share) just means "not here".
    ResolveDgnLibEntry = ""
End Function

' Full path of the ARES DGNLib inside sFolder, or "" when it is not there. Dir is always called WITH an
' argument, so it starts a fresh search and cannot disturb an enumeration running elsewhere.
Private Function LibraryFileIn(ByVal sFolder As String) As String
    On Error GoTo ErrorHandler

    LibraryFileIn = ""
    If Len(Trim(sFolder)) = 0 Then Exit Function

    Dim sPath As String
    sPath = sFolder & "\" & DGNLIB_FILE_NAME
    If Len(Dir(sPath)) > 0 Then LibraryFileIn = sPath
    Exit Function

ErrorHandler:
    ' Silent fail-closed (like ResolveDgnLibEntry): a path Dir cannot even probe is simply not a hit.
    LibraryFileIn = ""
End Function

' True when sPath is the design file currently open (case-insensitive full-path compare - FullName is
' path + name + extension, mvba-docs/04-properties/FullName_Property.md).
Private Function IsActiveDesignFile(ByVal sPath As String) As Boolean
    On Error GoTo ErrorHandler

    IsActiveDesignFile = False
    If ActiveDesignFile Is Nothing Then Exit Function

    IsActiveDesignFile = (StrComp(ActiveDesignFile.FullName, sPath, vbTextCompare) = 0)
    Exit Function

ErrorHandler:
    ' Silent fail-closed: "cannot tell" means "not the active file", so the caller opens it - harmless.
    IsActiveDesignFile = False
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

' Read-only attach check: True when El carries the named ItemType from LibraryName. Thin wrapper over
' Element.Items.HasItems (verified in mvba-docs/03-methods/HasItems_Method.md, signature
' Boolean = object.HasItems(Libname [, ItemTypename])) after a cache Refresh (mvba-docs/03-methods/
' Refresh_Method.md, Items.Refresh Libname) so a same-pass attach is visible. Unlike inferring absence
' from GetPropertyValueFromElement returning Null (which cannot distinguish "not attached" from
' "attached but empty"), this reports the unambiguous ATTACHMENT state - the frontier the value engine
' (PropertyCalculation) uses to write a value only where the target property is already attached.
' No model write (Refresh is a cache refresh only). Standard error pattern -> False on fault.
Public Function IsItemAttachedToElement(ByVal El As element, ByVal ItemName As String, Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE) As Boolean
    On Error GoTo ErrorHandler

    IsItemAttachedToElement = False
    If El Is Nothing Then Exit Function
    If Len(ItemName) = 0 Then Exit Function

    El.Items.Refresh LibraryName
    IsItemAttachedToElement = El.Items.HasItems(LibraryName, ItemName)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.IsItemAttachedToElement"
    IsItemAttachedToElement = False
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
' Tolerant of a hand-authored DGNLib whose real property name differs from the ItemType name: it
' tries the caller's access string first (fast path), and only if that RAISES or yields Null does it
' fall back to the ItemType definition's actual property name(s). ARES item types carry a single
' property, so "the first property that yields a value" is unambiguous. A genuinely value-less item
' returns Null SILENTLY (the normal "no value" case) — no parasitic log.
' ItemName omitted defaults to PropertyName (ARES convention: the ItemType name IS the property name),
' so the read addresses the property's OWN item — never "the first attached item" of a multi-item element.
'
' bNoFallback = True suppresses that single-property fallback and returns Null instead. Required by any
' MULTI-property ItemType (ARES_SYS/ARES_Render carries SchemaVersion + Entries): the fallback returns
' "the first property that yields a value", so reading an EMPTY SchemaVersion would silently hand back
' Entries. It is the LAST parameter on purpose — the hot call sites are positional
' (PropertyCalculation.bas:1620), so inserting it before LibraryName would land the item name in it.
Public Function GetPropertyValueFromElement(ByVal El As element, ByVal PropertyName As String, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ByVal bNoFallback As Boolean = False) As Variant
    On Error GoTo ErrorHandler

    GetPropertyValueFromElement = Null
    If Len(PropertyName) = 0 Then Exit Function    ' fail-closed: no property named, nothing to address

    ' Address PropertyName's own item when the caller named none: resolving the FIRST attached ARES item
    ' (Items.Find "*") reads the WRONG item on an element carrying several ARES properties.
    If Len(ItemName) = 0 Then ItemName = PropertyName

    Dim oHandler As ItemTypePropertyHandler
    Set oHandler = GetItemTypePropertyHandlerFromElement(El, ItemName, LibraryName)
    If oHandler Is Nothing Then Exit Function

    ' Fast path: the caller's access string. GetPropertyValue RAISES on an unknown access string, so
    ' isolate it under On Error Resume Next (the mismatch is expected for some DGNLibs — stay silent).
    Dim vVal As Variant
    vVal = Null
    On Error Resume Next
    vVal = oHandler.GetPropertyValue(PropertyName)
    On Error GoTo ErrorHandler
    If Not IsNull(vVal) Then
        GetPropertyValueFromElement = vVal
        Exit Function
    End If

    ' Strict mode: no fallback, the absent/empty property reads as Null (never another property's value).
    If bNoFallback Then Exit Function

    ' Fallback: resolve the real property name from the ItemType definition and retry.
    GetPropertyValueFromElement = GetFirstPropertyValue(oHandler, LibraryName)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetPropertyValueFromElement"
    GetPropertyValueFromElement = Null
End Function

' Fallback for GetPropertyValueFromElement: iterate the attached ItemType's real ItemTypeProperty
' names (from the definition) and return the value of the first one the handler can read. Resolves
' the ItemType from the handler's own ItemTypeName (robust when the caller passed no ItemName). Each
' GetPropertyValue is isolated (silent) since a mismatch/absence must not log. Returns Null when no
' property yields a value. ARES item types are single-property, so the first hit is unambiguous.
Private Function GetFirstPropertyValue(ByVal oHandler As ItemTypePropertyHandler, ByVal LibraryName As String) As Variant
    On Error GoTo ErrorHandler

    GetFirstPropertyValue = Null

    Dim ITL As ItemTypeLibrary
    Set ITL = FindItemTypeLibrary(LibraryName)
    If ITL Is Nothing Then Exit Function

    Dim oItem As ItemType
    Set oItem = ITL.GetItemTypeByName(oHandler.ItemTypeName)
    If oItem Is Nothing Then Exit Function

    Dim oProp As ItemTypeProperty
    Dim vVal  As Variant
    Do
        Set oProp = oItem.Find("*", oProp)
        If oProp Is Nothing Then Exit Do
        vVal = Null
        On Error Resume Next
        vVal = oHandler.GetPropertyValue(oProp.PropertyName)
        On Error GoTo ErrorHandler
        If Not IsNull(vVal) Then
            GetFirstPropertyValue = vVal
            Exit Function
        End If
    Loop
    Exit Function

ErrorHandler:
    GetFirstPropertyValue = Null
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetFirstPropertyValue"
End Function

' Write a property value to an element. Returns True on success.
' Tolerant of a hand-authored DGNLib whose real property name differs from the ItemType name (the
' write-side mirror of GetPropertyValueFromElement): it tries the caller's access string first (fast
' path), and only if that RAISES or returns False does it fall back to the ItemType definition's
' actual property name(s). ARES item types carry a single property, so "the first property that
' accepts the write" is unambiguous. Returns False only when neither the given name nor any real
' property name accepts the value (a genuinely constrained property — picklist / type mismatch).
' ItemName omitted defaults to PropertyName (ARES convention: the ItemType name IS the property name),
' so the write addresses the property's OWN item — never "the first attached item" of a multi-item
' element, where the fallback would land the value in the WRONG property.
'
' bNoFallback = True suppresses that fallback and returns False instead — mandatory on a MULTI-property
' ItemType (ARES_SYS/ARES_Render), where "the first property that accepts the write" can silently land
' the value in the wrong field AND still report success. LAST parameter on purpose: the hot call sites
' are positional (PropertyCalculation.bas:1626/:1638), so any earlier position would break them.
Public Function SetPropertyValueToElement(ByVal El As element, ByVal PropertyName As String, ByVal PropertyValue As Variant, Optional ByVal ItemName As String = "", Optional ByVal LibraryName As String = ARESConstants.ARES_NAME_LIBRARY_TYPE, Optional ByVal bNoFallback As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    SetPropertyValueToElement = False
    If Len(PropertyName) = 0 Then Exit Function    ' fail-closed: no property named, nothing to address

    ' Address PropertyName's own item when the caller named none (see GetPropertyValueFromElement).
    If Len(ItemName) = 0 Then ItemName = PropertyName

    Dim oHandler As ItemTypePropertyHandler
    Set oHandler = GetItemTypePropertyHandlerFromElement(El, ItemName, LibraryName)
    If oHandler Is Nothing Then Exit Function

    ' Fast path: the caller's access string. SetPropertyValue RAISES on an unknown access string, so
    ' isolate it under On Error Resume Next (the mismatch is expected for some DGNLibs — stay silent).
    Dim bOk As Boolean
    bOk = False
    On Error Resume Next
    bOk = oHandler.SetPropertyValue(PropertyName, PropertyValue)
    On Error GoTo ErrorHandler
    If bOk Then
        SetPropertyValueToElement = True
        Exit Function
    End If

    ' Strict mode: no fallback, a refused write stays refused (never redirected to another property).
    If bNoFallback Then Exit Function

    ' Fallback: resolve the real property name from the ItemType definition and retry.
    SetPropertyValueToElement = SetFirstPropertyValue(oHandler, LibraryName, PropertyValue)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.SetPropertyValueToElement"
    SetPropertyValueToElement = False
End Function

' Fallback for SetPropertyValueToElement: iterate the attached ItemType's real ItemTypeProperty names
' (from the definition) and write PropertyValue to the first one the handler accepts. Resolves the
' ItemType from the handler's own ItemTypeName (robust when the caller passed no ItemName). Each
' SetPropertyValue is isolated (silent) since a wrong name RAISES and must not log. Returns True on the
' first accepted write, False when no property accepts the value. ARES item types are single-property,
' so the first success is unambiguous. Structural mirror of GetFirstPropertyValue.
Private Function SetFirstPropertyValue(ByVal oHandler As ItemTypePropertyHandler, ByVal LibraryName As String, ByVal PropertyValue As Variant) As Boolean
    On Error GoTo ErrorHandler

    SetFirstPropertyValue = False

    Dim ITL As ItemTypeLibrary
    Set ITL = FindItemTypeLibrary(LibraryName)
    If ITL Is Nothing Then Exit Function

    Dim oItem As ItemType
    Set oItem = ITL.GetItemTypeByName(oHandler.ItemTypeName)
    If oItem Is Nothing Then Exit Function

    Dim oProp As ItemTypeProperty
    Dim bOk   As Boolean
    Do
        Set oProp = oItem.Find("*", oProp)
        If oProp Is Nothing Then Exit Do
        bOk = False
        On Error Resume Next
        bOk = oHandler.SetPropertyValue(oProp.PropertyName, PropertyValue)
        On Error GoTo ErrorHandler
        If bOk Then
            SetFirstPropertyValue = True
            Exit Function
        End If
    Loop
    Exit Function

ErrorHandler:
    SetFirstPropertyValue = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.SetFirstPropertyValue"
End Function
