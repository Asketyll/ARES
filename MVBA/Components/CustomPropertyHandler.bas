' Module: CustomPropertyHandler
' Description: Attaches, reads and writes ARES custom properties (MicroStation Item Types) on
'              elements, with silent error handling. Definitions/value lists live in a DGNLib (the
'              "ARES" ItemTypeLibrary), authored via the Item Types dialog - not created from VBA.
'              The DGNLib IS the list: GetCustomPropertyNames enumerates it directly, no config var
'              to keep in sync. Also owns the MicroStation-side refresh (RefreshItemTypes) and the
'              DGNLib edit round trip (OpenCustomPropertyLibrary).
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

' Every ItemType of the "ARES" ItemTypeLibrary, via the documented successive-Find idiom
' (Find("*", previous)). Order is the library's own authoring order, not guaranteed stable. Always
' returns a 0-based array; missing library or no ItemType -> the one-empty-entry array [""].
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

' Forces MicroStation to re-read Item Type state (it only scans MS_DGNLIBLIST at boot) AND attaches the
' two ARES libraries to the active model. Idempotent and deliberately silent - a background consistency
' step, no status message.
Public Sub RefreshItemTypes()
    On Error GoTo ErrorHandler

    ' The UPDATEALL key-in only takes effect while the Item Types dialog is OPEN - hence the open /
    ' update / close sandwich. Works inline from the DGN-open event too.
    CadInputQueue.SendKeyin "DIALOG ITEMTYPE OPEN"
    CadInputQueue.SendKeyin "ITEMTYPE DIALOG UPDATEALL"

    ' UPDATEALL only refreshes what the model already knows about: a DGN that has never seen the ARES
    ' libraries still shows nothing. SELECT + SAVE brings each library INTO the active model, which is
    ' what makes the properties usable there. Both are needed - ARES for the user-facing properties,
    ' ARES_SYS for the render bindings - and both are re-run on every DGN open, harmlessly, because
    ' the operation is idempotent and a file that already carries them is left as it is.
    CadInputQueue.SendKeyin "ITEMTYPE DIALOG SELECT " & ARESConstants.ARES_NAME_LIBRARY_TYPE
    CadInputQueue.SendKeyin "ITEMTYPE DIALOG SELECT " & ARESConstants.ARES_NAME_LIBRARY_SYS
    CadInputQueue.SendKeyin "ITEMTYPE DIALOG SAVE"

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

' Opens the DGNLib holding the ARES ItemTypes and raises the Item Types dialog on it for editing.
' Returns False when the library file cannot be located. MicroStation supports only ONE open design
' file, so this closes the working file first; DGNOpenClose.OnDesignFileOpened already calls
' RefreshItemTypes when the user reopens it, so the edit loop closes itself.
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

' Full path of the ARES DGNLib, or "" when not found. Resolved through MS_DGNLIBLIST (the MVBA
' ItemTypeLibrary object exposes no source path of its own). Each entry may be a file, folder or
' wildcard, probed both ways; falls back to the installer's deployment folder.
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

' Makes sure MS_DGNLIBLIST declares the ARES DGNLib, and returns True when it does (already, or after
' being added). MicroStation loads Item Type libraries ONLY from the DGNLibs on that list: on a station
' where it was never set - or set by a workspace that knows nothing of ARES - no ARES property exists at
' all, whatever the file on disk says.
'
' APPENDS, never replaces: the list is the site's, and ARES has no business dropping someone else's
' libraries. Two reads, deliberately: the EXPANDED value to decide whether the library is already
' covered (paths must resolve to be compared), and the RAW DEFINITION to append to, so that a site
' writing "$(SOME_LIB_DIR)/*.dgnlib" keeps its reference live instead of having it frozen to whatever
' it happened to expand to today.
'
' Nothing is written when an entry already resolves to the library (idempotent across sessions), nor
' when the file is missing from the installer's folder - there would be nothing to declare. Note that
' AddConfigurationVariable writes a User-level value, which takes precedence over a Project or System
' definition of the same variable from then on; appending the raw definition is what keeps that
' precedence harmless.
'
' Call at boot, BEFORE the DGN-open handler runs RefreshItemTypes - that key-in sandwich is what makes
' MicroStation act on the change within the running session.
Public Function EnsureLibraryInDgnLibList() As Boolean
    On Error GoTo ErrorHandler

    EnsureLibraryInDgnLibList = False

    Dim sList  As String
    Dim sRaw   As String
    Dim entries() As String
    Dim i      As Long
    Dim sEntry As String

    sList = Config.GetVar(MS_DGNLIBLIST_VAR)          ' expanded; ARES_NAVD when undefined

    ' Already covered? Any entry that resolves to the library is enough - that is exactly what
    ' MicroStation itself reads, so matching on the resolved file avoids duplicating an entry that
    ' names the same folder in another form (trailing slash, wildcard, forward slashes).
    If sList <> ARESConstants.ARES_NAVD Then
        entries = Split(sList, DGNLIBLIST_SEPARATOR)
        For i = LBound(entries) To UBound(entries)
            If Len(ResolveDgnLibEntry(entries(i))) > 0 Then
                EnsureLibraryInDgnLibList = True
                Exit Function
            End If
        Next i
    End If

    ' Not declared anywhere. Only the installer's folder can be declared blindly - if the library is not
    ' there either, stay silent rather than point the list at a file that does not exist.
    If Len(LibraryFileIn(DGNLIB_FALLBACK_DIR)) = 0 Then Exit Function

    ' Same wildcard form the installer writes, so an ARES-added entry is indistinguishable from its own.
    sEntry = DGNLIB_FALLBACK_DIR & "\*.dgnlib"

    sRaw = Config.GetVar(MS_DGNLIBLIST_VAR, False)
    If sRaw = ARESConstants.ARES_NAVD Or Len(Trim(sRaw)) = 0 Then
        EnsureLibraryInDgnLibList = Config.SetVar(MS_DGNLIBLIST_VAR, sEntry)
    Else
        EnsureLibraryInDgnLibList = Config.SetVar(MS_DGNLIBLIST_VAR, sRaw & DGNLIBLIST_SEPARATOR & sEntry)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.EnsureLibraryInDgnLibList"
    EnsureLibraryInDgnLibList = False
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

' Read-only attach check: True when El carries the named ItemType (after a cache Refresh so a same-pass
' attach is visible). Unlike a Null value read, this distinguishes "not attached" from "attached but
' empty" - the frontier PropertyCalculation uses to write only where a property is already attached.
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

' True when oItem has EXACTLY 2 members, named (case-insensitively, matching IsKnownProperty's convention)
' "X" and "Y" in either order - the shape a split-coordinate ItemType must have (Coord/CellCoord calc
' sources write X/Y independently into such an item instead of one combined "X;Y" string; PropertyRendering's
' "Prop[Name:X]"/"Prop[Name:Y]" token syntax reads one field of it). False for a 1-member item (today's
' single-field shape - unchanged behaviour, by construction, since every caller only special-cases a True
' result), False for 3+ members, False for 2 members not named X/Y. sXMember/sYMember are set to the
' member's REAL property name (its own casing, as declared in the DGNLib) so callers pass the exact access
' string on to GetPropertyValueFromElement/SetPropertyValueToElement - never assume the literal "X"/"Y"
' casing the caller asked about is what the ItemType actually declares.
' Single source of truth for "is this a split coordinate item" - do not re-derive this shape test elsewhere.
Public Function GetXYSplitMembers(ByVal oItem As ItemType, ByRef sXMember As String, ByRef sYMember As String) As Boolean
    On Error GoTo ErrorHandler

    GetXYSplitMembers = False
    sXMember = ""
    sYMember = ""
    If oItem Is Nothing Then Exit Function

    Dim oProp As ItemTypeProperty
    Dim nCount As Long
    Dim sName As String

    nCount = 0
    Do
        Set oProp = oItem.Find("*", oProp)
        If oProp Is Nothing Then Exit Do

        nCount = nCount + 1
        If nCount > 2 Then                            ' 3+ members - not a split-coordinate item
            sXMember = ""
            sYMember = ""
            Exit Function
        End If

        sName = oProp.PropertyName
        If StrComp(sName, "X", vbTextCompare) = 0 Then
            sXMember = sName
        ElseIf StrComp(sName, "Y", vbTextCompare) = 0 Then
            sYMember = sName
        Else
            sXMember = ""                             ' a member not named X or Y - not this shape
            sYMember = ""
            Exit Function
        End If
    Loop

    GetXYSplitMembers = (nCount = 2) And (Len(sXMember) > 0) And (Len(sYMember) > 0)
    If Not GetXYSplitMembers Then
        sXMember = ""
        sYMember = ""
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandler.GetXYSplitMembers"
    GetXYSplitMembers = False
    sXMember = ""
    sYMember = ""
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

' Reads a property value; Null when absent. Tolerant of a hand-authored DGNLib whose real property
' name differs from the ItemType name: tries the caller's access string first, then the ItemType's
' actual property name(s). bNoFallback=True suppresses that fallback - mandatory for a MULTI-property
' ItemType like ARES_Render, where it could silently return the wrong field.
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

' Fallback for GetPropertyValueFromElement: returns the value of the first real ItemTypeProperty the
' handler can read (ARES item types are single-property, so the first hit is unambiguous).
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

' Write-side mirror of GetPropertyValueFromElement: tries the caller's access string first, falls
' back to the ItemType's real property name(s) on a hand-authored DGNLib. bNoFallback=True suppresses
' that fallback - mandatory on a MULTI-property ItemType, where it could silently write the wrong field.
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

' Fallback for SetPropertyValueToElement, structural mirror of GetFirstPropertyValue: writes
' PropertyValue to the first real ItemTypeProperty the handler accepts.
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
