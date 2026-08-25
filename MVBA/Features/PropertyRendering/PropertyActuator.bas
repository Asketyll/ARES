' Module: PropertyActuator
' Description: Fourth engine (epic 16) - reverses the property system's read doctrine: writes a graphic
'              attribute (Color/Level) FROM a custom property value, SELF reaction only.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler,
'               PropertyCalculation, GetElements, LangManager, ErrorHandlerClass (global ErrorHandler),
'               CallStackClass (global CallStack)
Option Explicit

'######################################################################################################################
'                                          DOCTRINE
'######################################################################################################################
' Pilot properties are FIXED/RESERVED (ARES_PROP_COLOR/ARES_PROP_LVL), never user-configurable - a
' configurable picker once let a length value get painted as a raw color index in production. Attach via
' PropertyTagging's "|" grammar, no dedicated rules var. Any attach/detach call here is a review BLOCKER.

'######################################################################################################################
'                                          ONE-SHOT STATUS GUARDS
'######################################################################################################################
' Status-bar only, never logged (expected refusals) - reset at the start of every ProcessElement.
Private mbColorInvalidShown As Boolean
Private mbLevelInvalidShown As Boolean
Private mbSelfRatchetShown As Boolean

' Session-wide fail-closed latch: a repeated attribute-write failure (locked DGN, rights) disables the
' actuator instead of retrying every pass. Cleared by RefreshActuatorState.
Private mbWriteDisabled As Boolean

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

' Master switch for the Color attribute. Lazily initialises ARESConfig, fail-closed False.
Public Function IsColorEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsColorEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_ACTUATE_COLOR Is Nothing Then Exit Function
    IsColorEnabled = CBool(ARESConfig.ARES_ACTUATE_COLOR.Value)
    Exit Function

ErrorHandler:
    IsColorEnabled = False
End Function

' Master switch for the Level attribute. Lazily initialises ARESConfig, fail-closed False.
Public Function IsLevelEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsLevelEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_ACTUATE_LEVEL Is Nothing Then Exit Function
    IsLevelEnabled = CBool(ARESConfig.ARES_ACTUATE_LEVEL.Value)
    Exit Function

ErrorHandler:
    IsLevelEnabled = False
End Function

' Combined master switch - the term IsAnyFeatureEnabled needs so a config running ONLY the actuator still
' reaches ElementChangeHandler.ProcessElement.
Public Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler
    IsEnabled = IsColorEnabled() Or IsLevelEnabled()
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

' Clears the write-disabled latch - used by the options panel (a user just toggled a switch) and the
' test harness. No rules to re-parse (pilot properties are fixed constants).
Public Sub RefreshActuatorState()
    mbWriteDisabled = False
End Sub

' Depth-0 hook, called from ElementChangeHandler.ProcessElement (after PropertyCalculation, before
' PropertyRendering) when IsEnabled. Resets one-shot flags, then reacts SELF: read own pilot property,
' compare to own attribute, write only if different, once per enabled attribute.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler
    Dim bStackPushed As Boolean

    mbColorInvalidShown = False
    mbLevelInvalidShown = False
    mbSelfRatchetShown = False

    If mbWriteDisabled Then Exit Sub
    If oEl Is Nothing Then Exit Sub
    If Not oEl.IsGraphical Then Exit Sub
    If oEl.IsLocked Then Exit Sub    ' never paint a locked or reference-owned element
    If Not IsColorEnabled() And Not IsLevelEnabled() Then Exit Sub

    ' A trigger must never be painted by the value it is itself the source of. NOT redundant with
    ' IsSelfSourceRatchet below: a Lvl*-fed rule's SourceKind never equals csColor/csLvl, so ratchet alone
    ' misses a Lvl*-trigger that also carries the pilot property.
    If IsProtectedTriggerCell(oEl) Then Exit Sub
    If IsProtectedTriggerLevel(oEl) Then Exit Sub

    CallStack.Push "PropertyActuator.ProcessElement", oEl
    bStackPushed = True

    If IsColorEnabled() Then ActuateColor oEl, ARES_PROP_COLOR
    If IsLevelEnabled() Then ActuateLevel oEl, ARES_PROP_LVL
    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ProcessElement"
    If bStackPushed Then CallStack.Pop
End Sub

'######################################################################################################################
'                                          TRIGGER EXCLUSION (own fail-closed wrappers, Cell and Level)
'######################################################################################################################

' PropertyCalculation.IsTriggerCell fails OPEN on error (paints by default). This wrapper fails CLOSED
' instead - the actuator's own copy of the guard, protecting the element on any fault.
Private Function IsProtectedTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsProtectedTriggerCell = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    IsProtectedTriggerCell = PropertyCalculation.IsTriggerCell(oEl)
    Exit Function

ErrorHandler:
    IsProtectedTriggerCell = True    ' fail-closed: assume trigger, protect it
End Function

' Same fail-closed wrapper for PropertyCalculation.IsTriggerLevel. No element-type restriction: a Level
' trigger can be any graphical element, not just a cell.
Private Function IsProtectedTriggerLevel(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsProtectedTriggerLevel = False
    If oEl Is Nothing Then Exit Function
    IsProtectedTriggerLevel = PropertyCalculation.IsTriggerLevel(oEl)
    Exit Function

ErrorHandler:
    IsProtectedTriggerLevel = True    ' fail-closed: assume trigger, protect it
End Function

'######################################################################################################################
'                                          ATTRIBUTE ACTUATION
'######################################################################################################################

' Writes P's color onto oEl.Color - EXCEPT a CellElement's header, never written (see below); on a cell,
' only sub-elements still matching the header's OLD color are repainted, preserving FillMode=2 fill.
'
' A cell header's Color is never written: it triggers a MicroStation-native ByCellColor cascade that
' silently resets sub-elements' FillMode/FillColor. Already tried and failed: handle refresh, descendant
' snapshot/restore. Root cause of a real fill-corruption bug - do not reinstate or retry those workarounds
' (cahier-des-charges-groupcolor-fillmode.md).
Private Sub ActuateColor(ByVal oEl As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, P) Then Exit Sub

    ' Refuse a self-fed pilot property - see IsSelfSourceRatchet.
    If IsSelfSourceRatchet(P, oEl, PropertyCalculation.csColor) Then
        ReportSelfRatchet
        Exit Sub
    End If

    Dim vVal As Variant
    vVal = CustomPropertyHandler.GetPropertyValueFromElement(oEl, P, P)
    If IsNull(vVal) Then Exit Sub
    If Len(CStr(vVal)) = 0 Then Exit Sub

    If Not IsNumeric(vVal) Then
        ReportColorInvalid
        Exit Sub
    End If

    Dim targetColor As Long
    targetColor = CLng(vVal)
    If targetColor < 0 Or targetColor > 255 Then
        ReportColorInvalid
        Exit Sub
    End If

    Dim oldColor As Long
    oldColor = oEl.Color

    If Not oEl.IsCellElement Then
        If oldColor <> targetColor Then
            oEl.Color = targetColor
            oEl.Rewrite
        End If
    End If

    If oEl.IsCellElement Then
        Dim ELEnum As ElementEnumerator
        Dim subEl As element
        Set ELEnum = oEl.AsCellElement.GetSubElements
        Do While ELEnum.MoveNext
            Set subEl = ELEnum.Current
            If Not subEl Is Nothing Then
                If subEl.IsGraphical Then
                    If subEl.Color = oldColor And oldColor <> targetColor Then
                        ' Color/FillColor are independent properties - preserve the fill across the write.
                        If subEl.IsClosedElement Then
                            If subEl.AsClosedElement.FillMode = 2 Then
                                Dim savedFillColor As Long
                                savedFillColor = subEl.AsClosedElement.fillcolor
                                subEl.Color = targetColor
                                subEl.AsClosedElement.fillcolor = savedFillColor
                            Else
                                subEl.Color = targetColor
                            End If
                        Else
                            subEl.Color = targetColor
                        End If
                        subEl.Rewrite
                    End If
                End If
            End If
        Loop
    End If
    Exit Sub

ErrorHandler:
    DisableAfterWriteFailure
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ActuateColor"
End Sub

' Writes P's value onto oEl.Level (a Level name, never auto-created), same repaint rule as ActuateColor.
' Unlike ActuateColor, the header Level IS written - Level has no ByCellColor-style cascade sentinel.
' An unresolvable name (-> ReportLevelInvalid) and a resolved-but-unwritable level (-> this Sub's
' ErrorHandler) are separate failure paths, not one.
Private Sub ActuateLevel(ByVal oEl As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, P) Then Exit Sub

    If IsSelfSourceRatchet(P, oEl, PropertyCalculation.csLvl) Then
        ReportSelfRatchet
        Exit Sub
    End If

    Dim vVal As Variant
    vVal = CustomPropertyHandler.GetPropertyValueFromElement(oEl, P, P)
    If IsNull(vVal) Then Exit Sub
    Dim sLevelName As String
    sLevelName = CStr(vVal)
    If Len(sLevelName) = 0 Then Exit Sub

    Dim targetLevel As Level
    Set targetLevel = GetElements.GetLevel(sLevelName, False)
    If targetLevel Is Nothing Then
        ReportLevelInvalid
        Exit Sub
    End If

    Dim oldLevelName As String
    If Not oEl.Level Is Nothing Then oldLevelName = oEl.Level.Name

    If oldLevelName <> targetLevel.Name Then
        oEl.Level = targetLevel
        oEl.Rewrite
    End If

    If oEl.IsCellElement Then
        Dim ELEnum As ElementEnumerator
        Dim subEl As element
        Dim sSubLevelName As String
        Set ELEnum = oEl.AsCellElement.GetSubElements
        Do While ELEnum.MoveNext
            Set subEl = ELEnum.Current
            If Not subEl Is Nothing Then
                If subEl.IsGraphical Then
                    sSubLevelName = ""
                    If Not subEl.Level Is Nothing Then sSubLevelName = subEl.Level.Name
                    If sSubLevelName = oldLevelName And oldLevelName <> targetLevel.Name Then
                        subEl.Level = targetLevel
                        subEl.Rewrite
                    End If
                End If
            End If
        Loop
    End If
    Exit Sub

ErrorHandler:
    DisableAfterWriteFailure
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ActuateLevel"
End Sub

' True when P's calc source on oEl is the same SELF source as TargetKind - the pilot property would then
' only ever reflect what this actuator itself last wrote. Checked per element/attribute at write time,
' since a calc rule's source can vary by condition.
Private Function IsSelfSourceRatchet(ByVal P As String, ByVal oEl As element, ByVal TargetKind As CalcSource) As Boolean
    On Error GoTo ErrorHandler

    IsSelfSourceRatchet = False

    Dim SourceKind As CalcSource
    Dim SourceArg As String
    Dim sCanonical As String
    If PropertyCalculation.GetCalcRuleForProperty(P, oEl, SourceKind, SourceArg, sCanonical) Then
        IsSelfSourceRatchet = (SourceKind = TargetKind)
    End If
    Exit Function

ErrorHandler:
    IsSelfSourceRatchet = True    ' fail-closed: refuse the write if we cannot tell
End Function

'######################################################################################################################
'                                          STATUS REPORTING (one-shot, status bar only, never logged)
'######################################################################################################################

Private Sub ReportColorInvalid()
    On Error Resume Next
    If mbColorInvalidShown Then Exit Sub
    mbColorInvalidShown = True
    LangManager.ShowStatusT "ActuatorColorInvalid"
End Sub

Private Sub ReportLevelInvalid()
    On Error Resume Next
    If mbLevelInvalidShown Then Exit Sub
    mbLevelInvalidShown = True
    LangManager.ShowStatusT "ActuatorLevelInvalid"
End Sub

Private Sub ReportSelfRatchet()
    On Error Resume Next
    If mbSelfRatchetShown Then Exit Sub
    mbSelfRatchetShown = True
    LangManager.ShowStatusT "ActuatorSelfRatchetRefused"
End Sub

' Session-wide fail-closed latch on a repeated attribute-write failure: idempotent, logs one English
' line, then every subsequent ProcessElement short-circuits until RefreshActuatorState clears it.
Private Sub DisableAfterWriteFailure()
    On Error Resume Next
    If mbWriteDisabled Then Exit Sub
    mbWriteDisabled = True
    ErrorHandler.HandleError "Property actuator: attribute write failed, actuation disabled for this session", 0, "", "PropertyActuator.ActuateAttribute"
End Sub
