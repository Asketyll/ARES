' Module: PropertyActuator
' Description: Fourth engine (epic 16) - reverses the read doctrine of Tagging/Calculation/Rendering:
'              writes a graphic attribute (Color/Level) FROM a custom property value, SELF only.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler,
'               PropertyCalculation, GetElements, LangManager, ErrorHandlerClass (global ErrorHandler)
Option Explicit

'######################################################################################################################
'                                          DOCTRINE
'######################################################################################################################
' Tagging attaches. Calculation computes values. Rendering writes text. PropertyActuator writes graphic
' ATTRIBUTES (Color, Level in v1) - the fourth and last leg of the custom-property system, and the first
' one that reverses direction: property VALUE -> graphic attribute, instead of attribute -> property.
'
' It is a pure SELF reaction, not a group engine. Its source, CellColor[pattern]/CellLvl[pattern], is
' already a PUSHABLE Cell* source in PropertyCalculation: PushCellDerivedValuesToMembers already keeps the
' pilot property fresh on every OTHER group member that carries it attached (frontier IsItemAttachedToElement).
' By the time this module's ProcessElement runs on a given element, that element's own pilot property is
' already up to date - so the actuator only has to read ITS OWN pilot property and reflect it onto ITS OWN
' attribute. No PUSH, no PULL, no Link.GetLink here - that group propagation stays PropertyCalculation's job.
'
' Pilot properties are FIXED and RESERVED (ARESConstants.ARES_PROP_COLOR/ARES_PROP_LVL), not user-configurable
' (revised 2026-08-20 - see cahier des charges §8): a real-world test with a LENGTH property picked as the
' color pilot via an earlier configurable-picker design painted the length value itself as a raw color index.
' A fixed, reserved property name closes that error class by construction. Attachment goes through
' PropertyTagging's existing "|" multi-property grammar on the user's own tagging rules (e.g.
' "Lvl[WALLS]=Commune|ARES_Color") - no dedicated actuator rules variable.
'
' Any AttachItemToElement/RemoveItemFromElement/direct SetPropertyValueToElement call in this module is a
' review BLOCKER, same doctrine line as the other three engines.

'######################################################################################################################
'                                          ONE-SHOT STATUS GUARDS
'######################################################################################################################
' Reset at the start of every ProcessElement (same shape as PropertyCalculation's
' mbRejectedShown/mbNoTargetShown - status-bar only, never logged, an EXPECTED refusal per project doctrine).
Private mbColorInvalidShown As Boolean
Private mbLevelInvalidShown As Boolean
Private mbSelfRatchetShown As Boolean

' Session-wide fail-closed latch (cahier §4.6): a repeated COM attribute-write failure (locked DGN, rights)
' disables the actuator for the rest of the session instead of retrying/spamming on every hot-path pass.
' Mirrors PropertyRendering.mbWriteDisabled. Cleared by RefreshActuatorState (same entry point tests use to
' reset the module's state).
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

' Clears the write-disabled latch (§4.6) - the same "retry point" convention as
' PropertyRendering.RefreshRenderCaches, used by both the options panel (a user just toggled a switch) and
' the test harness. No rules to re-parse (pilot properties are fixed constants, revised 2026-08-20).
Public Sub RefreshActuatorState()
    mbWriteDisabled = False
End Sub

' Depth-0 hook, called from ElementChangeHandler.ProcessElement (after PropertyCalculation, before
' PropertyRendering) when IsEnabled. Resets one-shot flags, then reacts SELF: read own pilot property,
' compare to own attribute, write only if different, once per enabled attribute.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    mbColorInvalidShown = False
    mbLevelInvalidShown = False
    mbSelfRatchetShown = False

    ' A repeated COM write failure on this DGN (locked file, rights) must not spam the user nor retry
    ' indefinitely on the hot path (cahier §4.6) - same fail-closed session latch as
    ' PropertyRendering.mbWriteDisabled, cleared only by RefreshActuatorState.
    If mbWriteDisabled Then Exit Sub

    If oEl Is Nothing Then Exit Sub
    If Not oEl.IsGraphical Then Exit Sub

    ' Containment exclusion (cahier §4.6): never paint a locked or reference-owned element - same guard,
    ' same place in the sequence, as PropertyRendering.bas:534-536 (El.IsLocked).
    If oEl.IsLocked Then Exit Sub

    If Not IsColorEnabled() And Not IsLevelEnabled() Then Exit Sub

    ' Fail-closed exclusion of the trigger cell (cahier des charges §4.2/§6.2): a cell that matches
    ' CellColor[pattern]/CellLvl[pattern] must never be painted by the value it is itself the source of,
    ' or the non-looping invariant (source set / painted set disjoint by construction) no longer holds.
    If IsProtectedTriggerCell(oEl) Then Exit Sub

    ' Same exclusion, for a LEVEL trigger (epic 16 follow-up, LvlColor/LvlStyle/LvlWeight, 2026-08-24): an
    ' element whose OWN Level matches a pushable Lvl* pattern is the self-inclusion candidate
    ' FindFirstMatchingLevelInGroup can legitimately return for ITSELF. IsSelfSourceRatchet (below) does NOT
    ' catch this case - it compares the calc rule's SourceKind to csColor/csLvl, but a Lvl*-fed rule's
    ' SourceKind is csLvlColor/csLvlStyle/csLvlWeight, a DIFFERENT enum value, so the ratchet's equality
    ' test never fires. Without this guard, an element that is BOTH the Lvl*-trigger AND a carrier of
    ' ARES_Color would have its own ByLevel-symbolic Color silently frozen into the resolved literal index
    ' by this very pass (BearingPass computes ARES_Color from Level.ElementColor, then this Sub reads it
    ' back and overwrites the element's Color with that literal - not a loop, since the second pass is a
    ' no-op once stable, but a surprising, unwanted mutation of an element that should never be painted by
    ' its own authority, symmetric to the cell case above).
    If IsProtectedTriggerLevel(oEl) Then Exit Sub

    If IsColorEnabled() Then ActuateColor oEl, ARES_PROP_COLOR
    If IsLevelEnabled() Then ActuateLevel oEl, ARES_PROP_LVL
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ProcessElement"
End Sub

'######################################################################################################################
'                                          TRIGGER EXCLUSION (own fail-closed wrappers, Cell and Level)
'######################################################################################################################

' PropertyCalculation.IsTriggerCell is scoped GLOBALLY (any pushable Cell* rule in the WHOLE config, not
' just this actuator's own pilot rule) and its OWN ErrorHandler resolves to False (does NOT protect on a
' fault - it would paint the element by default). This wrapper is the actuator's OWN, protecting on fault
' like PropertyRendering.FeedsCellSource does - NEVER call IsTriggerCell bare from anywhere else in this
' module, and NEVER delegate to FeedsCellSource itself (that one is scoped to the renderer's own concern).
' Accepted compromise (cahier §4.2 point 1): a cell that matches an UNRELATED rule's pattern is also
' excluded here (false negative, no visible consequence - it was never in the painted set anyway).
Private Function IsProtectedTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsProtectedTriggerCell = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    IsProtectedTriggerCell = PropertyCalculation.IsTriggerCell(oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: if we cannot tell, assume the cell is a trigger and protect it from being painted.
    IsProtectedTriggerCell = True
End Function

' Mirrors IsProtectedTriggerCell exactly, for PropertyCalculation.IsTriggerLevel (Lvl* sources, epic 16
' follow-up) instead of IsTriggerCell - same fail-closed-on-fault wrapper rationale, own ErrorHandler
' protecting rather than PropertyCalculation.IsTriggerLevel's own (which resolves False on fault). NO
' element-type restriction here (unlike IsProtectedTriggerCell's IsCellElement gate), mirroring
' IsTriggerLevel itself - a Level trigger can be any graphical element, not just a cell.
Private Function IsProtectedTriggerLevel(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsProtectedTriggerLevel = False
    If oEl Is Nothing Then Exit Function
    IsProtectedTriggerLevel = PropertyCalculation.IsTriggerLevel(oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: if we cannot tell, assume the element is a trigger and protect it from being painted.
    IsProtectedTriggerLevel = True
End Function

'######################################################################################################################
'                                          ATTRIBUTE ACTUATION
'######################################################################################################################

' Reads P's current value on oEl, resolves it as a MicroStation color index, writes oEl.Color only when
' different - UNLESS oEl is a CellElement, whose own header Color is NEVER written (see below). On a cell,
' only repaints original sub-elements (cahier §4.4): only the sub-elements whose CURRENT color equals the
' header's OWN color BEFORE any change - a sub-element already holding a deliberately different color is left
' alone (re-derived from the retired StringsInEl.bas:266 rule, since AutoLengths left no surviving code to
' call). A FillMode=2 sub-element has its own fill color saved and restored around the write (retired
' ONLY_COLOR hook parity).
'
' WHY a CellElement's header Color is never written (2026-08-24, cahier-des-charges-groupcolor-fillmode.md
' investigation - do NOT "fix" this by reinstating the write, it is the root cause of a real bug, not an
' oversight): writing it triggers a MicroStation-NATIVE cascade onto any sub-element whose own Color/FillColor
' is set to the ByCellColor sentinel (mvba-docs Color_Property.md / FillColor_Property.md /
' ByCellColor_Method.md) - the header write silently resets such a sub-element's resolved FillMode/FillColor
' (2->1) as a side effect of MicroStation's own colour re-derivation, corrupting the very fill this Sub is
' trying to preserve, with no reliable way for MVBA code to intercept or undo it after the fact (several
' attempts tried and failed/regressed on a real repro: refreshing the element handle right after the write,
' snapshotting and restoring descendant Color/FillMode/FillColor state around it - a cell sub-element/
' ComplexShapeElement component has no independently retrievable model id either, so a restore keyed by id is
' a silent no-op, see StringsInEl.RefreshElementHandle's own comment on that). The retired ONLY_COLOR hook
' (formerly ElementChangeHandler.cls Branch 1, now removed) never had this problem for the simple reason it
' never wrote a CellElement's own header Color either - it went straight to a sub-element repaint loop,
' exactly like the one below.
Private Sub ActuateColor(ByVal oEl As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, P) Then Exit Sub

    ' SELF-ratchet refusal (cahier §6.1), checked against THIS element's actual governing calc rule -
    ' GetCalcRuleForProperty needs a real element to resolve Lvl/Cell/Type conditions, so this cannot be
    ' done once ahead of time. A pilot property whose calc source is the element's OWN Color feeding back
    ' into that same Color is a stable, meaningless fixed point.
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

    ' Non-cell elements have no sub-elements to repaint - writing their own Color IS the pilot property's
    ' whole effect, and carries none of the cascade risk described above. A CellElement's header is
    ' deliberately left untouched here; only the repaint loop below acts on it.
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
                        ' FillColor preservation (retired ONLY_COLOR hook parity, cahier-des-charges-groupcolor-
                        ' fillmode.md §4): a FillMode=2 sub-element's own fill color is saved and restored
                        ' around the write, since Color/FillColor are documented as independent properties.
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
    ' A fault here is most likely the attribute write/Rewrite itself (the only COM call in this sub that
    ' can fail on a locked/rights-restricted DGN despite the upstream IsLocked guard - e.g. a reference or
    ' server-side lock taken between the guard and the write). Latch the session-wide kill switch (§4.6)
    ' rather than retry this element again next pass.
    DisableAfterWriteFailure
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ActuateColor"
End Sub

' Reads P's current value on oEl, resolves it as a Level NAME (never auto-created - cahier §7.0), writes
' oEl.Level only when different. Same "original sub-elements only" repaint rule as ActuateColor, generalized
' to Level (compare on the sub-element's Level NAME, since Level objects are not directly comparable).
' An UNRESOLVABLE level NAME gets the generic ReportLevelInvalid status (cahier §7.0: "niveau verrouille/
' reference externe = meme chemin d'echec generique") - a name GetElements.GetLevel simply cannot find.
' A level that resolves but then FAILS TO WRITE (locked/reference-owned target) is a DIFFERENT path in
' practice: it faults inside the .Level = / .Rewrite call below, caught by this sub's own ErrorHandler,
' which logs once (English) and latches the session-wide kill switch (§4.6) rather than showing
' ReportLevelInvalid repeatedly - reviewer-4 flagged this wording mismatch (2026-08-19): both cases are
' "no dedicated handling" in spirit (no bespoke per-cause status), but they are NOT literally the same
' status/code path. Left as-is (not worth a bespoke unified status for a rare case); noted here so a future
' reader does not go looking for a single shared branch that does not exist.
Private Sub ActuateLevel(ByVal oEl As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, P) Then Exit Sub

    ' SELF-ratchet refusal (cahier §6.1) - see ActuateColor for the full rationale.
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

    ' Unlike ActuateColor, a CellElement's header Level IS written here: Level has no ByCellColor-style
    ' cascade sentinel (mvba-docs confirms ByLevelColor/ByLevelLineStyle/ByLevelLineWeight exist for
    ' Color/LineStyle/LineWeight, none for Level itself), so this write carries none of the corruption risk
    ' that ActuateColor's header write does on a cell - see that Sub's header comment for the full rationale.
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
    ' See ActuateColor's ErrorHandler for the rationale.
    DisableAfterWriteFailure
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyActuator.ActuateLevel"
End Sub

' True when P's calc source ON THIS ELEMENT (PropertyCalculation.ARES_Calc_Rules, conditions resolved
' against oEl) is the SAME SELF source as TargetKind (csColor for a Color rule, csLvl for a Level rule) -
' the ratchet cahier §6.1 forbids: the pilot property would only ever reflect what THIS actuator itself
' last wrote onto the very attribute it drives. Checked per element, per attribute, right before the write -
' PropertyCalculation.GetCalcRuleForProperty needs a REAL element to resolve a rule's Lvl/Cell/Type
' conditions (it returns "no rule" outright on Nothing), and a calc rule's SourceKind for P can legitimately
' vary by condition from one element to the next, so this cannot be checked once ahead of time.
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
    ' Fail-closed (reviewer-4 finding: an internal fault here used to fail OPEN - i.e. allow the write -
    ' which was inconsistent with IsProtectedTriggerCell's assumed fail-closed doctrine right next to it.
    ' When we cannot tell whether P is SELF-sourced from this very attribute, refuse the write: the cost of
    ' a wrongly-skipped legitimate write is a status the user can investigate; the cost of a wrongly-allowed
    ' one is the silent, symptomless ratchet §6.1 exists to prevent).
    IsSelfSourceRatchet = True
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

' Session-wide fail-closed latch on a repeated attribute-write failure (cahier §4.6), mirroring
' PropertyRendering.DisableAfterWriteFailure (PropertyRendering.bas:2941-2946): idempotent, logs ONE
' English line (never translated/status-bar - this is a genuine fault, not an expected refusal), then every
' subsequent ProcessElement short-circuits at the top until RefreshActuatorState clears it.
Private Sub DisableAfterWriteFailure()
    On Error Resume Next
    If mbWriteDisabled Then Exit Sub
    mbWriteDisabled = True
    ErrorHandler.HandleError "Property actuator: attribute write failed, actuation disabled for this session", 0, "", "PropertyActuator.ActuateAttribute"
End Sub
