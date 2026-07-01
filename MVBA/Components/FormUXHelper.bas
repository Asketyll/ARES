' Module: FormUXHelper
' Description: Shared UX plumbing for ARES UserForms - control lock, true inline-edit cancel,
'              non-blocking feedback, localized tooltips, restore-to-default persistence.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARES_MS_VAR_Class
Option Explicit

' Named result of an inline-edit key press - replaces the literal 13/27 key codes.
Public Enum FormUXInlineKey
    FormUXKeyNone = 0
    FormUXKeyCommit = 1
    FormUXKeyCancel = 2
End Enum

' Enable/disable every actionable control on a form, recursing into containers.
' Explicit state - replaces each form's toggle Locked()/CheckControlForLock pair.
Public Sub SetControlsLocked(ByVal oForm As Object, ByVal bLocked As Boolean)
    On Error GoTo ErrorHandler
    Dim oCtrl As Control
    For Each oCtrl In oForm.Controls
        LockControl oCtrl, bLocked
    Next oCtrl
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.SetControlsLocked"
End Sub

' Recursive worker for SetControlsLocked.
Private Sub LockControl(ByVal oCtrl As Control, ByVal bLocked As Boolean)
    On Error GoTo ErrorHandler
    Select Case TypeName(oCtrl)
        Case "CommandButton", "CheckBox", "SpinButton", "ComboBox"
            oCtrl.Enabled = Not bLocked
        Case "Frame", "MultiPage", "Page"
            Dim oSub As Control
            For Each oSub In oCtrl.Controls
                LockControl oSub, bLocked
            Next oSub
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.LockControl"
End Sub

' Classify an inline-edit key press. Enter (no Shift) commits, Esc cancels, anything else is none.
Public Function InlineEditKey(ByVal KeyCode As Integer, ByVal Shift As Integer) As FormUXInlineKey
    On Error GoTo ErrorHandler
    InlineEditKey = FormUXKeyNone
    If Shift <> 0 Then Exit Function
    Select Case KeyCode
        Case vbKeyReturn
            InlineEditKey = FormUXKeyCommit
        Case vbKeyEscape
            InlineEditKey = FormUXKeyCancel
    End Select
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.InlineEditKey"
    InlineEditKey = FormUXKeyNone
End Function

' Commit a hidden-textbox / reveal-button inline editor: write through only when the value
' actually changed, then hide the box and show the button. Returns True iff it wrote (so the
' caller can run a side-effect, e.g. PropertyTagging.RefreshRules).
Public Function CommitInlineEdit(ByVal oBox As MSForms.TextBox, ByVal oBtn As MSForms.CommandButton, _
                                 ByVal oVar As ARES_MS_VAR_Class) As Boolean
    On Error GoTo ErrorHandler
    CommitInlineEdit = False
    If oBox.Value <> oVar.Value Then
        oVar.Value = oBox.Value
        CommitInlineEdit = True
    End If
    oBox.Visible = False
    oBtn.Visible = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.CommitInlineEdit"
    CommitInlineEdit = False
End Function

' The Esc true-cancel: reset the box to the stored value so the ensuing commit/teardown
' sees no change and writes nothing.
Public Sub RevertInlineEdit(ByVal oBox As MSForms.TextBox, ByVal oVar As ARES_MS_VAR_Class)
    On Error GoTo ErrorHandler
    oBox.Value = oVar.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.RevertInlineEdit"
End Sub

' Non-blocking "finish or Esc" cue - a translated status line plus refocus. No Sleep, no loop.
Public Sub NudgeActiveEdit(ByVal oBox As MSForms.TextBox)
    On Error GoTo ErrorHandler
    LangManager.ShowStatusT "FormFinishEditFirst"
    oBox.SetFocus
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.NudgeActiveEdit"
End Sub

' Set a localized tooltip on any control.
Public Sub SetTip(ByVal oCtrl As Object, ByVal sKey As String)
    On Error GoTo ErrorHandler
    oCtrl.ControlTipText = GetTranslation(sKey)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.SetTip"
End Sub

' Persist a config var back to its default (Property Let path -> Config.SetVar). NOT ResetToDefault,
' which only sets the in-memory value and does not persist. Public surface for the AC-12 restore
' handler wired once the per-form Reset button exists (Track B).
Public Sub PersistDefault(ByVal oVar As ARES_MS_VAR_Class)
    On Error GoTo ErrorHandler
    oVar.Value = oVar.DefaultValue
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FormUXHelper.PersistDefault"
End Sub
