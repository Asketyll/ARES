' Module: PropertyRendering_Types
' Description: Shared kernel for the PropertyRendering split - the RenderEntry record and the pure
'              array helpers over it. Public so RenderEntry can be passed ByRef across module
'              boundaries; no module-scope state, no dependency on any other module.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: none

Option Explicit

' One parsed ARES_Render entry: which sub-text it drives, the Template, and the value last rendered for
' each of the Template's tokens. ValNames/ValValues are parallel, bounded by nVals. Dropped marks an
' entry the state machine released; serialisation skips it, which avoids ever copying a UDT holding
' dynamic arrays just to compact the list.
Public Type RenderEntry
    SubId As Long
    Template As String
    ValNames() As String
    ValValues() As String
    nVals As Long
    Dropped As Boolean
End Type

' Outcomes of the per-entry state machine.
Public Const ENTRY_UNCHANGED As Long = 0
Public Const ENTRY_UPDATED As Long = 1
Public Const ENTRY_DROP As Long = 2

' Look up a token's value in a name/value pair set (case-insensitive). Absent or empty both yield "",
' which Expand renders as the literal token - the two cases are deliberately indistinguishable there.
Public Function LookupValue(ByVal sName As String, ByRef names() As String, ByRef values() As String, ByVal n As Long) As String
    Dim i As Long
    LookupValue = ""
    If n <= 0 Then Exit Function
    For i = 0 To n - 1
        If StrComp(names(i), sName, vbTextCompare) = 0 Then
            LookupValue = values(i)
            Exit Function
        End If
    Next i
End Function

' Does an ENTRY exist for this token (regardless of whether its value is empty)? The unset case turns on
' this distinction: an existing-but-empty entry means "rendered as the literal token last time".
Public Function HasValueEntry(ByVal sName As String, ByRef names() As String, ByVal n As Long) As Boolean
    Dim i As Long
    HasValueEntry = False
    If n <= 0 Then Exit Function
    For i = 0 To n - 1
        If StrComp(names(i), sName, vbTextCompare) = 0 Then
            HasValueEntry = True
            Exit Function
        End If
    Next i
End Function

' Append one name/value pair, growing both parallel arrays together.
Public Sub AppendValue(ByVal sName As String, ByVal sValue As String, ByRef names() As String, ByRef values() As String, ByRef n As Long)
    ReDim Preserve names(0 To n)
    ReDim Preserve values(0 To n)
    names(n) = sName
    values(n) = sValue
    n = n + 1
End Sub

' Replace an entry's stored LastValues wholesale.
Public Sub SetEntryValues(ByRef ents() As RenderEntry, ByVal idx As Long, ByRef names() As String, ByRef values() As String, ByVal n As Long)
    Dim i As Long
    ents(idx).nVals = n
    If n <= 0 Then Exit Sub
    ReDim ents(idx).ValNames(0 To n - 1)
    ReDim ents(idx).ValValues(0 To n - 1)
    For i = 0 To n - 1
        ents(idx).ValNames(i) = names(i)
        ents(idx).ValValues(i) = values(i)
    Next i
End Sub

' How many entries would actually be written back (a released entry no longer counts). Relocated here
' (not listed under any module in the split plan) because it has two callers in two different new
' modules - RenderBoundElement (StateMachine) and TryFirstAuthor (Authoring) - so it cannot travel with
' "its one consumer" the way EntryIsConsistent/VariantToPlainString did; it is a pure array helper over
' RenderEntry with no module state, exactly Module Types' stated purpose.
Public Function CountLiveEntries(ByRef ents() As RenderEntry, ByVal nEnts As Long) As Long
    Dim i As Long
    CountLiveEntries = 0
    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then CountLiveEntries = CountLiveEntries + 1
    Next i
End Function
