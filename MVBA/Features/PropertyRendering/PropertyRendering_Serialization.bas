' Module: PropertyRendering_Serialization
' Description: ARES_Render metadata read/write and the Entries blob (de)serialisation. Full mechanism:
'              see _bmad/docs/property-rendering-mechanics.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, CustomPropertyHandler, ErrorHandlerClass (global ErrorHandler),
'               PropertyRendering_Types, PropertyRendering_Reporting

Option Explicit

' The two String properties of the ARES_Render ItemType. Every access names BOTH of them and the library
' explicitly - see ReadRenderMetadata for why an omitted ItemName silently breaks every read and write.
Private Const PROP_SCHEMA As String = "SchemaVersion"
Private Const PROP_ENTRIES As String = "Entries"

' Read and parse the ARES_Render metadata of an element. False = attached but unusable - the caller must
' refuse, never guess. True with nEnts = 0 is the legal "freshly attached, nothing stored yet" state. EVERY
' access names ItemName AND LibraryName explicitly (two independent CustomPropertyHandler traps otherwise).
' Full rationale: see "ReadRenderMetadata" in property-rendering-mechanics.md.
Public Function ReadRenderMetadata(ByVal El As element, ByRef ents() As RenderEntry, ByRef nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim vSchema As Variant
    Dim vEntries As Variant
    Dim sSchema As String
    Dim sEntries As String

    ReadRenderMetadata = False
    nEnts = 0

    vSchema = CustomPropertyHandler.GetPropertyValueFromElement(El, PROP_SCHEMA, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True)
    vEntries = CustomPropertyHandler.GetPropertyValueFromElement(El, PROP_ENTRIES, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True)

    sSchema = VariantToPlainString(vSchema)
    sEntries = VariantToPlainString(vEntries)

    If Len(sSchema) = 0 Then
        ' Freshly attached: both fields empty is legal and means "nothing stored yet". A version-less
        ' item that nevertheless carries entries has been tampered with.
        If Len(sEntries) > 0 Then
            ReportMetadataUnreadable
            Exit Function
        End If
        ReadRenderMetadata = True
        Exit Function
    End If

    If sSchema <> ARES_RENDER_SCHEMA Then
        ' A newer ARES wrote a shape this build cannot read. Refuse to interpret it and NEVER rewrite it.
        ReportSchemaUnsupported
        Exit Function
    End If

    If Len(sEntries) = 0 Then
        ReadRenderMetadata = True
        Exit Function
    End If

    ReadRenderMetadata = DeserializeEntries(sEntries, ents, nEnts)
    If Not ReadRenderMetadata Then ReportMetadataUnreadable
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Serialization.ReadRenderMetadata"
    ReportMetadataUnreadable
    ReadRenderMetadata = False
    nEnts = 0
End Function

' Write SchemaVersion + Entries back. Both writes are strict (bNoFallback) and both name their item and
' library. Called only AFTER the corresponding text writes succeeded.
Public Function WriteRenderMetadata(ByVal El As element, ByRef ents() As RenderEntry, ByVal nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim sEntries As String

    WriteRenderMetadata = False
    sEntries = SerializeEntries(ents, nEnts)

    If Not CustomPropertyHandler.SetPropertyValueToElement(El, PROP_SCHEMA, ARES_RENDER_SCHEMA, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True) Then Exit Function
    If Not CustomPropertyHandler.SetPropertyValueToElement(El, PROP_ENTRIES, sEntries, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True) Then Exit Function

    ' No Rewrite on the bearer: an item write goes straight to the file (mvba-docs/03-methods/
    ' SetPropertyValue_Method.md - "always writes a change back to the file immediately"). The targeted
    ' text write does its own Rewrite on the sub-element it touched.
    WriteRenderMetadata = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Serialization.WriteRenderMetadata"
    WriteRenderMetadata = False
End Function

' Records separated by Chr(1), fields by Chr(2), LastValues by Chr(3) (a flat name,value,name,value
' sequence), line breaks inside a Template by Chr(4). These four cannot be typed into a MicroStation text
' nor produced by a normal property value, which buys one fail-closed rejection rule instead of a whole
' escaping machinery: a VALUE carrying any of them (or CR/LF) is refused at the render choke point.
Private Function SerializeEntries(ByRef ents() As RenderEntry, ByVal nEnts As Long) As String
    On Error GoTo ErrorHandler

    Dim sOut As String
    Dim sRec As String
    Dim i As Long, j As Long

    SerializeEntries = ""
    If nEnts = 0 Then Exit Function

    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then
            sRec = CStr(ents(i).SubId) & SepField() & EncodeTemplate(ents(i).Template) & SepField()
            For j = 0 To ents(i).nVals - 1
                If j > 0 Then sRec = sRec & SepPair()
                sRec = sRec & ents(i).ValNames(j) & SepPair() & ents(i).ValValues(j)
            Next j
            If Len(sOut) > 0 Then sOut = sOut & SepRecord()
            sOut = sOut & sRec
        End If
    Next i

    SerializeEntries = sOut
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_Serialization.SerializeEntries"
    SerializeEntries = ""
End Function

' Parse the Entries blob. False on ANY malformation - hostile input is never partially trusted.
Private Function DeserializeEntries(ByVal sBlob As String, ByRef ents() As RenderEntry, ByRef nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim recs() As String
    Dim fields() As String
    Dim pairs() As String
    Dim names() As String
    Dim values() As String
    Dim n As Long
    Dim i As Long, j As Long

    DeserializeEntries = False
    nEnts = 0

    recs = Split(sBlob, SepRecord())
    ' A zero-length blob yields UBound = -1; iterating LBound..UBound is the only safe form (never
    ' index parts(0) directly).
    For i = LBound(recs) To UBound(recs)
        If Len(recs(i)) > 0 Then
            fields = Split(recs(i), SepField())
            If UBound(fields) - LBound(fields) <> 2 Then Exit Function

            If Not IsNumeric(fields(LBound(fields))) Then Exit Function

            n = 0
            If Len(fields(LBound(fields) + 2)) > 0 Then
                pairs = Split(fields(LBound(fields) + 2), SepPair())
                If ((UBound(pairs) - LBound(pairs) + 1) Mod 2) <> 0 Then Exit Function
                For j = LBound(pairs) To UBound(pairs) Step 2
                    AppendValue pairs(j), pairs(j + 1), names, values, n
                Next j
            End If

            ReDim Preserve ents(0 To nEnts)
            ents(nEnts).SubId = CLng(fields(LBound(fields)))
            ents(nEnts).Template = DecodeTemplate(fields(LBound(fields) + 1))
            ents(nEnts).Dropped = False
            SetEntryValues ents, nEnts, names, values, n
            nEnts = nEnts + 1
        End If
    Next i

    DeserializeEntries = True
    Exit Function

ErrorHandler:
    ' Malformed metadata is expected hostile input, not a fault worth logging: the caller surfaces a
    ' status and refuses to render.
    DeserializeEntries = False
    nEnts = 0
End Function

' A Template legitimately carries line breaks (a multi-line TextNode's Template is exactly what
' GetTextAtSubId joins with vbLf), so they are serialised rather than refused - the refusal rule applies
' to VALUES only.
Private Function EncodeTemplate(ByVal s As String) As String
    EncodeTemplate = Replace(s, vbLf, SepLine())
End Function

Private Function DecodeTemplate(ByVal s As String) As String
    DecodeTemplate = Replace(s, SepLine(), vbLf)
End Function

' VBA forbids a function call in a Const initialiser, so the four delimiters are trivial functions.
Private Function SepRecord() As String
    SepRecord = Chr(1)
End Function

Private Function SepField() As String
    SepField = Chr(2)
End Function

Private Function SepPair() As String
    SepPair = Chr(3)
End Function

Private Function SepLine() As String
    SepLine = Chr(4)
End Function

' True when a string carries one of the four serialisation delimiters. None of them can be typed into a
' MicroStation text nor produced by a normal property value, which is exactly why they were chosen.
Public Function ContainsSerialisationDelimiter(ByVal s As String) As Boolean
    ContainsSerialisationDelimiter = True
    If InStr(1, s, SepRecord()) > 0 Then Exit Function
    If InStr(1, s, SepField()) > 0 Then Exit Function
    If InStr(1, s, SepPair()) > 0 Then Exit Function
    If InStr(1, s, SepLine()) > 0 Then Exit Function
    ContainsSerialisationDelimiter = False
End Function

' The strict rule for VALUES: no delimiter and no line break of any kind. A Template is deliberately NOT
' held to it - a multi-line TextNode's Template legitimately contains vbLf, which is what SepLine
' serialises.
Public Function ValueHasIllegalChar(ByVal s As String) As Boolean
    ValueHasIllegalChar = True
    If ContainsSerialisationDelimiter(s) Then Exit Function
    If InStr(1, s, vbCr) > 0 Then Exit Function
    If InStr(1, s, vbLf) > 0 Then Exit Function
    ValueHasIllegalChar = False
End Function

' Convert an item read to a plain String WITHOUT any locale-dependent transform: Null and any non-string
' type yield "". Used for the metadata's own String fields. Relocated here with its one caller
' (ReadRenderMetadata) - same "travels with its one consumer" principle as EntryIsConsistent in
' Module D (StateMachine).
Private Function VariantToPlainString(ByVal v As Variant) As String
    VariantToPlainString = ""
    If IsNull(v) Then Exit Function
    If IsArray(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    VariantToPlainString = v
End Function
