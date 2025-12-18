Attribute VB_Name = "GetElements"
' Module: GetElements
' Description: This module provides functions to validate levels and retrieve elements from MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass, MicroStationDefinition, ARESConstants
Option Explicit

' === PUBLIC FUNCTIONS ===

' Function to validate if a level name exists in the active design file
Public Function IsValidLevelName(ByVal LevelName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim Level As Level
    Dim LevelTable As Levels

    IsValidLevelName = False

    ' Check if there is an active model reference
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "GetElements.IsValidLevelName", "WARNING"
        Exit Function
    End If

    ' Get the level table from the active design file
    Set LevelTable = ActiveDesignFile.Levels

    ' Try to find the level by name
    On Error Resume Next
    Set Level = LevelTable.Find(LevelName)

    ' If no error occurred and level is not Nothing, the level exists
    If Err.Number = 0 And Not Level Is Nothing Then
        IsValidLevelName = True
    End If

    On Error GoTo ErrorHandler

    Exit Function

ErrorHandler:
    IsValidLevelName = False
    ErrorHandler.HandleError Err.Description, Err.Number, "GetElements.IsValidLevelName", "ERROR"
End Function

' Function to get all graphical elements from specified levels (excluding rasters)
Public Function GetElementsByLevels(Lvls() As String, _
                                   Optional FilterByTypes As Variant, _
                                   Optional IncludeRasters As Boolean = False) As Variant
    On Error GoTo ErrorHandler

    Dim Elements() As Element
    Dim Esc As ElementScanCriteria
    Dim ee As ElementEnumerator
    Dim CurrentElement As Element
    Dim i As Long
    Dim Count As Long
    Dim ValidLevels() As String
    Dim ValidLevelCount As Long
    Dim MSDEType() As MsdElementType

    ' Check if there is an active model reference
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "GetElements.GetElementsByLevels", "WARNING"
        ReDim Elements(0)
        GetElementsByLevels = Elements
        Exit Function
    End If

    ' Validate and collect valid level names
    ValidLevelCount = 0
    ReDim ValidLevels(LBound(Lvls) To UBound(Lvls))

    For i = LBound(Lvls) To UBound(Lvls)
        If IsValidLevelName(Lvls(i)) Then
            ValidLevels(ValidLevelCount) = Lvls(i)
            ValidLevelCount = ValidLevelCount + 1
        Else
            ErrorHandler.HandleError "Invalid level name: " & Lvls(i), 0, "GetElements.GetElementsByLevels", "WARNING"
        End If
    Next i

    ' If no valid levels, return empty array
    If ValidLevelCount = 0 Then
        ErrorHandler.HandleError "No valid levels found", 0, "GetElements.GetElementsByLevels", "WARNING"
        ReDim Elements(0)
        GetElementsByLevels = Elements
        Exit Function
    End If

    ' Resize array to actual valid level count
    ReDim Preserve ValidLevels(0 To ValidLevelCount - 1)

    ' Initialize element scan criteria
    Set Esc = New ElementScanCriteria

    ' Include only graphical elements
    Esc.ExcludeNonGraphical

    ' Exclude raster types unless explicitly included
    If Not IncludeRasters Then
        Esc.ExcludeType msdElementTypeRasterHeader
        Esc.ExcludeType msdElementTypeRasterComponent
        Esc.ExcludeType msdElementTypeRasterReference
        Esc.ExcludeType msdElementTypeRasterReferenceComponent
        Esc.ExcludeType msdElementTypeRasterFrame
    End If

    ' Apply type filter if provided
    If Not IsMissing(FilterByTypes) Then
        MSDEType = EnsureElementTypeArray(FilterByTypes)
        Esc.ExcludeAllTypes
        For i = LBound(MSDEType) To UBound(MSDEType)
            If IsValidElementType(MSDEType(i)) And MSDEType(i) <> ARES_MSDETYPE_ERROR Then
                Esc.IncludeType MSDEType(i)
            End If
        Next i
    End If

    ' Scan and collect elements from all valid levels
    ReDim Elements(0)
    Count = 0

    For i = LBound(ValidLevels) To UBound(ValidLevels)
        ' Set level filter
        Esc.IncludeOnlyLevel ActiveDesignFile.Levels.Find(ValidLevels(i))

        ' Scan for elements
        Set ee = ActiveModelReference.Scan(Esc)

        ' Count and collect elements
        Do While ee.MoveNext
            Set CurrentElement = ee.Current
            If CurrentElement.IsGraphical Then
                Count = Count + 1
                ReDim Preserve Elements(1 To Count)
                Set Elements(Count) = CurrentElement
            End If
        Loop
    Next i

    GetElementsByLevels = Elements
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "GetElements.GetElementsByLevels", "ERROR"
    ReDim Elements(0)
    GetElementsByLevels = Elements
End Function

' === PRIVATE HELPER FUNCTIONS ===

' Private function to ensure a variant is an array of MsdElementType
Private Function EnsureElementTypeArray(ByVal Value As Variant) As Variant
    On Error GoTo ErrorHandler
    Dim tempArray() As MsdElementType
    Dim i As Long

    If IsArray(Value) Then
        ' Check each element in the array and convert if necessary
        ReDim tempArray(LBound(Value) To UBound(Value))
        For i = LBound(Value) To UBound(Value)
            Select Case VarType(Value(i))
                Case vbString
                    tempArray(i) = StringToMsdElementType(Value(i))
                Case vbLong
                    tempArray(i) = Value(i)
                Case vbInteger
                    tempArray(i) = CLng(Value(i))
                Case Else
                    ReDim tempArray(0)
                    tempArray(0) = ARES_MSDETYPE_ERROR
                    EnsureElementTypeArray = tempArray
                    Exit Function
            End Select
        Next i
        EnsureElementTypeArray = tempArray
    ElseIf Not IsMissing(Value) And Not IsEmpty(Value) Then
        ' Create a single-element array
        ReDim tempArray(0)
        Select Case VarType(Value)
            Case vbString
                tempArray(0) = StringToMsdElementType(Value)
            Case vbLong
                tempArray(0) = Value
            Case vbInteger
                tempArray(0) = CLng(Value)
            Case Else
                ReDim tempArray(0)
                tempArray(0) = ARES_MSDETYPE_ERROR
                EnsureElementTypeArray = tempArray
                Exit Function
        End Select
        EnsureElementTypeArray = tempArray
    Else
        EnsureElementTypeArray = Array(ARES_MSDETYPE_ERROR)
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "GetElements.EnsureElementTypeArray", "ERROR"
    ReDim tempArray(0)
    tempArray(0) = ARES_MSDETYPE_ERROR
    EnsureElementTypeArray = tempArray
End Function
