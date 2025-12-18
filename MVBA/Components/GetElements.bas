Attribute VB_Name = "GetElements"
' Module: GetElements
' Description: This module provides functions to get ElementEnumerator
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass, ARESConstants, MicroStationDefinition
Option Explicit

Public Function ByEE(Optional Levels As Variant, Optional Range As Variant, Optional CellName As String = Empty, Optional GraphicGroup As Long = -1, Optional AllowNoGraphicGroup As Boolean = False, Optional ElTypes As Variant, Optional Colors As Variant, Optional LineStyles As Variant, Optional LineWeights As Variant) As ElementEnumerator
    On Error GoTo ErrorHandler

    Dim ByLevel As Boolean
    Dim ByRange As Boolean
    Dim ByCellName As Boolean
    Dim ByGG As Boolean
    Dim ByType As Boolean
    Dim ByColor As Boolean
    Dim ByLineStyle As Boolean
    Dim ByLineWeight As Boolean
    Dim oLevel() As Level
    Dim oLineStyle() As LineStyle
    Dim oElType() As MsdElementType
    Dim oColor() As Long
    Dim oLineWeight() As Long
    ReDim oLevel(0)
    ReDim oLineStyle(0)
    ReDim oElType(0)
    ReDim oColor(0)
    ReDim oLineWeight(0)
    Dim oRange As Range3d
    Dim i As Integer
    Dim esc As New ElementScanCriteria

    If Not IsMissing(Levels) Then
        ByLevel = True
    End If
    If Not IsMissing(Range) Then
        ByRange = True
    End If
    If CellName <> "" Then
        ByCellName = True
    End If
    If GraphicGroup <> -1 Then
        ByGG = True
    End If
    If Not IsMissing(ElTypes) Then
        ByType = True
    End If
    If Not IsMissing(Colors) Then
        ByColor = True
    End If
    If Not IsMissing(LineStyles) Then
        ByLineStyle = True
    End If
    If Not IsMissing(LineWeights) Then
        ByLineWeight = True
    End If
    If Not ByLevel And Not ByRange And Not ByCellName And Not ByGG And Not ByType And Not ByColor And Not ByLineStyle And Not ByLineWeight Then
        Set ByEE = ActiveModelReference.Scan(esc)
        Exit Function
    End If

    ' Process Levels parameter
    If ByLevel Then
        esc.ExcludeAllLevels
        If IsArray(Levels) Then
            For i = LBound(Levels) To UBound(Levels)
                If IsValidLevelName(Levels(i)) Then
                    If oLevel(UBound(oLevel)) Is Nothing Then
                        Set oLevel(UBound(oLevel)) = ActiveDesignFile.Levels(Levels(i))
                    Else
                        ReDim Preserve oLevel(UBound(oLevel) + 1)
                        Set oLevel(UBound(oLevel)) = ActiveDesignFile.Levels(Levels(i))
                    End If
                End If
            Next i
        Else
            If IsValidLevelName(Levels) Then
                Set oLevel(0) = ActiveDesignFile.Levels(Levels)
            End If
        End If
        If Not oLevel(UBound(oLevel)) Is Nothing Then
            For i = LBound(oLevel) To UBound(oLevel)
                esc.IncludeLevel oLevel(i)
            Next i
        End If
    End If

    If ByRange Then
        oRange = Range
        esc.IncludeOnlyWithinRange oRange
    End If

    If ByCellName Then
        esc.IncludeOnlyCell CellName
        esc.ExcludeAllTypes
        esc.IncludeType msdElementTypeCellHeader
    End If

    If ByGG Then
        ' Check if GraphicGroup is 0 when AllowNoGraphicGroup is False
        If GraphicGroup = ARESConstants.ARES_DEFAULT_GRAPHIC_GROUP_ID And Not AllowNoGraphicGroup Then
            ' Skip graphic group filter
        Else
            esc.IncludeOnlyGraphicGroup GraphicGroup
        End If
    End If

    ' Process ElTypes parameter
    If ByType Then
        If Not ByCellName Then
            esc.ExcludeAllTypes
        End If
        If IsArray(ElTypes) Then
            For i = LBound(ElTypes) To UBound(ElTypes)
                If MicroStationDefinition.IsValidElementType(ElTypes(i)) Then
                    If oElType(UBound(oElType)) = 0 Then
                        oElType(UBound(oElType)) = ElTypes(i)
                    Else
                        ReDim Preserve oElType(UBound(oElType) + 1)
                        oElType(UBound(oElType)) = ElTypes(i)
                    End If
                End If
            Next i
        Else
            If MicroStationDefinition.IsValidElementType(ElTypes) Then
                oElType(0) = ElTypes
            End If
        End If
        If oElType(UBound(oElType)) <> 0 Then
            For i = LBound(oElType) To UBound(oElType)
                esc.IncludeType oElType(i)
            Next i
        End If
    End If

    ' Process Colors parameter
    If ByColor Then
        esc.ExcludeAllColors
        If IsArray(Colors) Then
            For i = LBound(Colors) To UBound(Colors)
                esc.IncludeColor Colors(i)
            Next i
        Else
            esc.IncludeColor Colors
        End If
    End If

    ' Process LineStyles parameter
    If ByLineStyle Then
        esc.ExcludeAllLineStyles
        If IsArray(LineStyles) Then
            For i = LBound(LineStyles) To UBound(LineStyles)
                If IsValidLineStyleName(LineStyles(i)) Then
                    If oLineStyle(UBound(oLineStyle)) Is Nothing Then
                        Set oLineStyle(UBound(oLineStyle)) = ActiveDesignFile.LineStyles(LineStyles(i))
                    Else
                        ReDim Preserve oLineStyle(UBound(oLineStyle) + 1)
                        Set oLineStyle(UBound(oLineStyle)) = ActiveDesignFile.LineStyles(LineStyles(i))
                    End If
                End If
            Next i
        Else
            If IsValidLineStyleName(LineStyles) Then
                Set oLineStyle(0) = ActiveDesignFile.LineStyles(LineStyles)
            End If
        End If
        If Not oLineStyle(UBound(oLineStyle)) Is Nothing Then
            For i = LBound(oLineStyle) To UBound(oLineStyle)
                esc.IncludeLineStyle oLineStyle(i)
            Next i
        End If
    End If

    ' Process LineWeights parameter
    If ByLineWeight Then
        esc.ExcludeAllLineWeights
        If IsArray(LineWeights) Then
            For i = LBound(LineWeights) To UBound(LineWeights)
                esc.IncludeLineWeight LineWeights(i)
            Next i
        Else
            esc.IncludeLineWeight LineWeights
        End If
    End If

    Set ByEE = ActiveModelReference.Scan(esc)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.ByEE"
    Dim esc2 As New ElementScanCriteria
    esc2.ExcludeAllTypes
    esc2.ExcludeAllLevels
    Set ByEE = ActiveModelReference.Scan(esc2)
End Function

Public Function IsValidLevelName(ByVal levelName As String) As Boolean
    IsValidLevelName = False
    On Error GoTo ErrorHandler

    Dim oLevel As Level
    Set oLevel = ActiveDesignFile.Levels(levelName)

    If oLevel Is Nothing Then
        IsValidLevelName = False
    Else
        IsValidLevelName = True
    End If

    Set oLevel = Nothing
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 5:     '   Level not found
        Resume Next
        Case -2147024809:
        Resume Next
    Case Else
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.IsValidLevelName"
    End Select
End Function

Public Function IsValidLineStyleName(ByVal lineStyleName As String) As Boolean
    IsValidLineStyleName = False
    On Error GoTo ErrorHandler

    Dim oLineStyle As LineStyle
    Set oLineStyle = ActiveDesignFile.LineStyles(lineStyleName)

    If oLineStyle Is Nothing Then
        IsValidLineStyleName = False
    Else
        IsValidLineStyleName = True
    End If

    Set oLineStyle = Nothing
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 5:     '   LineStyle not found
        Resume Next
    Case Else
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.IsValidLineStyleName"
    End Select
End Function
