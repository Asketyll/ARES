' Module: GetElements
' Description: This module provides functions to get ElementEnumerator
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass, ARESConstants, MicroStationDefinition, CustomPropertyHandler
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
    Dim Esc As New ElementScanCriteria
    
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
        Set ByEE = ActiveModelReference.Scan(Esc)
        Exit Function
    End If
    
    ' Process Levels parameter
    If ByLevel Then
        Esc.ExcludeAllLevels
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
                Esc.IncludeLevel oLevel(i)
            Next i
        End If
    End If
    
    If ByRange Then
        oRange = Range
        Esc.IncludeOnlyWithinRange oRange
    End If
    
    If ByCellName Then
        Esc.IncludeOnlyCell CellName
        Esc.ExcludeAllTypes
        Esc.IncludeType msdElementTypeCellHeader
    End If
    
    If ByGG Then
        ' Check if GraphicGroup is 0 when AllowNoGraphicGroup is False
        If GraphicGroup = ARESConstants.ARES_DEFAULT_GRAPHIC_GROUP_ID And Not AllowNoGraphicGroup Then
            ' Skip graphic group filter
        Else
            Esc.IncludeOnlyGraphicGroup GraphicGroup
        End If
    End If
    
    ' Process ElTypes parameter
    If ByType Then
        If Not ByCellName Then
            Esc.ExcludeAllTypes
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
                Esc.IncludeType oElType(i)
            Next i
        End If
    End If
    
    ' Process Colors parameter
    If ByColor Then
        Esc.ExcludeAllColors
        If IsArray(Colors) Then
            For i = LBound(Colors) To UBound(Colors)
                Esc.IncludeColor Colors(i)
            Next i
        Else
            Esc.IncludeColor Colors
        End If
    End If
    
    ' Process LineStyles parameter
    If ByLineStyle Then
        Esc.ExcludeAllLineStyles
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
                Esc.IncludeLineStyle oLineStyle(i)
            Next i
        End If
    End If
    
    ' Process LineWeights parameter
    If ByLineWeight Then
        Esc.ExcludeAllLineWeights
        If IsArray(LineWeights) Then
            For i = LBound(LineWeights) To UBound(LineWeights)
                Esc.IncludeLineWeight LineWeights(i)
            Next i
        Else
            Esc.IncludeLineWeight LineWeights
        End If
    End If
    
    Set ByEE = ActiveModelReference.Scan(Esc)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.ByEE"
    Dim esc2 As New ElementScanCriteria
    esc2.ExcludeAllTypes
    esc2.ExcludeAllLevels
    Set ByEE = ActiveModelReference.Scan(esc2)
End Function

Public Function IsValidLevelName(ByVal LevelName As String) As Boolean
    IsValidLevelName = False
    On Error GoTo ErrorHandler
    
    Dim oLevel As Level
    Set oLevel = ActiveDesignFile.Levels(LevelName)
    
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

' DistanceToRange
' Distance from a point to a range box - zero when the point is inside it.
'
' Measuring to the box, not to its centre, is what makes SearchRadius mean "how far from the element"
' rather than "how far from its middle". A 6 x 4.5 m cell has its centre up to 3.75 m from its own
' edge, so a cable end sitting INSIDE such a cell was rejected by a 1.5 m radius and its Repere came
' out blank - while a small cell twice as far away would have been accepted. Enlarging the radius
' hides that, it does not fix it: the radius would then have to grow with the biggest cell around.
Private Function DistanceToRange(ByRef Pt As Point3d, ByRef r As Range3d) As Double
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    If Pt.X < r.Low.X Then dx = r.Low.X - Pt.X
    If Pt.X > r.High.X Then dx = Pt.X - r.High.X
    If Pt.Y < r.Low.Y Then dy = r.Low.Y - Pt.Y
    If Pt.Y > r.High.Y Then dy = Pt.Y - r.High.Y
    If Pt.Z < r.Low.Z Then dz = r.Low.Z - Pt.Z
    If Pt.Z > r.High.Z Then dz = Pt.Z - r.High.Z

    DistanceToRange = Sqr(dx * dx + dy * dy + dz * dz)
End Function


' FindNearestElement
' Scans a bbox of SearchRadius (master units) around Pt for the closest candidate (by distance to
' the candidate's EXTENT, zero when Pt is inside it - see DistanceToRange), optionally
' restricted to ElTypes/Levels. When RequirePropertyName is non-empty, a candidate must carry a
' non-Null, non-blank value for that custom property (CustomPropertyHandler.GetPropertyValueFromElement)
' to qualify - without it, "nearest element of any kind" would catch unrelated annotation/symbol
' cells sharing the radius in a real drawing. Returns Nothing when no qualifying candidate lies
' within SearchRadius.
Public Function FindNearestElement(ByRef Pt As Point3d, ByVal SearchRadius As Double, _
                                   Optional ElTypes As Variant, _
                                   Optional Levels As Variant, _
                                   Optional RequirePropertyName As String = "") As Element
    On Error GoTo ErrorHandler

    Set FindNearestElement = Nothing
    If SearchRadius <= 0 Then Exit Function

    Dim oRange As Range3d
    oRange = Range3dFromXYZXYZ(Pt.X - SearchRadius, Pt.Y - SearchRadius, Pt.Z - SearchRadius, _
                               Pt.X + SearchRadius, Pt.Y + SearchRadius, Pt.Z + SearchRadius)

    Dim ee As ElementEnumerator
    Set ee = ByEE(Levels:=Levels, Range:=oRange, ElTypes:=ElTypes)
    If ee Is Nothing Then Exit Function

    Dim bRequireProp As Boolean
    bRequireProp = (Len(Trim(RequirePropertyName)) > 0)

    Dim oEl     As Element
    Dim oBest   As Element
    Dim dBest   As Double
    Dim dDist   As Double
    Dim rCand   As Range3d
    Dim vProp   As Variant
    Dim bQualifies As Boolean

    dBest = SearchRadius
    Do While ee.MoveNext
        Set oEl = ee.Current
        bQualifies = True
        If bRequireProp Then
            vProp = CustomPropertyHandler.GetPropertyValueFromElement(oEl, RequirePropertyName, RequirePropertyName)
            bQualifies = Not IsNull(vProp)
            If bQualifies Then bQualifies = (Len(Trim(CStr(vProp))) > 0)
        End If
        If bQualifies Then
            rCand = oEl.Range
            dDist = DistanceToRange(Pt, rCand)
            If dDist <= dBest Then
                dBest = dDist
                Set oBest = oEl
            End If
        End If
    Loop

    Set FindNearestElement = oBest
    Exit Function

ErrorHandler:
    Set FindNearestElement = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.FindNearestElement"
End Function

Public Function GetLevel(ByVal LevelName As String, Optional CanCreateLevel As Boolean = True) As Level
    On Error GoTo ErrorHandler
    Set GetLevel = Nothing
    
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "", "GetElements.GetLevel"
        Exit Function
    End If

    ' Get or create the output level
    If GetElements.IsValidLevelName(LevelName) Then
        Set GetLevel = ActiveDesignFile.Levels(LevelName)
    ElseIf CanCreateLevel = True Then
        Set GetLevel = ActiveDesignFile.AddNewLevel(LevelName)
        ActiveDesignFile.Levels.Rewrite
    End If
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GetElements.GetLevel"
End Function
