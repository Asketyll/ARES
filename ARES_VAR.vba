' Module: ARES_VAR
' Description: This module provides configuration variables, absolute values, and explanations of each variable and where they are used.

' Dependencies: LangManager, ARES_MS_VAR_Class

'######################################################################################################################
'                       CAN'T BE EDITED IN MS VARIABLES ENVIRONMENT, YOU CAN MODIFY HERE
'######################################################################################################################

' Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const ARES_DEFAULT_GRAPHIC_GROUP_ID As Long = 0 ' Constant for no graphic group

' Used in Link and MicroStationDefinition module for check if a MsdElementType is unknow or raise a error
Public Const ARES_MSDETYPE_ERROR As Long = 44 ' If you use Type 44, you can replace with another MsdElementType not used

' Used in StringsInEl and AutoLengths module for separate a list in a environment variables
Public Const ARES_VAR_DELIMITER As String = "|" ' Delimiter used in :ARES_LENGTH_TRIGGER

' Used in Config and ARES_VAR
Public Const ARES_NAVD As String = "NaVD" ' Constant for undefined MS configuration variables :Not a Variable Defined

' Used in Length
Public Const ARES_RND_ERROR_VALUE As Byte = 255 ' Constant for error in ARES_ROUNDS and ARES_LENGTH_ROUND

'######################################################################################################################
'                             CAN BE MODIFIED IN MS ENVIRONMENT VARIABLES DO NOT MODIFY HERE
'                                   USE CONFIG MODULE TO GET, SET AND REMOVE A VALUE
'######################################################################################################################

' Declare a global collection to store ARES_MS_VAR variables
Dim MSVarsCollection As New Collection

' Used in Length module for default rounding
Public ARES_ROUNDS As ARES_MS_VAR_Class 'Default Value: 2  Range 0 to 254 (Byte -1)    #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in ElementChangeHandler ClassModule for automatically enable adding length to a text if the conditions are met.
Public ARES_AUTO_LENGTHS As ARES_MS_VAR_Class 'Dafault Value: True (ARES_AUTO_LENGTHS_DEFAULT)     True or False (Boolean)

' Used in Auto_Lengths module For length-specific rounding in text.
Public ARES_LENGTH_ROUND As ARES_MS_VAR_Class 'Default Value: 1 (ARES_LENGTH_RND_DEFAULT) Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in Auto_Lengths and StringsInEl module For Triggers in text.
Public ARES_LENGTH_TRIGGER As ARES_MS_VAR_Class 'Default Value: (Xx_m) (ARES_LENGTH_TRIGGER_DEFAULT) can a array use | (ARES_VAR_DELIMITER) like (Xx_m)|(Xx_cm)|(Xx_km)

' Used in AutoLengths and StringsInEl module for replace this triger with the length of element
Public ARES_LENGTH_TRIGGER_ID As ARES_MS_VAR_Class 'Default Value: Xx_  (ARES_LENGTH_TRIGGER_ID_DEFAULT)

' Used in CustomPropertyHandler module for default name of Library Type object
Public ARES_NAME_LIBRARY_TYPE As ARES_MS_VAR_Class 'Default Value: ARES

' Used in CustomPropertyHandler module for default name of Item Type object
Public ARES_NAME_ITEM_TYPE As ARES_MS_VAR_Class 'Default Value: ARESAutoLengthObject

' Used in LangManager module to force language if CONNECTUSER_LANGUAGE configuration variable is not set
Public ARES_LANGUAGE As ARES_MS_VAR_Class 'No Default Value

Public Function InitMSVars()
    Set ARES_ROUNDS = New ARES_MS_VAR_Class
    Set ARES_AUTO_LENGTHS = New ARES_MS_VAR_Class
    Set ARES_LENGTH_ROUND = New ARES_MS_VAR_Class
    Set ARES_LENGTH_TRIGGER = New ARES_MS_VAR_Class
    Set ARES_LENGTH_TRIGGER_ID = New ARES_MS_VAR_Class
    Set ARES_NAME_LIBRARY_TYPE = New ARES_MS_VAR_Class
    Set ARES_NAME_ITEM_TYPE = New ARES_MS_VAR_Class
    Set ARES_LANGUAGE = New ARES_MS_VAR_Class
    
    InitializeMSVar ARES_ROUNDS, "ARES_Round", "2"
    InitializeMSVar ARES_AUTO_LENGTHS, "ARES_Auto_Lengths", "True"
    InitializeMSVar ARES_LENGTH_ROUND, "ARES_Length_Round", "1"
    InitializeMSVar ARES_LENGTH_TRIGGER, "ARES_Length_Triggers", "(Xx_m)"
    InitializeMSVar ARES_LENGTH_TRIGGER_ID, "ARES_Length_Trigger_ID", "(Xx_)"
    InitializeMSVar ARES_NAME_LIBRARY_TYPE, "ARES_Library_Type_Name", "ARES"
    InitializeMSVar ARES_NAME_ITEM_TYPE, "ARES_Item_Type_Name", "ARESAutoLengthObject"
    InitializeMSVar ARES_LANGUAGE, "ARES_Language", ""
End Function

Private Function KeyExistsInCollection(key As String) As Boolean
    Dim i As Integer
    On Error GoTo ErrorHandler
    For i = 1 To MSVarsCollection.count
        If MSVarsCollection.Item(i).key = key Then
            KeyExistsInCollection = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    KeyExistsInCollection = False
End Function

Private Function InitializeMSVar(ByRef msVar As ARES_MS_VAR_Class, key As String, defaultValue As String)
    On Error GoTo ErrorHandler

    msVar.key = key
    msVar.Default = defaultValue
    msVar.Value = Config.GetVar(key)
    If msVar.Value = ARES_NAVD Then
        If Not ResetMSVar(msVar) Then GoTo ErrorHandler
    End If

    ' Check if the key already exists in the collection
    If Not KeyExistsInCollection(key) Then
        ' Add the variable to the collection
        MSVarsCollection.Add msVar, key
    End If

    Exit Function

ErrorHandler:
    MsgBox GetTranslation("VarInitializeMSVarfailed"), vbOKOnly
End Function

' Private function to get ARES_MS_VAR from a Variant
Private Function GetMSVarFromVariant(ByVal var As Variant) As ARES_MS_VAR_Class
    Dim key As String
    Dim i As Integer

    ' Check if the argument is of type ARES_MS_VAR_Class
    If TypeName(var) = "ARES_MS_VAR_Class" Then
        Set GetMSVarFromVariant = var
    ' Check if the argument is a String
    ElseIf VarType(var) = vbString Then
        key = var
        ' Find the variable in the collection
        For i = 1 To MSVarsCollection.count
            If MSVarsCollection.Item(i).key = key Then
                Set GetMSVarFromVariant = MSVarsCollection.Item(i)
                Exit Function
            End If
        Next i
        ShowStatus GetTranslation("VarKeyNotInCollection", key)
    Else
        ShowStatus GetTranslation("VarInvalidArgument")
    End If
End Function

' Function to reset a variable in the collection
Public Function ResetMSVar(ByVal var As Variant) As Boolean
    Dim msVar As ARES_MS_VAR_Class
    Set msVar = GetMSVarFromVariant(var)
    If msVar.key = "" Then
        ShowStatus GetTranslation("VarKeyNotFound", msVar.key)
        ResetMSVar = False
        Exit Function
    End If

    On Error GoTo ErrorHandler

    If Config.SetVar(msVar.key, msVar.Default) Then
        msVar.Value = Config.GetVar(msVar.key)
        ShowStatus GetTranslation("VarResetSuccess", msVar.Default)
        ResetMSVar = True
        Exit Function
    End If

ErrorHandler:
    ShowStatus "Impossible de réinitialiser la variable."
    ResetMSVar = False
End Function

' Function to remove a variable from the collection
Public Function RemoveMSVar(ByVal var As Variant, Optional showConfirmation As Boolean = True) As Boolean
    Dim msVar As ARES_MS_VAR_Class
    Set msVar = GetMSVarFromVariant(var)
    If msVar.key = "" Then
        ShowStatus GetTranslation("VarKeyNotFound", msVar.key)
        RemoveMSVar = False
        Exit Function
    End If

    On Error GoTo ErrorHandler

    ' Ask for confirmation before removing the variable
    If showConfirmation Then
        If MsgBox(GetTranslation("VarRemoveConfirm", msVar.key), vbYesNo) = vbNo Then
            RemoveMSVar = False
            Exit Function
        End If
    End If

    If Config.RemoveValue(msVar.key) Then
        ShowStatus GetTranslation("VarRemoveSuccess")
        RemoveMSVar = True
        Exit Function
    End If

ErrorHandler:
    ShowStatus GetTranslation("VarRemoveError")
    RemoveMSVar = False
End Function

Public Function ResetAllMSVar() As Boolean
    Dim var As Variant
    Dim success As Boolean
    success = True

    For Each var In MSVarsCollection
        If Not ResetMSVar(var.key) Then
            success = False
        End If
    Next var

    ResetAllMSVar = success
End Function

Public Function RemoveAllMSVar() As Boolean
    Dim var As Variant
    Dim success As Boolean

    ' Ask for confirmation before removing all variables
    If MsgBox(GetTranslation("VarsRemoveConfirm"), vbYesNo) = vbNo Then
        RemoveAllMSVar = False
        Exit Function
    End If

    success = True
    For Each var In MSVarsCollection
        If Not RemoveMSVar(var.key, False) Then
            success = False
        End If
    Next var

    RemoveAllMSVar = success
End Function
