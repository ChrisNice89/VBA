VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GBool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")
Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericValue

Private Const TRUE_STRING As String = "True"
Private Const FALSE_STRING As String = "False"
Private Const TRUE_VALUE As Byte = 1
Private Const FALSE_VALUE As Byte = 0

Public Enum BooleanFormat
    [_First]
    Default_bool
    YesNo_Bool
    WahrFalsch_Bool
    JaNein_Bool
    EinsNull_Bool
    [_Last]
End Enum

Private Type Member
    Value As Byte
 
End Type
Private myclass As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is GBool: End Property
Public Property Get IsTrue() As Boolean: IsTrue = myclass.Value = TRUE_VALUE: End Property
Public Property Get Value() As Byte: Value = myclass.Value: End Property
Public Property Get Invert() As GBool: Set Invert = GBool.Build(IIf(myclass.Value = TRUE_VALUE, False, True)): End Property

Private Static Function BuildTrue() As GBool
    
    Dim TrueValue As GBool
    If TrueValue Is Nothing Then
        myclass.Value = TRUE_VALUE
        Set TrueValue = New GBool
        myclass = EmptyStructure
    End If
    
    Set BuildTrue = TrueValue
    
End Function

Private Static Function BuildFalse() As GBool
    
    Dim FalseValue As GBool
    If FalseValue Is Nothing Then
        myclass.Value = FALSE_VALUE
        Set FalseValue = New GBool
        myclass = EmptyStructure
    End If
    
    Set BuildFalse = FalseValue
    
End Function

Public Function Build(ByVal Value As Boolean) As GBool
Attribute Build.VB_UserMemId = 0
    
    If Not (Me Is GBool) Then _
        Set Build = GBool.Build(Value): Exit Function
        
    If Value = True Then
        Set Build = BuildTrue
    Else
        Set Build = BuildFalse
    End If
    
End Function

Private Sub Class_Initialize()
    myclass.Value = GBool.Value
End Sub

Public Function ToBoolean() As Boolean: ToBoolean = IIf(myclass.Value = TRUE_VALUE, True, False): End Function
Public Function ToFormat(ByVal Format As BooleanFormat) As String

    Select Case Format
        Case BooleanFormat.Default_bool: ToFormat = IIf(myclass.Value = TRUE_VALUE, TRUE_STRING, FALSE_STRING)
        Case BooleanFormat.YesNo_Bool: ToFormat = IIf(myclass.Value = TRUE_VALUE, "Yes", "No")
        Case BooleanFormat.JaNein_Bool: ToFormat = IIf(myclass.Value = TRUE_VALUE, "Ja", "Nein")
        Case BooleanFormat.WahrFalsch_Bool: ToFormat = IIf(myclass.Value = TRUE_VALUE, "Wahr", "Falsch")
        Case BooleanFormat.EinsNull_Bool: ToFormat = IIf(myclass.Value = TRUE_VALUE, "1", "0")
    End Select

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GBool: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GBool):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function

    If Not TypeOf Other Is GBool Then _
        IGeneric_Equals = False: Exit Function
   
    Dim otherBool As GBool: Set otherBool = Other
    IGeneric_Equals = (myclass.Value = otherBool.Value)
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Me.ToFormat(Format:=Default_bool): End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbBoolean: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Binary_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericComparer: End Property

Private Function IGenericValue_Convert(ByVal Value As Variant) As IGenericValue

    If IsNull(Value) Then _
        Set IGenericValue_Convert = IGeneric_Default: Exit Function
    
    Select Case VarType(Value)
        Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong
            If Value = TRUE_VALUE Then
                Set IGenericValue_Convert = GBool.Build(True)
            Else
                Set IGenericValue_Convert = GBool.Build(False)
            End If
        Case VbVarType.vbBoolean
            If Value = True Then
                Set IGenericValue_Convert = GBool.Build(True)
            Else
                Set IGenericValue_Convert = GBool.Build(False)
            End If
        Case VbVarType.vbString
            If VBA.StrComp(VBA.Trim$(Value), TRUE_STRING, vbTextCompare) = 0 Then
                Set IGenericValue_Convert = GBool.Build(True)
            Else
                Set IGenericValue_Convert = GBool.Build(False)
            End If
        Case Else
            If CBool(Value) Then
                Set IGenericValue_Convert = GBool.Build(True)
            Else
                Set IGenericValue_Convert = GBool.Build(False)
            End If
    End Select

End Function

Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As CompareResult
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim otherBool As GBool: Set otherBool = Other
    
    If myclass.Value = otherBool.Value Then _
        IGenericValue_CompareTo = IsEqual: Exit Function
    
    If myclass.Value = FALSE_VALUE Then _
        IGenericValue_CompareTo = IsLower: Exit Function
        
    IGenericValue_CompareTo = IsGreater
     
End Function

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = IGenericValue_VarType
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = P
    
End Function



