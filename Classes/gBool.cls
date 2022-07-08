VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gBool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Primitives"
Option Explicit

Implements IGeneric
Implements IGenericValue

Private Const DEFAULT_VALUE As Boolean = False
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
Private This As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is gBool: End Property
Public Property Get IsTrue() As Boolean: IsTrue = This.Value = TRUE_VALUE: End Property
Public Property Get Invert() As gBool: Set Invert = gBool.Of(VBA.IIf(This.Value = TRUE_VALUE, False, True)): End Property

Public Property Get ToByte() As Byte: ToByte = This.Value: End Property
'@DefaultMember
Public Property Get ToBoolean() As Boolean: ToBoolean = VBA.IIf(This.Value = TRUE_VALUE, True, False): End Property
Attribute ToBoolean.VB_UserMemId = 0
Private Static Function BuildTrue() As gBool
    
    Dim TrueValue As gBool
    If TrueValue Is Nothing Then
        This.Value = TRUE_VALUE
        Set TrueValue = New gBool
        This = EmptyStructure
    End If
    
    Set BuildTrue = TrueValue
    
End Function

Private Static Function BuildFalse() As gBool
    
    Dim FalseValue As gBool
    If FalseValue Is Nothing Then
        This.Value = FALSE_VALUE
        Set FalseValue = New gBool
        This = EmptyStructure
    End If
    
    Set BuildFalse = FalseValue
    
End Function

Public Function Of(ByRef Value As Variant) As gBool
    
    If Not (Me Is gBool) Then _
        Set Of = gBool.Of(Value): Exit Function
        
    Select Case VBA.VarType(Value)
        Case VBA.VbVarType.vbNull, VBA.VbVarType.vbEmpty
            Set Of = gBool
        Case VBA.VbVarType.vbBoolean
            If Value = True Then
                Set Of = BuildTrue
            Else
                Set Of = BuildFalse
            End If
        Case VBA.VbVarType.vbByte, VBA.VbVarType.vbInteger, VBA.VbVarType.vbLong
            If Value = TRUE_VALUE Then
                Set Of = BuildTrue
            Else
                Set Of = BuildFalse
            End If
        Case VBA.VbVarType.vbString
            If VBA.StrComp(VBA.Trim$(Value), TRUE_STRING, vbTextCompare) = 0 Then
                Set Of = BuildTrue
            Else
                Set Of = BuildFalse
            End If
        Case Else
            If CBool(Value) Then
                Set Of = BuildTrue
            Else
                Set Of = BuildFalse
            End If
    End Select
    
End Function

Private Sub Class_Initialize(): This.Value = gBool.ToByte: End Sub

Public Function MaxOf(ByVal A As Byte, ByVal B As Byte) As gBool: Set MaxOf = gBool.Of(VBA.IIf(A > B, A, B)): End Function
Public Function MinOf(ByVal A As Byte, ByVal B As Byte) As gBool: Set MinOf = gBool.Of(VBA.IIf(A < B, A, B)): End Function
Public Function EqualsTo(ByVal Value As Byte) As Boolean: EqualsTo = This.Value = Value: End Function
Public Function CompareTo(ByVal Value As Byte) As CompareResult: CompareTo = Compare(This.Value, Value): End Function
Public Function Compare(ByVal A As Byte, ByVal B As Byte) As CompareResult: Compare = (A > B) - (A < B): End Function

Public Function IsBetween(ByVal A As Byte, B As Byte) As Boolean
    
    If (A > This.Value) Then _
        Exit Function
    
    If (B < This.Value) Then _
        Exit Function

    IsBetween = True
    
End Function

Public Function ToFormat(ByVal Format As BooleanFormat) As String

    Select Case Format
        Case BooleanFormat.Default_bool: ToFormat = IIf(This.Value = TRUE_VALUE, TRUE_STRING, FALSE_STRING)
        Case BooleanFormat.YesNo_Bool: ToFormat = IIf(This.Value = TRUE_VALUE, "Yes", "No")
        Case BooleanFormat.JaNein_Bool: ToFormat = IIf(This.Value = TRUE_VALUE, "Ja", "Nein")
        Case BooleanFormat.WahrFalsch_Bool: ToFormat = IIf(This.Value = TRUE_VALUE, "Wahr", "Falsch")
        Case BooleanFormat.EinsNull_Bool: ToFormat = IIf(This.Value = TRUE_VALUE, "1", "0")
    End Select

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = gBool: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = This.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean
    
    If Not (Other Is Nothing) Then
        IGeneric_IsRelatedTo = (TypeOf Other Is gBool)
    Else
        IGeneric_IsRelatedTo = True
    End If
        
End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean
    
    If (Other Is Nothing) Then _
        Exit Function
        
    If TypeOf Other Is gBool Then
        Dim otherBool As gBool: Set otherBool = Other
        IGeneric_Equals = (This.Value = otherBool.ToByte)
    End If
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Me.ToFormat(Format:=Default_bool): End Function

'IGenericValue
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_GetValueOrDefault(Optional ByVal DefaultValue As Variant = Null) As Variant: IGenericValue_GetValueOrDefault = VBA.IIf(Me Is gBool, DefaultValue, This.Value): End Property
Private Property Get IGenericValue_HasValue() As Boolean: IGenericValue_HasValue = Not (Me Is gBool): End Property
Private Property Get IGenericValue_VarType() As VBA.VbVarType: IGenericValue_VarType = VBA.VbVarType.vbBoolean: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(This.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Binary_Type: End Property
Private Property Get IGenericValue_GetValue() As Variant: IGenericValue_GetValue = This.Value: End Property
Private Property Get IGenericValue_ToString() As String: IGenericValue_ToString = This.Value: End Property
Private Function IGenericValue_HashCode() As Long: IGenericValue_HashCode = This.Value: End Function
Private Function IGenericValue_Of(ByRef Value As Variant) As IGenericValue: Set IGenericValue_Of = gBool.Of(Value): End Function

Private Function IGenericValue_Equals(ByVal Other As IGenericValue) As Boolean: IGenericValue_Equals = (This.Value = Other.GetValue): End Function
Private Function IGenericValue_IsLessThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsLessThan = (This.Value < Other.GetValue): End Function
Private Function IGenericValue_IsGreaterThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsGreaterThan = (This.Value > Other.GetValue): End Function

Private Function IGenericValue_IsBetween(ByVal A As IGenericValue, ByVal B As IGenericValue) As Boolean
    
    If (A Is Nothing) Then _
        Exit Function
    
    If (B Is Nothing) Then _
        Exit Function
    
    If (A.HasValue) Then _
        If (B.HasValue) Then _
            IGenericValue_IsBetween = Me.IsBetween(A.GetValue, B.GetValue)
    
End Function

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = IGenericValue_VarType
        .Name = Name
        .Value = This.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function
