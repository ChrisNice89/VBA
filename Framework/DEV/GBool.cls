VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gbool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IGeneric
Implements GenericValue

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
Private myclass As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is Gbool: End Property
Public Property Get IsTrue() As Boolean: IsTrue = myclass.Value = TRUE_VALUE: End Property
Public Property Get Value() As Byte: Value = myclass.Value: End Property
Public Property Get Invert() As Gbool: Set Invert = Gbool.Of(IIf(myclass.Value = TRUE_VALUE, False, True)): End Property

Private Static Function BuildTrue() As Gbool
    
    Dim TrueValue As Gbool
    If TrueValue Is Nothing Then
        myclass.Value = TRUE_VALUE
        Set TrueValue = New Gbool
        myclass = EmptyStructure
    End If
    
    Set BuildTrue = TrueValue
    
End Function

Private Static Function BuildFalse() As Gbool
    
    Dim FalseValue As Gbool
    If FalseValue Is Nothing Then
        myclass.Value = FALSE_VALUE
        Set FalseValue = New Gbool
        myclass = EmptyStructure
    End If
    
    Set BuildFalse = FalseValue
    
End Function

Public Function Of(ByVal Value As Boolean) As Gbool
Attribute Of.VB_UserMemId = 0
    
    If Not (Me Is Gbool) Then _
        Set Of = Gbool.Of(Value): Exit Function
        
    If Value = True Then
        Set Of = BuildTrue
    Else
        Set Of = BuildFalse
    End If
    
End Function

Private Sub Class_Initialize()
    myclass.Value = Gbool.Value
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
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = Gbool: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is Gbool):  End Property

Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function

    If Not TypeOf Other Is Gbool Then _
        IGeneric_Equals = False: Exit Function
   
    Dim otherBool As Gbool: Set otherBool = Other
    IGeneric_Equals = (myclass.Value = otherBool.Value)
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Me.ToFormat(Format:=Default_bool): End Function

'GenericValue
Private Property Get GenericValue_VarType() As VbVarType: GenericValue_VarType = vbBoolean: End Property
Private Property Get GenericValue_Size() As Byte: GenericValue_Size = LenB(myclass.Value): End Property
Private Property Get GenericValue_SqlType() As ISqlDataType: GenericValue_SqlType = ISqlDataType.Binary_Type: End Property
Private Property Get GenericValue_ToValue() As Variant: GenericValue_ToValue = myclass.Value: End Property
Private Property Get GenericValue_Instance() As IGeneric: Set GenericValue_Instance = Me: End Property
Private Property Get GenericValue_Comparer() As IGenericComparer: Set GenericValue_Comparer = IGenericComparer: End Property

Private Function GenericValue_Convert(ByRef Value As Variant) As GenericValue

    If VBA.IsNull(Value) Then _
        Set GenericValue_Convert = IGeneric_Default: Exit Function
        
    Select Case VarType(Value)
        Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong
            If Value = TRUE_VALUE Then
                Set GenericValue_Convert = Gbool.Of(True)
            Else
                Set GenericValue_Convert = Gbool.Of(False)
            End If
        Case VbVarType.vbBoolean
            If Value = True Then
                Set GenericValue_Convert = Gbool.Of(True)
            Else
                Set GenericValue_Convert = Gbool.Of(False)
            End If
        Case VbVarType.vbString
            If VBA.StrComp(VBA.Trim$(Value), TRUE_STRING, vbTextCompare) = 0 Then
                Set GenericValue_Convert = Gbool.Of(True)
            Else
                Set GenericValue_Convert = Gbool.Of(False)
            End If
        Case Else
            If CBool(Value) Then
                Set GenericValue_Convert = Gbool.Of(True)
            Else
                Set GenericValue_Convert = Gbool.Of(False)
            End If
    End Select

End Function

Private Function GenericValue_CompareTo(ByVal Other As GenericValue) As CompareResult
    
    If Other Is Nothing Then _
        GenericValue_CompareTo = isgreater: Exit Function
    
    Dim otherBool As Gbool: Set otherBool = Other
    
    If myclass.Value = otherBool.Value Then _
        GenericValue_CompareTo = IsEqual: Exit Function
    
    If myclass.Value = FALSE_VALUE Then _
        GenericValue_CompareTo = islower: Exit Function
        
    GenericValue_CompareTo = isgreater
     
End Function

Private Function GenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = GenericValue_VarType
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set GenericValue_ToSqlParameter = P
    
End Function
