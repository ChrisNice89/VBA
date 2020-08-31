VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IPrototype
Implements IValue

Public Enum TBooleanFormat
    [_First]
    DefaulTBoolean
   
    [_Last]
End Enum

Private Type Member
    Value As Boolean
    Format As TBooleanFormat
   
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Object() As IValue: Set Object = Me: End Property
Public Property Get Value() As Boolean: Value = myClass.Value: End Property

Public Function Build( _
                        ByVal Value As Boolean, _
                        Optional ByVal Format As TBooleanFormat = TBooleanFormat.DefaulTBoolean) As GBoolean
Attribute Build.VB_UserMemId = 0
With myClass
    .Value = Value
    .Format = Format
End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal other As IGeneric) As Boolean

    If other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not other.IsRelatedTo(Me) Then IGeneric_Equals = False: Exit Function

    Dim b As GBoolean: Set b = other
   
    IGeneric_Equals = (myClass.Value = b.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal other As IGeneric) As ComparerEnum
    Dim b As GBoolean
    
    If other.IsRelatedTo(Me) Then
        Set b = other
        Select Case True
            Case myClass.Value > b.Value: IGeneric_CompareTo = IsGreater
            Case myClass.Value < b.Value: IGeneric_CompareTo = IsLower
            Case Else: IGeneric_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = 0: End Property
Private Function IGeneric_IsRelatedTo(ByVal other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf other Is GBoolean): End Function
Private Function IGeneric_ToString() As String

    Select Case myClass.Format
        
        
    End Select
    
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GBoolean: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GBoolean: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

