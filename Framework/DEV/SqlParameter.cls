VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "<Sql>Class"
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IObject
Implements IPrototype

Private Type Member
    TValue As IValue
    Value As IValue
    Values As VBA.Collection
    Name As String
    Index As Byte
    Hashfunction As Hashing
    Parameter As ADODB.Parameter
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type

Private Virtual As Internal

Public Property Get Count() As Long
    Count = myClass.Values.Count
End Property

Public Property Get Index() As Byte
    Index = myClass.Index
End Property

Public Property Get Name() As String
    Name = myClass.Parameter.Name
End Property

Public Property Get TValue() As IValue
    Set TValue = myClass.TValue
End Property

Public Property Get Value(ByVal Index As Byte) As IValue
    Set Value = myClass.Values(Index)
End Property

Public Property Get Parameter() As ADODB.Parameter
    Set Parameter = myClass.Parameter
End Property

Public Sub Update(ByVal Value As IValue)
    With myClass
        If Not .TValue.Object.IsRelatedTo(Value) Then
        
        End If
        Set .Value = Value
        'Edit
        .Parameter.Size = Len(.Value.Value)
        .Parameter.Value = .Value.Value
   End With
End Sub

Public Function UseValue(ByVal Index As Byte) As IValue
    With myClass
        Call Me.Update(.Values(Index))
        Set UseValue = .Value
    End With
End Function

Public Property Get CurrentValue() As IValue
    Set CurrentValue = myClass.Value
End Property

Public Function AddValue(ByRef Value As IValue) As SqlParameter
    
    If myClass.Value.Object.IsRelatedTo(Value) Then
        Call myClass.Values.Add(Value)
    End If
    
    Set AddValue = Me
    
End Function

Public Function Build( _
                        ByRef TValue As IValue, _
                        Optional ByVal Name As String = "?") As SqlParameter
    With myClass
        Set .TValue = TValue
        Set .Parameter = New ADODB.Parameter
        With .Parameter
            .Name = Name
            .Type = myClass.TValue.DatabaseType
            .Size = Len(myClass.TValue.Value)
            .Value = myClass.TValue.Value
        End With
        .Index = Index
    End With
    
    Set Build = New SqlParameter
    'Set myClass.Value = New VBA.Collection
    'myclass.Values.Add Value
    'Set Build = Skynet.Create(Me)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    
    If Not Me Is SqlParameter Then
         With myClass
            Set .Parameter = SqlParameter.Parameter
            Set .TValue = SqlParameter.TValue
            Set .Values = New VBA.Collection
            .Index = SqlParameter.Index
        End With
        Call Skynet.CastObject(SqlParameter).Dispose
    End If

End Sub

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

'TODO TMap IObject_Dispose
Private Sub IObject_Dispose()
    myClass = EmptyStructure
End Sub

Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Dim p As SqlParameter: Set p = Other
    'If Not p.Count = Me.Count Then IObject_Equals = False: Exit Function
    
    IObject_Equals = (Me.CurrentValue.Object.Equals(p.CurrentValue.Object))
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum

    Dim i As Long
    
    If Other.IsRelatedTo(Me) Then
        Dim p As SqlParameter: Set p = Other
        IObject_CompareTo = myClass.Value.Object.CompareTo(p.CurrentValue)
    End If
    
End Function

Private Property Get IObject_HashValue() As Long:
    IObject_HashValue = 0
    'Skynet.HashValue(StrConv(myClass.Value, vbFromUnicode), myClass.HashFunction): End Property

End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is SqlParameter): End Function

Private Function IObject_ToString() As String: IObject_ToString = "": End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New SqlParameter: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

