VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database

Implements IGeneric
Implements IPrototype

Private Type Member
    Key As IGeneric
    Value As Object
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Key() As IGeneric: Set Key = myClass.Key: End Property
Public Property Get Value() As Object: Set Value = myClass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As Object) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Key is nothing")
    
    With myClass
        Set .Key = Key: Set .Value = Value
    End With
    
    Set Build = New GenericPair
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    With GenericPair
        Set myClass.Key = .Key: Set myClass.Value = .Value
    End With
End Sub

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal other As IGeneric) As Boolean
    
    Dim pair As GenericPair
    
    If other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If TypeOf other Is GenericPair Then
        Set pair = other
        If Not myClass.Value Is Nothing Then
            IGeneric_Equals = myClass.Key.Equals(pair.Key) And myClass.Value.Equals(pair.Value): Exit Function
        Else
            IGeneric_Equals = myClass.Key.Equals(pair.Key) And pair.Value Is Nothing: Exit Function
        End If
    End If
    
    IGeneric_Equals = myClass.Key.Equals(other)
    
End Function

Private Function IGeneric_CompareTo(ByVal other As IGeneric) As ComparerEnum
    
    Dim pair As GenericPair
    
    If other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    If TypeOf other Is GenericPair Then
        Set pair = other
        IGeneric_CompareTo = myClass.Key.CompareTo(pair.Key): Exit Function
    End If
    
    IGeneric_CompareTo = myClass.Key.CompareTo(other)
     
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf other Is GenericPair): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & vbNewLine & _
                        "[" & _
                            myClass.Key.ToString & _
                            ", " & _
                            myClass.Value.ToString & _
                         "]"
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPair: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericPair: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


