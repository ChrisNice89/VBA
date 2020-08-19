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

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    With GenericPair
        Set myClass.Key = .Key: Set myClass.Value = .Value
    End With
End Sub

Public Function Build(ByVal Key As IGeneric, ByVal Value As Object) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    With myClass
        Set .Key = Key: Set .Value = Value
    End With
    
    Set Build = New GenericPair
    
End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    Dim pair As GenericPair: Set pair = Other
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If Not myClass.Value Is Nothing Then _
        If TypeOf myClass.Value Is IGeneric Then _
            IGeneric_Equals = myClass.Key.Equals(pair.Key) And myClass.Value.Equals(pair.Value): Exit Function
 
    IGeneric_Equals = myClass.Key.Equals(pair.Key) And myClass.Value Is pair.Value
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    Dim pair As GenericPair: Set pair = Other
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    IGeneric_CompareTo = myClass.Key.CompareTo(pair.Key)
    
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPair): End Function
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


