VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IPrototype

Private Type Member
    Key As IGeneric
    Value As IGeneric
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Key() As IGeneric: Set Key = myClass.Key: End Property
Public Property Get Value() As Object: Set Value = myClass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As IGeneric) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Key is nothing")
    
    If (Me Is GenericPair) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set myClass.Key = Key: Set myClass.Value = Value
  
    Set Build = New GenericPair 'Skynet.CreateInstance(New GenericPair, Virtual.Pointer, Virtual.Size)
    myClass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
    With GenericPair
        Set myClass.Key = .Key: Set myClass.Value = .Value
    End With
End Sub

'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericMap, VarPtr(C), Virtual.Size)
  
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherPair As GenericPair
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    Set otherPair = Other
    If myClass.Key.Equals(otherPair.Key) Then
        If (myClass.Value Is Nothing) = False Then
            IGeneric_Equals = myClass.Value.Equals(otherPair.Value)
        Else
            IGeneric_Equals = otherPair.Value Is Nothing
        End If
    End If

End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function

    Set otherPair = Other
    IGeneric_CompareTo = myClass.Key.CompareTo(otherPair.Key)
    
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPair): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "[" & _
                            myClass.Key.ToString & _
                            ", "
                         
                        If (myClass.Value Is Nothing) = False Then
                            IGeneric_ToString = IGeneric_ToString & myClass.Value.ToString & "]"
                        Else
                            IGeneric_ToString = IGeneric_ToString & " NULL]"
                        End If
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPair: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericPair: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


