VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericPairXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")

Option Explicit
'@PredeclaredId

Option Compare Database

Implements IGeneric
Implements IPrototype

Private Type Member
    Key As IGeneric
    Value As Object
End Type
Private myClass As Member

Private Type Internal
    size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Key() As IGeneric: Set Key = myClass.Key: End Property
Public Property Get Value() As Object: Set Value = myClass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As Object) As GenericPairXY
Attribute Build.VB_UserMemId = 0
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Key is nothing")
    
    With myClass
        Set .Key = Key: Set .Value = Value
    End With
    
    Set Build = Skynet.CreateInstance(New GenericPairXY, Virtual.Pointer, Virtual.size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.size = LenB(myClass)
End Sub

Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericMap, VarPtr(C), Virtual.size)
  
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    

    Dim otherPair As GenericPairXY
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If TypeOf Other Is GenericPairXY Then
        If Not myClass.Key Is Nothing Then
            If Not myClass.Value Is Nothing Then
                IGeneric_Equals = myClass.Key.Equals(otherPair.Key) And myClass.Value.Equals(otherPair.Value): Exit Function
            Else
                IGeneric_Equals = myClass.Key.Equals(otherPair.Key) And otherPair.Value Is Nothing: Exit Function
            End If
        End If
        
    End If
    
    IGeneric_Equals = myClass.Key.Equals(Other)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    Dim Pair As GenericPairXY
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    If TypeOf Other Is GenericPairXY Then
        Set Pair = Other
        IGeneric_CompareTo = myClass.Key.CompareTo(Pair.Key): Exit Function
    End If
    
    IGeneric_CompareTo = myClass.Key.CompareTo(Other)
     
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPairXY): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & vbNewLine & _
                        "[" & _
                            myClass.Key.ToString & _
                            ", "
                         
                        If Not myClass.Value Is Nothing Then
                            IGeneric_ToString = IGeneric_ToString & myClass.Value.ToString & "]"
                        Else
                            IGeneric_ToString = IGeneric_ToString & " Nothing]"
                        End If
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPairXY: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericPairXY: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


