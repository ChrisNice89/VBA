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
Implements IGenericComparer

Private Type Member
    Key As IGeneric
    Value As IGeneric
End Type
Private myclass As Member

Public Property Get Key() As IGeneric: Set Key = myclass.Key: End Property
Public Property Get Value() As Object: Set Value = myclass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As IGeneric) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    If (Me Is GenericPair) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    If (Key Is Nothing) Or (Value Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("(Key Is Nothing) Or (Value Is Nothing)")
        
    Set myclass.Key = Key: Set myclass.Value = Value
    Set Build = New GenericPair: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    With GenericPair
        Set myclass.Key = .Key: Set myclass.Value = .Value
    End With
End Sub

Private Sub Class_Terminate()
    myclass = EmptyStructure
End Sub

Private Function IGenericComparer_Compare(ByVal a As IGeneric, ByVal b As IGeneric) As ComparerEnum

    Dim x As GenericPair, y As GenericPair
    Set x = a: Set y = b
    'default comparer
    IGenericComparer_Compare = IGenericValue.Comparer.Compare(x.Key, y.Key)
    
End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric
    With New IGeneric 'Bug, caused while GenericPair implements IGeneric
        Set IGeneric_Default = GenericPair.Build(.Default, .Default)
    End With
End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long
    
    Const MAXLONG = 2 ^ 31 - 1
    Static Hash As Currency
    
    If Hash = 0 Then _
        Hash = (13 * CDbl(13 * 27 + myclass.Key.HashValue)) + myclass.Value.HashValue
    
    IGeneric_HashValue = CLng(Hash - (Fix(Hash / MAXLONG) * MAXLONG))
        
End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean
    
    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        Exit Sub
        
    Set otherPair = Other
    
    If myclass.Key.IsRelatedTo(otherPair.Key) Then _
        IGeneric_IsRelatedTo = myclass.Value.IsRelatedTo(otherPair.Value)
    
End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = GenericPair.Build(myclass.Key, myclass.Value): End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        Exit Function
        
    Set otherPair = Other
    
    If myclass.Key.Equals(otherPair.Key) Then _
        IGeneric_Equals = myclass.Value.Equals(otherPair.Value)

End Function

Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "[" & _
                            myclass.Key.ToString & _
                            "," & myclass.Value.ToString & "]"
                                               
End Function
