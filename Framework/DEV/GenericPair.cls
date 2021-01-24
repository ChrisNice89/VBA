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
Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As IGeneric) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    If (Me Is GenericPair) = False Then _
        Call System.Throw(Me, "Public Function Build").BuildExeption
'
    Set myclass.Key = Key: Set myclass.Value = Value
    Set Build = New GenericPair: Set myclass.Key = Nothing: Set myclass.Value = Nothing
    
End Function

Private Sub Class_Initialize()
    With GenericPair
        Set myclass.Key = .Key: Set myclass.Value = .Value
    End With
End Sub

Private Sub Class_Terminate()
    myclass = EmptyStructure
End Sub
'IGenericComparer -Key comparer
Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long
    Dim pair As GenericPair
    Set pair = Obj
    IGenericComparer_GetHashCode = System.DefaultComparer.GetHashCode(pair.Key)
End Function

Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult

    Dim x As GenericPair, y As GenericPair
    Set x = A: Set y = B
    
    IGenericComparer_Compare = System.DefaultComparer.Compare(x.Key, y.Key)
    
End Function
'
Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean

    Dim x As GenericPair, y As GenericPair
    Set x = A: Set y = B

    IGenericComparer_Equals = System.DefaultComparer.Equals(x.Key, y.Key)

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericPair: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericPair: End Property
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

Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        Exit Function
    
    If Not TypeOf Other Is GenericPair Then _
        IGeneric_Equals = False: Exit Function
        
    Set otherPair = Other
    
    If myclass.Key Is Nothing = False Then
        If myclass.Key.Equals(otherPair.Key) = False Then _
            Exit Function
    Else
        If otherPair.Key Is Nothing = False Then _
            Exit Function
    End If
    
    If myclass.Value Is Nothing = False Then
        If myclass.Value.Equals(otherPair.Value) = False Then _
            Exit Function
    Else
        If otherPair.Value Is Nothing = False Then _
            Exit Function
    End If
    
    IGeneric_Equals = True
    
End Function

Private Function IGeneric_ToString() As String
    
        IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "["
        If Not myclass.Key Is Nothing Then
            IGeneric_ToString = IGeneric_ToString & myclass.Key.ToString
        Else
            IGeneric_ToString = IGeneric_ToString & "NULL"
        End If
        
        IGeneric_ToString = IGeneric_ToString & ";"
        
        If Not myclass.Value Is Nothing Then
            IGeneric_ToString = IGeneric_ToString & myclass.Key.ToString
        Else
            IGeneric_ToString = IGeneric_ToString & "NULL"
        End If
        
        IGeneric_ToString = IGeneric_ToString & "]"
                                               
End Function
