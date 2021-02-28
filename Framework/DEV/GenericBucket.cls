VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric

Private Type Member
    Key As IGeneric
    Value As IGeneric
    HashCode As Long
    After As GenericBucket
    Before As GenericBucket
    State As Byte
End Type
Private myclass As Member

Public Property Get HashCode() As Long: HashCode = myclass.HashCode: End Property
Public Property Let HashCode(ByVal Value As Long):  myclass.HashCode = Value: End Property

Public Property Get State() As Byte: State = myclass.State: End Property
Public Property Let State(ByVal Value As Byte):  myclass.State = Value: End Property

Public Property Get After() As GenericBucket: Set After = myclass.After: End Property
Public Property Set After(ByVal B As GenericBucket): Set myclass.After = B: End Property

Public Property Get Before() As GenericBucket: Set Before = myclass.Before: End Property
Public Property Set Before(ByVal B As GenericBucket): Set myclass.Before = B: End Property

Public Property Get Key() As IGeneric: Set Key = myclass.Key: End Property
Public Property Set Key(ByVal K As IGeneric): Set myclass.Key = K: End Property

Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property
Public Property Set Value(ByVal V As IGeneric): Set myclass.Value = V: End Property

Public Sub Invalidate()
    myclass = EmptyStructure
End Sub

'Public Function Build( _
'                        ByVal Key As IGeneric, _
'                        ByVal Value As IGeneric, _
'                        Optional ByVal Before As GenericBucket) As GenericNode
'
'    If (Me Is GenericBucket) = False Then _
'        Call System.Throw(Me, "Public Function Build").BuildExeption
'
'    Dim newBucket As GenericBucket
'    Set newBucket = New GenericBucket
'
'    With newBucket
'        Set .Key = Key
'        Set .Value = Value
'        Set .Before = Before
'    End With
'
'    Set Build = newBucket
'
'End Function

Public Function ToPair() As GenericPair: Set ToPair = GenericPair(myclass.Key, myclass.Value): End Function

Private Sub Class_Terminate()
    myclass = EmptyStructure
'    Debug.Print "Node_Terminate"
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericBucket: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericBucket: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long
    
    Const MAXLONG = 2 ^ 31 - 1
    Dim Hash As Currency
    Hash = 13
    
    If Not myclass.Key Is Nothing Then _
        Hash = (Hash * CDbl(Hash * 27 + myclass.Key.HashValue))
    
    If Not myclass.Value Is Nothing Then _
        Hash = Hash + myclass.Value.HashValue
    
    IGeneric_HashValue = CLng(Hash - (Fix(Hash / MAXLONG) * MAXLONG))
        
End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean:  IGeneric_IsRelatedTo = TypeOf Other Is GenericBucket: End Property

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric

End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

     
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


