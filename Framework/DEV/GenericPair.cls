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

#If Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef src As Any, ByVal Length As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef src As Any, ByVal Length As Long)
#End If

Private Type Member
    Key As IGeneric
    Value As IGeneric
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Key() As IGeneric: Set Key = myclass.Key: End Property
Public Property Get Value() As Object: Set Value = myclass.Value: End Property

Public Function Build(ByVal Key As IGeneric, ByVal Value As IGeneric) As GenericPair
Attribute Build.VB_UserMemId = 0
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Key is nothing")
    
    If (Me Is GenericPair) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set myclass.Key = Key: Set myclass.Value = Value
'    Set Build = Skynet.CreateInstance(New GenericPair, Virtual.Pointer, Virtual.Size)
    Set Build = New GenericPair: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
    Set myclass.Key = GenericPair.Key
    Set myclass.Value = GenericPair.Value
End Sub

Private Sub Class_Terminate()
    myclass = EmptyStructure
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericPair: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long
    
    Const MAXLONG = 2 ^ 31 - 1
    Dim Hash As Currency
    
    If myclass.Value Is Nothing Then _
        IGeneric_HashValue = myclass.Key.HashValue: Exit Function
   
    Hash = (13 * CDbl(13 * 27 + myclass.Key.HashValue)) + myclass.Value.HashValue
    IGeneric_HashValue = CLng(Hash - (Fix(Hash / MAXLONG) * MAXLONG))
        
End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPair):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericMap, VarPtr(c), Virtual.Size)
  
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    Set otherPair = Other
    If myclass.Key.Equals(otherPair.Key) Then
        If (myclass.Value Is Nothing) = False Then
            IGeneric_Equals = myclass.Value.Equals(otherPair.Value)
        Else
            IGeneric_Equals = otherPair.Value Is Nothing
        End If
    End If

End Function

Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "[" & _
                            myclass.Key.ToString & _
                            ", "
                         
                        If (myclass.Value Is Nothing) = False Then
                            IGeneric_ToString = IGeneric_ToString & myclass.Value.ToString & "]"
                        Else
                            IGeneric_ToString = IGeneric_ToString & " NULL]"
                        End If
 
End Function

