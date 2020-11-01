VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "<T>Class"
Option Explicit

Implements IGeneric

Private Type Member
    Object As Object
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Object() As Object: Set Object = myclass.Object: End Property
Public Function Build(ByVal Value As Object) As GObject
    
    Set myclass.Object = Value
    Set Build = New GObject
    Set myclass.Object = Nothing
    'Set Build = Skynet.CreateInstance(New GObject, Virtual.Pointer, Virtual.Size)
 
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
    Set myclass.Object = GObject.Object
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GObject: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GObject):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GObject, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    Dim o As GObject: Set o = Other

    IGeneric_Equals = (myclass.Object Is o.Object)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    Dim o As GObject: Set o = Other
    
    Dim i As LongPtr
    i = ObjPtr(o.Object)
    
    Select Case True
        Case ObjPtr(myclass.Object) > i: IGeneric_CompareTo = IsGreater
        Case ObjPtr(myclass.Object) < i: IGeneric_CompareTo = IsLower
        Case Else: IGeneric_CompareTo = IsEqual
    End Select
    
End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = TypeName$(Me) & " :: " & TypeName$(myclass.Object)
End Function
