VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IObject

Private Type Member
    Value As String
    O As Object
End Type
Private m As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(ByVal Value As String) As TString
    m.Value = Value
    Set m.O = New VBA.Collection
    Set Build = System.Create(New TString, Me, LenB(m))
End Function

Public Property Get O() As Object
    Set O = m.O
End Property
Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(m)
End Sub

Public Property Get Value() As String
    Value = m.Value
End Property

Private Function IObject_Clone() As IObject
    
    Dim Clone As TString
    Set Clone = New TString
    
    Dim c As Member
    
    With System
'        Call .CopyData(VarPtr(c), VarPtr(m), LenB(m))
'        internal = VarPtr(c)
'        Set IObject_Clone = .Create(Clone, Me, LenB(m))
'        internal = VarPtr(m)
        
        Call .CopyData(VarPtr(c), Virtual.Pointer, .IObject(Me).Size)
        Virtual.Pointer = VarPtr(c)
        Set IObject_Clone = .Create(Clone, Me, .IObject(Me).Size)
        Virtual.Pointer = VarPtr(m)
    End With
End Function

Private Function IObject_Serialize() As Byte()
    Dim c As Member
    Dim b() As Byte
    ReDim b(1 To LenB(Member))
    
    Call System.CopyData(VarPtr(b), VarPtr(m), UBound(m))
    IObject_Serialize = b

End Function

Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
 
End Function

Private Sub IObject_Dispose()

End Sub

Private Function IObject_Equals(Other As IObject) As Boolean
    IObject_Equals = TypeOf Other Is TString
End Function

Private Property Get IObject_HashValue() As Long

End Property

Private Function IObject_ToString() As String

End Function

Private Property Get IObject_VirtualPointer() As LongPtr
    IObject_VirtualPointer = Virtual.Pointer
End Property
