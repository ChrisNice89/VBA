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
Private internal As Long

Public Function Build(ByVal Value As String) As TString
    m.Value = Value
    Set m.O = New VBA.Collection
    Set Build = System.Create(New TString, Me, LenB(m))
End Function

Public Property Get O() As Object
    Set O = m.O
End Property
Private Sub Class_Initialize()
    internal = VarPtr(m)
End Sub

Public Property Get Value() As String
    Value = m.Value
End Property

Private Function IObject_Clone() As IObject
    
    Dim Clone As TString
    Set Clone = New TString
    
    Dim c As Member
    
    With System
        'internal = VarPtr(c)
        Call .CopyData(VarPtr(c), VarPtr(m), LenB(m))
        internal = VarPtr(c)
        Set IObject_Clone = .Create(Clone, Me, LenB(m))
        internal = VarPtr(m)
    End With
End Function

Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
 
End Function

Private Function IObject_Equals(Other As IObject) As Boolean
    IObject_Equals = TypeOf Other Is TString
End Function

Private Property Get IObject_HashValue() As Long

End Property

Private Function IObject_ToString() As String

End Function

Private Property Get IObject_VirtualPointer() As Long
    IObject_VirtualPointer = internal
End Property
