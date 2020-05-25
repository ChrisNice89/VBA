VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements IObject

Private Type Member
    Value As String
    Internal As LongPtr
End Type
Private m As Member

Public Function Build(ByVal Value As String) As TString
    m.Value = Value
    Set Build = System.Create(New TString, Me, LenB(m))
End Function

Private Sub Class_Initialize()
    m.Internal = VarPtr(m)
End Sub

Private Function IObject_Clone() As IObject
    
    Dim Clone As TString
    Set Clone = New TString
    
    Dim c As Member
    
    With System
        c.Internal = .IObject(Clone).VirtualPointer
        c.Value = m.Value
        m.Internal = c.Internal

        IObject_Clone = .Create(Clone, Me, LenB(m))
    End With
    
    m.Internal = VarPtr(m)
    
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

Private Property Get IObject_VirtualPointer() As LongLong
    IObject_VirtualPointer = m.Internal
End Property
