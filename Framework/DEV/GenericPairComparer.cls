VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericPairComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")
Option Explicit
Implements IGenericComparer

Private Type Member
    KeyComparer As IGenericComparer
End Type
Private myclass As Member

Public Property Get KeyComparer() As IGenericComparer
    If myclass.KeyComparer Is Nothing Then
        Set KeyComparer = GenericValueComparer.Default
    Else
        Set KeyComparer = myclass.KeyComparer
    End If
End Property

Public Property Get Default(Optional ByVal KeyComparer As IGenericComparer) As IGenericComparer
    
    If KeyComparer Is Nothing = False Then _
        Set myclass.KeyComparer = KeyComparer
        
    Set Default = New GenericPairComparer
    
End Property

Private Sub Class_Initialize()
    Set myclass.KeyComparer = GenericPairComparer.KeyComparer
End Sub

Public Function Compare(ByVal x As GenericPair, ByVal y As GenericPair) As ComparerEnum: Compare = myclass.KeyComparer.Compare(x.Key, y.Key): End Function
Private Function IGenericComparer_Compare(ByVal a As IGeneric, ByVal b As IGeneric) As ComparerEnum: IGenericComparer_Compare = Me.Compare(a, b): End Function




