VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericRowComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")
Option Explicit
Implements IGenericComparer

Private Type Member
    ElementComparer As IGenericComparer
    SelectedIndex As Long
End Type
Private myclass As Member

Public Property Get SelectedIndex() As Long
    SelectedIndex = myclass.SelectedIndex
End Property
Public Property Get ElementComparer() As IGenericComparer
    If myclass.ElementComparer Is Nothing Then
        Set ElementComparer = GenericValueComparer.Default
    Else
        Set ElementComparer = myclass.ElementComparer
    End If
End Property

Public Property Get Default(ByVal SortByIndex As Long, Optional ByVal ElementComparer As IGenericComparer) As IGenericComparer
    
    If ElementComparer Is Nothing = False Then _
        Set myclass.ElementComparer = ElementComparer
    
    myclass.SelectedIndex = SortByIndex
    Set Default = New GenericRowComparer
    
End Property

Private Sub Class_Initialize()
    myclass.SelectedIndex = GenericRowComparer.SelectedIndex
    Set myclass.ElementComparer = GenericRowComparer.ElementComparer
End Sub

Public Function Compare(ByVal x As IGenericReadOnlyList, ByVal y As IGenericReadOnlyList) As ComparerEnum: Compare = myclass.ElementComparer.Compare(x.ElementAt(myclass.SelectedIndex), y.ElementAt(myclass.SelectedIndex)): End Function
Private Function IGenericComparer_Compare(ByVal a As IGeneric, ByVal b As IGeneric) As ComparerEnum: IGenericComparer_Compare = Me.Compare(a, b): End Function






