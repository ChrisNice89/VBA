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
    Set ElementComparer = myclass.ElementComparer
End Property

Public Function Build(ByVal SortByIndex As Long, Optional ByVal ElementComparer As IGenericComparer) As IGenericComparer
    
    myclass.SelectedIndex = SortByIndex
    Set myclass.ElementComparer = ElementComparer
    Set Build = New GenericRowComparer
    
End Function

Private Sub Class_Initialize()
    myclass.SelectedIndex = GenericRowComparer.SelectedIndex
    Set myclass.ElementComparer = GenericRowComparer.ElementComparer
    If myclass.ElementComparer Is Nothing = False Then _
        Set myclass.ElementComparer = IGenericValue.Comparer
End Sub

Public Function Compare(ByVal x As IGenericReadOnlyList, ByVal y As IGenericReadOnlyList) As CompareResult: Compare = myclass.ElementComparer.Compare(x.ElementAt(myclass.SelectedIndex), y.ElementAt(myclass.SelectedIndex)): End Function
Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult: IGenericComparer_Compare = Me.Compare(A, B): End Function

Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean

End Function

Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long

End Function
