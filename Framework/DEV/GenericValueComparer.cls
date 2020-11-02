VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericValueComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "<T>Class"
Option Explicit
Implements IGenericComparer

Public Property Get Default() As IGenericComparer: Set Default = New GenericValueComparer: End Property

Public Function Compare(ByVal x As IGenericValue, ByVal y As IGenericValue) As ComparerEnum
    
    If x Is Nothing = False Then _
        Compare = x.CompareTo(y): Exit Function
    
    If y Is Nothing = False Then _
        Compare = -1: Exit Function
    
    Compare = 0: Exit Function
     
End Function

Private Function IGenericComparer_Compare(ByVal a As IGeneric, ByVal b As IGeneric) As ComparerEnum: IGenericComparer_Compare = Me.Compare(a, b): End Function

