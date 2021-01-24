VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Base"
Option Compare Database

Public Enum SortOrder
    Descending = -1
    Ascending = 1
End Enum

Public Enum CompareResult
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

'IGenericCompare Default comparer
Public Function GetHashCode(ByVal Obj As IGeneric) As Long
    If Not Obj Is Nothing Then _
        GetHashCode = Obj.HashValue
End Function
Public Function Compare(ByVal a As IGeneric, ByVal B As IGeneric) As CompareResult
    
    If a Is Nothing Then _
        IGenericComparer_Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.IsLower): Exit Function
    
    If B Is Nothing Then _
        IGenericComparer_Compare = CompareResult.IsGreater: Exit Function
        
    Dim x As IGenericValue, y As IGenericValue
    'try cast
    Set x = a: Set y = B
    Compare = x.CompareTo(y): Exit Function

End Function

Public Function Equals(ByVal a As IGeneric, ByVal B As IGeneric) As Boolean
    
    If a Is Nothing = False Then _
        IGenericComparer_Equals = a.Equals(B): Exit Function
    
    If B Is Nothing = False Then _
        IGenericComparer_Equals = False: Exit Function
    
    Equals = True
    
End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericComparer Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

