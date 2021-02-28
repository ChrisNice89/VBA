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
Option Explicit

Public Enum SortOrder
    descending = -1
    ascending = 1
End Enum

Public Enum CompareResult
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

'IGenericCompare Default comparer
Public Function GetHashCode(ByVal Obj As IGeneric) As Long: GetHashCode = Obj.HashValue: End Function
Public Function Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult
    
    If A Is Nothing Then _
        Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.IsLower): Exit Function

    If B Is Nothing Then _
        Compare = CompareResult.IsGreater: Exit Function
        
    Dim x As IGenericValue, y As IGenericValue
    'try cast
    Set x = A: Set y = B
    Compare = x.CompareTo(y): Exit Function

End Function
'Dont change equality compare! See IGenericCollection -> Contains
Public Function Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean
    
    If A Is Nothing Then _
        Equals = B Is Nothing: Exit Function
    
    If A Is B Then _
        Equals = True: Exit Function
        
    If B Is Nothing Then _
        Equals = False: Exit Function
    
    Equals = A.Equals(B)
    
End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericComparer Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

