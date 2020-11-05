VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "Base"
Option Compare Database

Public Enum SortOrder
    descending = -1
    Random = 0
    ascending = 1
End Enum

Public Enum ComparerEnum
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

Public Function Compare(ByVal a As IGeneric, ByVal b As IGeneric) As ComparerEnum: End Function

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub

