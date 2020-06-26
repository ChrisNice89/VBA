VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "Base"
Public Enum ComparerEnum
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

Public Property Get HashValue() As Long: End Property
Public Function IsRelatedTo(ByRef Other As IObject) As Boolean: End Function
Public Function Equals(ByRef Other As IObject) As Boolean: End Function
Public Function CompareTo(ByRef Other As IObject) As ComparerEnum: End Function
Public Function ToString() As String: End Function
Public Function Clone() As IObject: End Function
Public Sub Dispose(): End Sub


'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub
