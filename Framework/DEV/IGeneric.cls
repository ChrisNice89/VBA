VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGeneric"
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
Public Property Get IsMutable() As Boolean: End Property
Public Property Get HashValue() As Long: End Property
Public Function IsRelatedTo(ByVal Other As IGeneric) As Boolean: End Function
Public Function Equals(ByVal Other As IGeneric) As Boolean: End Function
Public Function CompareTo(ByVal Other As IGeneric) As ComparerEnum: End Function
Public Function ToString() As String: End Function
Attribute ToString.VB_UserMemId = 0
Public Function Clone() As IGeneric: End Function
Public Sub Dispose(): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub
