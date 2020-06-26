VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"
Public Property Get Current() As Variant: End Property
Public Function MoveNext() As Boolean: End Function
Public Sub Reset(): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize").Instantiation.Raise
'End Sub

