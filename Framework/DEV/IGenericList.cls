VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Elements() As IGenericCollection: End Property

Public Function GetAt(ByVal Index As Long) As IGeneric: End Function
Public Function Add(ByVal Element As IGeneric) As Boolean: End Function
Public Sub AddAll(ByVal Collection As IGenericCollection): End Sub

Public Function Pop() As IGeneric: End Function
Public Function IndexOf(ByVal Element As IGeneric) As Long: End Function
Public Function Contains() As Boolean: End Function

Public Function Remove(ByVal Element As IGeneric) As Boolean: End Function
Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long: End Function
Public Sub RemoveAt(ByVal Index As Long): End Sub

Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = 0): End Sub

'Private Sub Class_Initialize()
'    If Not Me Is IGenericList Then _
'        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub

