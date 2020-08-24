VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Public Property Get Item(ByVal Key As IGeneric) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): End Property
Public Property Get Elements() As IGenericCollection: End Property
Public Property Get Keys() As IGenericCollection: End Property
Public Property Get Values() As IGenericCollection: End Property

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: End Function
'Public Function ContainsValue(ByVal Value As IGeneric) As Boolean: End Function
Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean: End Function
Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric): End Sub
Public Sub Remove(ByVal Key As IGeneric): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub

