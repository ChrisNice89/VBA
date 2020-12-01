VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Public Enum DictionaryData
    KeyData = 0
    ValueData
    PairData
End Enum

Public Property Get Count() As Long: End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): End Property

Public Function ContainsValue(ByVal Value As IGeneric) As Boolean: End Function
Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: End Function
Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean: End Function
Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric): End Sub
Public Sub Remove(ByVal Key As IGeneric): End Sub
Public Sub Clear(): End Sub

Public Sub CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): End Sub
Public Function IteratorOf(ByVal t As DictionaryData) As IGenericIterator: End Function
    

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub

