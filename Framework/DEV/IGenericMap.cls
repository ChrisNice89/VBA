VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericMap
'*
'* Purpose:
'*
'*
'*
'* Notes:
'*
'*
'*
'*
'*
'*======================================================================================================================

Option Explicit

Public Property Get Elements() As IGenericCollection:  End Property
Public Property Get Comparer() As IGenericComparer: End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Item As IGeneric): End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0

Public Function GetKeys() As GenericArray: End Function
Public Function GetValues() As GenericArray: End Function
Public Function Add(ByVal Key As IGeneric, ByVal Value As IGeneric, Optional ByVal Overwrite As Boolean) As Boolean: End Function
Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean: End Function
Public Function Remove(ByVal Key As IGeneric) As Boolean: End Function
Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long: End Function
Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: End Function
Public Sub AddAll(ByVal Map As IGenericMap, Optional ByVal Overwrite As Boolean): End Sub

Private Sub Class_Initialize()
    If Not Me Is IGenericMap Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub
