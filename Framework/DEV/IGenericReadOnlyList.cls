VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericReadOnlyList"
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
'* Class:  IGenericReadOnlyList
'*
'* Purpose: The IGenericReadOnlyList class provides an easy way to produce a read-only list from any type
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

'@Folder("<T>Enumerable")
Option Explicit

Public Property Get Elements() As IGenericCollection: End Property
Public Property Get Last() As Long: End Property
Public Property Get First() As Long: End Property

Public Function GetAt(ByVal Index As Long) As IGeneric: End Function
Attribute GetAt.VB_UserMemId = 0
Public Function IndexOf(ByVal Item As IGeneric) As Long: End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericReadOnlyList Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub
