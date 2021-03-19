VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericList"
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
'* Class:  IGenericList
'*
'* Purpose: IGenericList, represents an ordered sequence of objects. The elements contained in a IGenericList can be accessed, iterated and removed according to the order in which they appear internally in the IGenericList.
'*          The ordering of the elements is why this data structure is called a IGenericList.
'*
'*
'*
'* Notes:   The user can access elements by their integer index (position in the list), and search for elements in the list.
'*          Some list implementations have restrictions on the elements that they may contain. For example, some implementations prohibit null elements,each element can occur only once
'*          and some have restrictions on the order of their elements.
'*
'*
'*
'*
'*======================================================================================================================

Option Explicit

Public Property Get Elements() As IGenericCollection: End Property

Public Function GetAt(ByVal Index As Long) As IGeneric: End Function
Attribute GetAt.VB_UserMemId = 0
Public Function Add(ByVal Element As IGeneric) As Boolean: End Function
Public Function Pop() As IGeneric: End Function
Public Function IndexOf(ByVal Element As IGeneric) As Long: End Function
Public Function AddAll(ByVal Collection As IGenericCollection) As IGenericList: End Function

Public Function Remove(ByVal Element As IGeneric) As Boolean: End Function
Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long: End Function
Public Function RetainAll(ByVal Collection As IGenericCollection) As Long: End Function

Public Sub RemoveAt(ByVal Index As Long): End Sub
Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = 0): End Sub

Private Sub Class_Initialize()
    If Not Me Is IGenericList Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub



