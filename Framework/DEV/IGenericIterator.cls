VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericIterator
'*
'* Purpose: An iterator is a method in this Code which is used in an array or in collections like the list, etc.
'*          to retrieve elements one by one. Or in other words, we can say that an iterator is used to perform
'*          an iteration over the collections. It uses the Do-Loop statement to return the element from the collection
'*          at a time and it always remembers the current location of the iterator, so when the next iteration
'*          takes place it will return the next element of the given collection.
'*          The return type of the HasNext method is IGeneric.
'*
'*
'*======================================================================================================================

'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Public Property Get Current() As IGeneric: End Property
Public Function HasNext(ByRef Out As IGeneric) As Boolean: End Function
Attribute HasNext.VB_UserMemId = 0
Public Sub Reset(): End Sub


Private Sub Class_Initialize()
    Call System.Throw(Me, "Private Sub Class_Initialize").Instantiation
End Sub

