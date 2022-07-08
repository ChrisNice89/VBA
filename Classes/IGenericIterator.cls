VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule ImplementedInterfaceMember
Option Explicit

'*===========================================================================================
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
'*          The return type of the HasNext method is Boolean.
'*
'* Notes:   Read only access
'*          Provides the ability to iterate through a sequence by exposing a Current property.
'*          Only a move forward cursor
'*
'*
'*
'*======================================================================================================================

Public Property Get Current() As Object: End Property
Public Function HasNext(Optional ByRef Out As Object) As Boolean: End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericIterator Then _
        Call GenericError.Build(Me, "Private Sub Class_Initialize").Instantiation
End Sub


