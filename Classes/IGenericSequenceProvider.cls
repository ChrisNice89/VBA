VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericSequenceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ImplementedInterfaceMember
Option Explicit
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericSequenceProvider
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


Implements IGenericSequence
Implements IGenericIterator

Public Function Count() As Long: End Function
Public Function First() As Object: End Function
Public Function Last() As Object: End Function
Public Function GetAt(ByVal Index As Long) As Object: End Function
Public Function GuessCount() As Long: End Function
Public Function ToArray() As GenericArray: Set ToArray = GenericArray: End Function
Public Function ToList() As GenericOrderedList: Set ToList = GenericOrderedList.Build:  End Function

'IGenericIterator
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGenericIterator_Current() As Object: End Property
'@Ignore NonReturningFunction, EmptyMethod
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: End Function
'IGenericSequence
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me: End Function
