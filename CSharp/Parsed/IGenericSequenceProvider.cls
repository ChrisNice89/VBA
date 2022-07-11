VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericSequenceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Interface
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, FunctionReturnValueDiscarded, ImplementedInterfaceMember
'@Folder "Sequences.Abstract"
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

Public Function Count() As Long: Count = 0: End Function
Public Function First() As Object: Set First = Nothing: End Function
Public Function Last() As Object: Set Last = Nothing: End Function
Public Function GetAt(ByVal Index As Long) As Object: Set GetAt = Nothing: End Function
Public Function GuessCount() As Long: GuessCount = 0: End Function
Public Function ToArray() As GenericArray: Set ToArray = GenericArray: End Function
Public Function ToList() As GenericOrderedList: Set ToList = GenericOrderedList.Build: End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = Nothing: End Property
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: IGenericIterator_HasNext = False: End Function
'IGenericSequence
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me: End Function
