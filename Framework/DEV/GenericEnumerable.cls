VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId


Implements IGenericIterator

Public Event Where(ByRef Condition As Boolean, ByVal Current As IGeneric)

Private Enum IteratorType
    DefaultIterator
    WhereIterator
    
End Enum

Private Type Member
    Sequence As IGenericIterator
    Iterator As IteratorType
    Delegate As Boolean
    Current As IGeneric
End Type
Private myclass As Member

Public Property Get Sequence() As IGenericIterator
    Set Sequence = myclass.Sequence
End Property

Public Property Get IteratorType() As Integer
    IteratorType = myclass.Iterator
End Property

Public Function ForEach(ByVal Sequence As IGenericIterator, Optional ByVal Where As Boolean = True) As GenericEnumerable
    
    If Where Then _
        myclass.Iterator = WhereIterator
    
    Set myclass.Sequence = Sequence
    Set ForEach = New GenericEnumerable
    
End Function

Public Function ToList() As GenericList

    Dim List As GenericList
    Set List = GenericList.Build

    Call List.AddAll(myclass.Sequence)
    Set ToList = List
    
End Function

Private Sub Class_Initialize()
    
    Set myclass.Sequence = GenericEnumerable.Sequence
    myclass.Iterator = GenericEnumerable.IteratorType
    
End Sub

Public Function Count() As Long
    
    Dim Element As IGeneric
    Do While myclass.Sequence.HasNext(Element): Count = Count + 1: Loop
    
End Function
'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim Result As Boolean
    
    Do
        If myclass.Sequence.HasNext(myclass.Current) = False Then _
            Exit Function
            
'        Select Case myclass.Iterator
'            Case IteratorType.DefaultIterator
'                Result = True
'            Case IteratorType.WhereIterator
'                RaiseEvent Where(Result, myclass.Current)
'
'        End Select
'
    Loop While Result = False
    
    Set Out = myclass.Current
    IGenericIterator_HasNext = Result
     
End Function

Private Sub IGenericIterator_Reset()
    Set myclass.Current = Nothing
End Sub



