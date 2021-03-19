VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericCollection"
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
'* Class:  IGenericCollection
'*
'* Purpose: A IGenericCollection is a set of related objects and provide methods that perform basic operations on the whole collection
'*          It has methods that tell you how many elements are in the collection (Count), methods that check whether a given object is in the collection (Contains),
'*          methods that copy elements from the collection (Copy,CopyTo, ToArray), and methods that provide an iterator over the collection (Iterator).
'*
'*
'* Notes: The IGenericCollection does about what you'd expect given that a Collection represents a group of objects.
'*
'*
'*
'*
'*
'*======================================================================================================================

'@Exposed
'@Folder "<T>Enumerable"

Option Explicit

Implements IGenericComparer

Private Const MAXVALUE As Double = 2 ^ 31

Public Property Get Count() As Long: End Property

Public Function Iterator() As IGenericIterator: End Function
Attribute Iterator.VB_UserMemId = 0
Public Function ToArray() As GenericArray: End Function
Public Function Copy() As IGenericCollection: End Function
Public Function Contains(ByVal Element As IGeneric) As Boolean: End Function
Public Function ContainsAll(ByVal Collection As IGenericCollection) As Boolean: End Function
Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): End Sub
Public Sub Clear(): End Sub

Private Sub Class_Initialize()
    If Not Me Is IGenericCollection Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'IGenericCompare Collection comparer
'Does deep hashcode
Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long

    If (Obj Is Nothing) Then _
        Exit Function
    
    Dim C As IGenericCollection
    Dim H As Double
    
    Set C = Obj
    With C.Iterator
        Do While .HasNext(Obj)
            If Not Obj Is Nothing Then
                H = 31 * H + Obj.HashValue
                If H > MAXVALUE Then _
                     H = CLng(H - (Fix(H / MAXVALUE) * MAXVALUE))
            End If
        Loop
    End With

    IGenericComparer_GetHashCode = H
    
End Function
'TODO
Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult
    
    Dim i As Long, j As Long
    
    If A Is B Then _
        IGenericComparer_Compare = CompareResult.IsEqual: Exit Function
        
    If A Is Nothing Then _
        IGenericComparer_Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.islower): Exit Function
    
    If B Is Nothing Then _
        IGenericComparer_Compare = CompareResult.isgreater: Exit Function
        
    i = A.HashValue
    j = B.HashValue
    
    Select Case True
        Case i > j: IGenericComparer_Compare = isgreater
        Case i < j: IGenericComparer_Compare = islower
        Case Else: IGenericComparer_Compare = IsEqual
    End Select

End Function
'Does deep equal
Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean
    
    Dim First As IGenericCollection, Second As IGenericCollection
    Dim x As IGenericIterator, y As IGenericIterator
    Dim i As IGeneric, j As IGeneric
    
    If A Is B Then _
        IGenericComparer_Equals = True: Exit Function
        
    If A Is Nothing Or B Is Nothing Then _
        IGenericComparer_Equals = False: Exit Function
    
    Set First = A: Set Second = B
    
    If Not First.Count = Second.Count Then _
        Exit Function
    
    Set x = First.Iterator: Set y = Second.Iterator
    
    Do While (x.HasNext(i) And y.HasNext(j))
    
        If Not i Is j Then
            If Not (i Is Nothing) Then
                If Not i.Equals(j) Then _
                    Exit Function
            Else
                If Not (j Is Nothing) Then _
                    Exit Function
            End If
        End If
        
    Loop
    
    IGenericComparer_Equals = True

End Function


