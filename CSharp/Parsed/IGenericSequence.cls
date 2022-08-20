VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
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
'* Class:  IGenericSequence
'*
'* Purpose: Exposes an Iterator instance, which supports a simple iteration over a sequence.
'*          IGenericSequence contains a single method 'Iterator', which returns an IGenericIterator instance.
'*          The returned IGenericIterator provides the ability to iterate through the sequence by exposing a 'HasNext(Out Element)' method.
'*
'*
'*
'* Notes: Determines whether two sequences are equal.
'*        The Equals method enumerates the two source sequences in parallel and compares corresponding elements by using the default equality method
'*
'*
'*
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericHashProvider
Implements IGenericPredicator
Implements IGenericIterator

Private Const MaxValue As Long = 2 ^ 31 - 1

Public Function Iterator() As IGenericIterator: Set Iterator = Me: End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = Nothing: End Property
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: IGenericIterator_HasNext = False: End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericSequence Then _
        Call GenericError.Build(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'Sequence equality check
Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean
    
    Dim IteratorA As IGenericIterator, IteratorB As IGenericIterator
    Dim x As IGenericValue, y As IGenericValue
    
    If A Is B Then _
        IGenericPredicator_IsTrue = True: Exit Function
        
    If A Is Nothing Then _
        Exit Function
        
    If B Is Nothing Then _
        Exit Function
    
    Set IteratorA = CastSequence(A).Iterator
    Set IteratorB = CastSequence(B).Iterator
    
    With IteratorA
        Do While .HasNext(x)
        
            If Not IteratorB.HasNext(y) Then _
                Exit Function
        
            If Not (x Is y) Then
            
                If (x Is Nothing) Then _
                    Exit Function

                If Not x.Equals(y) Then _
                    Exit Function
                
            End If
            
        Loop
    End With
    
    IGenericPredicator_IsTrue = Not IteratorB.HasNext(y)
    
End Function

Private Function IGenericHashProvider_Equality(ByVal A As Object, ByVal B As Object) As Boolean: IGenericHashProvider_Equality = IGenericPredicator_IsTrue(A, B): End Function

'Does deep hashcode
Private Function IGenericHashProvider_HashOf(ByVal obj As Object) As Long
    
    Dim h As Double
    Dim Element As IGenericValue
    
    If (obj Is Nothing) Then _
        Exit Function
    
    With CastSequence(obj).Iterator
        Do While .HasNext(Element)
            If Not obj Is Nothing Then
                h = 31 * h + Element.HashCode
                h = CLng(h - (VBA.Fix(h / MaxValue) * MaxValue))
            End If
        Loop
    End With

    IGenericHashProvider_HashOf = h
    
End Function

Private Function CastSequence(ByVal s As IGenericSequence) As IGenericSequence: Set CastSequence = s: End Function

'IGeneric for IGenericSequence
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = Me: End Property
Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = 0: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean

    If TypeOf Other Is IGenericSequence Then _
        IGeneric_IsRelatedTo = True: Exit Property
        
End Property

Private Function IGeneric_ToString() As String: IGeneric_ToString = "Class: " & VBA.TypeName$(Me) & " :: Default": End Function

Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = 0: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = 0: End Property


