VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule ImplementedInterfaceMember
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericComparer
'*
'* Purpose: The IGenericComparer interface supports equality comparisons.
'*          The Equals(Object, Object) method is reflexive, symmetric, and transitive.
'*          That is, it returns true if used to compare an object with itself; true for two objects x and y if it is true for y and x; and true for two objects x and z if it is true for x and y and also true for y and z.
'*
'*          The AscendingOrdered method compares two objects and returns a value indicating whether one is strict less than than the other.
'*          AscendingOrdered(X,Y) is used with the List.Sort and List.BinarySearch methods.
'*          It provides a way to customize the sort order of a collection.
'*          Implement this class to provide a custom implementation of the IGenericComparer interface for use with collection classes such as the SortedList/ SortedSet classes.
'*
'* Notes:   Provides a base comparison class for implementations for Objects including nulls.
'*          Comparing null with any reference type is allowed and does not generate an exeption.
'*          AscendingOrdered must be a strict weak ordering over the elements. That is, for any elements a, b, and c, the following conditions must hold:
'*          AscendingOrdered(a, a) is always false. (Irreflexivity)
'*          If AscendingOrdered(a, b) and AscendingOrdered(b, c) are both true, then AscendingOrdered(a, c) is also true. (Transitive comparability)
'*          Two elements are incomparable if neither is ordered before the other according to the predicate. If a and b are incomparable, and b and c are incomparable, then a and c are also incomparable. (Transitive incomparability)
'*
'*
'*======================================================================================================================

Implements IGeneric

'returns true if its first argument should be ordered before its second argument; otherwise, false.
Public Function AscendingOrdered(ByVal A As Object, ByVal B As Object) As Boolean
      
    If A Is Nothing Then _
        Exit Function

    If B Is Nothing Then _
        AscendingOrdered = True: Exit Function
        
    AscendingOrdered = ObjPtr(A) < ObjPtr(B)

End Function

Public Function Equality(ByVal A As Object, ByVal B As Object) As Boolean: Equality = A Is B: End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericComparer Then _
        Call GenericError.Build(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'IGeneric for comparer
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property

Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = Me: End Property
Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = True: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = TypeOf Other Is IGenericComparer: End Property
Private Function IGeneric_ToString() As String: IGeneric_ToString = "Class: " & VBA.TypeName$(Me) & " :: Default": End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGeneric_VirtualPointer() As Long: End Property
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGeneric_VirtualSize() As Long: End Property


