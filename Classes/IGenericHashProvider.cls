VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericHashProvider"
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
'* Class:  IGenericHashProvider
'*
'* Purpose: The IGenericHashProvider interface supports equality comparisons.
'*          Implement HashOf method to provide customized hash codes for objects, corresponding to the customized equality comparison provided by the Equality method.
'*          The Equality(Object, Object) method is reflexive, symmetric, and transitive.
'*          That is, it returns true if used to compare an object with itself; true for two objects x and y if it is true for y and x; and true for two objects x and z if it is true for x and y and also true for y and z.
'*
'*          If an object that is used as a key in a hash table does not provide a useful implementation of HashCode, you can specify a hash code provider by supplying an IGenericHashProvider implementation

'* Notes:   Provides a base comparison class for implementations for Objects including nulls.
'*          HashOf should return CLng(0) for Nothing
'*
'*          Implementations are required to ensure that if the Equality(A, B) method returns true for two objects A and B, then the value returned by the HashOf(A) method for A must equal the value returned for B.
'*          If your class overrides the Equals method and does value equality then you should override HashOf too!
'*
'*          Supplies a hash code for an object, using a custom hash function.
'*          Comparing null with any reference type is allowed and does not generate an exeption.
'*          This interface supports only equality comparisons. Customization of comparisons for sorting and ordering is provided by the IComparer interface.
'*
'*
'*======================================================================================================================

Implements IGenericPredicator

'A hash code is a numeric value that is used to insert and identify an object in a hash-based collection such as the HashMap or HashSet class. The HashOf method provides this hash code for algorithms that need quick checks of object equality.
Public Function HashOf(ByVal Element As Object) As Long: HashOf = ObjPtr(Element): End Function
'Two objects that are equal return hash codes that are equal. However, the reverse is not true: equal hash codes do not imply object equality, because different (unequal) objects can have identical hash codes.
Public Function Equality(ByVal A As Object, ByVal B As Object) As Boolean: Equality = A Is B: End Function

Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean: IGenericPredicator_IsTrue = Me.Equality(A, B): End Function
