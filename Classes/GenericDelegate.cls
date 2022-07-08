VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericDelegate
'*
'* Purpose: A delegate is a type that represents references to methods with a particular parameter list and return type.
'*          When you instantiate a delegate, you can associate its instance with any method with a compatible signature and return type.
'*          You can invoke (or call) the method through the delegate instance.
'*
'*          This ability to refer to a method as a parameter makes delegates ideal for defining callback methods.
'*          You can write a method that compares two objects in your application.
'*          That method can be used in a delegate for a sort algorithm. Because the comparison code is separate from the library, the sort method can be more general.
'*
'* Notes: Delegates are similar to pointers to functions, in C or C++. A delegate is a reference type variable that holds the reference to a method. The reference can be changed at runtime!
'*        Delegates allow methods to be passed as parameters.
'*        Delegates encapsulate both an object instance and a method.
'*
'*
'*
'*======================================================================================================================

Implements IGenericPredicate
Implements IGenericPredicator
Implements IGenericHashProvider

Public Event Match(ByVal Element As Object, ByRef Result As Boolean)
Public Event IsTrue(ByVal A As Object, ByVal B As Object, ByRef Result As Boolean)

Public Event HashCode(ByVal Element As Object, ByRef HashCode As Long)
Public Event Equality(ByVal A As Object, ByVal B As Object, ByRef Result As Boolean)

Public Function Build() As GenericDelegate: Set Build = New GenericDelegate: End Function

'IGenericPredicate
Private Function IGenericPredicate_Match(ByVal Element As Object) As Boolean: RaiseEvent Match(Element, IGenericPredicate_Match): End Function
'IGenericPredicator
Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean: RaiseEvent IsTrue(A, B, IGenericPredicator_IsTrue): End Function
'IGenericHashProvider
Private Function IGenericHashProvider_Equality(ByVal A As Object, ByVal B As Object) As Boolean: RaiseEvent Equality(A, B, IGenericHashProvider_Equality): End Function
Private Function IGenericHashProvider_HashOf(ByVal Element As Object) As Long: RaiseEvent HashCode(Element, IGenericHashProvider_HashOf): End Function

