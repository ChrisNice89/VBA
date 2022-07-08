VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericMapper"
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
'* Class:  GenericMapper
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

Public Event Mapper(ByRef Element As Object)

Public Function Build() As GenericMapper
    Set Build = New GenericMapper
End Function

Public Function Map(ByVal Element As Object) As Object: Set Map = Element: RaiseEvent Mapper(Map): End Function ' to avoid null error. should return at least identity

