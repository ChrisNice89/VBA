VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGeneric"
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
'* Class:  IGeneric
'*
'* Purpose: Base Class of all instances in this Framework
'*
'*
'*
'*        Equals - Supports equality comparisons between objects.
'*        HashValue - Generates a number corresponding to the value of the object to support the use in a hash-based collection such as the HashSet/ Map class
'*        ToString - Manufactures a human-readable text string that describes an instance of the class.
'*        IsMutable - Mutable types are those whose data members can be changed after the instance is created but Immutable types are those whose data members can not be changed after the instance is created
'*
'*
'*
'*
'*======================================================================================================================

Public Property Get ClassName() As String: ClassName = VBA.TypeName$(Me): End Property
Public Property Get IsDefault() As Boolean: IsDefault = True: End Property
Public Property Get GetType() As IGeneric: Set GetType = Me: End Property
Public Property Get VirtualPointer() As LongPtr: End Property 'obsole
Public Property Get VirtualSize() As Long: End Property 'obsole
Public Property Get IsMutable() As Boolean: IsMutable = False: End Property
Public Property Get HashCode() As Long: End Property
Public Property Get IsRelatedTo(ByVal Other As Object) As Boolean: IsRelatedTo = True: End Property

Public Function Equals(ByVal Other As Object) As Boolean: Equals = Other Is Me: End Function
'@DefaultMember
Public Function ToString() As String: ToString = "Class: " & VBA.TypeName$(Me) & " :: Default": End Function
Attribute ToString.VB_UserMemId = 0

Private Sub Class_Initialize()
    If Not Me Is IGeneric Then _
        Call GenericError.Build(Me, "Private Sub Class_Initialize()").Instantiation
End Sub
