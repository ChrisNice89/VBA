VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Interface
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, FunctionReturnValueDiscarded, ImplementedInterfaceMember
'@Folder "Collections.Abstract"
Option Explicit

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
'*        Provides a custom implementation of the IGenericComparer interface for use with collection classes such as the Map class, or with methods such as List.Sort.
'*
'*
'*
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericHashProvider
Implements IGenericIterator

Private Const MaxValue As Long = 2 ^ 31 - 1

Public Property Get HasRandomAccess() As Boolean: End Property
Public Property Get Count() As Long: End Property
Public Function GetAt(ByVal Index As Long) As Object: End Function
Public Function Contains(ByVal Element As Object) As Boolean: End Function

Public Function Iterator() As IGenericIterator: Set Iterator = Me: End Function
Public Function ToArray() As GenericArray: Set ToArray = GenericArray: End Function
Public Function Copy() As IGenericCollection: Set Copy = Me: End Function
Public Sub CopyTo(ByVal Target As GenericArray, ByVal Index As Long): End Sub
Public Sub Clear(): End Sub

Private Sub Class_Initialize()
    If Not Me Is IGenericCollection Then _
        Call GenericError.Build(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'IGenericIterator
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGenericIterator_Current() As Object: End Property
'@Ignore NonReturningFunction, EmptyMethod
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: End Function

'IGenericHashProvider
Private Function IGenericHashProvider_HashOf(ByVal obj As Object) As Long
   
    Dim Collection As IGenericCollection
    Dim h As Double, HashCode As Long
    
    Dim Element As Object
    Dim Casted As IGenericValue
    
    If (obj Is Nothing) Then _
        Exit Function
    
    Set Collection = obj 'Try cast
    With Collection.Iterator
        Do While .HasNext(Element)
            If Not Element Is Nothing Then
                If TypeOf Element Is IGenericValue Then
                    Set Casted = Element
                    HashCode = Casted.HashCode
                Else
                    HashCode = ObjPtr(Element)
                End If
                h = 31 * h + HashCode
                IGenericHashProvider_HashOf = CLng(h - (VBA.Fix(h / MaxValue) * MaxValue))
            End If
        Loop
    End With
    
End Function

'Does deep equal
Private Function IGenericHashProvider_Equality(ByVal A As Object, ByVal B As Object) As Boolean
    
    Dim First As IGenericCollection, Second As IGenericCollection
    Dim Element As Object
    
    If A Is Nothing Then _
        IGenericHashProvider_Equality = VBA.IIf(B Is Nothing, True, False): Exit Function
    
    If B Is Nothing Then _
        Exit Function
    'Try Cast
    Set First = A: Set Second = B
    
    If Not (First.Count = Second.Count) Then _
        Exit Function
    
    With First.Iterator
        Do While .HasNext(Element)
            If Not Second.Contains(Element) Then _
                Exit Function
        Loop
    End With
    
    IGenericHashProvider_Equality = True

End Function

'IGeneric
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGeneric_VirtualPointer() As Long: End Property
'@Ignore NonReturningFunction, EmptyMethod
Private Property Get IGeneric_VirtualSize() As Long:: End Property
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = Me: End Property
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = True: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean

    If TypeOf Other Is IGenericCollection Then _
        IGeneric_IsRelatedTo = True: Exit Property
    
End Property
Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = "Class: " & VBA.TypeName$(Me) & " :: Default": End Function


