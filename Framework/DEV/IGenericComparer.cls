VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericComparer"
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
'* Class:  IGenericComparer
'*
'* Purpose: The IGenericComparer interface supports equality comparisons. Provides a base class for implementations of the IGenericComparer interface.
'*          Customization of comparisons for sorting and ordering is provided as well.
'*          Implement GetHashCode method to provide customized hash codes for objects, corresponding to the customized
'*          equality comparison provided by the Equals method.
'*          The Equals(Object, Object) method is reflexive, symmetric, and transitive.
'*          That is, it returns true if used to compare an object with itself; true for two objects x and y if it is true for y and x; and true for two objects x and z if it is true for x and y and also true for y and z.
'*
'*          Compare(X,Y) is used with the List.Sort and List.BinarySearch methods.
'*          It provides a way to customize the sort order of a collection.
'*
'* Notes:   Used as default class if no comparer is specified
'*          GetHashCode should return CLng(0) for Nothing
'*          Compare(A,B)=IsEqual != Equals(A,B)=True
'*          Implementations are required to ensure that if the Equals(Object, Object) method returns true for two objects x and y,
'*          then the value returned by the GetHashCode(Object) method for x must equal the value returned for y.
'*          If your class overrides the Equals method and does value equality then you should override GetHashCode too!
'*
'*          If you have an array of types (such as string or integer) that already support IGenericComparer,
'*          you can sort that array without providing any explicit reference to IComparer.
'*          In that case, the elements of the array are cast to the default implementation of IGenericComparer
'*
'*          Comparing null with any reference type is allowed and does not generate an exeption.
'*          A null reference is considered to be less than any reference that is not null.
'*
'*
'*
'*
'*======================================================================================================================

Implements IGeneric

Public Enum SortOrder
    descending = -1
    Ascending = 1
End Enum

Public Enum CompareResult
    islower = -1
    IsEqual = 0
    isgreater = 1
End Enum

'IGenericCompare Default comparer
Public Function GetHashCode(ByVal Obj As IGeneric) As Long
    If (Obj Is Nothing) = False Then _
        GetHashCode = Obj.HashValue
End Function

Public Function Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult

    If A Is B Then _
        Compare = CompareResult.IsEqual: Exit Function
        
    If A Is Nothing Then _
        Compare = CompareResult.islower: Exit Function
'Works too!
'    If A Is Nothing Then _
'        Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.islower): Exit Function
    
'    If B Is Nothing Then _
'        Compare = CompareResult.isgreater: Exit Function

    Dim Value As GenericValue
    'try cast
    Set Value = A
    Compare = Value.CompareTo(B)

End Function

Public Function Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean
    
    If A Is B Then _
        Equals = True: Exit Function
    
    If A Is Nothing Then _
        Equals = False: Exit Function 'Or B is nothing
   
    Equals = A.Equals(B)
    
End Function

Private Sub Class_Initialize()
    If Not Me Is IGenericComparer Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'IGeneric for comparer
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = IGenericComparer: End Property
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean: IGeneric_Equals = Other Is Me: End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is IGenericComparer: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = TypeOf Other Is IGenericComparer: End Property
Private Function IGeneric_ToString() As String: IGeneric_ToString = "Class: " & TypeName$(Me) & " :: Default": End Function
Private Property Get IGeneric_VirtualPointer() As Long: End Property
Private Property Get IGeneric_VirtualSize() As Long:: End Property
