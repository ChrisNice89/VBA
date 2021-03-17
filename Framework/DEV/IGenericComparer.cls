VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericComparer"
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
'* Class:  IGenericComparer
'*
'* Purpose: GenericComparer & GenericEquality
'*          Used as default class if no comparer is specified
'*
'*
'* Notes:   GetHashCode should return CLng(0) for Nothing
'*
'*
'*
'*
'*
'*======================================================================================================================

'@Folder "Base"
Option Explicit

Implements IGeneric

Public Enum SortOrder
    descending = -1
    ascending = 1
End Enum

Public Enum CompareResult
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

'IGenericCompare Default comparer
Public Function GetHashCode(ByVal Obj As IGeneric) As Long
    If (Obj Is Nothing) = False Then _
        GetHashCode = Obj.HashValue
End Function

Public Function Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult
    
    If A Is Nothing Then _
        Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.IsLower): Exit Function

    If B Is Nothing Then _
        Compare = CompareResult.IsGreater: Exit Function
        
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
