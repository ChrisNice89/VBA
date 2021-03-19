VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGeneric"
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
'* Class:  IGeneric
'*
'* Purpose: Base Class of all instances in this Framework
'*
'*
'*
'* Notes:
'*
'*
'*
'*
'*
'*======================================================================================================================

'@Exposed
Option Explicit
'@Folder "Base"

Public Property Get ClassName() As String: ClassName = TypeName$(Me): End Property
Public Property Get IsDefault() As Boolean: IsDefault = True: End Property
Public Property Get Default() As IGeneric: Set Default = IGeneric: End Property
Public Property Get VirtualPointer() As LongPtr: End Property
Public Property Get VirtualSize() As Long: End Property
Public Property Get IsMutable() As Boolean: IsMutable = False: End Property
Public Property Get HashValue() As Long: End Property
Public Property Get IsRelatedTo(ByVal Other As IGeneric) As Boolean: IsRelatedTo = True: End Property

Public Function Equals(ByVal Other As IGeneric) As Boolean: Equals = Other Is Me: End Function
Public Function ToString() As String: ToString = "Class: " & TypeName$(Me) & " :: Default": End Function
Attribute ToString.VB_UserMemId = 0
Public Function Clone() As IGeneric: Set Clone = IGeneric: End Function

Private Sub Class_Initialize()
    If Not Me Is IGeneric Then _
        Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

