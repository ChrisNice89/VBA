VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, SuspiciousPredeclaredInstanceAccess
'@Folder "Base"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericPair
'*
'* Purpose:
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

Implements IGeneric

Private Type Member
    Key As Object
    Value As Object
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Key() As Object: Set Key = This.Key: End Property
Public Property Get Value() As Object: Set Value = This.Value: End Property

Public Function Build(ByVal Key As Object, ByVal Value As Object) As GenericPair
    
    '@Ignore FunctionReturnValueDiscarded
    If (Me Is GenericPair) = False Then _
        Call GenericError.Build(Me, "Public Function Build").BuildExeption
'
    Set This.Key = Key: Set This.Value = Value
    Set Build = New GenericPair: Set This.Key = Nothing: Set This.Value = Nothing
    
End Function

Private Sub Class_Initialize(): Set This.Key = GenericPair.Key: Set This.Value = GenericPair.Value: End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericPair: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericPair: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long
    
    Const MAXLONG As Long = 2 ^ 31 - 1
    Static Hash As Long
    
    If (Hash = 0) Then
        
        Dim h As Double
        Dim Casted As IGenericValue
        
        If Not This.Key Is Nothing Then
            If TypeOf This.Key Is IGeneric Then
                Set Casted = This.Key
                h = CDbl(Casted.HashCode) * 31
            Else
                h = ObjPtr(This.Key)
            End If
        End If
        
        If Not This.Value Is Nothing Then
            If TypeOf This.Value Is IGeneric Then
                Set Casted = This.Value
                h = h + CDbl(Casted.HashCode)
            Else
                h = ObjPtr(This.Value)
            End If
        End If
            
        Hash = CLng(h - (Fix(h / MAXLONG) * MAXLONG))
    
    End If
    
    IGeneric_HashCode = Hash
        
End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean
    
    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        Exit Property
        
    Set otherPair = Other

    If This.Key.IsRelatedTo(otherPair.Key) Then _
        IGeneric_IsRelatedTo = This.Value.IsRelatedTo(otherPair.Value)
    
End Property

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean

    Dim otherPair As GenericPair
    
    If (Other Is Nothing) Then _
        Exit Function
    
    If Not TypeOf Other Is GenericPair Then _
        Exit Function
        
    Set otherPair = Other
    
    If This.Key Is Nothing = False Then
        If This.Key.Equals(otherPair.Key) = False Then _
            Exit Function
    ElseIf otherPair.Key Is Nothing = False Then _
            Exit Function
    End If
    
    If This.Value Is Nothing = False Then
        If This.Value.Equals(otherPair.Value) = False Then _
            Exit Function
    ElseIf otherPair.Value Is Nothing = False Then _
            Exit Function
    End If
    
    IGeneric_Equals = True
    
End Function

Private Function IGeneric_ToString() As String
        
        Dim Casted As IGenericValue
        
        IGeneric_ToString = "["
        If Not This.Key Is Nothing Then
            If TypeOf This.Key Is IGeneric Then
                Set Casted = This.Key
                IGeneric_ToString = IGeneric_ToString & Casted.ToString
            Else
                IGeneric_ToString = IGeneric_ToString & VBA.TypeName$(This.Key)
            End If
        Else
            IGeneric_ToString = IGeneric_ToString & "NULL"
        End If
        
        IGeneric_ToString = IGeneric_ToString & ";"
        
        If Not This.Value Is Nothing Then
            If TypeOf This.Value Is IGeneric Then
                Set Casted = This.Value
                IGeneric_ToString = IGeneric_ToString & Casted.ToString
            Else
                IGeneric_ToString = IGeneric_ToString & VBA.TypeName$(This.Key)
            End If
        Else
            IGeneric_ToString = IGeneric_ToString & "NULL"
        End If
        
        IGeneric_ToString = IGeneric_ToString & "]"
                                               
End Function


