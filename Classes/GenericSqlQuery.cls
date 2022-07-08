VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSqlQuery"
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
'* Class:  GenericSqlQuery
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

Public Enum ISqlQueryOptions
    Prepared = 2 ^ 0
    StoredProc = 2 ^ 1
    CommandText = 2 ^ 2 'Default
    NamedParameter = 2 ^ 3
    FullTable = 2 ^ 4
End Enum

Private Type Member
    Name As gString
    Desc As gString
    Sql As gString
    Parameters As GenericOrderedList
    SqlManager As GenericSqlManager
    Options As ISqlQueryOptions
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Options() As ISqlQueryOptions: Options = This.Options: End Property

Public Property Get Name() As gString

    If (This.Name Is Nothing) Then _
        Set This.Name = gString
        
    Set Name = This.Name

End Property

Public Property Get Sql() As gString

    If (This.Sql Is Nothing) Then _
        Set This.Sql = gString
        
    Set Sql = This.Sql

End Property

Public Property Get SqlManager() As GenericSqlManager: Set SqlManager = This.SqlManager: End Property
Public Property Get HasParameter() As Boolean: HasParameter = Me.Count > 0: End Property

Public Property Get Count() As Long
    
    If Not (This.Parameters Is Nothing) Then _
        Count = This.Parameters.Elements.Count
    
End Property

Public Property Get IsReady() As Boolean

    If (This.Sql Is Nothing) Then _
        Exit Property
    
    If (This.Sql.IsEmptyOrWhiteSpace) Then _
        Exit Property
        
    IsReady = True
    
End Property

Public Function Build( _
                        ByVal Sql As gString, _
                        Optional ByVal SqlManager As GenericSqlManager, _
                        Optional ByVal Name As gString, _
                        Optional ByVal Options As ISqlQueryOptions = ISqlQueryOptions.CommandText) As GenericSqlQuery
  
    Dim newQuery As Member

    If (Sql Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Build").IsNothing("Sql")
    
    If (Sql.IsEmptyOrWhiteSpace) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("Sql.IsEmptyOrWhiteSpace")
    
    If (Name Is Nothing) Then _
        Set Name = gString
            
    newQuery.Options = Options
    Set newQuery.Name = Name
    Set newQuery.Sql = Sql.Trim
    Set newQuery.SqlManager = SqlManager
    Set newQuery.Parameters = GenericOrderedList.Build
  
    Set Build = New GenericSqlQuery
    Call System.Inject(Build, VarPtr(newQuery), LenB(newQuery))
    
End Function

Public Sub ReplaceAt(ByVal Index As Long, ByVal Value As Variant)
    
    Dim Pair As GenericPair
    Dim SqlValue As IGenericValue
    
    Set Pair = This.Parameters.ElementAt(Index)
    Set SqlValue = Pair.Value
    Set Pair = GenericPair.Build(Pair.Key, SqlValue.Of(Value))
    Set This.Parameters.ElementAt(Index) = Pair
    
End Sub

Public Sub Clear()
   
    If Not (This.Parameters Is Nothing) Then _
        Call This.Parameters.Elements.Clear
        
End Sub
 
 Public Property Get GetAt(ByVal Index As Long) As GenericPair: Set GetAt = This.Parameters.ElementAt(Index): End Property

Public Property Get ParameterAt(ByVal Index As Long) As IGenericValue
    Set ParameterAt = Me.GetAt(Index).Value
End Property

Public Property Set ParameterAt(ByVal Index As Long, ByVal Value As IGenericValue)
    Set This.Parameters.ElementAt(Index) = GenericPair.Build(Me.GetAt(Index).Key, Value)
End Property

Public Function RemoveAt(ByVal Index As Long) As GenericSqlQuery
        
    Call This.Parameters.RemoveAt(Index)
    Set RemoveAt = Me
    
End Function

Public Function RemoveLast() As GenericSqlQuery
        
    Call This.Parameters.Pop
    Set RemoveLast = Me
    
End Function

Public Function IndexOf(ByVal Name As IGenericValue) As Long
    
    For IndexOf = This.Parameters.First To This.Parameters.Last
        If Name.Equals(Me.GetAt(IndexOf).Key) Then _
            Exit Function
    Next
    
    IndexOf = -1

End Function

Public Function Add(ByVal Value As IGenericValue, Optional ByVal Name As IGenericValue) As GenericSqlQuery
    
    If Name Is Nothing Then _
        Set Name = gString.Of("?")
    Call This.Parameters.Add(GenericPair.Build(Name, Value))
    Set Add = Me
    
End Function

Public Function AddAll(ByVal Values As IGenericSequence) As GenericSqlQuery
    
    '@Ignore VariableNotAssigned
    Dim Value As IGenericValue, Name As gString
  
    Dim Collection As IGenericCollection, Count As Long
    
    If (Values Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function AddValues").IsNothing("Values")
    
    Set AddAll = Me
    
    If TypeOf Values Is IGenericCollection Then
        Set Collection = Values
        Count = Collection.Count
        
        If (Count = 0) Then _
            Exit Function
     
        Call This.Parameters.EnsureCapacity(Count)
    
    End If
    
    With Values.Iterator
        If Not .HasNext(Value) Then _
            Exit Function
        
        Set Name = gString.Of("?")
        
        Do: Call This.Parameters.Add(GenericPair.Build(Name, Value)): Loop While .HasNext(Value)
    End With
    
End Function

Public Function Parameters() As IGenericIterator: Set Parameters = This.Parameters.Range: End Function

Public Function Fetch(Optional ByVal ASYNC As Boolean = True) As GenericSqlStream
  
    If (This.SqlManager Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Fetch").IsNothing("SqlManager")
    
    If Not (This.SqlManager.HasCredentials) Then _
        Call GenericError.Build(Me, "Public Function Fetch").InvalidOperation("Not (This.SqlManager.HasCredentials)")
        
    Set Fetch = This.SqlManager.Fetch(Me, ASYNC)
    
End Function

Public Function Execute(Optional ByVal ASYNC As Boolean = True) As Long
    
    If (This.SqlManager Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Execute").IsNothing("SqlManager")
    
    If Not (This.SqlManager.HasCredentials) Then _
        Call GenericError.Build(Me, "Public Function Execute").InvalidOperation("Not (This.SqlManager.HasCredentials)")
        
    Execute = This.SqlManager.Execute(Me, ASYNC)

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSqlQuery: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericSqlQuery: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSqlQuery):  End Property

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function
Private Function IGeneric_ToString() As String

        IGeneric_ToString = _
                        "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                        "Sql: " & Me.Sql.ToString & vbNewLine & _
                        "Parameters: " & Me.Count
    
End Function

