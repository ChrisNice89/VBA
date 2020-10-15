VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Database")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IPrototype

Public Event Result(ByVal Data As ADODB.Recordset)

Private Type RecordsetProperties
    Setting As ISqlRecodsetType
    CursorType As ISqlCursorType
    CursorLocation As ISqlCursorLocation
    LockType As ISqlLockType
End Type

Private Type Member
    Connection As ADODB.Connection
    Command As ADODB.Command
    ParameterList As GenericOrderedMap
    Parameter As Long
    Sql As String
    SqlOptions As ISqlQueryType
    IsPrepared As Boolean
    IsAsync As Boolean
    HasResult As Boolean
    CloseConection As Boolean
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Sql() As String
    Sql = myclass.Sql
End Property

Public Property Get ID() As Long
    ID = ObjPtr(myclass.Command)
End Property

Public Function Build( _
                        ByVal Connection As ADODB.Connection, _
                        ByVal Sql As String, _
                        ByVal Parameter As Long, _
                        Optional ByVal Options As ISqlQueryType = ISqlQueryType.Asynch + ISqlQueryType.ReturnRecords + ISqlQueryType.NamedParameter) As GenericSql

    If Sql = vbNullString Then
        'error
    End If
    
    If Connection Is Nothing Then
    'error
    End If
    
    If Parameter > 0 Then
    'error
    End If

    With myclass
        Set .Connection = Connection
        .Sql = Sql
        .SqlOptions = Options
        .Parameter = Parameter
    End With
    
    Set Build = Skynet.CreateInstance(New GenericSql, Virtual.Pointer, Virtual.Size)
 
End Function

Public Sub Prepare()
    
    Dim P As GenericPair
    
    If myclass.Parameter > 0 Then
        If myclass.ParameterList Is Nothing Then
            'error
        End If
        If Not myclass.ParameterList.Count = myclass.Parameter Then
            'error
        End If
    End If
    
    myclass.Sql = VBA.Trim$(myclass.Sql)
    Set myclass.Command = New ADODB.Command
    
    With myclass.Command
        If (myclass.SqlOptions And ISqlQueryType.NamedParameter) = ISqlQueryType.NamedParameter Then _
            .NamedParameters = True
        
        If (myclass.SqlOptions And ISqlQueryType.Prepared) = ISqlQueryType.Prepared Then _
            .Prepared = True
            
        If (myclass.SqlOptions And ISqlQueryType.StoredProc) = ISqlQueryType.StoredProc Then
            .CommandText = myclass.Sql
            .CommandType = ISqlCommandType.CmdStoredProc
        Else
            .CommandText = IIf(VBA.Right$(myclass.Sql, 1) = ";", myclass.Sql, myclass.Sql & ";")
            .CommandType = ISqlCommandType.CmdText
        End If
               
        Set .ActiveConnection = myclass.Connection
    
    End With
    
    Dim Pair As GenericPair
    Dim Value As ISqlValue
    
    If Not myclass.ParameterList Is Nothing Then
        With myclass.ParameterList.Iterator(Pairs_)
            Do While .HasNext(Pair)
                Set Value = Pair.Value
                Call Value.AppendTo(myclass.Command)
            Loop
        End With
    End If
    
    With myclass
        .HasResult = (myclass.SqlOptions And ISqlQueryType.ReturnRecords) = ISqlQueryType.ReturnRecords
        .IsAsync = (myclass.SqlOptions And ISqlQueryType.Asynch) = ISqlQueryType.Asynch
        .IsPrepared = True
    End With
    
End Sub

Public Function Execute(Optional ByVal Asynchronous As Boolean = False) As Boolean
    
    Dim Affected As Long
    Dim Rs As ADODB.Recordset
    
    If Not myclass.IsPrepared Then _
        Call Me.Prepare
    
    If myclass.HasResult Then
        'Call Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset).Open(myclass.Command, , , , IIf(Asynchronous, 16, 0))
        Set Rs = Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset)
        Call Rs.Open(myclass.Command, , , , IIf(Asynchronous, 16, 0))
    Else
        Call myclass.Command.Execute(Affected, , 128 + IIf(Asynchronous, 16, 0))
    End If
       
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function Add( _
                      ByVal Value As ISqlValue, _
                      Optional ByVal Name As String = vbNullString, _
                      Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As GenericSql

    If myclass.Parameter = 0 Then
        'Error
    End If
    
    If myclass.ParameterList Is Nothing Then _
        Set myclass.ParameterList = GenericOrderedMap.Build(myclass.Parameter)
    
    If myclass.ParameterList.Count > myclass.Parameter Then
        'error
    End If
    
    If Name = vbNullString Then _
        Name = "P" & myclass.ParameterList.Count + 1
    
    Call myclass.ParameterList.Add(GString(Name), Value.Prepare(Name, Direction))

    Set Add = Me
    
End Function

Public Function Update(ByVal Index As Variant, ByVal NewValue As ISqlValue) As GenericSql
    
    Dim Parameter As ADODB.Parameter
    
    If myclass.Parameter = 0 Then
        'Error
    End If
    
    With myclass.ParameterList
        If VarType(Index) = vbString Then _
            Index = .IndexOfKey(GString(Index))
        
        If Not .ElementAt(Index).IsRelatedTo(NewValue) Then
            'Error false type
        End If
        
        Set .ElementAt(Index) = NewValue
    End With
    
    If myclass.IsPrepared Then
        Set Parameter = myclass.Command.Parameters(Index - 1)
        With NewValue
'            If Not .IsNumeric Then _
'                Parameter.Size = Len(.Primitive)
'
'            Parameter.Value = .Primitive
        End With
    End If
    
    Set Update = Me
    
End Function

Public Function CreateParameter( _
                                    ByVal Name As String, _
                                    ByVal Value As ISqlValue, _
                                    Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As ADODB.Parameter

    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Name = Name
        .Attributes = adParamNullable '64
        
'        If Value.IsNumeric Then
'            .Precision = 4
'            .NumericScale = 4
'        Else
'            .Size = Len(Value.Primitive)
'        End If
        
        .Type = Value.SqlType
        .Value = Value.Primitive
        
    End With
    
    Set CreateParameter = P

End Function

Public Function CreateGenericParameter(ByVal Name As GString, ByVal Value As ISqlValue)
    Set CreateGenericParameter = Me.CreateParameter(Name.Value, Value, Value.Direction)
End Function

Public Function CreateRecordset(Optional ByVal Setting As ISqlRecodsetType = ISqlRecodsetType.Default_Recordset) As ADODB.Recordset

    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    With Rs
        Select Case Setting
            Case ISqlRecodsetType.ReadOnly_Recordset
                .CursorLocation = ISqlCursorLocation.UseServer
                .CursorType = ISqlCursorType.OpenForwardOnly
                .LockType = ISqlLockType.LockReadOnly
    
            Case ISqlRecodsetType.UpdateBatch_Recordset
                .CursorLocation = ISqlCursorLocation.UseClient
                .CursorType = ISqlCursorType.OpenForwardOnly
                .LockType = ISqlLockType.LockBatchOptimistic
    
            Case ISqlRecodsetType.Update_Recordset
                .CursorLocation = ISqlCursorLocation.UseClient 'CursorLocation.adUseServer
                .CursorType = ISqlCursorType.OpenForwardOnly 'CursorType.adOpenKeyset
                .LockType = ISqlLockType.LockOptimistic
    
            Case ISqlRecodsetType.Disconnected_Recordset
                .CursorLocation = ISqlCursorLocation.UseClient  'Setting the cursor location to client side is important to get a disconnected recordset.
                .CursorType = ISqlCursorType.OpenForwardOnly
                .LockType = ISqlLockType.LockBatchOptimistic
    
            Case ISqlRecodsetType.Default_Recordset
                .CursorLocation = ISqlCursorLocation.UseServer 'CursorLocation.adUseClient  'Client-side cursor to enable AbsolutePosition property
                .CursorType = ISqlCursorType.OpenStatic
                .LockType = ISqlLockType.LockBatchOptimistic
    
            Case Else
    
                Call Err.Raise(1#, TypeName(Me), "CreateRecordset - Type not specified")
        End Select
    End With

    Set CreateRecordset = Rs
   
End Function
'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean
    IGeneric_IsMutable = False
End Property

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then _
        Exit Function
        
    If TypeOf Other Is GenericSql Then _
        IGeneric_IsRelatedTo = True
        
End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = TypeName$(Me) & " :: {" & myclass.Sql & "}"
End Function

Private Function IGeneric_Clone() As IGeneric

End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

End Function
 
Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum

End Function

Private Sub IGeneric_Dispose()

End Sub
Private Function EmptyStructure() As Member: End Function
Private Property Get IGeneric_HashValue() As Long
    IGeneric_HashValue = ObjPtr(Me)
End Property
'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericSql: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericSql: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


