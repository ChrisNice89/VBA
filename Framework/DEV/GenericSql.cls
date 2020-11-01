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
Implements IGenericIterator

Public Event Result(ByVal Data As ADODB.Recordset)

Private Const ASYNC_OPTION As Long = 16&
Private Const EXECUTE_ASYNC As Boolean = True
Private Const CLOSE_CONNECTION_IMMEDIATELY  As Boolean = True

Private Type Iterator
    Records As ADODB.Recordset
    Converter() As IGenericValue
    Current As IGeneric
    HasFetchComplete As Boolean
    HasInitialized As Boolean
End Type

Private Type Member
    Credentials As SqlCredentials
    Command As ADODB.Command
    Result As ADODB.Recordset
    ParameterList As GenericOrderedMap
    Parameter As Long
    Sql As String
    SqlOptions As ISqlQueryType
    IsPrepared As Boolean
    AffectedRecords As Long
    IsInExecution As Boolean
    Iterator As Iterator
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Private WithEvents Database As ADODB.Connection
Attribute Database.VB_VarHelpID = -1

Public Property Get Sql() As String
    Sql = myclass.Sql
End Property

Public Property Get ID() As Long
    ID = ObjPtr(myclass.Command)
End Property

Public Property Get IsPrepared() As Boolean
    IsPrepared = myclass.IsPrepared
End Property

Public Property Get Result() As ADODB.Recordset

    If Me.IsReady Then _
        Set Result = myclass.Result

End Property

Public Property Get HasResult() As Boolean
    If Me.IsReady Then _
        HasResult = Not myclass.Result Is Nothing
End Property
Public Property Get IsReady() As Boolean
    IsReady = Not IsFetching(myclass.Command)
End Property

Public Function Build( _
                        ByVal Credentials As SqlCredentials, _
                        Optional ByVal Sql As String, _
                        Optional ByVal Parameter As Long, _
                        Optional ByVal Options As ISqlQueryType = ISqlQueryType.ReturnRecords + ISqlQueryType.NamedParameter + ISqlQueryType.Prepared) As GenericSql

    If Sql = vbNullString Then
        'error
    End If
    
    If Credentials Is Nothing Then
    'error
    End If
    
    If Parameter < 0 Then
    'error
    End If

    With myclass
        Set .Credentials = Credentials
        .Sql = Sql
        .SqlOptions = Options
        .Parameter = Parameter
    End With
    
    Set Build = Skynet.CreateInstance(New GenericSql, Virtual.Pointer, Virtual.Size)
 
End Function

Public Sub Prepare(ByVal Sql As String)
    
    Dim P As GenericPair
    
    If Disconnect(myclass.Command) Then
        Set myclass.Command = New ADODB.Command
    Else
        Debug.Print "Error Public Sub Prepare()"
    End If

    myclass.Sql = VBA.Trim$(Sql)
    
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
    End With
    
    Dim Pair As GenericPair
    Dim Value As IGenericValue
    
    If myclass.Parameter > 0 Then
        If myclass.ParameterList Is Nothing Then
            'error
        End If
        
        With myclass.ParameterList.Iterator(Pairs_)
            Do While .HasNext(Pair)
                Set Value = Pair.Value
                Call myclass.Command.Parameters.Append(Value.ToSqlParameter)
            Loop
        End With
    End If
    
    myclass.IsPrepared = True
    
End Sub

Public Function ExecuteSchema(ByVal Schema As ISqlDBSchema, ParamArray Args() As Variant) As ADODB.Recordset
    
    Dim rs As ADODB.Recordset
    
    With CreateConnection(myclass.Credentials)
        Call .Open
        
        If IsMissing(Args) Then
            Set rs = .OpenSchema(Schema)
        Else
            Set rs = .OpenSchema(Schema, Args)
        End If
        
        Call .Close
        
    End With
    
    RaiseEvent Result(rs)
    
    Set ExecuteSchema = rs
    
End Function

Public Sub Execute()

    Set Database = CreateConnection(myclass.Credentials, True)
    
    If Not myclass.IsPrepared Then _
        Call Me.Prepare(myclass.Sql)
    
    Do While Database.State And ISqlConnectionState.Connection_IsConnecting
    Loop
    
    Set myclass.Command.ActiveConnection = Database
    myclass.IsInExecution = True

    On Error Resume Next
        If (myclass.SqlOptions And ISqlQueryType.ReturnRecords) = ISqlQueryType.ReturnRecords Then
            Call Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset).Open(myclass.Command, , , , IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
        Else
            Call myclass.Command.Execute(myclass.AffectedRecords, , 128 + IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
        End If
    On Error GoTo 0
    
    
End Sub

Public Function ExecuteMapper() As IGenericIterator

    Dim rs As ADODB.Recordset
    
    If Not myclass.IsPrepared Then _
        Call Me.Prepare(myclass.Sql)
    
    If Connect(myclass.Command, myclass.Credentials) Then
        Set Database = myclass.Command.ActiveConnection
    Else
        Debug.Print "error If Not Connect(myclass.Command, myclass.Credentials, Database)"
    End If
    
    Set myclass.Result = Nothing
    
    Set rs = Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset)
    myclass.IsInExecution = True
    Call rs.Open(myclass.Command, , , , IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
    
    Set ExecuteMapper = Me.RowMapper(rs)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function Add( _
                      ByVal Value As IGenericValue, _
                      Optional ByVal Name As String = vbNullString, _
                      Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As GenericSql

    If myclass.IsPrepared Then
        'error
    End If
    
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
    
'    Call myclass.ParameterList.Add(GString(Name), Value.Prepare(Name, Direction))

    Set Add = Me
    
End Function

Private Function Indexes(ByVal Table As String) As ADODB.Recordset
    Set Indexes = Me.ExecuteSchema(SchemaIndexes, Empty, Empty, Empty, Empty, Table)
End Function

Private Function TableDetails(ByVal Table As String) As ADODB.Recordset
    Set TableDetails = Me.ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty)
End Function

Private Function TableExists(ByVal Table As String) As Boolean
    TableExists = Me.ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty).RecordCount > 0
End Function

Private Function ColumnExists(ByVal Table As String, ByVal Field As String) As Boolean
    ColumnExists = Me.ExecuteSchema(SchemaColumns, Empty, Empty, Table, Field).RecordCount > 0
End Function

Private Function Tables() As ADODB.Recordset
    Set Tables = Me.ExecuteSchema(SchemaTables, Empty, Empty, Empty, "TABLE")
End Function

Private Function TableStructure(ByVal Table As String) As ADODB.Recordset
    Set TableStructure = Me.ExecuteSchema(SchemaColumns, Empty, Empty, Table, Empty) '.ToArray("COLUMN_NAME"))
End Function

Public Function Update(ByVal Index As Variant, ByVal NewValue As IGenericValue) As GenericSql
    
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
                                    ByVal Value As IGenericValue, _
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
        .Value = Value.RawValue
        
    End With
    
    Set CreateParameter = P

End Function

Public Function CreateRecordset(Optional ByVal Setting As ISqlRecodsetType = ISqlRecodsetType.Default_Recordset) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    With rs
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

    Set CreateRecordset = rs
   
End Function

Private Function ImportQuery(ByVal Source As ADODB.Connection, ByVal Sql As String, ByVal Table As String) As String
    
    Dim vCommandParts As Variant
    Dim External As Variant

    vCommandParts = Split(Sql, " FROM ", 2, vbTextCompare)
    External = Split(vCommandParts(1), " WHERE ", 2)
    External(0) = ExternalLink(Source, External(0))
    vCommandParts(1) = Join$(External, " ")
    
'    If Adapter.TableExists(Table) Then
'        vCommandParts(0) = "INSERT INTO " & Table & Space$(1) & vCommandParts(0)
'    Else
'        vCommandParts(0) = vCommandParts(0) & " INTO " & Table
'    End If
    
    ImportQuery = Join$(vCommandParts, " FROM ")
    
End Function

Private Function ExportQuery(ByVal Source As ADODB.Connection, ByVal Sql As String, ByVal Table As String) As String
    
    Dim vCommandParts As Variant
    Dim External As Variant
     
    vCommandParts = Split(Sql, " FROM ", 2, vbTextCompare)
     
'    If Source.CreateAdapter.TableExists(Table) Then
'        vCommandParts(0) = "INSERT INTO " & ExternalLink(Source, Table) & Space$(1) & vCommandParts(0)
'    Else
'        vCommandParts(0) = vCommandParts(0) & " INTO " & Source.ExternalLink(Table)
'    End If
    
    ExportQuery = Join$(vCommandParts, " FROM ")
    
End Function

Private Function ExternalLink(ByVal Connection As ADODB.Connection, ByVal Table As String) As String
    With Connection
        ExternalLink = "[" & .Properties("Extended Properties") & ";DATABASE=" & .Properties("Data Source") & "].[" & VBA.Trim$(Table) & "]"
    End With
End Function

'internal stuff
Private Function CreateConnection(ByVal Credentials As SqlCredentials, Optional ByVal OpenConnection As Boolean = False) As ADODB.Connection
        
    If Credentials Is Nothing Then
    'erro
    End If
    
    Dim oConnection As ADODB.Connection
    Set oConnection = New ADODB.Connection
    
    With oConnection
        .CursorLocation = ISqlConnectionCursorLocation.Connection_UseServer
        .CommandTimeout = 0
        .Mode = ISqlConnectionMode.ReadWrite_ConnectionMode
        
        Select Case Credentials.ProviderTyp
            Case ISqlConnectionProvider.Excel_Provider
                Select Case VBA.Mid$(Credentials.DataSource, InStrRev(Credentials.DataSource, ".") + 1)
                    Case "xlsx"  'Excel 2007-2010 Workbook (.xlsx)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
                    
                    Case "xlsm" 'Excel 2007-2010 Workbook (.xlsm)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Macro;HDR=YES"
                    
                    Case "xls" 'Excel 97-2003 Workbook (.xls)'nur für 32-Bit OfficeVersionen
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Properties("Extended Properties") = "Excel 8.0;HDR=YES"
                    
                    Case Else
                        Call Err.Raise(1, TypeName$(Me), "Not defined File:" & Credentials.DataSource)
                        '.ConnectionString = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBQ & ";HDR=Yes';"
                End Select
            
            Case ISqlConnectionProvider.CSV_File_Provider
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Properties("Extended Properties") = "Text;HDR=Yes;FMT=Delimited();"
            
            Case ISqlConnectionProvider.Sql_Provider
                .Provider = "SQLOLEDB.1"
            
            Case ISqlConnectionProvider.Access97_2003_Provider
                .Provider = "Microsoft.Jet.OLEDB.4.0"
            
            Case ISqlConnectionProvider.Access2007_2013_Provider
                .Provider = "Microsoft.ACE.OLEDB.12.0"
            
            Case Else
                Call Err.Raise(1, TypeName$(Me), "Not defined File:" & Credentials.DataSource)
            
        End Select
        
        'Diese Property findet man eigentlich bei fast allen DB-Treibern wieder. Somit hat sie auch dementsprechend verschiedene Hintergründe. Bei Access beschreibt der DataSource den Pfad zu der physischen MDB-Datei. Beim MS-SQL-Server beschreibt der DataSource den Namen des Servers (nicht die Datenbank!). Bei Informix z.B. beschreibt diese Eigenschaft wieder den Server und den Namen der Datenbank.
        .Properties("Data Source") = Credentials.DataSource
        
        If Credentials.SSPI Then
            .Properties("Integrated Security") = "SSPI"
        Else
            If LenB(Credentials.UserID) > 0 Then _
                .Properties("User id") = Credentials.UserID
            
            If LenB(Credentials.UserPassword) > 0 Then
                .Properties("Password") = Credentials.UserPassword
                'nur notwendig, wenn Sie "User ID" und "Password" setzen
                .Properties("Encrypt Password") = True
            End If
        End If
        'Diese Eigenschaft gibt an, ob nach dem Aufbau der Connection zur Datenbank die Userinformationen aus dem ConnectionString und den Connection-Properties gelöscht werden sollen
        'oder gehalten werden sollen. Die Default-Einstellung ist unverständlicher Weise True (behalten). Dieser Wert sollte also explizit auf False gesetzt werden.
        .Properties("Persist Security Info") = False
        .Properties("Prompt") = Credentials.Prompt
        
        'Datenbankpasswort einer Accessdatenbank. Dieses darf nicht mit dem Userpasswort aus den Logininformationen einer MDW-Sicherheitsdatei verwechselt werden.
        If LenB(Credentials.DatabasePassword) > 0 Then _
            .Properties("Jet OLEDB:Database Password") = Credentials.DatabasePassword
        
        If LenB(Credentials.IntitialCatalog) > 0 Then _
            .Properties("Initial Catalog") = Credentials.IntitialCatalog
        
        '.Properties("Mode") = Credentials.Mode
        '.Properties("Jet OLEDB:System database") =  'Pfad zu der System-MDW (Sicherheitdatei).
        '.Properties("Location") =
        
    End With
    
    If OpenConnection Then _
        Call oConnection.Open(Options:=IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
  
   Set CreateConnection = oConnection
    
End Function

Private Function Connect(ByVal Command As ADODB.Command, ByVal Credentials As SqlCredentials) As Boolean
   
    If Command Is Nothing Then
        'Error
    End If
    
    With Command
        Set .ActiveConnection = CreateConnection(Credentials, OpenConnection:=True)
        Connect = IsConnected(.ActiveConnection)
    End With
    
End Function

Private Function IsConnected(ByVal Connection As ADODB.Connection) As Boolean
    Debug.Print Connection.State
    If Not Connection Is Nothing Then _
        IsConnected = (Connection.State And ISqlConnectionState.Connection_IsOpen) = ISqlConnectionState.Connection_IsOpen

End Function

Private Function IsFetching(ByVal Command As ADODB.Command) As Boolean
    
    Const BREAK As Long = 5000000
    Dim Counter As Long
    
    If Not Command Is Nothing Then
        If Not Command.ActiveConnection Is Nothing Then
            Do While (Command.State And ISqlCommandState.StateExecuting) = ISqlCommandState.StateExecuting 'Do While Connection.State = ISqlConnectionState.Connection_IsOpen + ISqlConnectionState.Connection_IsExecuting
'                If Counter Mod BREAK = 0 Then
'                    Debug.Print ("Command is executing ...")
'                    DoEvents
'                    Counter = 0
'                End If
'                Counter = Counter + 1
            Loop
        End If
'        Do While (Command.State And ISqlCommandState.StateExecuting) = ISqlCommandState.StateExecuting
'        Loop
         DoEvents '!!!!!
    End If
    
    IsFetching = False

End Function

'Private Function IsReady(ByVal Connection As ADODB.Connection) As Boolean
'
'    If Not Connection Is Nothing Then
'        Do While Connection.State = ISqlConnectionState.Connection_IsOpen + ISqlConnectionState.Connection_IsExecuting
''            Debug.Print ("Command is executing2 ...")
'        Loop
'    End If
'
'    IsReady = IsConnected(Connection)
'
'End Function

Private Function Disconnect(ByVal Command As ADODB.Command) As Boolean
    
    If Not Command Is Nothing Then
        If Not IsFetching(Command) Then
            Call Command.ActiveConnection.Close
            Set Command.ActiveConnection = Nothing
        End If
    End If
    
    Disconnect = True
    
End Function

Private Sub Database_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Database_ConnectComplete"
End Sub

Private Sub Database_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Set Database = Nothing
    Debug.Print "Private Sub Database_Disconnect"
End Sub

'Private Sub Database_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
'    Debug.Print "Private Sub Database_WillExecute"
'End Sub

Private Sub Database_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)

    Debug.Print "Database_ExecuteComplete"
 
    If pError Is Nothing Then
        If Not pRecordset Is Nothing Then
            With myclass
                Set .Result = pRecordset
                Set .Result.ActiveConnection = Nothing
                .AffectedRecords = .Result.RecordCount
            End With
        End If
        
        If Not myclass.Result Is Nothing Then _
            RaiseEvent Result(myclass.Result)
            
    Else
        Debug.Print "Sql Error"
        Debug.Print pError.Description
    End If
    
    If CLOSE_CONNECTION_IMMEDIATELY Then
        Call pCommand.ActiveConnection.Close
        Set pCommand.ActiveConnection = Nothing
    End If
    
    myclass.IsInExecution = False
       
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericSql: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSql):  End Property

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

'IGenericIterator
Public Function RowMapper(ByVal Records As ADODB.Recordset) As IGenericIterator

    Dim Field As ADODB.Field
    Dim i As Long
  
    If (Records Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function RowMapper").InvalidInput("(Records Is Nothing)")
    
    If (Records.State = ISqlCommandState.StateClosed) Then _
        Call Skynet.Throw(Me, "Public Function RowMapper").InvalidInput("(Records.State = ISqlCommandState.StateClosed)")
    
    With myclass.Iterator
        Set .Records = Records
        .HasFetchComplete = .Records.BOF And .Records.EOF
        
        Do While (.Records.State And ISqlCommandState.StateExecuting) = ISqlCommandState.StateExecuting
            Debug.Print "Records fetching...."
        Loop

        ReDim .Converter(.Records.Fields.Count - 1)
        ReDim Columns(.Records.Fields.Count - 1) As IGeneric
        
        For i = 0 To .Records.Fields.Count - 1
            Set Columns(i) = GString(.Records.Fields(i).Name)
            Select Case .Records.Fields(i).Type
                Case Is = ISqlDataType.Boolean_Type: Set .Converter(i) = GString '("Yes/No")
                Case Is = ISqlDataType.Binary_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.Numeric_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.Single_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.Currency_Type: Set .Converter(i) = GString '("Currency")
                Case Is = ISqlDataType.Double_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.Decimal_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.SmallInt_Type: Set .Converter(i) = GString '("#")
                Case Is = ISqlDataType.Integer_Type: Set .Converter(i) = GString '("#")
                Case Is = ISqlDataType.UnsignedTinyInt_Type: Set .Converter(i) = GString '("#")
                Case Is = ISqlDataType.BigInt_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.Date_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.VarWChar_Type: Set .Converter(i) = GString
                Case Is = ISqlDataType.LongVarWChar_Type: Set .Converter(i) = GString '203 Gibt eine lange Null-terminierte Unicode-Zeichenfolgenwert an.
                Case Else
                    Set .Converter(i) = GString
            End Select
        Next
        
        Set .Current = GenericArray.BuildFrom(Columns)
        
    End With
    
    Set RowMapper = Skynet.CreateInstance(New GenericSql, Virtual.Pointer, Virtual.Size)
    
End Function

Private Property Get IGenericIterator_Current() As IGeneric
    Set IGenericIterator_Current = myclass.Iterator.Current
End Property

Private Function IGenericIterator_HasNext(Out As IGeneric) As Boolean
    
    Dim i As Long
    
    With myclass.Iterator
       .HasFetchComplete = .Records.EOF
       
        If .HasFetchComplete Then
            'Close clear ect all records fetched
            Call .Records.Close
            Set .Records = Nothing
            Exit Function
        End If
        
        ReDim Elements(.Records.Fields.Count - 1) As IGeneric
        
        For i = 0 To .Records.Fields.Count - 1
            Set Elements(i) = .Converter(i).Create(.Records.Fields(i).Value)
        Next
        
        Set .Current = GenericArray.BuildFrom(Elements)
        Set Out = .Current
        
        .Records.MoveNext
        
        IGenericIterator_HasNext = Not .HasFetchComplete
        
    End With
    
End Function

Private Sub IGenericIterator_Reset()

End Sub
