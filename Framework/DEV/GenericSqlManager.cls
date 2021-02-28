VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSqlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Sql"

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericIterator

Public Event result(ByVal Data As ADODB.Recordset)

Private Const ASYNC_OPTION As Long = 16&
Private Const EXECUTE_ASYNC As Boolean = False
Private Const CLOSE_CONNECTION_IMMEDIATELY  As Boolean = True
Private Const DEBUG_MODE  As Boolean = True

Public Enum ISqlCursorLocation
    UseClient = 3 'Bedeutet, daß die Daten komplett auf den Client gezogen werden sollen. Dies bringt jedoch einen eventuell großen Speicherbedarf auf dem Client.
    UseServer = 2 'Die Daten werden nicht komplett auf den Client gezogen, sondern nur diese die vom Client aktuell benötigt werden. Das spart viel Speicherbedarf auf dem Client. Nachteil ist jedoch, daß bei manchen Datenanzeige-Controlls wie z.B. DataGrid oder MSHFlexGrid die Daten mit einem serverseitigen Cursor keine Daten angezeigt werden, bzw. ein paar Funktionen nicht mehr verfügbar sind.
End Enum
'
'Der CommandType gibt an was ausgeführt werden soll. Dafür liegt das CommandTypeEnum zu Grunde.
Public Enum ISqlCommandType
    CmdUnspecified = -1 'Does not specify the command type argument.
    CMDText = 1 'Evaluates CommandText as a textual definition of a command or stored procedure call.
    CmdTable = 2 'Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
    CmdStoredProc = 4 'Evaluates CommandText as a stored procedure name.
    CmdUnknown = 8 'Indicates that the type of command in the CommandText property is not known.
    CmdFile = 256 'Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
    'Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect.
    'This value cannot be combined with the ExecuteOptionEnum value adAsyncExecute.
     CmdTableDirect = 512
End Enum
'
Public Enum ISqlQueryType
    ReturnRecords = 2 ^ 0
    Prepared = 2 ^ 1
    StoredProc = 2 ^ 3
    NamedParameter = 2 ^ 4
End Enum
'
Public Enum ISqlCursorType
    OpenForwardOnly = 0 'Dieser Cursor ist der schnellste. Er verbraucht am wenigsten Speicher und Netzresourcen. Cursor kann jedoch nur vorwärts bewegt werden. Änderungen sind Zeilenweise möglich, jedoch nicht für andere User sichtbar.
    OpenKeyset = 1 'Dieser Cursor erstellt für jeden Datensatz einen Schlüssel. Er ist schneller als der Static- und der Dynamic-Cursor, da er nicht die ganzen Datensätze lädt sondern nur einen Schlüssel auf den Datensatz. Änderungen und löschen sind möglich und werden anderen Usern angezeigt. Einzig anfügen wird anderen Usern nicht angezeigt.
    OpenDynamic = 2 'Dieser Cursor ist der flexibelste, er ermöglicht alle Bewegungen. Alle Änderungen werden allen angezeigt. Er ist aber auch der ressourcenbeanspruchenste.
    OpenStatic = 3 'Ruft eine Kopie der abgefragten Daten ab. Daten können nicht geändert werden. Cursor in alle Richtungen beweglich.
End Enum
'
Public Enum ISqlLockType
    LockReadOnly = 1 'Nur Lesen, Daten können nicht verändert werden (Default)
    LockPessimistic = 2 'Daten werden beim Öffnen des Recordsets für andere User gesperrt.
    LockOptimistic = 3 'Daten werden erst beim Update gesperrt.
    LockBatchOptimistic = 4 'Dieser Modus wird nur für Stapelaktuallisierung benützt.
End Enum

Public Enum ISqlRecodsetType
    ReadOnly_Recordset
    Disconnected_Recordset
    Default_Recordset
    UpdateBatch_Recordset
    Update_Recordset
End Enum

'ModeEnum
Public Enum ISqlConnectionMode
    Unknown_ConnectionMode = 0          '"Unkown or unset permissions." -Keine Beschränkungen [= Default]
    Read_ConnectionMode = 1             '"User cannot read data."
    Write_ConnectionMode = 2            '"User cannot write data."
    ReadWrite_ConnectionMode = 3        '"User cannot read nor write data."
    ShareDenyRead_ConnectionMode = 4    '"Other users cannot read data."
    ShareDenyWrite_ConnectionMode = 8   '"Other users cannot write data."
    ShareExclusive_ConnectionMode = 12  '"Other users cannot read or write data."
    ShareDenyNone_ConnectionMode = 16   '"Other users cannot do anything with data."
End Enum
'
Public Enum ISqlObjectState
    IsClosed_State = 0        'Indicates that the object is closed.
    IsOpen_State = 1          'Indicates that the object is open.
    IsConnecting_State = 2    'Indicates that the object is connecting.
    IsExecuting_State = 4     'Indicates that the object is executing a command.
    IsFetching_State = 8      'Indicates that the rows of the object are being retrieved.
End Enum
'
Public Enum ISqlConnectionCursorLocation
    Connection_UseClient = 3 'Bedeutet, daß die Daten komplett auf den Client gezogen werden sollen. Dies bringt jedoch einen eventuell großen Speicherbedarf auf dem Client.
    Connection_UseServer = 2 'Die Daten werden nicht komplett auf den Client gezogen, sondern nur diese die vom Client aktuell benötigt werden. Das spart viel Speicherbedarf auf dem Client. Nachteil ist jedoch, daß bei manchen Datenanzeige-Controlls wie z.B. DataGrid oder MSHFlexGrid die Daten mit einem serverseitigen Cursor keine Daten angezeigt werden, bzw. ein paar Funktionen nicht mehr verfügbar sind.
End Enum

'ConnectPromptEnum
'Specifies whether a dialog box should be displayed to prompt for missing parameters when opening a connection to a data source.
Public Enum ISqlConnectionPrompt
    PromptAlways = 1 'Das Fenster wird auf jeden Fall gezeigt (auch wenn die Anmeldedaten eigentlich komplett sein könnten).
    PromptComplete = 2  'Das Fenster wird angezeigt wenn nicht alle Daten vorhanden sind. Es sind alle Optionen möglich.
    PromptCompleteRequired = 3  'Das Fenster wird gezeigt wenn nicht alle Daten vorhanden sind. Es können nur User-Id, Passwort und Servername (SQL-Server) angegeben werden.
    PromptNever = 4 'Erlaubt im MultiUser-Bereich das Öffnen anderer Verbindungen ohne Beschränkungen (Lesen/Schreiben/Beides).
End Enum

Public Enum ISqlDBSchema
    SchemaAsserts = 0 ''Returns the assertions defined in the catalog that are owned by a given user.
    '(ASSERTIONS Rowset) CONSTRAINT_CATALOG CONSTRAINT_SCHEMA CONSTRAINT_NAME
    SchemaCatalogs = 1 ' 'Returns the physical attributes associated with catalogs accessible from the DBMS.
    '(CATALOGS Rowset) CATALOG_NAME
    SchemaCharacterSets = 2 'Returns the character sets defined in the catalog that are accessible to a given user.
    '(CHARACTER_SETS Rowset) CHARACTER_SET_CATALOG CHARACTER_SET_SCHEMA CHARACTER_SET_NAME
    SchemaCheckConstraints = 5 'Returns the check constraints defined in the catalog that are owned by a given user.
    '(CHECK_CONSTRAINTS) Rowset) CONSTRAINT_CATALOG CONSTRAINT_SCHEMA CONSTRAINT_NAME
    SchemaCollations = 3 'Returns the character collations defined in the catalog that are accessible to a given user.
    '(COLLATIONS Rowset) COLLATION_CATALOG COLLATION_SCHEMA COLLATION_NAME
    SchemaColumnPrivileges = 13 'Returns the privileges on columns of tables defined in the catalog that are available to, or granted by, a given user.
    '(COLUMN_PRIVILEGES Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME GRANTOR GRANTEE
    SchemaColumns = 4 'Returns the columns of tables '(including views) defined in the catalog that are accessible to a given user.
    '(COLUMNS Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
    SchemaColumnsDomainUsage = 11 'Returns the columns defined in the catalog that are dependent on a domain defined in the catalog and owned by a given user.
     '(COLUMN_DOMAIN_USAGE Rowset) DOMAIN_CATALOG DOMAIN_SCHEMA DOMAIN_NAME COLUMN_NAME
    SchemaConstraintColumnUsage = 6 'Returns the columns used by referential constraints, unique constraints, check constraints, and assertions, defined in the catalog and owned by a given user.
    '(CONSTRAINT_COLUMN_USAGE Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
    SchemaConstraintTableUsage = 7 'Returns the tables that are used by referential constraints, unique constraints, check constraints, and assertions defined in the catalog and owned by a given user.
    '(CONSTRAINT_TABLE_USAGE Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME
    SchemaCubes = 32 'Returns information about the available cubes in a schema '(or the catalog, if the provider does not support schemas).
    '(CUBES Rowset*) CATALOG_NAME SCHEMA_NAME CUBE_NAME
    SchemBInfoKeywords = 30 'Returns a list of provider-specific keywords.
    '(IDBInfo::GetKeywords) <None>
    SchemBInfoLiterals = 31 'Returns a list of provider-specific literals used in text commands.
    '(IDBInfo::GetLiteralInfo) <None>
    Schemimensions = 33 'Returns information about the dimensions in a given cube. It has one row for each dimension.
    '(DIMENSIONS Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME DIMENSION_NAME DIMENSION_UNIQUE_NAME
    SchemaForeignKeys = 27 'Returns the foreign key columns defined in the catalog by a given user.
    '(FOREIGN_KEYS Rowset) PK_TABLE_CATALOG PK_TABLE_SCHEMA PK_TABLE_NAME FK_TABLE_CATALOG FK_TABLE_SCHEMA FK_TABLE_NAME
    SchemaHierarchies = 34 'Returns information about the hierarchies available in a dimension.
    '(HIERARCHIES Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME DIMENSION_UNIQUE_NAME HIERARCHY_NAME HIERARCHY_UNIQUE_NAME
    SchemaIndexes = 12 'Returns the indexes defined in the catalog that are owned by a given user.
    '(INDEXES Rowset) TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME
    SchemaKeyColumnUsage = 8 'Returns the columns defined in the catalog that are constrained as keys by a given user.
    '(KEY_COLUMN_USAGE Rowset) CONSTRAINT_CATALOG CONSTRAINT_SCHEMA CONSTRAINT_NAME TABLE_CATALOG TABLE_SCHEMA TABLE_NAME COLUMN_NAME
    SchemaLevels = 35 'Returns information about the levels available in a dimension.
    '(LEVELS Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME DIMENSION_UNIQUE_NAME HIERARCHY_UNIQUE_NAME LEVEL_NAME LEVEL_UNIQUE_NAME
    SchemaMeasures = 36 'Returns information about the available measures.
    '(MEASURES Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME MEASURE_NAME MEASURE_UNIQUE_NAME
    SchemaMembers = 38 'Returns information about the available members.
    '(MEMBERS Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME DIMENSION_UNIQUE_NAME HIERARCHY_UNIQUE_NAME LEVEL_UNIQUE_NAME LEVEL_NUMBER MEMBER_NAME MEMBER_UNIQUE_NAME MEMBER_CAPTION MEMBER_TYPE Tree operator. For more information, see OLE DB for Online Analytical Processing '(OLAP).
    SchemaPrimaryKeys = 28 'Returns the primary key columns defined in the catalog by a given user.
    '(PRIMARY_KEYS Rowset) PK_TABLE_CATALOG PK_TABLE_SCHEMA PK_TABLE_NAME
    SchemaProcedureColumns = 29 'Returns information about the columns of rowsets returned by procedures.
    '(PROCEDURE_COLUMNS Rowset) PROCEDURE_CATALOG PROCEDURE_SCHEMA PROCEDURE_NAME COLUMN_NAME
    SchemaProcedureParameters = 26 'Returns information about the parameters and return codes of procedures.
     '(PROCEDURE_PARAMETERS Rowset) PROCEDURE_CATALOG PROCEDURE_SCHEMA PROCEDURE_NAME PARAMETER_NAME
    SchemaProcedures = 16 'Returns the procedures defined in the catalog that are owned by a given user.
    '(PROCEDURES Rowset) PROCEDURE_CATALOG PROCEDURE_SCHEMA PROCEDURE_NAME PROCEDURE_TYPE
    SchemaProperties = 37 'Returns information about the available properties for each level of the dimension.
    '(PROPERTIES Rowset) CATALOG_NAME SCHEMA_NAME CUBE_NAME DIMENSION_UNIQUE_NAME HIERARCHY_UNIQUE_NAME LEVEL_UNIQUE_NAME MEMBER_UNIQUE_NAME PROPERTY_TYPE PROPERTY_NAME
    SchemaProviderSpecific = -1 'Used if the provider defines its own nonstandard schema queries. <Provider specific>
    SchemaProviderTypes = 22 'Returns the '(base) data types supported by the data provider.
    '(PROVIDER_TYPES Rowset) DATA_TYPE BEST_MATCH
    SchemaReferentialConstraints = 9 'Returns the referential constraints defined in the catalog that are owned by a given user.
    '(REFERENTIAL_CONSTRAINTS Rowset) CONSTRAINT_CATALOG CONSTRAINT_SCHEMA CONSTRAINT_NAME
    SchemaSchemata = 17 'Returns the schemas '(database objects) that are owned by a given user.
    '(SCHEMATA Rowset) CATALOG_NAME SCHEMA_NAME SCHEMA_OWNER
    SchemaSQLLanguages = 18 'Returns the conformance levels, options, and dialects supported by the SQL-implementation processing data defined in the catalog.
    '(SQL_LANGUAGES Rowset) <None>
    SchemaStatistics '19 'Returns the statistics defined in the catalog that are owned by a given user.
    '(STATISTICS Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME
    SchemaTableConstraints = 10 'Returns the table constraints defined in the catalog that are owned by a given user.
    '(TABLE_CONSTRAINTS Rowset) CONSTRAINT_CATALOG CONSTRAINT_SCHEMA CONSTRAINT_NAME TABLE_CATALOG TABLE_SCHEMA TABLE_NAME CONSTRAINT_TYPE
    SchemaTablePrivileges = 14 'Returns the privileges on tables defined in the catalog that are available to, or granted by, a given user.
    '(TABLE_PRIVILEGES Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME GRANTOR GRANTEE
    SchemaTables = 20 'Returns the tables '(including views) defined in the catalog that are accessible to a given user.
    '(TABLES Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME TABLE_TYPE
    SchemaTranslations = 21 'Returns the character translations defined in the catalog that are accessible to a given user.
    '(TRANSLATIONS Rowset) TRANSLATION_CATALOG TRANSLATION_SCHEMA TRANSLATION_NAME
    SchemaTrustees = 39 'Reserved for future use.
    SchemaUsagePrivileges = 15 'Returns the USAGE privileges on objects defined in the catalog that are available to, or granted by, a given user.
     '(USAGE_PRIVILEGES Rowset) OBJECT_CATALOG OBJECT_SCHEMA OBJECT_NAME OBJECT_TYPE GRANTOR GRANTEE
    SchemaViewColumnUsage = 24 'Returns the columns on which viewed tables, defined in the catalog and owned by a given user, are dependent.
     '(VIEW_COLUMN_USAGE Rowset) VIEW_CATALOG VIEW_SCHEMA VIEW_NAME
    SchemaViews = 23 'Returns the views defined in the catalog that are accessible to a given user.
     '(VIEWS Rowset) TABLE_CATALOG TABLE_SCHEMA TABLE_NAME
    SchemaViewTableUsage = 25 'Returns the tables on which viewed tables, defined in the catalog and owned by a given user, are dependent.
    '(VIEW_TABLE_USAGE Rowset) VIEW_CATALOG VIEW_SCHEMA VIEW_NAME )
End Enum

Private Type SqlField
    Name As IGeneric
    SqlType As IGenericValue
End Type

Private Type Iterator
    Records As ADODB.Recordset
    Fields() As SqlField
    Buffer As GenericOrderedMap
    current As IGeneric
    HasRows As Boolean
    HasLoaded  As Boolean
End Type

'Provider Database
Public Enum ISqlConnectionProvider
    Default_Provider
    Excel_Provider
    Access97_2003_Provider
    Access2007_2013_Provider
    Sql_Provider
    MySql_Provider
    Oracle_Provider
    PostgreSql_Provider
    DBMS_Provider
    CSV_File_Provider
End Enum

Private Type Credentials
    ProviderTyp As ISqlConnectionProvider
    Prompt As ISqlConnectionPrompt
    Mode As ISqlConnectionMode
    DataSource As String
    UserPassword As String
    UserID As String
    Filepassword As String 'Datenbankpasswort einer Accessdatenbank
    HDR As Boolean
    ConnectionString As String
    SSPI As Boolean
    IntitialCatalog As String
End Type

Private Type Member
    Credentials As Credentials
    Locks As Long
    ParameterList As GenericOrderedList
    Iterator As Iterator
End Type
Private myclass As Member

Private WithEvents Database As ADODB.Connection
Attribute Database.VB_VarHelpID = -1

Private Property Get ParameterList() As GenericOrderedList
    Set ParameterList = GenericOrderedList.Build(64)
End Property

Public Function BuildSqlConnection( _
                                    ByVal ServerName As String, _
                                    ByVal InitialCatalog As String, _
                                    Optional ByVal SSPI As Boolean = False, _
                                    Optional ByVal User As String = "SA", _
                                    Optional ByVal Password As String = vbNullString) As GenericSqlManager
    
    With myclass.Credentials
        .ProviderTyp = Sql_Provider
        .DataSource = ServerName
        .IntitialCatalog = InitialCatalog
        .SSPI = SSPI
        If .SSPI = False Then
            .UserID = User
            .UserPassword = Password
        End If
        .Prompt = PromptNever
    End With
   
    Set myclass.ParameterList = ParameterList
    Set BuildSqlConnection = System.CreateInstance(New GenericSqlManager, VarPtr(myclass), LenB(myclass))

End Function

Public Function BuildAccessConnection( _
                                        ByVal Path As String, _
                                        Optional ByVal User As String = "Admin", _
                                        Optional ByVal Password As String, _
                                        Optional ByVal Filepassword As String) As GenericSqlManager
    
    With myclass.Credentials
        .ProviderTyp = Access2007_2013_Provider
        .DataSource = Path
        .UserID = User
        .Filepassword = Filepassword
        .Prompt = PromptNever
    End With
    
    Set myclass.ParameterList = ParameterList
    Set BuildAccessConnection = System.CreateInstance(New GenericSqlManager, VarPtr(myclass), LenB(myclass))

End Function

Public Function BuildCsvConnection(ByVal File As String, Optional ByVal Filepassword As String) As GenericSqlManager
    
    With myclass.Credentials
        .ProviderTyp = CSV_File_Provider
        .DataSource = VBA.Left$(File, InStrRev(File, "\") - 1)
        .Filepassword = Filepassword
        .SSPI = True
        .Prompt = PromptNever
    End With
    
    Set myclass.ParameterList = ParameterList
    Set BuildCsvConnection = System.CreateInstance(New GenericSqlManager, VarPtr(myclass), LenB(myclass))

End Function

Public Function BuildExcelConnection(ByVal File As String, Optional ByVal Filepassword As String) As GenericSqlManager
    
    With myclass.Credentials
        .ProviderTyp = Excel_Provider
        .DataSource = File
        .Filepassword = Filepassword
        .SSPI = True
        .Prompt = PromptNever
    End With
    
    Set myclass.ParameterList = ParameterList
    Set BuildExcelConnection = System.CreateInstance(New GenericSqlManager, VarPtr(myclass), LenB(myclass))
    
End Function

Public Function ExecuteSchema(ByVal Schema As ISqlDBSchema, ParamArray Args() As Variant) As ADODB.Recordset
    
    Dim Rs As ADODB.Recordset
    
    With CreateConnection(myclass.Credentials)
        Call .Open
        
        If IsMissing(Args) Then
            Set Rs = .OpenSchema(Schema)
        Else
            Set Rs = .OpenSchema(Schema, Args)
        End If
        
        Call .Close
        
    End With
    
    RaiseEvent result(Rs)
    
    Set ExecuteSchema = Rs
    
End Function

Public Function SelectWhere(ByVal Table As GString, ByVal Fields As IGenericReadOnlyList, ByVal Where As GenericPair) As IGenericIterator
   
    Const SELECT_STATEMENT As String = "SELECT  FROM [] "
    Dim Sql As GString
    
    With GString
        Set Sql = .Build(SELECT_STATEMENT) _
                    .Insert(15, Table) _
                    .Insert(8, .Join(Fields, ",")) _
                    .Concat(.Build("WHERE " & Where.Key & "=?"))
    End With
    
    Set SelectWhere = Me.RowMapper(Me.Query(Sql.Value, myclass.ParameterList.Push(Where.Value)), Fields)

End Function

Public Sub UpdateWhere(ByVal Table As GString, ByVal Updates As IGenericMap, ByVal Where As GenericPair)
    
    Const UPDATE_STATEMENT As String = "UPDATE [] SET "
    Dim Sql As GString
    
    With GString
        Set Sql = .Build(UPDATE_STATEMENT) _
                    .Insert(15, .Join(Updates.GetKeys, "=?,")) _
                    .Insert(9, Table) _
                    .Concat(.Build(" WHERE " & Where.Key & "=?"))
    End With
    
    Call myclass.ParameterList.AddAll(Updates.GetValues)
    Call Me.Execute(Sql.Value, myclass.ParameterList.Push(Where.Value))

End Sub

Public Sub InsertInto(ByVal Table As GString, ByVal Updates As IGenericMap)

    Const INSERT_STATEMENT As String = "INSERT INTO [] () VALUES ()"
    Dim Sql As GString
    
    'Used String Readonly here to avoid .Repeat(Parameters.Count, "?").ToArray !
    With GString
        Set Sql = .Build(INSERT_STATEMENT) _
                    .Insert(27, .Join(.Repeat(Updates.Elements.Count, "?"), ",")) _
                    .Insert(17, .Join(Updates.GetKeys, ",")) _
                    .Insert(14, Table)
    End With
      
    Call Me.Execute(Sql.Value, Updates.GetValues)
    
End Sub

Public Sub Execute(ByVal Sql As String, Optional ByVal Parameters As IGenericReadOnlyList)
  
    Dim Cmd As ADODB.Command
    Dim AffectedRecords As Long
    
    If Database Is Nothing Then _
        Set Database = CreateConnection(myclass.Credentials)
    
    If Not (Database.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
        Call Database.Open(Options:=IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
        
    Set Cmd = New ADODB.Command
    Call Me.Prepare(Cmd, Sql, 0, Parameters)
    
    Do While Database.State And ISqlObjectState.IsConnecting_State
        If DEBUG_MODE Then _
            Debug.Print "is connecting..."
    Loop
    
    Set Cmd.ActiveConnection = Database
    Call Cmd.Execute(AffectedRecords, , 128 + IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
 
End Sub

Public Function Query(ByVal Sql As String, Optional ByVal Parameters As IGenericReadOnlyList) As ADODB.Recordset
    
    Dim Cmd As ADODB.Command
    Dim Rs As ADODB.Recordset
    
    If Database Is Nothing Then _
        Set Database = CreateConnection(myclass.Credentials)
    
    If Not (Database.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
        Call Database.Open(Options:=IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
    
    Set Cmd = New ADODB.Command
    Call Me.Prepare(Cmd, Sql, ReturnRecords, Parameters)
    
    Do While Database.State And ISqlObjectState.IsConnecting_State
        If DEBUG_MODE Then _
            Debug.Print "is connecting..."
    Loop
    
    Set Cmd.ActiveConnection = Database
    
'    On Error Resume Next
        Set Rs = Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset)
        Call Rs.Open(Cmd, , , , IIf(EXECUTE_ASYNC, ASYNC_OPTION, 0))
    
'        Do While Cmd.State And ISqlObjectState.IsExecuting_State
'            If DEBUG_MODE Then _
'                Debug.Print "is executing..."
'        Loop
'    On Error GoTo 0
   
    Set Query = Rs

End Function

Public Sub Prepare(ByVal Command As ADODB.Command, ByVal Sql As String, ByVal Options As ISqlQueryType, Optional ByVal Parameters As IGenericReadOnlyList)

    Sql = VBA.Trim$(Sql)
    
    With Command
        If (Options And ISqlQueryType.NamedParameter) = ISqlQueryType.NamedParameter Then _
            .NamedParameters = True
        
        If (Options And ISqlQueryType.Prepared) = ISqlQueryType.Prepared Then _
            .Prepared = True
            
        If (Options And ISqlQueryType.StoredProc) = ISqlQueryType.StoredProc Then
            .CommandText = Sql
            .CommandType = ISqlCommandType.CmdStoredProc
        Else
            .CommandText = IIf(VBA.Right$(Sql, 1) = ";", Sql, Sql & ";")
            .CommandType = ISqlCommandType.CMDText
        End If
    End With
    
    Dim Value As IGenericValue
    If Not Parameters Is Nothing Then
        If Parameters.Elements.Count > 0 Then
            With Parameters.Elements.Iterator
                Do While .HasNext(Value)
                    Call Command.Parameters.Append(Value.ToSqlParameter)
                Loop
            End With
        End If
    End If

End Sub

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

Public Function CreateRecordset(ByVal Setting As ISqlRecodsetType) As ADODB.Recordset

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
Private Function CreateConnection(ByRef Credentials As Credentials) As ADODB.Connection
        
    Dim oConnection As ADODB.Connection
    Set oConnection = New ADODB.Connection
    
    With oConnection
        .CursorLocation = ISqlConnectionCursorLocation.Connection_UseServer
        .CommandTimeout = 10
        .Mode = ISqlConnectionMode.ReadWrite_ConnectionMode
        '.Properties("Mode") = Credentials.Mode
        
        Select Case Credentials.ProviderTyp
            Case ISqlConnectionProvider.Excel_Provider
                Select Case VBA.Mid$(Credentials.DataSource, InStrRev(Credentials.DataSource, ".") + 1)
                    Case "xlsx"  'Excel 2007-2010 Workbook (.xlsx)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Filepassword
                    
                    Case "xlsm" 'Excel 2007-2010 Workbook (.xlsm)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Macro;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Filepassword
                    
                    Case "xls" 'Excel 97-2003 Workbook (.xls)'nur für 32-Bit OfficeVersionen
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Properties("Extended Properties") = "Excel 8.0;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Filepassword
                    
                    Case Else
                        Call Err.Raise(1, TypeName$(Me), "Not defined File:" & Credentials.DataSource)
                        '.ConnectionString = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBQ & ";HDR=Yes';"
                End Select
            
            Case ISqlConnectionProvider.CSV_File_Provider
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Properties("Extended Properties") = "Text;HDR=Yes;FMT=Delimited();"
                'Provider=MSDASQL.1;Extended Properties="DefaultDir=C:\MyCSVFolder;Driver={Microsoft Text Driver (*.txt; *.csv)};DriverId=27;Extensions=asc,csv,tab,txt;"
            
            Case ISqlConnectionProvider.Sql_Provider
                .Provider = "SQLOLEDB.1"
                .Properties("Initial Catalog") = Credentials.IntitialCatalog
            
            Case ISqlConnectionProvider.Oracle_Provider 'Error
                .Provider = "MSDAORA"
                .Properties("Initial Catalog") = Credentials.IntitialCatalog
            
            Case ISqlConnectionProvider.PostgreSql_Provider 'Error
                .Provider = "PostgreSQL OLE DB Provider"
                .Properties("Initial Catalog") = Credentials.IntitialCatalog
                .Properties("location") = Credentials.DataSource
            
            Case ISqlConnectionProvider.MySql_Provider 'Error
                .Provider = "MSDASQL.1"
                
            Case ISqlConnectionProvider.Access97_2003_Provider
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Properties("Jet OLEDB:Database Password") = Credentials.Filepassword
                
            Case ISqlConnectionProvider.Access2007_2013_Provider
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Jet OLEDB:Database Password") = Credentials.Filepassword
            
            Case Else
                Call Err.Raise(1, TypeName$(Me), "Not defined File:" & Credentials.DataSource)
            
        End Select
        
        'Diese Property findet man eigentlich bei fast allen DB-Treibern wieder.
        'Somit hat sie auch dementsprechend verschiedene Hintergründe. '
        'Bei Access beschreibt der DataSource den Pfad zu der physischen MDB-Datei. '
        'Beim MS-SQL-Server beschreibt der DataSource den Namen des Servers (nicht die Datenbank!). Bei Informix z.B. beschreibt diese Eigenschaft wieder den Server und den Namen der Datenbank.
        .Properties("Data Source") = Credentials.DataSource
        
        If Credentials.SSPI Then
            .Properties("Integrated Security") = "SSPI"
        Else
            .Properties("User id") = Credentials.UserID
            .Properties("Password") = Credentials.UserPassword
            'nur notwendig, wenn Sie "User ID" und "Password" setzen
'            .Properties("Encrypt Password") = True
        End If
        'Diese Eigenschaft gibt an, ob nach dem Aufbau der Connection zur Datenbank die Userinformationen aus dem ConnectionString und den Connection-Properties gelöscht werden sollen
        'oder gehalten werden sollen. Die Default-Einstellung ist unverständlicher Weise True (behalten). Dieser Wert sollte also explizit auf False gesetzt werden.
        .Properties("Persist Security Info") = False
        .Properties("Prompt") = Credentials.Prompt
        
        Dim p As ADODB.Property
        If DEBUG_MODE Then
'            Debug.Print "Provider :: " & .Provider
'            For Each p In .Properties
'                Debug.Print p.Name & " :: " & p.Value
'            Next
        End If
        
    End With
    
    Set CreateConnection = oConnection
    
End Function

Private Function Connect(ByVal Command As ADODB.Command, ByRef Credentials As Credentials) As Boolean
   
    If Command Is Nothing Then
        'Error
    End If
    
    With Command
        Set .ActiveConnection = CreateConnection(Credentials)
        Call .ActiveConnection.Open
        Connect = IsConnected(.ActiveConnection)
    End With
    
End Function

Private Function IsConnected(ByVal Connection As ADODB.Connection) As Boolean
    
    If Not Connection Is Nothing Then _
        IsConnected = (Connection.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State

End Function

Private Function IsFetching(ByVal Command As ADODB.Command) As Boolean
    
    Const BREAK As Long = 5000000
    Dim Counter As Long
    
    If Not Command Is Nothing Then
        If Not Command.ActiveConnection Is Nothing Then
            Do While (Command.State And ISqlObjectState.IsExecuting_State) = ISqlObjectState.IsExecuting_State 'Do While Connection.State = ISqlConnectionState.Connection_IsOpen + ISqlConnectionState.Connection_IsExecuting
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

Private Sub Database_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_WillConnect"
        
End Sub

Private Sub Database_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_ConnectComplete"
        
End Sub

Private Sub Database_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_Disconnect"
    
    Set Database = Nothing
    
End Sub

Private Sub Database_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
      
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_WillExecute"
    
    myclass.Locks = myclass.Locks + 1
    
End Sub

Private Sub Database_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_ExecuteComplete"
    
    myclass.Locks = myclass.Locks - 1
    If pError Is Nothing Then
        
        If Not pRecordset Is Nothing Then
            Set pRecordset.ActiveConnection = Nothing
            RaiseEvent result(pRecordset)
        End If
            
    Else
    
        Debug.Print "Sql Error"
        Debug.Print pError.Description
        
    End If
    
    Set pCommand.ActiveConnection = Nothing
    
    If CLOSE_CONNECTION_IMMEDIATELY Then _
        If myclass.Locks = 0 Then _
            Call pConnection.Close
    
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSqlManager: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericSqlManager: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSqlManager):  End Property

Private Function IGeneric_ToString() As String
    IGeneric_ToString = TypeName$(Me) & " :: {" & myclass.Credentials.DataSource & "}"
End Function

Private Function IGeneric_Clone() As IGeneric

End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

End Function
 
Private Function EmptyStructure() As Member: End Function

'IGenericIterator
Public Function RowMapper(ByVal Records As ADODB.Recordset, ByVal Fields As IGenericReadOnlyList) As IGenericIterator
    
    Dim newClass As Member
    Dim Field As ADODB.Field
    Dim i As Long

    If (Records Is Nothing) Then _
        Call System.Throw(Me, "Public Function RowMapper").InvalidInput("(Records Is Nothing)")
    
    If (Records.State = ISqlObjectState.IsClosed_State) Then _
        Call System.Throw(Me, "Public Function RowMapper").InvalidInput("(Records.State = ISqlObjectState.IsClosed_State)")
    
    If (Fields Is Nothing) Then _
        Call System.Throw(Me, "Public Function RowMapper").InvalidInput("(Fields Is Nothing)")
    
    ReDim newClass.Iterator.Fields(Fields.Elements.Count - 1)
    Set newClass.Iterator.Records = Records
    Set newClass.Iterator.Buffer = GenericOrderedMap.Build(Fields.Elements.Count)
  
    With newClass.Iterator.Records
    
        Do While (.State And ISqlObjectState.IsExecuting_State) = ISqlObjectState.IsExecuting_State
            If DEBUG_MODE Then _
                Debug.Print "is fetching..."
        Loop
        
        If .State = ISqlObjectState.IsOpen_State Then _
            newClass.Iterator.HasRows = Not (.BOF Or .EOF)
        
        If newClass.Iterator.HasRows Then
        
            If Not (.Fields.Count = Fields.Elements.Count) Then _
                Call System.Throw(Me, "Public Function RowMapper").InvalidInput("number if fields that should mapped doesnt match with queried recordset fields")
            
            For i = 0 To .Fields.Count - 1
'                Debug.Print .Fields(i).Name & " :: " & .Fields(i).Type
                Set newClass.Iterator.Fields(i).Name = Fields.ElementAt(i)
                Call newClass.Iterator.Buffer.Add(Fields.ElementAt(i), Nothing)
                Select Case .Fields(i).Type
                    Case Is = ISqlDataType.Boolean_Type: Set newClass.Iterator.Fields(i).SqlType = GBool
                    Case Is = ISqlDataType.Char_Type: Set newClass.Iterator.Fields(i).SqlType = GString
                    Case Is = ISqlDataType.Numeric_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.Single_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.Currency_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.Double_Type: Set newClass.Iterator.Fields(i).SqlType = GString
                    Case Is = ISqlDataType.Decimal_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.Integer_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.BigInt_Type: Set newClass.Iterator.Fields(i).SqlType = GNumeric
                    Case Is = ISqlDataType.Date_Type: Set newClass.Iterator.Fields(i).SqlType = GDateTime
                    Case Is = ISqlDataType.VarWChar_Type: Set newClass.Iterator.Fields(i).SqlType = GString
                    Case Is = ISqlDataType.LongVarChar_Type: Set newClass.Iterator.Fields(i).SqlType = GString
                    Case Is = ISqlDataType.LongVarWChar_Type: Set newClass.Iterator.Fields(i).SqlType = GString '203 Gibt eine lange Null-terminierte Unicode-Zeichenfolgenwert an.
                    Case Else
                        Set newClass.Iterator.Fields(i).SqlType = GString
                End Select
            Next
        
        ElseIf DEBUG_MODE Then _
            Debug.Print "is fetching..."
            
        End If
    End With
    
    newClass.Iterator.HasLoaded = True
    Set RowMapper = System.CreateInstance(New GenericSqlManager, VarPtr(newClass), LenB(newClass))
    
End Function

Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.current: End Property

Private Function IGenericIterator_HasNext(Out As IGeneric) As Boolean
    
    Dim i As Long
    
    If myclass.Iterator.HasRows Then
        With myclass.Iterator.Records.Fields

            For i = 0 To .Count - 1
                Set myclass.Iterator.Buffer.Item(myclass.Iterator.Fields(i).Name) = myclass.Iterator.Fields(i).SqlType.Convert(.Item(i).Value)
            Next
            
            If DEBUG_MODE Then _
                Debug.Print "Row mapped"
                
        End With
        
        Call myclass.Iterator.Records.MoveNext
        myclass.Iterator.HasRows = Not myclass.Iterator.Records.EOF
        
        Set myclass.Iterator.current = myclass.Iterator.Buffer
        Set Out = myclass.Iterator.current 'Avoid converting twice
    
        IGenericIterator_HasNext = True
        Exit Function
        
    End If
        
    If myclass.Iterator.HasLoaded = False Then _
        Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
    
    Call myclass.Iterator.Buffer.Elements.Clear 'TODO Should we do that ?!
    Set myclass.Iterator.current = Nothing
    Set Out = Nothing

End Function

Private Sub IGenericIterator_Reset()

     With myclass.Iterator
        If Not .Records Is Nothing Then
            If .Records.State = ISqlObjectState.IsOpen_State Then
                .HasRows = Not (.Records.BOF And .Records.EOF)
                If .HasRows Then _
                    Call .Records.MoveFirst
            End If
        End If
    End With
    
End Sub

