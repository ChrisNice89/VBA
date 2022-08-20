VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSqlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, ParameterNotUsed, ProcedureNotUsed
'@PredeclaredId
'@Exposed
'@Folder "Sql"
Option Explicit
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSqlManager
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

Private Const OPEN_ASYNC As Boolean = True
Private Const CLOSE_CONNECTION_IMMEDIATELY  As Boolean = True
Private Const DEBUG_MODE  As Boolean = False
Private Const ASYNC_OPTION As Long = 16& 'Dont change!
Private Const NO_RECORDS As Long = 128 'Dont change!

Private Enum ISqlQueryStatus
    Cancel = 4
    CantDeny = 3
    ErrorOccurred = 2
    Ok = 1
    UnwantedEvent = 0
End Enum

Private Enum ISqlCursorLocation
    UseClient = 3 'Bedeutet, dass die Daten komplett auf den Client gezogen werden sollen. Dies bringt jedoch einen eventuell großen Speicherbedarf auf dem Client.
    UseServer = 2 'Die Daten werden nicht komplett auf den Client gezogen, sondern nur diese die vom Client aktuell benötigt werden. Das spart viel Speicherbedarf auf dem Client. Nachteil ist jedoch, dass bei manchen Datenanzeige-Controlls wie z.B. DataGrid oder MSHFlexGrid die Daten mit einem serverseitigen Cursor keine Daten angezeigt werden, bzw. ein paar Funktionen nicht mehr verfügbar sind.
End Enum
'
'Der CommandType gibt an was ausgeführt werden soll. Dafür liegt das CommandTypeEnum zu Grunde.
Private Enum ISqlCommandType
    CmdUnspecified = -1 'Does not specify the command type argument.
    CmdText = 1 'Evaluates CommandText as a textual definition of a command or stored procedure call.
    CmdTable = 2 'Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
    CmdStoredProc = 4 'Evaluates CommandText as a stored procedure name.
    CmdUnknown = 8 'Indicates that the type of command in the CommandText property is not known.
    CmdFile = 256 'Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
    'Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect.
    'This value cannot be combined with the ExecuteOptionEnum value adAsyncExecute.
     CmdTableDirect = 512
End Enum
'
Private Enum ISqlCursorType
    OpenForwardOnly = 0 'Dieser Cursor ist der schnellste. Er verbraucht am wenigsten Speicher und Netzresourcen. Cursor kann jedoch nur vorwÃ¤rts bewegt werden. Änderungen sind Zeilenweise mÃ¶glich, jedoch nicht für andere User sichtbar.
    OpenKeyset = 1 'Dieser Cursor erstellt für jeden Datensatz einen Schlüssel. Er ist schneller als der Static- und der Dynamic-Cursor, da er nicht die ganzen DatensÃ¤tze lädt sondern nur einen Schlüssel auf den Datensatz. Änderungen und löschen sind möglich und werden anderen Usern angezeigt. Einzig anfügen wird anderen Usern nicht angezeigt.
    OpenDynamic = 2 'Dieser Cursor ist der flexibelste, er ermöglicht alle Bewegungen. Alle Änderungen werden allen angezeigt. Er ist aber auch der ressourcenbeanspruchenste.
    OpenStatic = 3 'Ruft eine Kopie der abgefragten Daten ab. Daten können nicht geändert werden. Cursor in alle Richtungen beweglich.
End Enum
'
Private Enum ISqlLockType
    LockReadOnly = 1 'Nur Lesen, Daten kÃ¶nnen nicht verÃ¤ndert werden (Default)
    LockPessimistic = 2 'Daten werden beim Ã–ffnen des Recordsets für andere User gesperrt.
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
Private Enum ISqlConnectionMode
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
Private Enum ISqlObjectState
    IsClosed_State = 0        'Indicates that the object is closed.
    IsOpen_State = 1          'Indicates that the object is open.
    IsConnecting_State = 2    'Indicates that the object is connecting.
    IsExecuting_State = 4     'Indicates that the object is executing a command.
    IsFetching_State = 8      'Indicates that the rows of the object are being retrieved.
End Enum
'
Private Enum ISqlConnectionCursorLocation
    Connection_UseClient = 3 'Bedeutet, dass die Daten komplett auf den Client gezogen werden sollen. Dies bringt jedoch einen eventuell großen Speicherbedarf auf dem Client.
    Connection_UseServer = 2 'Die Daten werden nicht komplett auf den Client gezogen, sondern nur diese die vom Client aktuell benötigt werden. Das spart viel Speicherbedarf auf dem Client. Nachteil ist jedoch, dass bei manchen Datenanzeige-Controlls wie z.B. DataGrid oder MSHFlexGrid die Daten mit einem serverseitigen Cursor keine Daten angezeigt werden, bzw. ein paar Funktionen nicht mehr verfügbar sind.
End Enum

'ConnectPromptEnum
'Specifies whether a dialog box should be displayed to prompt for missing parameters when opening a connection to a data source.
Private Enum ISqlConnectionPrompt
    PromptAlways = 1 'Das Fenster wird auf jeden Fall gezeigt (auch wenn die Anmeldedaten eigentlich komplett sein kÃ¶nnten).
    PromptComplete = 2  'Das Fenster wird angezeigt wenn nicht alle Daten vorhanden sind. Es sind alle Optionen möglich.
    PromptCompleteRequired = 3  'Das Fenster wird gezeigt wenn nicht alle Daten vorhanden sind. Es können nur User-Id, Passwort und Servername (SQL-Server) angegeben werden.
    PromptNever = 4 'Erlaubt im MultiUser-Bereich das Öffnen anderer Verbindungen ohne Beschränkungen (Lesen/Schreiben/Beides).
End Enum

Private Enum ISqlDBSchema
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
'
'Provider Database
Public Enum ISqlConnectionProvider
    NoProviderSelected
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
    HasLoaded As Boolean
    ProviderTyp As ISqlConnectionProvider
    Prompt As ISqlConnectionPrompt
    Mode As ISqlConnectionMode
    DataSource As String
    UserPassword As String
    UserID As String
    Password As String
    HDR As Boolean
    SSPI As Boolean
    IntitialCatalog As String
End Type

Private Type Member
    Credentials As Credentials
    ConnectionString As String
    Locks As Long
    ExecutionList As GenericOrderedList
    BatchMode As Boolean
End Type
Private This As Member

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ReadDB As ADODB.Connection
Attribute ReadDB.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents WriteDB As ADODB.Connection
Attribute WriteDB.VB_VarHelpID = -1

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Locks() As Long: Locks = This.Locks: End Property
Public Property Get HasCredentials() As Boolean: HasCredentials = This.Credentials.HasLoaded: End Property
Public Property Get IsAsync() As Boolean: IsAsync = OPEN_ASYNC: End Property

'TODO async Connecting
Public Property Get IsConnected() As Boolean

    If Not ReadDB Is Nothing Then _
        If (ReadDB.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
            IsConnected = True
    
    If Not WriteDB Is Nothing Then _
        If (WriteDB.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
            IsConnected = True
        
End Property

Public Function BuildSqlConnection( _
                                    ByVal ServerName As String, _
                                    ByVal InitialCatalog As String, _
                                    Optional ByVal SSPI As Boolean = False, _
                                    Optional ByVal User As String = "SA", _
                                    Optional ByVal Password As String = vbNullString) As GenericSqlManager
    
    Dim newSqlManager As Member
    
    With newSqlManager.Credentials
        .ProviderTyp = Sql_Provider
        .DataSource = ServerName
        .IntitialCatalog = InitialCatalog
        .SSPI = SSPI
        If .SSPI = False Then
            .UserID = User
            .UserPassword = Password
        End If
        .Prompt = PromptNever
        .Mode = ReadWrite_ConnectionMode
        .HasLoaded = True
    End With

    Set BuildSqlConnection = New GenericSqlManager
    Call System.Inject(BuildSqlConnection, VarPtr(newSqlManager), LenB(newSqlManager))
    
End Function

Public Function BuildAccessConnection( _
                                        ByVal Path As String, _
                                        Optional ByVal User As String = "Admin", _
                                        Optional ByVal Password As String) As GenericSqlManager
    
    Dim newSqlManager As Member
    
    With newSqlManager.Credentials
        .ProviderTyp = Access2007_2013_Provider
        .DataSource = Path
        .UserID = User
        .Password = Password
        .Prompt = PromptNever
        .Mode = ReadWrite_ConnectionMode
        .HasLoaded = True
    End With
 
    Set BuildAccessConnection = New GenericSqlManager
    Call System.Inject(BuildAccessConnection, VarPtr(newSqlManager), LenB(newSqlManager))
    
End Function

Public Function BuildCsvConnection(ByVal File As String, Optional ByVal Password As String) As GenericSqlManager
    
    Dim newSqlManager As Member
    
    With newSqlManager.Credentials
        .ProviderTyp = CSV_File_Provider
        .DataSource = VBA.Left$(File, InStrRev(File, "\") - 1)
        .Password = Password
        .SSPI = True
        .Prompt = PromptNever
        .Mode = ReadWrite_ConnectionMode
        .HasLoaded = True
    End With
    
    Set BuildCsvConnection = New GenericSqlManager
    Call System.Inject(BuildCsvConnection, VarPtr(newSqlManager), LenB(newSqlManager))
    
End Function

Public Function BuildExcelConnection(ByVal File As String, Optional ByVal Password As String) As GenericSqlManager
    
    Dim newSqlManager As Member
    
    With newSqlManager.Credentials
        .ProviderTyp = Excel_Provider
        .DataSource = File
        .Password = Password
        .SSPI = True
        .Prompt = PromptNever
        .Mode = ReadWrite_ConnectionMode
        .HasLoaded = True
    End With

    Set BuildExcelConnection = New GenericSqlManager
    Call System.Inject(BuildExcelConnection, VarPtr(newSqlManager), LenB(newSqlManager))
    
End Function

Public Function Where(ByRef Field As String, ByRef Operator As String) As gString
    Set Where = gString.Of("WHERE [" & Field & "]" & Operator & "?")
End Function

Public Function SelectFrom(ByRef Table As String, ByVal Fields As IGenericSequence) As gString
    
    If (Fields Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function SelectFrom").IsNothing("Fields")
 
    Set SelectFrom = gString.Of("SELECT [] FROM []").Insert(16&, Table).Insert(8&, gString.Join(Fields, "],[").ToString)
            
End Function
'TODO Fields as IGenericSequence
Public Function InsertInto(ByRef Table As String, ByVal Fields As IGenericCollection) As gString
    
    If (Fields Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function InsertInto").IsNothing("Fields")
    '@Ignore ArgumentWithIncompatibleObjectType
    Set InsertInto = gString.Of("INSERT INTO [] () VALUES ()").Insert(26&, gString.Repeat(Fields.Count, "?", ",").ToString).Insert(16&, gString.Join(Fields, ",").ToString).Insert(13&, Table)
            
End Function

Public Function UpdateTable(ByRef Table As String, ByVal Fields As IGenericSequence) As gString
    
    If (Fields Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function UpdateTable").IsNothing("Fields")
 
    Set UpdateTable = gString.Of("UPDATE [] SET []=?").Insert(15&, gString.Join(Fields, "]=?,[").ToString).Insert(8&, Table)
            
End Function

Public Function Query( _
                        ByVal Sql As gString, _
                        Optional ByVal Options As ISqlQueryOptions = ISqlQueryOptions.CommandText) As GenericSqlQuery

    If (Sql Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Query").IsNothing("Sql")
        
    If (Sql.IsEmptyOrWhiteSpace) Then _
        Call GenericError.Build(Me, "Public Function Query").InvalidInput("Sql.IsEmptyOrWhiteSpace")
    
    If Not (Me.HasCredentials) Then _
        Call GenericError.Build(Me, "Public Function Query").InvalidOperation("Not (Me.HasCredentials)")
        
    Set Query = GenericSqlQuery.Build(Sql:=Sql, SqlManager:=Me, Options:=Options)
    
End Function

Public Function Fetch(ByVal Query As GenericSqlQuery, Optional ByVal FetchAsync As Boolean = True) As GenericSqlStream
    
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If (Query Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Fetch").InvalidOperation("Query")
    
    If Not (Query.IsReady) Then _
        Call GenericError.Build(Me, "Public Function Fetch").InvalidOperation("Not (Query.IsReady)")
        
    If ReadDB Is Nothing Then
        Set ReadDB = CreateConnection(This.Credentials, True)
        Call ReadDB.Open(ReadDB.ConnectionString, , , VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0&))
    End If
 
    Set RS = Me.CreateRecordset(ISqlRecodsetType.Disconnected_Recordset)
    Set cmd = CreateCommand(Query)
    
    If Not IsReady(ReadDB) Then _
        Call GenericError.Build(Me, "Public Function Fetch").InvalidOperation("(Cant open connection to database")
    
    Set cmd.ActiveConnection = ReadDB
    Call RS.Open(cmd, , , , VBA.IIf(FetchAsync, ASYNC_OPTION, 0))
    
    Set Fetch = GenericSqlStream.Of(RS)
    
End Function

'@Ignore NonReturningFunction
Public Function Execute(ByVal Query As GenericSqlQuery, Optional ByVal ExecAsync As Boolean = True) As Long

    If (Query Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Execute").InvalidOperation("Query")
    
    If Not (Query.IsReady) Then _
        Call GenericError.Build(Me, "Public Function Execute").InvalidOperation("Not (Query.IsReady)")
    
    If Me.IsConnected Then
        
        With CreateCommand(Query)
            Set .ActiveConnection = WriteDB
            Call .Execute(Execute, , NO_RECORDS)
        End With
        
    Else
    
        Call Me.Connect
        
        If ExecAsync Then
        
            Call This.ExecutionList.Add(CreateCommand(Query))
            
        Else
            
            If Not IsReady(WriteDB) Then _
                Call GenericError.Build(Me, "Public Function Execute").InvalidOperation("(Cant open connection to database")
            
            With CreateCommand(Query)
                Set .ActiveConnection = WriteDB
                Call .Execute(Execute, , NO_RECORDS)
            End With
            
       End If
        
    End If
    
End Function

Public Function Flush(Optional ByVal Disconnect As Boolean = True) As Long
    
    Dim Command As ADODB.Command
    
    If (This.ExecutionList Is Nothing) Then _
        Exit Function
    
    If (This.ExecutionList.Elements.Count = 0) Then _
        Exit Function
    
    Flush = This.ExecutionList.Elements.Count
    
    If Not Me.IsConnected Then _
        Call Me.Connect
        
    This.BatchMode = True
    With This.ExecutionList.Reverse
    
        If Not IsReady(WriteDB) Then _
            Call GenericError.Build(Me, "Public Function Flush").InvalidOperation("(Cant open connection to database")
    
        Do
            Set Command = .Pop
            Set Command.ActiveConnection = WriteDB
            Call Command.Execute(, , NO_RECORDS + ASYNC_OPTION)
        Loop While (.Elements.Count > 0)
        
    End With
    This.BatchMode = False
    
    Call VBA.DoEvents
    
    If Disconnect Then _
        Call Me.Disconnect
    
End Function

Private Function ExecuteSchema(ByVal Schema As ISqlDBSchema, ParamArray Args() As Variant) As ADODB.Recordset
    
    Dim RS As ADODB.Recordset
    
    With CreateConnection(This.Credentials)
        Call .Open
        
        If IsMissing(Args) Then
            Set RS = .OpenSchema(Schema)
        Else
            Set RS = .OpenSchema(Schema, Args)
        End If
        
'        Call .Close
        
    End With
    
    Set ExecuteSchema = RS
    
End Function

'Public Sub ResetID(ByVal Table As String)
'    Call Me.Execute("ALTER TABLE " & Table & " ALTER COLUMN ID COUNTER(1,1)")
'End Sub
Public Function Connect() As GenericSqlManager
    
    If Not (ReadDB Is Nothing) Then
    
        If (ReadDB.State = ISqlObjectState.IsClosed_State) Then _
            Call ReadDB.Open(, , , VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0&))
            
    Else
    
        Set ReadDB = CreateConnection(This.Credentials, True)
        Call ReadDB.Open(ReadDB.ConnectionString, , , VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0&))
        
    End If
    
    If Not (WriteDB Is Nothing) Then
    
        If (WriteDB.State = ISqlObjectState.IsClosed_State) Then _
            Call WriteDB.Open(, , , VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0&))
            
    Else
    
        Set WriteDB = CreateConnection(This.Credentials, False)
        Call WriteDB.Open(WriteDB.ConnectionString, , , VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0&))
        
    End If
    
    Set Connect = Me
    
End Function

Public Function Disconnect() As GenericSqlManager
    
    Dim i As Long
    
    If Me.IsConnected Then
    
        Do While This.Locks > 0
            If i Mod 10000 = 0 Then
                DoEvents
                i = 0
            End If
            i = i + 1
        Loop
    
        If (ReadDB.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
            Call ReadDB.Close
        
        If (WriteDB.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State Then _
            Call WriteDB.Close
       
    End If
    
    Set Disconnect = Me
     
End Function

Private Function Indexes(ByVal Table As String) As ADODB.Recordset
    Set Indexes = ExecuteSchema(SchemaIndexes, Empty, Empty, Empty, Empty, Table)
End Function

Private Function TableDetails(ByVal Table As String) As ADODB.Recordset
    Set TableDetails = ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty)
End Function

Private Function TableExists(ByVal Table As String) As Boolean
    TableExists = ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty).RecordCount > 0
End Function

Private Function ColumnExists(ByVal Table As String, ByVal Field As String) As Boolean
    ColumnExists = ExecuteSchema(SchemaColumns, Empty, Empty, Table, Field).RecordCount > 0
End Function

Private Function Tables() As ADODB.Recordset
    Set Tables = ExecuteSchema(SchemaTables, Empty, Empty, Empty, "TABLE")
End Function

Private Function TableStructure(ByVal Table As String) As ADODB.Recordset
    Set TableStructure = ExecuteSchema(SchemaColumns, Empty, Empty, Table, Empty) '.ToArray("COLUMN_NAME"))
End Function

Public Function ColumnsOf(ByVal Table As gString) As GenericOrderedList
    
    Dim Columns As GenericOrderedList
    
    With TableStructure(Table.ToString)
        Set Columns = GenericOrderedList.Build(255)
        Do While Not (.EOF)
            Call Columns.Add(gString.Of(.Fields.Item("COLUMN_NAME").Value))
            
'            Field = VBA.Replace$(.Fields.Item("COLUMN_NAME").Value, " ", "_")
'            Field = VBA.Replace$(Field, ":", "_")
'            Field = VBA.Replace$(Field, "-", "_")
'            Debug.Print Field & " AS GString"
'            Debug.Print "Set ." & Field & "=GString.Build(""" & .Fields.Item("COLUMN_NAME").Value & """)"
    
            .MoveNext
        Loop
    End With

    Set ColumnsOf = Columns
    
End Function

Public Function CreateRecordset(ByVal Setting As ISqlRecodsetType) As ADODB.Recordset

    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    With RS
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

    Set CreateRecordset = RS
   
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
Private Function CreateConnection(ByRef Credentials As Credentials, Optional ByVal ReadOnly As Boolean) As ADODB.Connection
        
    Dim oConnection As ADODB.Connection
    
    If (Credentials.HasLoaded = False) Then _
        Call GenericError.Build(Me, "Private Function CreateConnection").InvalidOperation("(Credentials.HasLoaded = False)")
        
    Set oConnection = New ADODB.Connection
    
    With oConnection
        .CursorLocation = ISqlConnectionCursorLocation.Connection_UseServer
        .CommandTimeout = 10
        .Mode = VBA.IIf(ReadOnly, ISqlConnectionMode.Read_ConnectionMode, ISqlConnectionMode.ReadWrite_ConnectionMode)
        '.Properties("Mode") = Credentials.Mode
        
        Select Case Credentials.ProviderTyp
            Case ISqlConnectionProvider.Excel_Provider
                Select Case VBA.mid$(Credentials.DataSource, VBA.InStrRev(Credentials.DataSource, ".") + 1)
                    Case "xlsx"  'Excel 2007-2010 Workbook (.xlsx)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Password
                    
                    Case "xlsm" 'Excel 2007-2010 Workbook (.xlsm)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Extended Properties") = "Excel 12.0 Macro;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Password
                    
                    Case "xls" 'Excel 97-2003 Workbook (.xls)'nur für 32-Bit OfficeVersionen
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Properties("Extended Properties") = "Excel 8.0;HDR=YES"
                        .Properties("Jet OLEDB:Database Password") = Credentials.Password
                    
                    Case Else
                        Call Err.Raise(1, VBA.TypeName$(Me), "Not defined File:" & Credentials.DataSource)
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
                .Properties("Jet OLEDB:Database Password") = Credentials.Password
                
            Case ISqlConnectionProvider.Access2007_2013_Provider
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Jet OLEDB:Database Password") = Credentials.Password
            
            Case Else
                Call Err.Raise(1, VBA.TypeName$(Me), "Not defined File:" & Credentials.DataSource)
            
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
'            .Properties("Encrypt Password") = True 'Error bei Access
        End If
        'Diese Eigenschaft gibt an, ob nach dem Aufbau der Connection zur Datenbank die Userinformationen aus dem ConnectionString und den Connection-Properties gelöscht werden sollen
        'oder gehalten werden sollen. Die Default-Einstellung ist unverständlicher Weise True (behalten). Dieser Wert sollte also explizit auf False gesetzt werden.
        .Properties("Persist Security Info") = False
        .Properties("Prompt") = Credentials.Prompt
        
'        Dim p As ADODB.Property
'        If DEBUG_MODE Then
'            Debug.Print "Provider :: " & .Provider
'            For Each p In .Properties
'                Debug.Print p.Name & " :: " & p.Value
'            Next
'        End If
        
    End With
    
    Set CreateConnection = oConnection
    
End Function

Private Function CreateCommand(ByVal Query As GenericSqlQuery) As ADODB.Command
    
    Dim Command As ADODB.Command
    Dim SqlParameter As ADODB.Parameter
    Dim i As Long
      
    Dim Sql As IGenericValue
    Dim Parameter As GenericPair
    Dim Name As IGenericValue
    Dim Value As IGenericValue

    Set Sql = Query.Sql
        
    If Not Sql.HasValue Then _
        Call GenericError.Build(Me, "Private Function CreateCommand").InvalidOperation("Sql is empty")
  
    Set Command = New ADODB.Command
    
    With Command
        
        .CommandType = ISqlCommandType.CmdText
        
        If (Query.Options And ISqlQueryOptions.NamedParameter) = ISqlQueryOptions.NamedParameter Then _
            .NamedParameters = True
        
        If (Query.Options And ISqlQueryOptions.Prepared) = ISqlQueryOptions.Prepared Then _
            .Prepared = True
            
        If (Query.Options And ISqlQueryOptions.StoredProc) = ISqlQueryOptions.StoredProc Then _
            .CommandType = ISqlCommandType.CmdStoredProc
            
        If (Query.Options And ISqlQueryOptions.FullTable) = ISqlQueryOptions.FullTable Then _
            .CommandType = ISqlCommandType.CmdTableDirect
      
        .CommandText = Sql.ToString
        
        If Query.HasParameter Then
            For i = 0 To Query.Count - 1
            
                Set Parameter = Query.GetAt(i)
                If Not (Parameter Is Nothing) Then
                
                    Set Value = Parameter.Value
                    If Not (Value Is Nothing) Then
                        Set SqlParameter = Value.ToSqlParameter
                    Else
                        Set SqlParameter = .CreateParameter(Type:=VBA.VbVarType.vbVariant, Direction:=ISqlParameterDirection.ParamInput, Size:=1, Value:=Null)
                    End If
                    
                    Set Name = Parameter.Key
                    If Not (Name Is Nothing) Then _
                        SqlParameter.Name = Name.ToString
                        
                    Call .Parameters.Append(SqlParameter)

                End If
             
            Next
        End If
       
    End With
   
    Set CreateCommand = Command
    
End Function

Private Function IsReady(ByVal Connection As ADODB.Connection) As Boolean
    
    If (Connection Is Nothing) Then _
        Exit Function
        
    With Connection
    
        If (.State = ISqlObjectState.IsClosed_State) Then _
            Call .Open(Options:=VBA.IIf(OPEN_ASYNC, ASYNC_OPTION, 0))
        
        If (.State And ISqlObjectState.IsConnecting_State) = ISqlObjectState.IsConnecting_State Then
            Do
                If DEBUG_MODE Then _
                    Debug.Print "is connecting..."
            Loop While (.State And ISqlObjectState.IsConnecting_State) = ISqlObjectState.IsConnecting_State
        End If
        
        IsReady = (.State And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State
        
    End With

End Function

Private Sub Class_Terminate()
    
    If Not (This.Locks = 0) Then _
        If DEBUG_MODE Then _
            Debug.Print "Private Sub Class_Terminate() :: Connection closed while locked!"
        
End Sub

Private Sub Database_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_BeginTransComplete"
End Sub

Private Sub Database_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_CommitTransComplete"
End Sub

Private Sub Database_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_RollbackTransComplete"
End Sub

Private Sub Database_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_WillConnect"
        
    If This.ExecutionList Is Nothing Then _
        Set This.ExecutionList = GenericOrderedList.Build
 
End Sub

Private Sub Database_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_ConnectComplete"
    
    Call Me.Flush(Disconnect:=False)
    
End Sub

Private Sub Database_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_Disconnect"
        
    If (This.Locks > 0) Then
        If DEBUG_MODE Then _
            Debug.Print "Locks active :: " & This.Locks
        VBA.DoEvents
    End If
    
    This.Locks = 0

End Sub

Private Sub Database_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
      
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_WillExecute"
    
    This.Locks = This.Locks + 1
'    Call Database.BeginTrans
    
End Sub

Private Sub Database_ExecuteComplete( _
                                        ByVal RecordsAffected As Long, _
                                        ByVal pError As ADODB.Error, _
                                        adStatus As ADODB.EventStatusEnum, _
                                        ByVal pCommand As ADODB.Command, _
                                        ByVal pRecordset As ADODB.Recordset, _
                                        ByVal pConnection As ADODB.Connection)
        
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_ExecuteComplete"
    
    If Not (pCommand Is Nothing) Then _
        Set pCommand.ActiveConnection = Nothing
        
    If (adStatus = ISqlQueryStatus.Ok) Then
'        Call Database.CommitTrans
        
        If Not pRecordset Is Nothing Then _
            Set pRecordset.ActiveConnection = Nothing
        
    Else
        If DEBUG_MODE Then _
            Debug.Print "Sql ExecuteError :: " & pError.Description
'        Call Database.RollbackTrans
    End If
    
    This.Locks = This.Locks - 1
    
    If This.BatchMode = False And CLOSE_CONNECTION_IMMEDIATELY Then _
        If (This.Locks = 0) Then _
            Call pConnection.Close
    
End Sub

Private Sub ReadDB_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
      
    If DEBUG_MODE Then _
        Debug.Print "Private Sub ReadDB_WillExecute"
    
    This.Locks = This.Locks + 1

End Sub

Private Sub ReadDB_ExecuteComplete( _
                                        ByVal RecordsAffected As Long, _
                                        ByVal pError As ADODB.Error, _
                                        adStatus As ADODB.EventStatusEnum, _
                                        ByVal pCommand As ADODB.Command, _
                                        ByVal pRecordset As ADODB.Recordset, _
                                        ByVal pConnection As ADODB.Connection)
        
    If DEBUG_MODE Then _
        Debug.Print "Private Sub ReadDB_ExecuteComplete"
    
    If Not (pCommand Is Nothing) Then _
        Set pCommand.ActiveConnection = Nothing
        
    If (adStatus = ISqlQueryStatus.Ok) Then
        Set pRecordset.ActiveConnection = Nothing
        
    ElseIf DEBUG_MODE Then
        Debug.Print "Sql ExecuteError :: " & pError.Description
        
    End If
    
    This.Locks = This.Locks - 1
    
End Sub

Private Sub ReadDB_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
        Debug.Print "Private Sub ReadDB_Disconnect"
        
    If (This.Locks > 0) Then
        If DEBUG_MODE Then _
            Debug.Print "Locks active :: " & This.Locks
    End If

End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSqlManager: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericSqlManager: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = ObjPtr(Me): End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSqlManager): End Property

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function

Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = _
                        "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                        "User: " & This.Credentials.UserID & vbNewLine & _
                        "Provider: " & This.Credentials.ProviderTyp & vbNewLine & _
                        "Datasource: " & This.Credentials.DataSource & vbNewLine & _
                        "Initial catalog: " & This.Credentials.IntitialCatalog & vbNewLine & _
                        "Locks: " & This.Locks
End Function


