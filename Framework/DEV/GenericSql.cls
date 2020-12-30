VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Sql"

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericIterator

Public Event Result(ByVal Data As ADODB.Recordset)

Private Const ASYNC_OPTION As Long = 16&
Private Const EXECUTE_ASYNC As Boolean = True
Private Const CLOSE_CONNECTION_IMMEDIATELY  As Boolean = True
Private Const DEBUG_MODE  As Boolean = False

'Gibt an in welche Richtung der Parameter agiert. Die Konstanten stammen aus dem ParameterDirectoinEnum.
Public Enum ISqlParameterDirection
    ParamUnknown = 0 'Die Richtung des Parameters ist nicht bekannt.
    ParamInput = 1 'Es handelt sich um einen Input-Parameter
    ParamOutput = 2 'Es handelt sich um einen Output-Parameter
    ParamInputOutput = 3 'Der Parameter ist sowohl Übergabeparameter, liefert aber auch einen Wert zurück.
    ParamReturnValue = 4 'Beschreibt einen Return-Wert
End Enum
'
Public Enum ISqlCursorLocation
    UseClient = 3 'Bedeutet, daß die Daten komplett auf den Client gezogen werden sollen. Dies bringt jedoch einen eventuell großen Speicherbedarf auf dem Client.
    UseServer = 2 'Die Daten werden nicht komplett auf den Client gezogen, sondern nur diese die vom Client aktuell benötigt werden. Das spart viel Speicherbedarf auf dem Client. Nachteil ist jedoch, daß bei manchen Datenanzeige-Controlls wie z.B. DataGrid oder MSHFlexGrid die Daten mit einem serverseitigen Cursor keine Daten angezeigt werden, bzw. ein paar Funktionen nicht mehr verfügbar sind.
End Enum
'
Public Enum ISqlCommandState
    StateClosed = 0        'Indicates that the object is closed.
    StateOpen = 1          'Indicates that the object is open.
    StateConnecting = 2    'Indicates that the object is connecting.
    StateExecuting = 4     'Indicates that the object is executing a command.
    StateFetching = 8      'Indicates that the rows of the object are being retrieved.
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

Private Type Iterator
    Records As ADODB.Recordset
    Converter() As IGenericValue
    Current As IGeneric
    HasRows As Boolean
    HasLoaded  As Boolean
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
                        Optional ByVal Options As ISqlQueryType) As GenericSql

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
        .SqlOptions = Options '= ISqlQueryType.ReturnRecords + ISqlQueryType.NamedParameter + ISqlQueryType.Prepared
        .Parameter = Parameter
    End With
    
    Set Build = System.CreateInstance(New GenericSql, Virtual.Pointer, Virtual.Size)
 
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
            .CommandType = ISqlCommandType.CMDText
        End If
    End With
    
    Dim pair As GenericPair
    Dim Value As IGenericValue
    
    If myclass.Parameter > 0 Then
        If myclass.ParameterList Is Nothing Then
            'error
        End If
        
        With myclass.ParameterList.IteratorOf(t:=PairData)
            Do While .HasNext(pair)
                Set Value = pair.Value
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
    
    Do While Database.STATE And ISqlObjectState.IsConnecting_State
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
        .Value = Value.ToValue
        
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
    Debug.Print Connection.STATE
    If Not Connection Is Nothing Then _
        IsConnected = (Connection.STATE And ISqlObjectState.IsOpen_State) = ISqlObjectState.IsOpen_State

End Function

Private Function IsFetching(ByVal Command As ADODB.Command) As Boolean
    
    Const BREAK As Long = 5000000
    Dim Counter As Long
    
    If Not Command Is Nothing Then
        If Not Command.ActiveConnection Is Nothing Then
            Do While (Command.STATE And ISqlCommandState.StateExecuting) = ISqlCommandState.StateExecuting 'Do While Connection.State = ISqlConnectionState.Connection_IsOpen + ISqlConnectionState.Connection_IsExecuting
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
    If DEBUG_MODE Then _
        Debug.Print "Database_ConnectComplete"
End Sub

Private Sub Database_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If DEBUG_MODE Then _
        Debug.Print "Private Sub Database_Disconnect"
    
    Set Database = Nothing
End Sub

'Private Sub Database_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
'    Debug.Print "Private Sub Database_WillExecute"
'End Sub

Private Sub Database_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    
    If DEBUG_MODE Then _
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
 
Private Sub IGeneric_Dispose()

End Sub
Private Function EmptyStructure() As Member: End Function

'IGenericIterator
Public Function RowMapper(ByVal Records As ADODB.Recordset) As IGenericIterator

    Dim Field As ADODB.Field
    Dim i As Long
  
    If (Records Is Nothing) Then _
        Call System.Throw(Me, "Public Function RowMapper").InvalidInput("(Records Is Nothing)")
    
    If (Records.STATE = ISqlCommandState.StateClosed) Then _
        Call System.Throw(Me, "Public Function RowMapper").InvalidInput("(Records.State = ISqlCommandState.StateClosed)")
    
    With myclass.Iterator
        Set .Records = Records
        
        Do While (.Records.STATE And ISqlCommandState.StateExecuting) = ISqlCommandState.StateExecuting
            Debug.Print "Records fetching...."
        Loop
        
        .HasRows = Not (.Records.BOF Or .Records.EOF)
        
        ReDim .Converter(.Records.Fields.Count - 1)
        
        For i = 0 To .Records.Fields.Count - 1
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
        
        .HasLoaded = True
    End With
    
    Set RowMapper = System.CreateInstance(New GenericSql, Virtual.Pointer, Virtual.Size)
    
End Function

Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Out As IGeneric) As Boolean
    
    Dim Row As GenericArray
    Dim i As Long, LB As Long
    
    If myclass.Iterator.HasRows Then
        With myclass.Iterator.Records
            With .Fields
                Set Row = GenericArray.Build(.Count)
                LB = Row.LowerBound
                
                For i = 0 To .Count - 1
                    Set Row.ElementAt(LB + i) = myclass.Iterator.Converter(i).Create(.Item(i).Value)
                Next
            End With
            
            Call .MoveNext
            myclass.Iterator.HasRows = Not .EOF
            
        End With
        
    Else
        
        If myclass.Iterator.HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

    End If
    
    Set myclass.Iterator.Current = Row
    Set Out = myclass.Iterator.Current 'Avoid converting twice
    
    IGenericIterator_HasNext = Not Row Is Nothing

End Function

Private Sub IGenericIterator_Reset()

     With myclass.Iterator
        If Not .Records Is Nothing Then
            If .Records.STATE = ISqlObjectState.IsOpen_State Then
                .HasRows = Not (.Records.BOF And .Records.EOF)
                If .HasRows Then _
                    Call .Records.MoveFirst
            End If
        End If
    End With
    
End Sub

