VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "<T>Class"

Public Enum ISqlDataType
    AUTOINCREMENT = -1
    Empty_Type = 0 'No value
    SmallInt_Type = 2 ' A 2-byte signed integer.
    Integer_Type = 3 'A 4-byte signed integer.
    Single_Type = 4 'A single-precision floating-point value.
    Double_Type = 5 'A double-precision floating-point value.
    Currency_Type = 6 'A currency value
    Date_Type = 7 'The number of days since December 30, 1899 + the fraction of a day.
    BSTR_Type = 8 'A null-terminated character string.
    IDispatch_Type = 9 'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
    Error_Type = 10 'A 32-bit error code
    Boolean_Type = 11 'A boolean value.
    Variant_Type = 12 'An Automation Variant. Note: Currently not supported by ADO.
    IUnknown_Type = 13 'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
    Decimal_Type = 14 'An exact numeric value with a fixed precision and scale.
    TinyInt_Type = 16 'A 1-byte signed integer.
    UnsignedTinyInt_Type = 17 'A 1-byte unsigned integer.
    UnsignedSmallInt_Type = 18 'A 2-byte unsigned integer.
    UnsignedInt_Type = 19 'A 4-byte unsigned integer.
    BigInt_Type = 20 'An 8-byte signed integer.
    UnsignedBigInt_Type = 21 'An 8-byte unsigned integer.
    FileTime_Type = 64 'The number of 100-nanosecond intervals since January 1,1601
    GUID_Type = 72 'A globally unique identifier (GUID)
    Binary_Type = 128 'A binary value.
    Char_Type = 129 'A string value.
    WChar_Type = 130 'A null-terminated Unicode character string.
    Numeric_Type = 131 'An exact numeric value with a fixed precision and scale.
    UserDefined_Type = 132 'A user-defined variable.
    DBDate_Type = 133 'A date value (yyyymmdd).
    DBTime_Type = 134 'A time value (hhmmss).
    DBTimeStamp_Type = 135 'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
    Chapter_Type = 136 'A 4-byte chapter value that identifies rows in a child rowset
    PropVariant_Type = 138 'An Automation PROPVARIANT.
    VarNumeric_Type = 139 'A numeric value (Parameter object only).
    VarChar_Type = 200 'A string value (Parameter object only).
    LongVarChar_Type = 201 'A long string value.
    VarWChar_Type = 202 'A null-terminated Unicode character string.
    LongVarWChar_Type = 203 'A long null-terminated Unicode string value.
    VarBinary_Type = 204 'A binary value (Parameter object only).
    LongVarBinary_Type = 205 'A long binary value.
    'AdArray = 0x2000 'A flag value combined with another data
End Enum
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
Public Enum ISqlConnectionState
    Connection_IsClosed = 0        'Indicates that the object is closed.
    Connection_IsOpen = 1          'Indicates that the object is open.
    Connection_IsConnecting = 2    'Indicates that the object is connecting.
    Connection_IsExecuting = 4     'Indicates that the object is executing a command.
    Connection_IsFetching = 8      'Indicates that the rows of the object are being retrieved.
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

Public Property Get RawValue() As Variant: End Property
Attribute RawValue.VB_UserMemId = 0
Public Property Get SqlType() As ISqlDataType: End Property
Public Property Get GetType() As VbVarType: End Property
Public Property Get Size() As Byte: End Property
Public Property Get Instance() As IGeneric: End Property

Public Function Create(ByVal Value As Variant, Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As IGenericValue: End Function
Public Function CompareTo(ByVal Other As IGenericValue) As ComparerEnum: End Function
Public Function ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput, Optional ByVal Name As String): End Function

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub
