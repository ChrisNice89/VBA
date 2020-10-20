VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISqlValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "<T>Class"

Public Enum ISqlDataType
    Autoincrement = -1
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
Public Enum ISqlState
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
Public Enum ISqlQueryType
    ReturnRecords = 2 ^ 0
    Prepared = 2 ^ 1
    Asynch = 2 ^ 2
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

Public Function Create(ByVal Value As Variant, Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As ISqlValue: End Function
Public Function IsRelatedTo(ByVal Other As ISqlValue) As Boolean: End Function
Public Function Prepare(Optional ByVal Name As String, Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As ISqlValue: End Function
Public Sub AppendTo(ByVal Command As ADODB.Command): End Sub

Public Property Get Primitive() As Variant: End Property
Attribute Primitive.VB_UserMemId = 0
Public Property Get SqlType() As ISqlDataType: End Property
Public Property Get PrimitiveType() As VbVarType: End Property
Public Property Get Size() As Byte: End Property
Public Property Get Direction() As ISqlParameterDirection: End Property

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub
