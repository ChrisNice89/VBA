VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------------------------------
' Module      : SqlFactory / Basisklasse
' Author      : Christoph Nitz(KPMG)
' Last Change : 19.07.2019
' Purpose     : Basisklasse für Connection/Command/Reader/Query/Transaction und Adapter Klasse
'               Datatable/Datarow are not implemented yet
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit

Implements IConnection
Implements ICommand
Implements IQuery
Implements IDataReader
Implements IAdapter
Implements ITransaction
Implements IQueryList
'
Private Const DEBUG_PRINT As Boolean = True
'
Public Event DataAvailable(ByVal Reader As IDataReader, ByRef Cancel As Boolean)
Public Event CommitTrans(ByVal CommandText As String, ByVal State, ByRef Cancel As Boolean)
Public Event StatementCompleted(ByRef Disconnect As Boolean, CommandText As String, ByVal IsError As Boolean)

Public Event Error( _
                        ByVal CommandText As String, _
                        ByVal Connectionstring As String, _
                        ByVal Parameter As ADODB.Parameters, _
                        ByVal ListOfError As ADODB.errors)
                           
Public Event BeforeImport( _
                             ByVal Table As String, _
                             ByVal Source As IConnection, _
                             ByRef Cancel As Boolean)
                            
Public Event AfterImport( _
                            ByVal Table As String, _
                            ByVal Source As IConnection, _
                            ByVal AffectedRecords As Long, _
                            ByVal Successfull As Boolean)
                            
Public Event BeforeExport( _
                             ByVal Table As String, _
                             ByVal Source As IConnection, _
                             ByRef Cancel As Boolean)
                            
Public Event AfterExport( _
                            ByVal Table As String, _
                            ByVal Source As IConnection, _
                            ByVal AffectedRecords As Long, _
                            ByVal Successfull As Boolean)
                                      
'Pointer
Private Const ZEROPOINTER As Long = 0

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As LongPtr)
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As LongPtr)
#End If

Private Type SAFEARRAYBOUND
    cElements    As Long
    lLbound      As Long
End Type

Private Type SafeArray
    cDims        As Integer
    fFeatures    As Integer
    cbElements   As Long
    cLocks       As Long
    pvData       As LongPtr
    Bounds() As SAFEARRAYBOUND
End Type

Private Type VariantTransfer
    #If Win64 Then
        bit(1 To 24) As Byte
    #Else
        bit(1 To 16) As Byte
    #End If
End Type

Private Enum OverrideEnum
    ovConnection = 1
    ovCommand
    ovReader
    ovAdapter
    ovConnectionString
    ovQueryList
    ovEvent
    ovTransaction
End Enum

Private Type RecordsetProperties
    Setting As ICommandRecodsetEnum
    CursorType As CursorTypeEnum
    CursorLocation As CursorLocationEnum
    LockType As LockTypeEnum
    EndOfRecordset As Boolean
End Type
'
Private Type ConnectionCredentials
    ProviderTyp As IConnectionProviderEnum
    Prompt As IConnectionPrompt
    Mode As IConnectionModeEnum
    DataSource As String
    UserPassword As String
    UserID As String
    DatabasePassword As String
    HDR As Boolean
    Connectionstring As String
End Type
'
Private Type QueryMember
   IsSaved As Boolean
   Name As String
   Description As String
   Options As IQueryOptionEnums
   ReturnRecords As Boolean
   Paramter() As ADODB.Parameter
   Async As Boolean
   AffectedRecords As Long
End Type
'
Private Type ReaderMember
    ReaderStatus As IReaderStatusEnum
    Data As ADODB.Recordset
    RowsCount As Long
    HasRows As Boolean
    IsClosed As Boolean
    Read As Boolean
    RowIndex As Long
End Type
'
Private Type CommandMember
    Object As ADODB.Command
    RecordsetType As ICommandRecodsetEnum
    Status As ICommandEventEnum
    ParameterValues As VBA.Collection
End Type
'
Private Type ConnectionMember
    Credentials As ConnectionCredentials
    Object As ADODB.Connection
    Connectionstring As String
    ExtendetProperties As String
    ProviderTyp As IConnectionProviderEnum
    Name As String
End Type
'
Private Type AdapterMember
    Source As IConnection
    Tablelist As scripting.Dictionary 'mscorlib.IDictionary
End Type
'
Private Type DataRowMember
    Index As Long
    Table As IDataTable
    EditMode As DatarowEditModeEnum
End Type
'
Private Type TransactionMember
    Cancel As Boolean
    IsError As Boolean
    Bridge As ITransaction
End Type
'
Private Type ColumnMember
    Index As Long
    Name As String
    Size As Long
    DataType As ICommandDataTypeEnum
End Type
'
Private Type DataTableMember
    Table As mscorlib.ArrayList 'CreateObject("System.Collections.HashTable")
    C As VBA.Collection
    Datarows As mscorlib.ArrayList
    ID As String
    Name As String
    Column() As ColumnMember
End Type
'
Private Type QueryListMember
    Name As String
    Collection As VBA.Collection
End Type
'
Private Type Members
    Instance As String
    GlobalDebug As Boolean
    Datarow As DataRowMember
    Event As TransactionMember
    DataTable As DataTableMember
    Connection As ConnectionMember
    Settings As RecordsetProperties
    Reader As ReaderMember
    Query As QueryMember
    Querylist As QueryListMember
    Command As CommandMember
    Adapter As AdapterMember
End Type

Private myClass As Members
Private WithEvents DatabaseEvent As ADODB.Connection
Attribute DatabaseEvent.VB_VarHelpID = -1

'class constructor
Public Function Override(ParamArray ArgList() As Variant) As SqlFactory
   Call Err.Raise(1#, TypeName$(Me) & "Public Function Override", TypeName$(Me) & "Public Function Override")
End Function

Private Function OverrideConnection(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
       Select Case GetReferencedType(Arg)
            Case VbVarType.vbObject
                Call ObjectToMember(Arg)
            Case VbVarType.vbString
                myClass.Connection.Connectionstring = Arg
        End Select
    Next
       
    With myClass.Connection
        If LenB(.Connectionstring) = 0 Then
            .Connectionstring = ExtractCredentials(.Object)
        End If
        
        .Name = Mid$(.Object.Properties("Data Source"), InStrRev(.Object.Properties("Data Source"), "\") + 1)
        Set DatabaseEvent = .Object
    End With
    
    Set OverrideConnection = Me
    On Error GoTo 0
    
End Function

Private Function OverrideCommand(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    Dim iType As VbVarType
    
    On Error Resume Next
    For Each Arg In ArgList
        Select Case GetReferencedType(Arg)
            Case VbVarType.vbObject
                Call ObjectToMember(Arg)
            Case VbVarType.vbString
                myClass.Connection.Connectionstring = Arg
        End Select
    Next
    
    Set OverrideCommand = Me
    On Error GoTo 0

End Function

Private Function OverrideReader(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant

    On Error Resume Next
    For Each Arg In ArgList
        Call ObjectToMember(Arg)
    Next
    
    With myClass.Reader
        .Read = False
        .RowIndex = -1
        .IsClosed = True
        .RowsCount = -1
        
        If Not .Data Is Nothing Then
            Set .Data = Arg
            .IsClosed = Not HasFlag(.Data.State, ICommandStateEnum.StateOpen)
    
            If Not CBool(.Data.BOF And .Data.EOF) Then
                .HasRows = True
                .RowIndex = 1
                
                If .Data.Supports(adApproxPosition) Then
                    .RowsCount = .Data.RecordCount
                End If
            
            End If
        End If
    End With
    
    Set OverrideReader = Me
    On Error GoTo 0
    
End Function

Private Function OverrideAdapter(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
        If TypeOf Arg Is IConnection Then
            Set myClass.Adapter.Source = Arg
        End If
    Next
    
    Set OverrideAdapter = Me
    On Error GoTo 0
    
End Function

Private Function OverrideConnectionstring(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
        If GetReferencedType(Arg) = VbVarType.vbString Then
            myClass.Connection.Connectionstring = Arg
        End If
    Next
    
    Set OverrideConnectionstring = Me
    On Error GoTo 0
    
End Function

Private Function OverrideQuerylist(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
        Select Case GetReferencedType(Arg)
            Case VbVarType.vbObject
                Call ObjectToMember(Arg)
            Case VbVarType.vbString
                myClass.Querylist.Name = Arg
        End Select
    Next
    
    Set OverrideQuerylist = Me
    On Error GoTo 0
    
End Function

Private Function OverrideEvent(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
        If TypeOf Arg Is ADODB.Connection Then Set DatabaseEvent = Arg
    Next
    
    Set OverrideEvent = Me
    On Error GoTo 0
    
End Function

Private Function OverrideTransaction(ParamArray ArgList() As Variant) As SqlFactory
    
    Dim Arg As Variant
    
    On Error Resume Next
    For Each Arg In ArgList
        If GetReferencedType(Arg) = VbVarType.vbObject Then
            Set myClass.Event.Bridge = Arg
        End If
    Next
    
    If myClass.Event.Bridge Is Nothing Then
        Set myClass.Event.Bridge = Me.CreateTransaction
    End If
        
    Set OverrideTransaction = Me
    On Error GoTo 0
    
End Function

Private Sub Class_Initialize()
    
    Me.GlobalDebug = DEBUG_PRINT
    
    With myClass.Command
        .RecordsetType = Disconnected_Recordset
        .Status = StatusCantDeny
    End With
    
    If Me Is SqlFactory Then
        Me.InstanceName = "Static Factory"
        Call PrintMessage(vbCrLf & myClass.Instance & " Class_Initialize")
    Else
        Me.InstanceName = SqlFactory.InstanceName
        Call PrintMessage(TypeName(Me) & " Instance :: " & myClass.Instance & " Class_Initialize")
    End If
    
End Sub

Private Sub Class_Terminate()
    Call PrintMessage(TypeName$(Me) & " Instance :: " & myClass.Instance & " Class_Terminate")
End Sub
Private Function Blank() As Members: End Function

Public Property Let GlobalDebug(ByVal This As String)
    myClass.GlobalDebug = This
End Property

Public Property Let InstanceName(ByVal This As String)
    myClass.Instance = This
End Property

Public Property Get InstanceName() As String
    InstanceName = myClass.Instance
End Property

Public Property Get Self() As SqlFactory
    Set Self = Me
End Property

Private Property Get ConnectionStatus() As IConnectionStateEnum
    
    Dim eReturn As IConnectionStateEnum

    On Error GoTo Err_handler
     If myClass.Connection.Object Is Nothing Then
        eReturn = Connection_not_initialized
     Else
        eReturn = Connection_initialized
     End If
    On Error GoTo 0

CleanExit:

    ConnectionStatus = eReturn
    Exit Property
    
Err_handler:

    Call ErrorHandler("Public Property Get Status")
    eReturn = Connection_not_initialized
    Resume CleanExit
    
End Property

Private Property Get IsConnected(ByVal Connection As ADODB.Connection) As Boolean

    Dim bReturn As Boolean
    
    If Not Connection Is Nothing Then
        bReturn = HasFlag(Connection.State, ICommandStateEnum.StateOpen)
    End If

    IsConnected = bReturn

End Property

Private Function OpenConnection(Optional ByVal Connectionstring As String = vbNullString) As ADODB.Connection
    
    If LenB(Connectionstring) > 0 Then myClass.Connection.Connectionstring = Connectionstring
    
    With myClass.Connection
        If .Object Is Nothing And LenB(Connectionstring) > 0 Then
            Set .Object = New ADODB.Connection
            .Connectionstring = Connectionstring
        End If
    End With
    
    Call Connect(myClass.Connection)
    Set OpenConnection = myClass.Connection.Object
    
End Function

Private Function CloseConnection() As Boolean
    Call Disconnect(myClass.Connection.Object)
End Function

'Connection class
Private Function Connect(ByRef Info As ConnectionMember) As Boolean
    
    On Error GoTo Err_handler
     If Not IsConnected(Info.Object) Then
        With Info
            Call .Object.Open(.Connectionstring)
        End With
     End If
    On Error GoTo 0
     
ReturnConnect:

    Connect = IsConnected(Info.Object)
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function Connect")
    Resume ReturnConnect
    
End Function

Private Sub Disconnect(ByVal Connection As ADODB.Connection)
   
    If Not Connection Is Nothing Then
        With Connection
            Do While Not HasFlag(.State, ICommandStateEnum.StateExecuting) And HasFlag(.State, ICommandStateEnum.StateOpen)
               Call .Close: Exit Do
            Loop
        End With
    End If
    
End Sub

Private Sub TerminateConnection()
    
    With myClass
        Call Disconnect(.Connection.Object)
        .Connection = BlankConnection
        .Event = BlankTransaction
    End With
    
    Set DatabaseEvent = Nothing
    
End Sub
Private Function BlankConnection() As ConnectionMember: End Function

'IConnection
Private Property Get IConnection_DataSource() As String
    IConnection_DataSource = myClass.Connection.Object.Properties("Data Source")
End Property

Private Property Get IConnection_IsError() As Boolean
    If Not myClass.Connection.Object Is Nothing Then
        IConnection_IsError = CBool(myClass.Connection.Object.errors.Count > 0)
    End If
End Property

Private Property Get IConnection_IsOpen() As Boolean
    IConnection_IsOpen = IsConnected(myClass.Connection.Object)
End Property

Private Property Get IConnection_ConnectionString() As String
    IConnection_ConnectionString = myClass.Connection.Object.Connectionstring
End Property

Private Property Get IConnection_ExtendedProperties() As String
    IConnection_ExtendedProperties = myClass.Connection.Object.Properties("Extended Properties")
End Property

Private Property Get IConnection_Provider() As String
    IConnection_Provider = myClass.Connection.Object.Provider
End Property

Private Property Get IConnection_ProviderTyp() As IConnectionProviderEnum
    IConnection_ProviderTyp = myClass.Connection.ProviderTyp
End Property

Private Property Get IConnection_Name() As String
    IConnection_Name = myClass.Connection.Name
End Property

Private Property Get IConnection_Status() As IConnectionStateEnum
    IConnection_Status = ConnectionStatus
End Property

Private Sub IConnection_Disconnect()
    Call Disconnect(myClass.Connection.Object)
End Sub

Private Function IConnection_Connect() As Boolean
    IConnection_Connect = Connect(myClass.Connection)
End Function

Private Function IConnection_ExternalLink(ByVal Table As String) As String
    With myClass.Connection.Object
        IConnection_ExternalLink = "[" & .Properties("Extended Properties") & ";DATABASE=" & .Properties("Data Source") & "].[" & Trim$(Table) & "]"
    End With
End Function

Public Function CreateAccessConnection(ByVal Path As String, Optional ByVal Password As String) As IConnection
    
    Dim Credentials As ConnectionCredentials
    With Credentials
        .ProviderTyp = Access2007_2013
        .DataSource = Path
        If LenB(Password) > 0 Then
            .DatabasePassword = Password
        End If
        .Prompt = PromptNever
        .UserID = "Admin"
    End With
   
    Set CreateAccessConnection = Me(BuildConnection(Credentials))

End Function

Public Function CreateCsvConnection(ByVal File As String, Optional ByVal Password As String) As IConnection
    
    Dim Credentials As ConnectionCredentials
    With Credentials
        .ProviderTyp = CSV_File
        .DataSource = Left$(File, InStrRev(File, "\") - 1)
        If LenB(Password) > 0 Then .DatabasePassword = Password
        .Prompt = PromptNever
        .UserID = vbNullString
    End With
    
    Set CreateCsvConnection = Me(BuildConnection(Credentials))

End Function

Public Function CreateExcelConnection(ByVal File As String, Optional ByVal Password As String) As IConnection
    
    Dim Credentials As ConnectionCredentials
    With Credentials
        .ProviderTyp = Excel_File
        .DataSource = File
        If LenB(Password) > 0 Then .DatabasePassword = Password
        .Prompt = PromptNever
        
    End With
    
    Set CreateExcelConnection = Me(BuildConnection(Credentials))
    
End Function

'internal stuff
Private Function BuildConnection(ByRef Credentials As ConnectionCredentials) As ADODB.Connection
    
    Dim oConnection As ADODB.Connection
    Set oConnection = New ADODB.Connection
    
    myClass.Connection.ProviderTyp = Credentials.ProviderTyp
    
    On Error GoTo Error_Handler
     With oConnection
         .CursorLocation = adUseClient
         .CommandTimeout = 0
         .Mode = ICommandConnectModeEnum.ModeReadWrite
        
         Select Case myClass.Connection.ProviderTyp
             Case IConnectionProviderEnum.Excel_File
                 Select Case Mid$(Credentials.DataSource, InStrRev(Credentials.DataSource, ".") + 1)
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
              
             Case IConnectionProviderEnum.CSV_File
                 .Provider = "Microsoft.Jet.OLEDB.4.0"
                 .Properties("Extended Properties") = "Text;HDR=Yes;FMT=Delimited();"
                 
             Case IConnectionProviderEnum.Sql
                 .Provider = "SQLOLEDB.1"
            
             Case IConnectionProviderEnum.Access97_2003
                 .Provider = "Microsoft.Jet.OLEDB.4.0"
                
             Case IConnectionProviderEnum.Access2007_2013
                 .Provider = "Microsoft.ACE.OLEDB.12.0"
             
             Case Else
                Call Err.Raise(1, TypeName$(Me), "Not defined File:" & Credentials.DataSource)
             
         End Select
            
         'Diese Property findet man eigentlich bei fast allen DB-Treibern wieder. Somit hat sie auch dementsprechend verschiedene Hintergründe. Bei Access beschreibt der DataSource den Pfad zu der physischen MDB-Datei. Beim MS-SQL-Server beschreibt der DataSource den Namen des Servers (nicht die Datenbank!). Bei Informix z.B. beschreibt diese Eigenschaft wieder den Server und den Namen der Datenbank.
         .Properties("Data Source") = Credentials.DataSource
         
         If LenB(Credentials.UserID) > 0 Then
            .Properties("User id") = Credentials.UserID
         End If
         
         If LenB(Credentials.UserPassword) > 0 Then
           ' .Properties("Password") = Credentials.UserPassword
            'nur notwendig, wenn Sie "User ID" und "Password" setzen
            .Properties("Encrypt Password") = True
         End If
            
         'Diese Eigenschaft gibt an, ob nach dem Aufbau der Connection zur Datenbank die Userinformationen aus dem ConnectionString und den Connection-Properties gelöscht werden sollen
         'oder gehalten werden sollen. Die Default-Einstellung ist unverständlicher Weise True (behalten). Dieser Wert sollte also explizit auf False gesetzt werden.
         '.Properties("Persist Security Info") = False
         .Properties("Prompt") = Credentials.Prompt
        
         'Datenbankpasswort einer Accessdatenbank. Dieses darf nicht mit dem Userpasswort aus den Logininformationen einer MDW-Sicherheitsdatei verwechselt werden.
         If LenB(Credentials.DatabasePassword) > 0 Then
            .Properties("Jet OLEDB:Database Password") = Credentials.DatabasePassword
         End If
         '.Properties("Mode") = Credentials.Mode
        
         '.Properties("Integrated Security") = "SSPI"
         '.Properties("Jet OLEDB:System database") =  'Pfad zu der System-MDW (Sicherheitdatei).
         '.Properties("Location") =
        
     End With
    On Error GoTo 0
            
CleanExit:
    
    Set BuildConnection = oConnection
    Exit Function
    
Error_Handler:

    Call ErrorHandler("Private Function BuildConnection")
    Resume CleanExit

End Function

Private Function IConnection_CreateCommand(Optional ByVal Transaction As ITransaction) As ICommand
    
     With New SqlFactory
        Call vTableSwap(ovTransaction)
        Call .Override(Transaction)
        Call vTableSwap(ovTransaction)
        
        Call vTableSwap(ovCommand)
        Set IConnection_CreateCommand = .Override(myClass.Connection.Object, New ADODB.Command, myClass.Connection.Connectionstring)
        Call vTableSwap(ovCommand)
    End With

End Function

Private Function IConnection_CreateAdapter(Optional ByVal Transaction As ITransaction) As IAdapter
        
    With New SqlFactory
        Dim Source As IConnection
        Set Source = Me
        
        Call vTableSwap(ovTransaction)
        Call .Override(Transaction)
        Call vTableSwap(ovTransaction)
        
        Call vTableSwap(ovAdapter)
        Set IConnection_CreateAdapter = .Override(Source)
        Call vTableSwap(ovAdapter)
    End With
    
End Function

Private Sub IConnection_Dispose()
    Call TerminateConnection
    Call PrintMessage("Private Sub IConnection_Dispose")
End Sub

'ConnectionEvents
Private Sub DatabaseEvent_WillConnect(Connectionstring As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Call PrintMessage(vbCrLf & TypeName$(Me) & " WillConnect " & pConnection.Connectionstring)
End Sub

Private Sub DatabaseEvent_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Call PrintMessage(TypeName$(Me) & " Disconnect from: " & pConnection.Connectionstring)
End Sub

Private Sub DatabaseEvent_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If adStatus = adStatusErrorsOccurred Then
        Call PrintMessage(TypeName(Me) & " Can not Connect " & pConnection.Connectionstring)
    Else
        Call PrintMessage(TypeName(Me) & " ConnectComplete " & pConnection.Connectionstring)
    End If
End Sub

Private Sub DatabaseEvent_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   
    If adStatus = adStatusErrorsOccurred Then
        Call ErrorHandler("Private Sub DatabaseEvent_ExecuteComplete")
        Call PrintMessage(pError.Description)
        If Not pCommand Is Nothing Then Call PrintMessage(pCommand.CommandText)
    End If

End Sub

'Return a fully formatted string containing any current ADO errors in the collection.
Private Function ADOError(ByVal pConnection As ADODB.Connection) As String

    Dim objAdoErr As ADODB.Error
    Dim strErr$
   
    strErr = vbNullString
    If Not pConnection Is Nothing Then
        With pConnection
            If .errors.Count > 0 Then
                For Each objAdoErr In .errors
                    With objAdoErr
                        strErr = "ADO Connection Error :  #" & .Number
                        strErr = strErr & vbCrLf & "Description: " & .Description
                        strErr = strErr & vbCrLf & "JET Error Number:    " & .SqlState
                        strErr = strErr & vbCrLf & "Native Error Number: " & .NativeError
                        strErr = strErr & vbCrLf & "Error reported by:   " & .Source
                        strErr = strErr & vbCrLf & "Help Context:        " & .HelpContext
                        strErr = strErr & vbCrLf & "Help file:           " & .HelpFile
                    End With
                Next objAdoErr
            End If
        End With
    End If
    
    pConnection.errors.Clear
    ADOError = strErr
    
End Function

Private Function ExtractCredentials(ByVal Cnn As ADODB.Connection) As String
    
    Dim Prp As ADODB.Property
    
    Dim oConnBuilder As ADODB.Connection
    Set oConnBuilder = New ADODB.Connection
    
    If Cnn Is Nothing Then Exit Function
    
    On Error GoTo Err_handler
    oConnBuilder.Provider = Cnn.Provider
    For Each Prp In oConnBuilder.Properties
        Prp.Value = Cnn.Properties(Prp.Name).Value
    Next
    
    On Error GoTo 0
    
    ExtractCredentials = oConnBuilder.Connectionstring
    Exit Function
    
Err_handler:

    'Call ErrorHandler("Private Function ExtractCredentials")
    Resume Next

End Function

'Transaction
Public Function CreateTransaction() As ITransaction
    
    SqlFactory.InstanceName = "ITransaction"
    
    With New SqlFactory
        Set CreateTransaction = .Self
    End With

End Function

'Connection
Public Function UseConnection(ByVal Connection As ADODB.Connection) As IConnection
Attribute UseConnection.VB_UserMemId = 0
     
     SqlFactory.InstanceName = "IConnection"
     
     With New SqlFactory
        Call vTableSwap(ovTransaction)
        Call .Override(Me.CreateTransaction)
        Call vTableSwap(ovTransaction)
        
        Call vTableSwap(ovConnection)
        Set UseConnection = .Override(Connection, ExtractCredentials(Connection))
        Call vTableSwap(ovConnection)
     End With
 
    Exit Function
    
End Function

'Command
Public Function CreateCommand(ByVal Connection As ADODB.Connection, ByVal Transaction As ITransaction) As ICommand
    
    SqlFactory.InstanceName = "ICommand"
    
    With New SqlFactory
        Call vTableSwap(ovTransaction)
        Call .Override(Transaction)
        Call vTableSwap(ovTransaction)
        
        Call vTableSwap(ovCommand)
        Set CreateCommand = .Override(Connection, ExtractCredentials(Connection), New ADODB.Command)
        Call vTableSwap(ovCommand)
    End With
    
End Function

'Adapter
Public Function CreateAdapter(ByVal Connection As ADODB.Connection, Optional ByVal Transaction As ITransaction) As IAdapter
    
    SqlFactory.InstanceName = "IDataAdapter"
    
    With New SqlFactory
        Dim Source As IConnection
        Set Source = .UseConnection(Connection)
    
        Call vTableSwap(ovTransaction)
        Call .Override(Transaction)
        Call vTableSwap(ovTransaction)
    
        Call vTableSwap(ovAdapter)
        Set CreateAdapter = .Override(Source)
        Call vTableSwap(ovAdapter)
    End With
    
End Function

'DataReader
Private Function CreateDataReader(ByVal Data As ADODB.Recordset) As IDataReader
    
    If HasFlag(ICommandStateEnum.StateClosed, Data.State) Then
        Call Data.Open
    End If
    
    SqlFactory.InstanceName = "IDataReader"
    
    With New SqlFactory
        Call vTableSwap(ovReader)
        Set CreateDataReader = .Override(Data)
        Call vTableSwap(ovReader)
    End With

End Function

'QueryList
Public Function CreateQueryList(ByVal Name As String) As IQueryList
    
    SqlFactory.InstanceName = "IQueryList"
   
    With New SqlFactory
        Call vTableSwap(ovQueryList)
        Set CreateQueryList = .Override(New VBA.Collection, Name)
        Call vTableSwap(ovQueryList)
    End With
    
End Function

'Ado Parameter
Public Function CreateParameter( _
                                    ByRef Name As String, _
                                    ByVal Value As Variant, _
                                    Optional ByVal Direction As ICommandParameterDirectionEnum = ICommandParameterDirectionEnum.ParamInput, _
                                    Optional ByVal Precision As Integer = 4, _
                                    Optional ByVal NumericScale As Integer = 4) As ADODB.Parameter

    Dim Param As ADODB.Parameter
    Set Param = New ADODB.Parameter
     
    On Error GoTo Err_handler
     With Param
         .Direction = Direction
         .Name = IIf(Left(Name, 1) = "@", Name, "@" & Name)
         .Attributes = adParamNullable
         .Value = Value
         .Type = GetParameterType(Parameter:=Value)
         .Size = Len(Nz(Value, " "))

         If .Type = ICommandDataTypeEnum.adCurrency Or .Type = ICommandDataTypeEnum.adDouble Or .Type = ICommandDataTypeEnum.adSingle Then
             .Precision = Precision
             .NumericScale = NumericScale
         End If

     End With
    On Error GoTo 0

CleanExit:

    'Set Param = Command.CreateParameter(.Name, .Type, .Direction, .Size, .value)
    Set CreateParameter = Param
    Exit Function

Err_handler:

    Call ErrorHandler("Public Function CreateParameter")
    Resume CleanExit

End Function

Private Function GetParameterType(ByVal Parameter As Variant) As ICommandDataTypeEnum

    Dim LongLong As Integer
    Dim eReturn As ICommandDataTypeEnum

    #If Win32 Then
        LongLong = 20
    #End If

    #If Win64 Then
        'LongLong = VBA.vbLongLong
    #End If

    On Error GoTo Err_handler
     Select Case VarType(Parameter)
         
         Case VbVarType.vbEmpty            '    0  Empty (uninitialized).
             Call Err.Raise(13&)
         
         Case VbVarType.vbNull             '    1  Null (no valid data).
             eReturn = ICommandDataTypeEnum.adError
         
         Case VbVarType.vbInteger          '    2  Integer.
             eReturn = ICommandDataTypeEnum.adInteger
         
         Case VbVarType.vbLong             '    3  Long integer.
             eReturn = ICommandDataTypeEnum.adInteger
         
         Case VbVarType.vbSingle           '    4  Single-precision floating-point number.
             eReturn = ICommandDataTypeEnum.adSingle
         
         Case VbVarType.vbDouble           '    5  Double-precision floating-point number.
             eReturn = ICommandDataTypeEnum.adDouble
         
         Case VbVarType.vbCurrency         '    6  Currency.
             eReturn = ICommandDataTypeEnum.adDecimal
         
         Case VbVarType.vbDate             '    7  Date.
             eReturn = ICommandDataTypeEnum.adDate
         
         Case VbVarType.vbString           '    8  String.
             eReturn = ICommandDataTypeEnum.adVarChar
            '.adVarWChar  'Assumes strings are not long
         
         Case VbVarType.vbObject           '    9  Object.
             Call Err.Raise(13&)
         
         Case VbVarType.vbError            '   10  Error.
             Call Err.Raise(13&)
         
         Case VbVarType.vbBoolean          '   11  Boolean.
             eReturn = ICommandDataTypeEnum.adBoolean
         
         Case VbVarType.vbVariant          '   12  Variant (used only with arrays of variants).
             Call Err.Raise(13&)
         
         Case VbVarType.vbDataObject       '   13  A data access object.
             Call Err.Raise(13&)
         
         Case VbVarType.vbDecimal          '   14  Decimal.
             eReturn = ICommandDataTypeEnum.adDecimal
         
         Case VbVarType.vbByte             '   17  Byte.
             eReturn = ICommandDataTypeEnum.adChar
         
         Case LongLong           '   20  LongLong integer (Valid on 64-bit platforms only).
             eReturn = ICommandDataTypeEnum.adBigInt
         
         Case VbVarType.vbUserDefinedType  '   36  Variants that contain user-defined types.
             Call Err.Raise(13&)
         
         Case VbVarType.vbArray            '   8192  Array.
             Call Err.Raise(13&)
         
         Case Else               '   For example array + variant = 8204
             Call Err.Raise(13&)
     
     End Select
    On Error GoTo 0

CleanExit:

    GetParameterType = eReturn
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function GetParameterType")
    eReturn = ICommandDataTypeEnum.adVarWChar
    Resume CleanExit
    
End Function

Public Function CreateRecordset(Optional ByVal Setting As ICommandRecodsetEnum = ICommandRecodsetEnum.Default_Recordset) As ADODB.Recordset

    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    With myClass.Settings
        .Setting = Setting
        Select Case Setting
            Case ICommandRecodsetEnum.ReadOnly_Recordset
                .CursorLocation = ICommandCursorLocationEnum.UseServer
                .CursorType = ICommandCursorTypeEnum.OpenForwardOnly
                .LockType = ICommandLockTypeEnum.LockReadOnly

            Case ICommandRecodsetEnum.UpdateBatch_Recordset
                .CursorLocation = ICommandCursorLocationEnum.UseClient
                .CursorType = ICommandCursorTypeEnum.OpenForwardOnly
                .LockType = ICommandLockTypeEnum.LockBatchOptimistic

            Case ICommandRecodsetEnum.Update_Recordset
                .CursorLocation = ICommandCursorLocationEnum.UseClient 'CursorLocationEnum.adUseServer
                .CursorType = ICommandCursorTypeEnum.OpenForwardOnly 'CursorTypeEnum.adOpenKeyset
                .LockType = ICommandLockTypeEnum.LockOptimistic

            Case ICommandRecodsetEnum.Disconnected_Recordset
                .CursorLocation = ICommandCursorLocationEnum.UseClient  'Setting the cursor location to client side is important to get a disconnected recordset.
                .CursorType = ICommandCursorTypeEnum.OpenForwardOnly
                .LockType = ICommandLockTypeEnum.LockBatchOptimistic

            Case ICommandRecodsetEnum.Default_Recordset
                .CursorLocation = ICommandCursorLocationEnum.UseServer 'CursorLocationEnum.adUseClient  'Client-side cursor to enable AbsolutePosition property
                .CursorType = ICommandCursorTypeEnum.OpenStatic
                .LockType = ICommandLockTypeEnum.LockBatchOptimistic

            Case Else

                Call Err.Raise(1#, TypeName(Me), "CreateRecordset - Type not specified")
        End Select
    End With

    With Rs
        .CursorLocation = myClass.Settings.CursorLocation 'CursorLocationEnum.adUseServer
        .CursorType = myClass.Settings.CursorType
        .LockType = myClass.Settings.LockType
    End With

    Set CreateRecordset = Rs
   
End Function

Private Function GetCleanString( _
                                    ByVal ThisText As String, _
                                    Optional ByVal Chars As String = "! @#$%^&*()\/{[]_.;:_äüö+?}") As String
    
    Dim i&
    
    For i = 1 To Len(Chars)
      ThisText = Replace$(ThisText, Mid$(Chars, i, 1), "")
    Next i
    
    GetCleanString = ThisText

End Function

Private Function ToArray(ByVal Data As ADODB.Recordset, ByVal Transposed As Boolean, ParamArray FieldsToSelect() As Variant) As Variant()

    Dim vRecords As Variant
  
    On Error GoTo Err_handler
     If Data Is Nothing Then
        ReDim vRecords(0, 0)
        Call Err.Raise(1#, , "Recordset is nothing")
     End If
     
     With Data
         If Not (.EOF And .BOF) Then
             Data.MoveFirst
            
             FieldsToSelect = DelegatedParamArray(FieldsToSelect)
            
             If IsMissing(FieldsToSelect) Then
                 vRecords = Data.GetRows
             Else
                 vRecords = Data.GetRows(Fields:=FieldsToSelect)
             End If
            
             If Transposed Then
                 Call Transpose(vRecords)
             End If
         Else
             ReDim vRecords(0, .Fields.Count - 1)
         
         End If
     End With
    On Error GoTo 0

CleanExit:

    ToArray = vRecords
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function ToArray")
    Resume CleanExit
    
End Function

Private Sub Transpose(ByRef vArray As Variant)
    
    Dim ptrSA As LongPtr
    Dim ptrT1 As LongPtr, ptrT2 As LongPtr
    
    Dim t1() As VariantTransfer, t2() As VariantTransfer
    
    Dim i As Long, j As Long, m As Long, n As Long
    Dim TotalSize As LongPtr
    
    If Not HasFlag(VarType(vArray), vbArray) Then Exit Sub
    
    ptrSA = GetSafeArrayPointer(vArray)
    
    With GetSafeArray(ptrSA)
       
        If .cDims <> 2 Then Exit Sub
    
        m = .Bounds(2).cElements
        n = .Bounds(1).cElements
        
        ReDim t1(1 To m, 1 To n)
        ReDim t2(1 To n, 1 To m)

        ptrT1 = VarPtr(t1(1, 1))
        ptrT2 = VarPtr(t2(1, 1))
        
        TotalSize = m * n * .cbElements
    
        Call CopyMemory(ByVal ptrT1, ByVal .pvData, TotalSize)
        
        For i = LBound(t1, 1) To UBound(t1, 1)
            For j = LBound(t1, 2) To UBound(t1, 2)
                t2(j, i) = t1(i, j)
            Next
        Next
    
        Call CopyMemory(ByVal .pvData, ByVal ptrT2, TotalSize)
        
        Dim Bounds() As SAFEARRAYBOUND
        ReDim Bounds(1 To 2)
        Bounds(1) = .Bounds(2)
        Bounds(2) = .Bounds(1)
        
        Call CopyMemory(ByVal ptrSA + 16&, Bounds(1), .cDims * 8&)
    
    End With
  
    Call CopyMemory(ByVal ptrT1, ZEROPOINTER, POINTERSIZE)
    Call CopyMemory(ByVal ptrT2, ZEROPOINTER, POINTERSIZE)

End Sub

Private Function ResizeToVector(ByRef vArray As Variant) As Variant
    
    Dim Out() As Variant
    
    Dim ptr As LongPtr
    ptr = GetSafeArrayPointer(vArray)
    
    Dim lElements As Long
    lElements = GetArrayLength(ptr)
    
    ReDim Out(1 To lElements)
    
    With GetSafeArray(ptr)
        Call CopyMemory(ByVal VarPtr(Out(1)), ByVal .pvData, lElements * .cbElements)
    End With
    
    Call CopyMemory(ByVal ptr, ZEROPOINTER, POINTERSIZE)
    
    ResizeToVector = Out
    
End Function

Private Function GetTypeLength(ByVal DataType As VbVarType) As Long

    Dim b As Long
    b = -1
  
    Select Case DataType
        Case vbVariant
            b = 16
        Case vbByte
            b = 1
        Case vbBoolean, vbInteger
            b = 2
        Case vbSingle, vbLong
            b = 4
        Case vbDouble, vbDate, vbCurrency
            b = 8
    End Select
    
    GetTypeLength = b
  
End Function

Private Function GetSafeArray(ByVal SafeArrayPointer As LongPtr) As SafeArray
    
    Dim SA As SafeArray
    Dim ptr As LongPtr

    If SafeArrayPointer = 0 Then Exit Function
     
    Call CopyMemory(ptr, ByVal SafeArrayPointer, POINTERSIZE)

    If ptr = 0 Then
        Exit Function
    End If

    Call CopyMemory(SA, ByVal SafeArrayPointer, 16&)
    ReDim SA.Bounds(1 To SA.cDims)
    Call CopyMemory(SA.Bounds(1), ByVal SafeArrayPointer + 16&, SA.cDims * 8&)
   
    GetSafeArray = SA

End Function

Private Function GetSafeArrayPointer(ByRef SafeArray As Variant) As LongPtr
    
    Const VT_BYREF As Long = &H4000
    Dim vType As Long
    Dim ptr As LongPtr
    
    Call CopyMemory(vType, ByVal VarPtr(SafeArray), 2&)
    If (vType And VT_BYREF) Then
        Debug.Print vType
    End If
    
    Call CopyMemory(ptr, ByVal VarPtr(SafeArray) + 8&, POINTERSIZE)
    If (vType And VT_BYREF) Then Call CopyMemory(ptr, ByVal ptr, POINTERSIZE)

    
    If IsArray(SafeArray) Then Call CopyMemory(GetSafeArrayPointer, ByVal VarPtr(SafeArray) + 8&, POINTERSIZE)

End Function

Private Function GetArrayDataType(ByVal Matrix As Variant) As VbVarType
    
    Dim DataType As VbVarType
    
    If (VarType(Matrix) And vbArray) Then
        DataType = (VarType(Matrix) And Not vbArray)
    Else
        DataType = vbError
    End If
    
    GetArrayDataType = DataType
End Function

Private Function GetArrayLength(ByVal SafeArrayPointer As LongPtr) As Long
     
    Dim cDims As Integer
    Dim lDimension() As SAFEARRAYBOUND
    Dim lElements As Long
    Dim i As Long
    
    Call CopyMemory(ByVal VarPtr(cDims), ByVal SafeArrayPointer, 2&)
    
    ReDim lDimension(1 To cDims)
    Call CopyMemory(ByVal VarPtr(lDimension(1)), ByVal SafeArrayPointer + 16&, cDims * 8&)
    
    lElements = 1
    For i = 1 To cDims
        lElements = lElements * lDimension(i).cElements
    Next
    
    GetArrayLength = lElements

End Function

Private Function GetArrayInfo(ByRef vArray As Variant, ArrayStructure As SafeArray) As Boolean

    Const VT_BY_REF As Long = &H4000&
    
    Dim ptr As LongPtr, iVType As Integer

    If Not IsArray(vArray) Then Exit Function
    
    With ArrayStructure
        CopyMemory iVType, vArray, 2
        CopyMemory ptr, ByVal VarPtr(vArray) + 8&, 4&
    
        If (iVType And VT_BY_REF) <> 0 Then
            CopyMemory ptr, ByVal ptr, 4
        End If
        
        CopyMemory ArrayStructure, ByVal ptr, 16&
'        only works if Sa.Bound(1 to 2) in the class header
'        If .cDims < 3 Then
'            CopyMemory .Bounds(1), ByVal ptr + 16&, .cDims * 8&
'            GetArrayInfo = True
'        End If
    End With
    
End Function

Private Function DelegatedParamArray(ParamArray ArgList() As Variant) As Variant
    
    Dim ptr As LongPtr
    Dim arrPrms() As Variant, arrWrk() As Variant
    Dim lowbound As Long
    
    arrPrms = ArgList

    On Error GoTo Err_handler
    Do While VarType(arrPrms(lowbound)) >= vbArray And UBound(arrPrms) < 1
        arrWrk = arrPrms(lowbound)
        arrPrms = arrWrk
        lowbound = LBound(arrPrms)
        If lowbound > UBound(arrPrms) Then Exit Do
    Loop
    On Error GoTo 0
    
    DelegatedParamArray = arrPrms
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function DelegatedParamArray")
    
End Function

Private Function Extract(ByVal Pointer As LongPtr) As Variant
    Dim retVal As Variant
    Call CopyMemory(Extract, ByVal Pointer, 16&)
End Function

Private Function ConvertValue(FieldType As ADODB.DataTypeEnum, FieldValue As Variant) As Variant

    On Error GoTo Err_handler
    Select Case FieldType
        Case ICommandDataTypeEnum.adUnsignedInt
            ConvertValue = CInt(FieldValue)
         
        Case ICommandDataTypeEnum.adUnsignedSmallInt
            ConvertValue = CInt(FieldValue)
            
        Case ICommandDataTypeEnum.adUnsignedTinyInt
            ConvertValue = CByte(FieldValue)
        
        Case ICommandDataTypeEnum.adDate
            ConvertValue = CDate(FieldValue)

        Case ICommandDataTypeEnum.adVarWChar
            ConvertValue = CStr(FieldValue)
        
        Case ICommandDataTypeEnum.adInteger
            ConvertValue = CLng(FieldValue)
        
        Case ICommandDataTypeEnum.adDouble
            ConvertValue = CDbl(FieldValue)
        
        Case ICommandDataTypeEnum.adSmallInt
            ConvertValue = CInt(FieldValue)
        
        Case ICommandDataTypeEnum.adCurrency
             ConvertValue = CCur(FieldValue)
        
        Case ICommandDataTypeEnum.adLongVarWChar
            ConvertValue = FieldValue
            
'        Case adNumeric
'            ConvertValue = CDec(FieldValue)

        Case ICommandDataTypeEnum.adBoolean
            ConvertValue = CBool(FieldValue)
            
        Case Else
            ConvertValue = FieldValue
            
    End Select
    On Error GoTo 0
    
    Exit Function
    
Err_handler:
   
    Call ErrorHandler("Private Function ConvertValue")
    ConvertValue = Null
     
End Function

Private Function ToExcelRange( _
                                ByRef vArray As Variant, _
                                ByRef Destination As Excel.Range, _
                                ByVal RangeName As String, _
                                ByRef Columns() As ColumnMember) As Boolean
    
    Dim HasHeader As Boolean
    Dim bError As Boolean
    Dim i As Long
    Dim ptr As LongPtr

    Call CopyMemory(ptr, ByVal VarPtrArray(Columns), POINTERSIZE)
    
    On Error GoTo Err_handler
     With GetSafeArray(ptr)
         If .cDims = 1 Then
             HasHeader = True
             Select Case .Bounds(1).lLbound
                 Case 0
                     For i = 0 To UBound(Columns)
                         Destination.Cells(1).Offset(0, i).Value = Columns(i).Name
                     Next i
                 Case 1
                     For i = 1 To UBound(Columns)
                        Destination.Cells(1).Offset(0, i - 1).Value = Columns(i).Name
                     Next i
             End Select
         End If
     End With
 
     With GetSafeArray(GetSafeArrayPointer(vArray))
        If .cDims = 2 Then
            If HasHeader Then
                Set Destination = Destination.Resize(1 + .Bounds(2).cElements + .Bounds(2).lLbound, .Bounds(1).cElements + .Bounds(1).lLbound)
                Destination.Cells(2, 1).Resize(.Bounds(2).cElements, .Bounds(1).cElements).Value2 = vArray
            Else
                Set Destination = Destination.Resize(.Bounds(2).cElements + .Bounds(2).lLbound, .Bounds(1).cElements + .Bounds(1).lLbound)
                Destination.Value2 = vArray
            End If
        Else
            Call Err.Raise(1#, , "Passed Array Dims: " & .cDims & "!=2")
        End If
     End With
     Destination.Name = GetCleanString(RangeName, "[ ]#'.§$")
    On Error GoTo 0
    
CleanExit:

    ToExcelRange = Not bError
    Exit Function
    
Err_handler:
    
    bError = True
    Call ErrorHandler("Private Function ToExcelRange")
    Resume CleanExit

End Function

Private Sub Format( _
                      ByRef ExportRange As Excel.Range, _
                      ByRef Columns() As ColumnMember, _
                      Optional ByVal HasHeader As Boolean = True)
    
    Dim i As Long
    Dim ptrSA As LongPtr
    Dim DataRange As Excel.Range
    
    On Error GoTo Err_handler
     Set DataRange = ExportRange
     DataRange.WrapText = False
     
     If HasHeader Then
         With DataRange
             .Rows(1).Columns.AutoFit
             .Rows(1).NumberFormat = "@"
             Set DataRange = .Offset(1).Resize(.Rows.Count - 1, .Columns.Count)
         End With
     End If
     
     Call CopyMemory(ptrSA, ByVal VarPtrArray(Columns), POINTERSIZE)
     
     With GetSafeArray(ptrSA)
         If .cDims = 1 Then
             For i = LBound(Columns) To UBound(Columns)
                 Select Case Columns(i).DataType
                     Case ICommandDataTypeEnum.adVarWChar, ICommandDataTypeEnum.adWChar
                         ExportRange.Columns(i).NumberFormat = "@"
                 
                      Case ICommandDataTypeEnum.adBoolean
                         ExportRange.Columns(i).NumberFormat = "0"
        
                      Case ICommandDataTypeEnum.adDBTimeStamp, ICommandDataTypeEnum.adDate, ICommandDataTypeEnum.adDBDate, ICommandDataTypeEnum.adDBTime
                          ExportRange.Columns(i).NumberFormat = "dd/mm/yyyy"
        
                      Case ICommandDataTypeEnum.adCurrency
                          ExportRange.Columns(i).NumberFormat = "#,##0.00 €"
                            
                      Case ICommandDataTypeEnum.adSingle, ICommandDataTypeEnum.adDouble
                          ExportRange.Columns(i).NumberFormat = "0.00"
                      
                      Case ICommandDataTypeEnum.adBigInt, ICommandDataTypeEnum.adInteger, ICommandDataTypeEnum.adSmallInt, ICommandDataTypeEnum.adTinyInt, _
                          ICommandDataTypeEnum.adUnsignedBigInt, ICommandDataTypeEnum.adUnsignedInt, ICommandDataTypeEnum.adUnsignedSmallInt, ICommandDataTypeEnum.adUnsignedTinyInt, _
                          ICommandDataTypeEnum.adDecimal, ICommandDataTypeEnum.adNumeric
        
                          ExportRange.Columns(i).NumberFormat = "0"
        
                      Case Else
                          ExportRange.Columns(i).NumberFormat = "@"
                  End Select
             Next i
         End If
     End With
    On Error GoTo 0
   
    Exit Sub

Err_handler:
    
    Call ErrorHandler("Private Sub Format")
    
End Sub

Private Function ToExcelTable( _
                                ByVal Data As ADODB.Recordset, _
                                ByRef Destination As Excel.Range, _
                                ByVal TableName As String, _
                                Optional ByVal HasHeader As Boolean = True) As Boolean

    
    Dim bError As Boolean
    
    Dim oLstObj As Excel.ListObject
    Dim oLstColumn As Excel.ListColumn
    
    On Error GoTo Err_handler
    With Destination
        Set oLstObj = .Parent.ListObjects.Add( _
                                                    SourceType:=xlSrcQuery, Source:=Data, _
                                                    XlListObjectHasHeaders:=IIf(HasHeader, 1, 0), _
                                                    Tablestylename:="TableStyleMedium9", _
                                                    Destination:=.Cells(1))
    End With
    
    With oLstObj
        .Name = GetCleanString(TableName)
        With .QueryTable
            .AdjustColumnWidth = True
            .Refresh
        End With
        .Range.Columns.AutoFit
        .Range.Interior.ColorIndex = -4142
        .HeaderRowRange.NumberFormat = "@"
        '.HeaderRowRange.Replace "#", "."
        If Not .DataBodyRange Is Nothing Then
            With .DataBodyRange
                .WrapText = False
            End With
        End If
        
        Set Destination = .Range
    End With
    On Error GoTo 0

    ToExcelTable = Not bError
    Exit Function

Err_handler:
    
    bError = True
    Call ErrorHandler("Private Function ToExcelTable")

End Function
Private Sub ClearParameter()
    
    myClass.Query.IsSaved = False
    
    If Not myClass.Command.Object Is Nothing Then
        With myClass.Command.Object.Parameters
            Do While .Count > 0
                .Delete .Count - 1
            Loop
        End With
    End If

End Sub

'ICommand
Private Property Get ICommand_Transaction() As ITransaction
    Set ICommand_Transaction = myClass.Event.Bridge
End Property

Private Function ICommand_Addparameter( _
                                          ByVal Name As String, _
                                          ByVal Value As Variant, _
                                          Optional ByVal Direction As ICommandParameterDirectionEnum = ICommandParameterDirectionEnum.ParamInput, _
                                          Optional ByVal Precision As Integer = 4, Optional ByVal NumericScale As Integer = 4) As ICommand
    
    On Error GoTo Err_handler
     With myClass.Command
         myClass.Query.IsSaved = False
          Name = IIf(LenB(Name) = 0, "P" & .Object.Parameters.Count + 1, Name)
         .Object.Parameters.Append Me.CreateParameter(Name, Value, Direction, Precision, NumericScale)
     End With
    On Error GoTo 0
    
CleanExit:

    Set ICommand_Addparameter = Me
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function ICommand_Addparameter")
    Resume CleanExit

End Function

Private Function ICommand_AddParameter2(ByVal Name As String, ParamArray Values() As Variant) As ICommand
    
    
    On Error GoTo Err_handler
     myClass.Query.IsSaved = False
     With myClass.Command
     
         If .ParameterValues Is Nothing Then
             Set .ParameterValues = New VBA.Collection
         End If
         Name = IIf(LenB(Name) = 0, "P" & .Object.Parameters.Count + 1, Name)
         
         .Object.Parameters.Append Me.CreateParameter(Name, Values(0), ParamInput, 4&, 4&)
         .ParameterValues.Add Values, Name
         
     End With
    On Error GoTo 0
    
CleanExit:

    Set ICommand_AddParameter2 = Me
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function ICommand_Addparameter2")
    Resume CleanExit

End Function

Private Sub ICommand_ClearParameter()
    Call ClearParameter
End Sub

Private Sub ICommand_Dispose()
    With myClass
        .Command = BlankCommand
        .Reader = BlankReader
        .Query = BlankQuery
        .Event = BlankTransaction
    End With
    Call PrintMessage("Private Sub ICommand_Dispose")
End Sub
Private Function BlankReader() As ReaderMember: End Function
Private Function BlankQuery() As QueryMember: End Function
Private Function BlankCommand() As CommandMember: End Function

Private Function ICommand_ExecuteNonQuery(ByVal Sql As String) As Long
    
    Dim lAffectedRecords As Long
    
    On Error GoTo 0
    Call Execute(Prepare(myClass.Command.Object, Sql), OneShot, lAffectedRecords)
    ICommand_ExecuteNonQuery = lAffectedRecords
    On Error GoTo 0
    
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function ICommand_ExecuteNonQuery")

End Function

Private Function ICommand_ExecuteReader(ByVal Sql As String) As IDataReader
    
    On Error GoTo Err_handler
    Set ICommand_ExecuteReader = Execute(Prepare(myClass.Command.Object, Sql), ReturnRecords)
    On Error GoTo 0
    
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function ICommand_ExecuteReader")

End Function

Private Function ICommand_ExecuteScalar(ByVal Sql As String) As Variant
    
    Dim Result As Variant
    
    On Error GoTo Err_handler
    With Execute(Prepare(myClass.Command.Object, Sql), ReturnRecords)
        If .HasRows Then
            Result = .Datarow(0)
        Else
            Result = Null
        End If
        .CloseReader
    End With
    On Error GoTo 0

CleanExit:

    ICommand_ExecuteScalar = Result
    Exit Function
    
Err_handler:
    
    Result = Null
    Call ErrorHandler("Private Function ICommand_ExecuteScalar")
    Resume CleanExit

End Function

Private Function ICommand_ExecuteSchema(ByVal Schema As ICommandSchema, ParamArray Args() As Variant) As IDataReader
        
    Dim Reader As IDataReader
    
    On Error GoTo Err_handler
     If IsMissing(Args) Then
        Set Reader = CreateDataReader(OpenConnection.OpenSchema(Schema))
     Else
        Set Reader = CreateDataReader(OpenConnection.OpenSchema(Schema, Args))
     End If
     
     Call myClass.Event.Bridge.RaiseDataAvailable(Reader)
    
    On Error GoTo 0

CleanExit:
    
    Call StatementCompleted
    Set ICommand_ExecuteSchema = Reader
    Exit Function
    
Err_handler:
    
    Set Reader = Me
    Call ErrorHandler("Private Function ICommand_ExecuteSchema")
    Resume CleanExit

End Function

Private Sub StatementCompleted()

    Dim Cancel As Boolean
    Dim Disconnect As Boolean
    
    Dim Cmd As ADODB.Command
    Dim Params As ADODB.Parameters
    Dim errors As ADODB.errors
    Dim Sql As String
    Dim sConnection As String
    
    On Error GoTo Err_handler
    
    myClass.Command.Status = IIf(myClass.Connection.Object.errors.Count > 0, ICommandEventEnum.StatusErrorsOccurred, ICommandEventEnum.StatusOK)
    
    Set errors = myClass.Connection.Object.errors
    sConnection = myClass.Connection.Connectionstring
    Set Cmd = myClass.Command.Object
    
    If Not Cmd Is Nothing Then
         Set Params = Cmd.Parameters
         Sql = Cmd.CommandText
    Else
        Call Err.Raise(vbObjectError, "StatementCompleted()", "Command is nothing")
    End If
    
    With myClass
        If .Command.Status = StatusErrorsOccurred Then
            Call .Event.Bridge.RaiseError(Sql, sConnection, Params, errors)
        End If
        
        Call .Event.Bridge.RaiseCommitTrans(Sql, .Command.Status, Cancel)
        If Cancel Then
             .Connection.Object.RollbackTrans
        End If
        
        Call .Event.Bridge.RaiseStatementCompleted(Disconnect, .Command.Object.CommandText, CBool(myClass.Connection.Object.errors.Count > 0))
        If Disconnect Then
            Call CloseConnection
        End If

    End With
    On Error GoTo 0
        
    Exit Sub

Err_handler:
    
    myClass.Command.Status = StatusErrorsOccurred
    Call ErrorHandler("Private Sub StatementCompleted")
    
    If Err.Number = vbObjectError Then
        Exit Sub
    Else
        Resume Next
    End If

End Sub

Private Property Get ICommand_IsConnected() As Boolean
    ICommand_IsConnected = CBool(myClass.Connection.Object.State = ICommandStateEnum.StateOpen)
End Property

Private Function ICommand_Prepare( _
                                    ByVal CommandText As String, _
                                    ByVal Options As IQueryOptionEnums, _
                                    ByVal Name As String, _
                                    Optional ByVal Description As String = vbNullString) As IQuery
    
    Dim i As Long
    
    With myClass.Command
        .Status = StatusCantDeny
        Set .Object = Prepare(.Object, CommandText, Options)
    End With
    
    With myClass.Query
        .Name = Name
        .Options = Options
        .Description = Description
        .ReturnRecords = HasFlag(Options, IQueryOptionEnums.ReturnRecords)
        .Async = HasFlag(Options, IQueryOptionEnums.Asynch)
        .IsSaved = True
        
        If myClass.Command.Object.Parameters.Count > 0 Then
            ReDim .Paramter(1 To myClass.Command.Object.Parameters.Count)
            For i = 0 To myClass.Command.Object.Parameters.Count - 1
                Set .Paramter(i + 1) = myClass.Command.Object.Parameters(i)
            Next i
        End If
    
    End With
    
    Set ICommand_Prepare = Me
    
End Function

Private Function Prepare( _
                            ByVal Command As ADODB.Command, _
                            ByVal CommandText As String, _
                            Optional ByVal Options As IQueryOptionEnums) As ADODB.Command
                            
                            
    CommandText = Trim$(CommandText)
    
    With Command
        .NamedParameters = True
        
        If HasFlag(Options, IQueryOptionEnums.StoredProc) Then
            .CommandText = CommandText
            .CommandType = adCmdStoredProc
        Else
            .CommandText = IIf(Right(CommandText, 1) = ";", CommandText, CommandText & ";")
            .CommandType = adCmdText
        End If
        
        If HasFlag(Options, IQueryOptionEnums.Prepared) Then
            .Prepared = True
        Else
            .Prepared = False
        End If

    End With
    
    Set Prepare = Command
        
End Function

Private Function Execute(ByVal Command As ADODB.Command, Options As IQueryOptionEnums, Optional ByRef AffectedRecords) As IDataReader
        
        Dim bError As Boolean
        Dim Rs As ADODB.Recordset
        Dim Reader As IDataReader
        
        On Error GoTo Err_handler
        Set Command.ActiveConnection = OpenConnection
        
        If HasFlag(Options, IQueryOptionEnums.ReturnRecords) Then
            Set Rs = Me.CreateRecordset(Disconnected_Recordset)
            Call Rs.Open(ByVal Command)
            Set Rs.ActiveConnection = Nothing
            
            If Not bError Then
                Set Reader = CreateDataReader(Rs)
                Call myClass.Event.Bridge.RaiseDataAvailable(Reader)
            End If
        
        ElseIf HasFlag(Options, IQueryOptionEnums.OneShot) Then
            If HasFlag(Options, IQueryOptionEnums.Asynch) Then
                Call Command.Execute(AffectedRecords, , 128 + 16) '+ 16 '/ adAsyncExecute
                   
                Do Until Command.State <> ICommandStateEnum.StateExecuting
                    'DoEvents
                    Call PrintMessage("Command is executing ...")
                Loop
                
            Else
                Call Command.Execute(AffectedRecords, , 128) '128 adExecuteNoRecords
            End If
        
        Else
        
            Call Err.Raise(1#, , "No excute options")
        End If
        On Error GoTo 0
        
CleanExit:
    
    Call StatementCompleted
    Set Execute = Reader
    Exit Function

Err_handler:
    
    Set Reader = Me
    bError = True
    Call ErrorHandler("Private Function Execute")
    Resume CleanExit
    
End Function

Private Sub IDataReader_CloseReader()
    With myClass.Reader
        If Not .Data Is Nothing Then
            If HasFlag(.Data.State, ICommandStateEnum.StateOpen) Then .Data.Close
        End If
        .IsClosed = True
        .RowIndex = -1
        .Read = False
        Set .Data = Nothing
    End With
End Sub

Private Property Get IDataReader_IsClosed() As Boolean
    IDataReader_IsClosed = myClass.Reader.IsClosed
End Property

Private Property Get IDataReader_Fields() As ADODB.Fields
    Set IDataReader_Fields = myClass.Reader.Data.Fields
End Property

Private Property Get IDataReader_HasRows() As Boolean
    IDataReader_HasRows = myClass.Reader.HasRows
End Property

Private Property Get IDataReader_DataRow(ByVal Index As Variant) As ADODB.Field
    Set IDataReader_DataRow = myClass.Reader.Data(Index)
End Property

Private Function IDataReader_Read() As Boolean

    With myClass.Reader
        If .Read Then .Data.MoveNext: .RowIndex = .RowIndex + 1
        .Read = Not .Data.EOF
        IDataReader_Read = .Read
    End With

End Function

Private Property Get IDataReader_RowIndex() As Long
    IDataReader_RowIndex = myClass.Reader.RowIndex
End Property

Private Property Get IDataReader_RowsCount() As Long
    IDataReader_RowsCount = myClass.Reader.RowsCount
End Property

Private Property Get IDataReader_FieldsCount() As Long
    IDataReader_FieldsCount = myClass.Reader.Data.Fields.Count
End Property

Private Function IDataReader_ToArray(ParamArray FieldsToSelect() As Variant) As Variant
    IDataReader_ToArray = ToArray(myClass.Reader.Data, True, FieldsToSelect)
End Function

Private Function IDataReader_ToComboBox(ByVal Control As Combobox, Optional ByVal ListRows As Integer = 10) As Combobox
    
    With Control
        .ListRows = ListRows
        .RowSource = vbNullString
        '.ListIndex = 0
        .RowSourceType = "Table/Query"
        If myClass.Reader.HasRows Then Set .Recordset = myClass.Reader.Data
    End With
    
    Set IDataReader_ToComboBox = Control
    
End Function
                                     
Private Function IDataReader_ToExcel( _
                                        ByRef Destination As Excel.Range, _
                                        ByVal RangeName As String, _
                                        Optional ByVal RangeToTable As Boolean = True, _
                                        Optional ByVal HasHeader As Boolean = True) As Boolean
                                   
    Dim retVal As Boolean
    Dim i As Long
    Dim Columns() As ColumnMember
    
    On Error GoTo Err_handler
    With myClass.Reader
        If .Data Is Nothing Then Exit Function
        
        ReDim Columns(1 To .Data.Fields.Count)
        For i = LBound(Columns) To UBound(Columns)
            Columns(i) = CreateColumn(.Data.Fields(i - 1), i)
        Next
        
        If RangeToTable Then
            retVal = ToExcelTable(.Data, Destination, RangeName, HasHeader)
        Else
            retVal = ToExcelRange(ToArray(.Data, True), Destination, RangeName, Columns)
        End If
        
        Call Format(Destination, Columns, HasHeader)
    
    End With
    On Error GoTo 0
   
    IDataReader_ToExcel = retVal
    Exit Function

Err_handler:

    Call ErrorHandler("Private Function IDataReader_ToExcel")
    
End Function

'IQuery
Private Property Get IQuery_CommandText() As String
    IQuery_CommandText = myClass.Command.Object.CommandText
End Property

Private Property Get IQuery_Description() As String
    IQuery_Description = myClass.Query.Description
End Property

Private Property Get IQuery_ParameterValues(ByVal Index As Variant) As Variant()
    
    Dim retVal() As Variant
    
    If VarType(Index) = vbString Then Index = IIf(Left(Index, 1) = "@", Index, "@" & Index)
    
    On Error GoTo Err_handler
    retVal = myClass.Command.ParameterValues(Index)
    On Error GoTo 0
    
CleanExit:

    IQuery_ParameterValues = retVal
    Exit Property

Err_handler:
    
    ReDim retVal(0)
    retVal(0) = vbNullString
    Resume CleanExit
    
End Property

Private Property Get IQuery_Transaction() As ITransaction
    Set IQuery_Transaction = myClass.Event.Bridge
End Property

Private Property Get IQuery_IsSaved() As Boolean
    IQuery_IsSaved = myClass.Query.IsSaved
End Property

Private Property Get IQuery_Name() As String
    IQuery_Name = myClass.Query.Name
End Property

Private Property Get IQuery_Parameters() As ADODB.Parameter()
    IQuery_Parameters = myClass.Query.Paramter
End Property

Private Property Get IQuery_Parameter(ByVal Index As Variant) As ADODB.Parameter
    If VarType(Index) = vbString Then Index = IIf(Left(Index, 1) = "@", Index, "@" & Index)
    Set IQuery_Parameter = myClass.Command.Object.Parameters(Index)
End Property

Private Property Get IQuery_Succesfull() As Boolean
    IQuery_Succesfull = CBool(myClass.Command.Status = StatusOK And myClass.Query.IsSaved)
End Property

Private Function IQuery_UpdateParameter(ByVal Index As Variant, NewValue As Variant) As IQuery

    On Error GoTo Err_handler
     If myClass.Query.IsSaved Then
         If VarType(Index) = vbString Then Index = IIf(Left(Index, 1) = "@", Index, "@" & Index)
         myClass.Command.Object.Parameters(Index) = NewValue
     End If
    On Error GoTo 0
      
CleanExit:

    Set IQuery_UpdateParameter = Me
    Exit Function

Err_handler:
    
    Call ErrorHandler("Private Function IQuery_UpdateParameter")
    Resume CleanExit

End Function

Private Function IQuery_Execute() As IDataReader
    
    Dim Reader As IDataReader
    
    On Error GoTo Err_handler
     With myClass.Query
        .AffectedRecords = -1
        If Not .IsSaved Then Call Err.Raise(1#, , "Query is not prepared")
        Set Reader = Execute(myClass.Command.Object, .Options, .AffectedRecords)
     End With
    On Error GoTo 0

CleanExit:
    
    Set IQuery_Execute = Reader
    Exit Function

Err_handler:
    
    Set Reader = Me
    Call ErrorHandler(myClass.Query.Name)
    Resume CleanExit
  
End Function

'IAdapter
Private Property Get Adapter() As IAdapter: Set Adapter = Me: End Property
Private Function IAdapter_CreateDataTable(ByVal Sql As String) As IDataTable
    
'    Dim ConnectionWasOpen As Boolean
'    ConnectionWasOpen = Adapter.Source.IsOpen
'
'    myClass.DataTable.Name = Split(Split(Sql, " FROM ", 2)(1), " WHERE ", 2)(0)
'
'    With Adapter.TableStructure(myClass.DataTable.Name)
'        Do While .Read
'            If CInt(.Datarow("COLUMN_FLAGS")) = 90 And CInt(.Datarow("DATA_TYPE")) = 3 Then
'                myClass.DataTable.ID = .Datarow("COLUMN_NAME").Value
'                Exit Do
'            End If
'        Loop
'    End With
'
'    With Adapter
'        If myClass.DataTable.ID = vbNullString Then
'            myClass.DataTable.ID = "ID"
'            Call .AddField(myClass.DataTable.Name, myClass.DataTable.ID, adAutoincrement)
'        End If
'
'        Sql = Left$(Sql, Len("SELECT") + 1) & myClass.DataTable.ID & " AS [Mapping ID]," & Mid$(Sql, Len("Select") + Len(myClass.DataTable.ID)) '& " ORDER BY 1 ASC"
'        Call ReaderToDataTable(.CreateCommand.ExecuteReader(Sql))
'        If Not ConnectionWasOpen Then Call .Source.Disconnect
'
'    End With
'
'    Set IAdapter_CreateDataTable = Me
    
End Function

Private Function IAdapter_CreateCommand() As ICommand
    Set IAdapter_CreateCommand = myClass.Adapter.Source.CreateCommand(myClass.Event.Bridge)
End Function

Private Property Get IAdapter_Transaction() As ITransaction
    Set IAdapter_Transaction = myClass.Event.Bridge
End Property

Private Property Get IAdapter_Source() As IConnection
    Set IAdapter_Source = myClass.Adapter.Source
End Property

Private Function IAdapter_GetColumns(ByVal Table As String) As Variant()
    IAdapter_GetColumns = ResizeToVector(Adapter.CreateCommand.ExecuteSchema(SchemaColumns, Empty, Empty, Table, Empty).ToArray("COLUMN_NAME"))
End Function

Private Function IAdapter_GetTables(Optional ByVal Export As Boolean = False) As Variant()
    If Export Then
        Select Case iCAT.Layout
            Case Genossenschaften
                IAdapter_GetTables = ResizeToVector(Array("tblG_00_Basis", "tblG_01_Stamm", "tblG_02_GrdSt", _
                    "tblG_03_GrdStEigent", "tblG_04_GrdStBelast", "tblG_05_GrdStZuord", "tblG_06_andSi", _
                    "tblG_07_andSiSiZuord", "tblG_08_Bürg", "tblG_09_BürgSiZuord", "tblG_10_Si", "tblG_11_SiKtoZuord", _
                    "tblG_12_BeurteilP", "tblG_13_BeurteilE", "tblG_14_KdStamm", "tblG_15_ZwgSt", "tblG_16_Einh", _
                    "tblG_17_EinhMitgl", "tblG_18_OblKto", "tblG_19_OblKd", "tblG_20_KBS", "tblG_HT_Konto", _
                    "tblG_HT_Kunden"))
            Case Sparkassen
                IAdapter_GetTables = ResizeToVector(Array("tblBBK", "tblSICH", "tblSIRE", _
                    "tblVMGO", "tblS_HT_Kunden"))
            Case SI
                IAdapter_GetTables = ResizeToVector(Array("tblSI_01_D", "tblSI_02_I-D", _
                    "tblSI_03_I", "tblSI_04_I-C", "tblSI_05_C"))
            Case Else
        End Select
    Else
        IAdapter_GetTables = ResizeToVector(Adapter.CreateCommand.ExecuteSchema(SchemaTables, Empty, Empty, Empty, "TABLE").ToArray("TABLE_NAME"))
    End If
End Function

Private Function IAdapter_RecordsCount(ByVal Table As String) As Long
    IAdapter_RecordsCount = Adapter.CreateCommand.ExecuteScalar("SELECT COUNT (*) FROM " & Table)
End Function

Private Function IAdapter_Indexes(ByVal Table As String) As IDataReader
    Set IAdapter_Indexes = Adapter.CreateCommand.ExecuteSchema(SchemaIndexes, Empty, Empty, Empty, Empty, Table)
End Function

Private Function IAdapter_TableDetails(ByVal Table As String) As IDataReader
    Set IAdapter_TableDetails = Adapter.CreateCommand.ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty)
End Function

Private Function IAdapter_TableExists(ByVal Table As String) As Boolean
    IAdapter_TableExists = Adapter.CreateCommand.ExecuteSchema(SchemaTables, Empty, Empty, Table, Empty).HasRows
End Function

Private Function IAdapter_ColumnExists(ByVal Table As String, ByVal Field As String) As Boolean
    IAdapter_ColumnExists = Adapter.CreateCommand.ExecuteSchema(SchemaColumns, Empty, Empty, Table, Field).HasRows
End Function

Private Function IAdapter_Tables() As IDataReader
    Set IAdapter_Tables = Adapter.CreateCommand.ExecuteSchema(SchemaTables, Empty, Empty, Empty, "TABLE")
End Function

Private Function IAdapter_TableStructure(ByVal Table As String) As IDataReader
    Set IAdapter_TableStructure = Adapter.CreateCommand.ExecuteSchema(SchemaColumns, Empty, Empty, Table, Empty)
End Function

Private Sub IAdapter_Reset(ParamArray Tables() As Variant)
    
    Dim Table As Variant
    
    With Adapter.CreateCommand
        For Each Table In DelegatedParamArray(Tables)
            Call .ExecuteNonQuery("DELETE FROM " & Table)
            Call .ExecuteNonQuery("ALTER TABLE " & Table & " ALTER COLUMN ID COUNTER(1,1)")
        Next Table
    End With
    
End Sub

Private Function IAdapter_AddField( _
                                        ByVal ToTable As String, _
                                        ByVal FieldName As String, _
                                        ByVal FieldType As ICommandDataTypeEnum, _
                                        Optional ByVal FieldLen As Long = 50&, _
                                        Optional ByVal FieldAllowsNull As Boolean = True) As IAdapter
    
    Dim FieldText As String
    Dim Sql As String
    
    Select Case FieldType
        Case ICommandDataTypeEnum.adInteger: FieldText = "Long"
        Case ICommandDataTypeEnum.adLongVarChar: FieldText = "Text(" & FieldLen & ")"
        Case ICommandDataTypeEnum.adBinary: FieldText = "BIT   DEFAULT 0"
        Case ICommandDataTypeEnum.adAutoincrement: FieldText = "AUTOINCREMENT PRIMARY KEY"
        Case Else:
    End Select
    
    Sql = "ALTER TABLE [" & ToTable & "] ADD COLUMN [" & FieldName & "] " & FieldText
    
    If FieldAllowsNull Then
       Sql = Sql & " NULL"
    Else
       Sql = Sql & " NOT NULL"
    End If
    
    Call Adapter.CreateCommand.ExecuteNonQuery(Sql)
    Set IAdapter_AddField = Me

End Function

Private Function IAdapter_CreateSchemaFile( _
                                                ByVal Table As String, _
                                                ByVal File As String, _
                                                ByVal Delimiter As String, _
                                                Optional ByVal TextDelimiter As String = "none", _
                                                Optional ByVal DecimalSymbol As String = ".", _
                                                Optional ByVal Columnheader As Boolean = True) As Boolean
    
    Dim iFF As Integer
    Dim bReturn As Boolean
    Dim FileName As String
    Dim Filefolder As String
    Dim Counter As Integer
    
    Dim Reader As IDataReader

    On Error GoTo Err_handler
     With New scripting.FileSystemObject
        If .FileExists(File) Then
            FileName = .GetFileName(File)
            Filefolder = .GetParentFolderName(File)
            If .FileExists(Filefolder & "\schema.ini") Then .DeleteFile Filefolder & "\schema.ini"

        Else
            Call Err.Raise(1#)

        End If
     End With

     iFF = FreeFile()
     Open Filefolder & "\schema.ini" For Output Access Write As #iFF

     Print #iFF, "[" & FileName & "]"
     Print #iFF, "CharacterSet = ANSI"
     Print #iFF, "DecimalSymbol=" & DecimalSymbol
     Print #iFF, "Format=Delimited(" & Delimiter & ")"
     Print #iFF, "TextDelimiter=" & TextDelimiter
     Print #iFF, "DateFormat=DD.MM.YYYY"
     Print #iFF, "DateTimeFormat =DD.MM.YYYY"
     Print #iFF, "ColNameHeader = " & IIf(Columnheader, "True", "False")
     Print #iFF, "MaxScanRows=1"
     
     Set Reader = Adapter.TableStructure(Table)
     Do While Reader.Read
        If CInt(Reader("COLUMN_FLAGS")) = 90 And CInt(Reader("DATA_TYPE")) = 3 Then
            Counter = Counter + 1
        End If
     Loop
     
     Set Reader = Adapter.TableStructure(Table)
     Do While Reader.Read
         Select Case CInt(Reader("DATA_TYPE"))
             Case ICommandDataTypeEnum.adBoolean  'dbBoolean
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Bit"
           
             Case ICommandDataTypeEnum.adUnsignedTinyInt 'dbByte
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Byte"
           
             Case ICommandDataTypeEnum.adSmallInt 'dbInteger
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Short"
           
             Case ICommandDataTypeEnum.adInteger 'dbLong
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Integer"
           
             Case ICommandDataTypeEnum.adCurrency 'dbCurrency
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Currency"
           
             Case ICommandDataTypeEnum.adSingle 'dbSingle
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Single"
           
             Case ICommandDataTypeEnum.adDouble 'dbDouble
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Double"
           
             Case ICommandDataTypeEnum.adDate 'dbDate
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Date"
           
             Case ICommandDataTypeEnum.adWChar, ICommandDataTypeEnum.adVarWChar 'dbText
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Char Width " & Reader("CHARACTER_MAXIMUM_LENGTH")
           
             Case ICommandDataTypeEnum.adLongVarBinary 'dbLongBinary
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "OLE"
           
             Case ICommandDataTypeEnum.adLongVarWChar 'dbMemo
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "LongChar"
           
             Case ICommandDataTypeEnum.adGUID 'dbGUID
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Char Width 16"
           
             Case Else
                 Print #iFF, "Col" & CInt(Reader("ORDINAL_POSITION")) - Counter & "=""" & Reader("COLUMN_NAME") & """" & Space$(1) & "Unknown DataType"
         End Select
     Loop
    On Error GoTo 0

    bReturn = True

CleanExit:

    IAdapter_CreateSchemaFile = bReturn
    Close iFF
    Exit Function

Err_handler:

    Print #iFF, "Some Error occurs"
    Print #iFF, "#(" & Err.Number & ")"
    Print #iFF, Err.Description
    Call ErrorHandler("Private Function IAdapter_CreateSchemaFile")
    bReturn = False: Resume CleanExit
    
End Function

Private Sub IAdapter_Dispose()
    myClass.Adapter.Source.Disconnect
    myClass.Adapter = BlankAdapter
    myClass.Event = BlankTransaction
    Call PrintMessage("Private Sub IAdapter_Dispose")
End Sub
Private Function BlankAdapter() As AdapterMember: End Function

Private Sub IAdapter_Drop(ByVal Table As String, Optional ByVal FieldName As String = vbNullString)
    
    Dim Sql As String
    
    If Not FieldName = vbNullString Then
        Sql = "ALTER TABLE [" & Table & "] DROP COLUMN [" & FieldName & "]"
    Else
        Sql = "DROP TABLE " & Table
    End If
        
    Call Adapter.CreateCommand.ExecuteNonQuery(Sql)
    
End Sub

Private Function IAdapter_ExternalExport(ByVal Source As IConnection, ByVal Sql As String, Optional ByVal Table As String = "tblExport", Optional ByVal ExecuteDropBefore As Boolean = True) As Boolean
    
    Dim Cancel As Boolean

    With Source
        If .Status = Connection_initialized Then
            If ExecuteDropBefore Then Call .CreateCommand.ExecuteNonQuery("DROP TABLE [" & Table & "]")
            With myClass.Event.Bridge
                Call .RaiseBeforeExport(Table, Source, Cancel)
                If Not Cancel Then Call .RaiseAfterExport(Table, Source, Adapter.CreateCommand.ExecuteNonQuery(ExportQuery(Source, Sql, Table)), (Not .IsError))
                IAdapter_ExternalExport = Not .IsError
            End With
         End If
     End With
     
End Function

Private Function ExportQuery(ByVal Source As IConnection, ByVal Sql As String, ByVal Table As String) As String
    
    Dim vCommandParts As Variant
    Dim External As Variant
     
    vCommandParts = Split(Sql, " FROM ", 2, vbTextCompare)
     
    If Source.CreateAdapter.TableExists(Table) Then
        vCommandParts(0) = "INSERT INTO " & Source.ExternalLink(Table) & Space$(1) & vCommandParts(0)
    Else
        vCommandParts(0) = vCommandParts(0) & " INTO " & Source.ExternalLink(Table)
    End If
    
    ExportQuery = Join$(vCommandParts, " FROM ")
    
End Function

Private Function IAdapter_ExternalImport(ByVal Source As IConnection, ByVal Sql As String, ByVal Table As String) As Boolean

    Dim Cancel As Boolean

    With myClass.Event.Bridge
        Call .RaiseBeforeImport(Table, Source, Cancel)
        If Not Cancel Then Call .RaiseAfterImport(Table, Source, Adapter.CreateCommand.ExecuteNonQuery(ImportQuery(Source, Sql, Table)), (Not .IsError))
        IAdapter_ExternalImport = Not .IsError
    End With
    
End Function

Private Function ImportQuery(ByVal Source As IConnection, ByVal Sql As String, ByVal Table As String) As String
    
    Dim vCommandParts As Variant
    Dim External As Variant

    vCommandParts = Split(Sql, " FROM ", 2, vbTextCompare)
    External = Split(vCommandParts(1), " WHERE ", 2)
    External(0) = Source.ExternalLink(External(0))
    vCommandParts(1) = Join$(External, " ")
    
    If Adapter.TableExists(Table) Then
        vCommandParts(0) = "INSERT INTO " & Table & Space$(1) & vCommandParts(0)
    Else
        vCommandParts(0) = vCommandParts(0) & " INTO " & Table
    End If
    
    ImportQuery = Join$(vCommandParts, " FROM ")
    
End Function

Private Function IAdapter_HasExclusiveAccess() As Boolean
    
    Dim bReturn As Boolean
    
    On Error GoTo ErrorHandler
    If Adapter.Source.Status = Connection_initialized Then
        With CreateObject("ADODB.Connection")
            .Mode = 12& ' adModeShareExclusive '
            .Open "Provider=" & Adapter.Source.Provider & ";Data Source=" & Adapter.Source.DataSource & ";"
            bReturn = True
            .Close
        End With
    End If
    On Error GoTo 0

CleanExit:

    On Error Resume Next
    IAdapter_HasExclusiveAccess = bReturn
    Exit Function

ErrorHandler:
    
    bReturn = False
    Resume CleanExit
    
End Function

Private Function IAdapter_TableMapping(TableColumns() As Variant) As Variant
    
    Dim Table As Variant
    Dim Column As Variant
    Dim bReturn As Boolean
    
    On Error GoTo Err_handler
    With myClass.Adapter
        If .Tablelist Is Nothing Then
            Set .Tablelist = New scripting.Dictionary
            For Each Table In Adapter.GetTables
                Call .Tablelist.Add(Table, New scripting.Dictionary)
                For Each Column In Adapter.GetColumns(Table)
                   .Tablelist(Table)(Column) = 0
                Next
            Next
        End If
        
        For Each Table In .Tablelist
            For Each Column In TableColumns
                If Not Column = vbNullString Then
                    If .Tablelist(Table).Exists(Column) Then
                        bReturn = True
                    Else
                        bReturn = False: Exit For
                    End If
                End If
           Next Column
           If bReturn Then Exit For
        Next Table
        
        If bReturn Then
            IAdapter_TableMapping = Table
        Else
            IAdapter_TableMapping = vbNullString
        End If
    End With
    On Error GoTo 0
    
    Exit Function

Err_handler:
    
    Call ErrorHandler("Private Function IAdapter_TableMapping")
    
End Function

'DataRow
Private Function BlankDataTable() As DataTableMember: End Function
Private Function CreateDeleteQuery(ByVal Table As String) As IQuery

    Dim Sql As String
    
    Dim Cmd As ICommand
    Set Cmd = Adapter.CreateCommand
    
    With myClass.DataTable.Column(0)
        Call Cmd.AddParameter(myClass.DataTable.ID, ConvertValue(.DataType, String(.Size, "0")))
    End With
    
    Sql = "DELETE * FROM " & Table & " WHERE " & myClass.DataTable.ID & "=?"
    
    Set CreateDeleteQuery = Cmd.Prepare(Sql, Prepared, "DeleteCommand")
    Exit Function

End Function

Private Function CreateInsertQuery(ByVal Table As String) As IQuery
    
    
    Dim Sql As String
    Dim i As Long
   
    Dim Cmd As ICommand
    Set Cmd = Adapter.CreateCommand
    
    With myClass.DataTable
        ReDim Fields(LBound(.Column) + 1 To UBound(.Column))
        ReDim Values(LBound(.Column) + 1 To UBound(.Column))
        
        For i = LBound(.Column) + 1 To UBound(.Column)
            With .Column(i)
                Call Cmd.AddParameter(.Name, ConvertValue(.DataType, String(.Size, "0")))
                Fields(i) = .Name & "=P" & i
                Values(i) = "P" & i
            End With
        Next i
    End With
    
    Sql = "INSERT INTO " & Table & " (" & Join$(Fields, ",") & ") VALUES (" & Join$(Values, ",") & ")"
    
    Set CreateInsertQuery = Cmd.Prepare(Sql, Prepared, "InsertCommand")
    Exit Function

End Function

Private Function CreateUpdateQuery(ByVal Table As String) As IQuery

    Dim Sql As String
    Dim i As Long
   
    Dim Cmd As ICommand
    Set Cmd = Adapter.CreateCommand
    
    With myClass.DataTable
        ReDim Fields(LBound(.Column) + 1 To UBound(.Column))
        
        For i = LBound(.Column) + 1 To UBound(.Column)
            With .Column(i)
                Call Cmd.AddParameter(.Name, ConvertValue(.DataType, String(.Size, "0")))
                Fields(i) = .Name & "=P" & i
            End With
        Next i
        
        With .Column(0)
            Call Cmd.AddParameter(myClass.DataTable.ID, ConvertValue(.DataType, String(.Size, "0")))
        End With
        
    End With
    
    Sql = "UPDATE " & Table & " SET " & Join$(Fields, ",") & " WHERE " & myClass.DataTable.ID & "=?"
    
    Set CreateUpdateQuery = Cmd.Prepare(Sql, Prepared, "UpdateCommand")
    Exit Function

End Function

Private Function BlankDataRow() As DataRowMember: End Function

'IQueryList
Private Property Get IQueryList_Self() As IQueryList
    Set IQueryList_Self = Me
End Property

Private Property Get IQueryList_Name() As String
    IQueryList_Name = myClass.Querylist.Name
End Property

Private Property Get IQueryList_NewEnum() As stdole.IUnknown
    Set IQueryList_NewEnum = myClass.Querylist.Collection.[_NewEnum]
End Property

Private Property Get IQueryList_Query(ByVal Index As Variant) As IQuery
    Set IQueryList_Query = myClass.Querylist.Collection(Index)
End Property

Private Property Get IQueryList_QueryList() As VBA.Collection
    Set IQueryList_QueryList = myClass.Querylist.Collection
End Property

Private Function IQueryList_Add(ByVal Query As IQuery) As IQuery
    
    If KeyExists(Query.Name) Then
        Call myClass.Querylist.Collection.Remove(Query.Name)
    End If
   
    Call myClass.Querylist.Collection.Add(Item:=Query, Key:=Query.Name)
    
    Set IQueryList_Add = Query
    
End Function

Private Property Get IQueryList_Count() As Long
    IQueryList_Count = myClass.Querylist.Collection.Count
End Property

Private Sub IQueryList_Dispose()
    myClass.Querylist = BlankQueryList
    Call PrintMessage("Private Sub IQueryList_Dispose")
End Sub
Private Function BlankQueryList() As QueryListMember: End Function

Private Function IQueryList_Exists(ByVal Key As String) As Boolean
    IQueryList_Exists = KeyExists(Key)
End Function

Private Function KeyExists(Key As String) As Boolean
    Dim oItem As Object
    
    Set oItem = Nothing: On Error Resume Next
    Set oItem = myClass.Querylist.Collection(Key): On Error GoTo 0
    
    KeyExists = Not CBool(oItem Is Nothing)

End Function

Private Sub IQueryList_Remove(Index As Variant)
    On Error Resume Next
    myClass.Querylist.Collection.Remove Index
    On Error GoTo 0
End Sub

Private Function BlankTransaction() As TransactionMember: End Function

Private Sub ITransaction_Dispose()
    myClass.Event = BlankTransaction
End Sub

Private Property Get ITransaction_Successful() As Boolean
    ITransaction_Successful = Not myClass.Event.Cancel
End Property

Private Property Get ITransaction_IsError() As Boolean
    ITransaction_IsError = myClass.Event.IsError
End Property

Private Sub ITransaction_RaiseAfterExport(ByVal Table As String, ByVal Source As IConnection, ByVal AffectedRecords As Long, ByVal Successfull As Boolean)
    RaiseEvent AfterImport(Table, Source, AffectedRecords, Successfull)
End Sub

Private Sub ITransaction_RaiseAfterImport(ByVal Table As String, ByVal Source As IConnection, ByVal AffectedRecords As Long, ByVal Successfull As Boolean)
    RaiseEvent AfterImport(Table, Source, AffectedRecords, Successfull)
End Sub

Private Sub ITransaction_RaiseBeforeExport(ByVal Table As String, ByVal Source As IConnection, Cancel As Boolean)
    RaiseEvent BeforeExport(Table, Source, Cancel)
End Sub

Private Sub ITransaction_RaiseBeforeImport(ByVal Table As String, ByVal Source As IConnection, Cancel As Boolean)
    RaiseEvent BeforeImport(Table, Source, Cancel)
End Sub

Private Sub ITransaction_RaiseCommitTrans(ByVal CommandText As String, ByVal State As ICommandEventEnum, Cancel As Boolean)
    RaiseEvent CommitTrans(CommandText, State, Cancel)
    myClass.Event.Cancel = Cancel
End Sub

Private Sub ITransaction_RaiseDataAvailable(ByVal Reader As IDataReader, Optional ByRef Cancel As Boolean = False)
    RaiseEvent DataAvailable(Reader, Cancel)
    myClass.Event.Cancel = Cancel
End Sub

Private Sub ITransaction_RaiseError( _
                                        ByVal CommandText As String, _
                                        ByVal Connectionstring As String, _
                                        ByVal Parameter As ADODB.Parameters, _
                                        ByVal ListOfError As ADODB.errors)
    myClass.Event.Cancel = True
    RaiseEvent Error(CommandText, Connectionstring, Parameter, ListOfError)
End Sub

Private Sub ITransaction_RaiseStatementCompleted(ByRef Disconnect As Boolean, ByVal CommandText As String, ByVal IsError As Boolean)
     RaiseEvent StatementCompleted(Disconnect, CommandText, IsError)
     myClass.Event.Cancel = IsError
End Sub

Private Sub PrintMessage(ByVal This As String)
    If myClass.GlobalDebug Then Debug.Print This
End Sub

'Error
Private Sub ErrorHandler(ByVal Proc As String)
    If myClass.GlobalDebug Then
        Call PrintMessage(vbCrLf & "Error " & TypeName$(Me) & vbCrLf & "Source: " & Proc & " :: (#" & Err.Number & ")" & vbCrLf & "Description: " & Err.Description)
    End If
End Sub

Private Function GetReferencedType(ByRef MyVar) As Long

    Const VT_BY_REF = &H4000&
    Dim iType As Integer

    Call CopyMemory(iType, MyVar, 2&)
    GetReferencedType = iType Mod VT_BY_REF

End Function

Private Sub ObjectToMember(ByVal Arg As Object)

    If Not Arg Is Nothing Then
        Select Case True
            Case TypeOf Arg Is ADODB.Command
                Set myClass.Command.Object = Arg
            
            Case TypeOf Arg Is ADODB.Connection
                Set myClass.Connection.Object = Arg
            
            Case TypeOf Arg Is ADODB.Recordset
                Set myClass.Reader.Data = Arg
            
            Case TypeOf Arg Is VBA.Collection
                Set myClass.Querylist.Collection = Arg
                
        End Select
    End If
    
End Sub

Private Function HasFlag(ByVal composite As Long, ByVal Flag As Long) As Boolean
    HasFlag = (composite And Flag) = Flag
End Function

Private Function CreateColumn(ByVal Field As ADODB.Field, ByVal Index As Long) As ColumnMember
    Dim C As ColumnMember
    With C
        .Name = Field.Name
        .Index = Index
        .DataType = Field.Type
        .Size = Field.DefinedSize
    End With
    CreateColumn = C
End Function

Private Sub vTableSwap(ByVal Ordinal As OverrideEnum)
    
    Const OffsetToVFT = &H1C + &H54
    
    Dim OffSetToFn As Long
    OffSetToFn = &H3C + (Ordinal - 1) * 4&
    
    Dim pVFT As Long
    Dim fnAddress As Long
    
    Call CopyMemory(pVFT, ByVal ObjPtr(Me), POINTERSIZE)
    Call CopyMemory(fnAddress, ByVal pVFT + OffsetToVFT + OffSetToFn, POINTERSIZE)
    
    Call CopyMemory(ByVal pVFT + OffsetToVFT + OffSetToFn, ByVal pVFT + OffsetToVFT, POINTERSIZE)
    Call CopyMemory(ByVal pVFT + OffsetToVFT, fnAddress, POINTERSIZE)
        
End Sub
