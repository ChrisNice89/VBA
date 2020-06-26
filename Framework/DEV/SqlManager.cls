VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Private Type Member
    Connection As SqlConnection
    Sql As TString
    Transactions As TMap
End Type

Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal
Public Event BeginTransaction(ByVal Name As String)

Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1

Public Function Prepare(ByVal Command As SqlCommand) As ADODB.Command

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim i As Long
    
    With Command
        cmd.CommandType = adCmdText
        cmd.CommandText = .Sql.Value
        For i = 1 To .CountParameter
            Call cmd.Parameters.Append(.Parameter(i).Parameter)
        Next
    End With
    Set Prepare = cmd
End Function

Public Sub ExecuteAsync(ByVal Command As SqlCommand)
    
    Call myClass.Transactions.Add(TNumeric(ObjPtr(Prepare(Command))), Command.Transaction)
    
End Sub

Private Sub AdoConnection_ExecuteComplete( _
                                            ByVal RecordsAffected As Long, _
                                            ByVal pError As ADODB.Error, _
                                            adStatus As ADODB.EventStatusEnum, _
                                            ByVal pCommand As ADODB.Command, _
                                            ByVal pRecordset As ADODB.Recordset, _
                                            ByVal pConnection As ADODB.Connection)

    Dim T As SqlTransaction
    If Not myClass.Transactions.TryGetValue(TNumeric(ObjPtr(pCommand)), T) Then
        'Error
    End If

    
End Sub

Private Sub Class_Initialize()
    Set myClass.Transactions = TMap.Build '(TNumeric, SqlTransaction, 10000)
End Sub
