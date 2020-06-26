VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Private Type Member
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

Private Sub Class_Initialize()
    Set myClass.Transactions = TMap.Build '(TNumeric, SqlTransaction, 10000)
End Sub

Public Function BeginTransaction(ByVal cmd As ADODB.Command) As SqlTransaction
    
    Dim T As SqlTransaction
    Set T = SqlTransaction.Build(cmd)
    Call myClass.Transactions.Add(T, T)
    
    'Execute
    
    Set BeginTransaction = T
    
End Function

Private Sub AdoConnection_ExecuteComplete( _
                                            ByVal RecordsAffected As Long, _
                                            ByVal pError As ADODB.Error, _
                                            adStatus As ADODB.EventStatusEnum, _
                                            ByVal pCommand As ADODB.Command, _
                                            ByVal pRecordset As ADODB.Recordset, _
                                            ByVal pConnection As ADODB.Connection)
    
    Dim T As SqlTransaction
    Set T = SqlTransaction.Build(pCommand)
    
    If Not myClass.Transactions.TryGetValue(T, T) Then

    Else
        Debug.Print "Private Sub AdoConnection_ExecuteComplete: Transaction not found"
    End If
    
    If Not pRecordset Is Nothing Then
        'Create Reader
    End If
    
    Call T.TransactionComplete(Nothing)
    Call myClass.Transactions.Remove(T)
    
End Sub

Private Function ValidateCommand(ByVal Command As SqlCommand) As Boolean

End Function

Private Function Prepare(ByVal Command As SqlCommand) As ADODB.Command

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
