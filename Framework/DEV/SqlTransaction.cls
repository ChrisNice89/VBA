VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Event Complete(ByVal Data As SqlReader)

Private Type Member
    Completed As Boolean
    Command As ADODB.Command
    Reader As SqlReader
    'Sql As TString
    Parameter As VBA.Collection
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(ByVal cmd As ADODB.Command) As SqlTransaction
    
    If cmd Is Nothing Then

    End If
    
    Set myClass.Command = cmd
End Function

Public Function WaitExecution() As Boolean
    If Not myClass.Command Is Nothing Then
        Do While (myClass.Command.State And ADODB.adStateExecuting) = ADODB.adStateExecuting
        Loop
        Set myClass.Command = Nothing
    End If
    WaitExecution = True
End Function

Public Sub TransactionComplete(ByVal Result As SqlReader)
   myClass.Completed = True
   RaiseEvent Complete(Result)
End Sub

