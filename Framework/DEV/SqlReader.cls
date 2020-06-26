VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

'Implements IDatarow

Private Type Member
    Data As Variant
    Fields As VBA.Collection
    'x As TString
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Private WithEvents Transaction As SqlTransaction
Attribute Transaction.VB_VarHelpID = -1

Public Function Build(ByVal Transaction As SqlTransaction) As SqlCommand
    
    Set Transaction = Transaction
    With myClass
        
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Transaction_Complete(ByVal Data As SqlReader)
    
    
    Set Transaction = Nothing
End Sub

Public Property Get HasFetched() As Boolean
    HasFetched = Transaction.WaitExecution
End Property

'Private Sub Build()
'
'    Dim Data As ADODB.Recordset
'    Dim f As ADODB.Field
'    Dim i As Long
'
'    Set myClass.Fields = New VBA.Collection
'    ReDim myClass.Names(1 To Data.Fields.Count)
'
'    For i = 0 To Data.Fields.Count - 1
'        With Data(i)
'            Set myClass.Names(i + 1) = TString(.Name)
'            Select Case .Type
'                Case adVarWChar
'                    Call myClass.Fields.Add(TString("", Function1, DefaultString), .Name)
'
'                Case adNumeric
'                    Call myClass.Fields.Add(TNumeric(0, DefaultNumber), .Name)
'
'                Case adDate
'                    Call myClass.Fields.Add(TDate(Now(), DefaultDate), .Name)
'
'                Case adBoolean
'                    Call myClass.Fields.Add(TBoolean(False, DefaulTBoolean), .Name)
'
'
'                Case Else
'
'            End Select
'        End With
'    Next
'
'End Sub

Private Sub IDatarow_Clear()

End Sub

Private Function IDatarow_Copy() As IDatarow
'    Dim c As Member
'
'    c.Names = myClass.Names
'    Set c.Fields = New VBA.Collection
'
'    Dim f As IDefaultConstructor
'    For Each f In myClass.Fields
'        c.Fields.Add f.Build()
'    Next
    
End Function
'
'Private Property Get IDatarow_Count() As Long
'
'End Property
'
'Private Property Get IDatarow_Field(ByVal Index As Variant) As IValue
'
'    Dim Value As Variant
'    Dim Field As IDefaultConstructor
'
'    Set Field = myClass.Fields(Index)
'
'    Set IDatarow_Field = Field.Build(Value)
'End Property
'
'Private Property Get IDatarow_Name(ByVal Index As Long) As String
'    Set IDatarow_Name = myClass.Names(Index).Value
'End Property
'
'Private Function IDatarow_Update(ByVal Index As Variant, NewValue As IValue) As IDatarow

'End Function
