VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Enum TableAEnum
    FieldA
    FieldB
End Enum

Public Enum TableBEnum
    FieldA
    FieldB
End Enum
Private TableA(2) As String
Private TableB(2) As String

Private Sub Class_Initialize()

    TableA(TableAEnum.FieldA) = "FieldA"
    TableA(TableAEnum.FieldB) = "FieldB"
    
    TableB(TableBEnum.FieldA) = "FieldA"
    TableB(TableBEnum.FieldB) = "FieldB"

End Sub

Public Function A(ByVal Field As TableAEnum) As String: A = TableA(Field): End Function
Public Function B(ByVal Field As TableAEnum) As String: B = TableB(Field): End Function
