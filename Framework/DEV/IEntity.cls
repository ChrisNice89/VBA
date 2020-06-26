VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Function Build(ByVal Datarow As IDatarow) As Object: End Function
Public Property Get Datarow() As IDatarow: End Property

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub


