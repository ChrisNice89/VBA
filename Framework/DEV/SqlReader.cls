VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IDatarow

Private Sub IDatarow_Clear()

End Sub

Private Function IDatarow_Copy() As IDatarow

End Function

Private Property Get IDatarow_Count() As Long

End Property

Private Property Get IDatarow_Field(ByVal Index As Variant) As IValue

End Property

Private Property Get IDatarow_Name(ByVal Index As Long) As String

End Property

Private Function IDatarow_Update(ByVal Index As Variant, NewValue As IValue) As IDatarow

End Function
