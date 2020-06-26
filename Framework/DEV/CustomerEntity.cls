VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IEntity

Private Type Member
    Datarow As IDatarow
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Private Function IEntity_Build(ByVal Datarow As IDatarow) As Object
    Set myClass.Datarow = Datarow
    Set IEntity_Build = Me
End Function

Private Property Get IEntity_Datarow() As IDatarow
    Set IEntity_Datarow = myClass.Datarow
End Property
    
Public Property Get Name() As IValue
    Set Name = myClass.Datarow(Table.A(TableAEnum.FieldA))
End Property

Public Property Get Adress() As IValue
    Set Adress = myClass.Datarow(Table.A(TableAEnum.FieldB))
End Property

