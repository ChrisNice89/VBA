VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<Sql>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IObject
Implements IPrototype

Public Event WillExecute()
Public Event Prepare()
Public Event Dispose()

Private Type Member
    Connection As SqlConnection
    sql As TString
    Parameter As VBA.Collection
End Type

Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type

Private Virtual As Internal

Public Sub Execute()

    RaiseEvent WillExecute
End Sub

Public Property Get sql() As TString
    Set sql = myClass.sql
End Property

Public Property Get Parameter(ByVal Index As Variant) As SqlParameter
    Set Parameter = myClass.Parameter(Index)
End Property

Public Function AddParameter(ByRef Parameter As SqlParameter) As SqlCommand
    
    Call myClass.Parameter.Add(Parameter, Parameter.Name)
    Set AddParameter = Me

End Function

Public Function CreateParameter(ByRef Value As IValue, Optional ByVal Name As String = vbNullString) As SqlParameter

    Dim p As SqlParameter
    
    If Name = vbNullString Then _
       Name = "p" & myClass.Parameter.Count + 1
    
    Set p = SqlParameter.Build(Value, myClass.Parameter.Count + 1, Name)
    Call Me.AddParameter(p)
    
    Set CreateParameter = p
    
End Function

Public Sub Prepare()

    RaiseEvent Prepare
End Sub

Public Function Build(ByVal sql As String, ByRef Connection As SqlConnection) As SqlCommand
    With myClass
        Set .sql = TString(sql)
        Set .Connection = Connection
        Set .Parameter = New VBA.Collection
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

'TODO TMap IObject_Dispose
Private Sub IObject_Dispose()
    myClass = EmptyStructure
End Sub

Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Dim cmd As SqlCommand: Set cmd = Other
   

    IObject_Equals = False
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim cmd As SqlCommand
    Dim i As Long
    
    If Other.IsRelatedTo(Me) Then
        Set cmd = Other
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = 0: End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is SqlCommand): End Function

Private Function IObject_ToString() As String: IObject_ToString = "": End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New SqlCommand: End Property

Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


