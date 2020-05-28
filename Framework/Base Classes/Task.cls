VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Implements IObject

Public Event Failed(ByRef raiseGlobal As Boolean)
Public Event Complete()

Private Enum ErrorNumber
    Invalid_
    OutOfRange_
    Overflow_
    NotEqual_
    PreDeclared_
End Enum

Private Type Member
    HasError As Boolean
    Source As Object
    Message As String
    Type As String
    Method As String
    Number As ErrorNumber
End Type
Private m As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get HasError() As Boolean
    HasError = m.HasError
End Property

Public Function Build(ByVal Source As Object, ByVal Method As String) As Task
    Set m.Source = Source
    m.Method = Method
    Set Build = System.Create(New Task, Me, LenB(m))
End Function

Public Sub Complete()
    RaiseEvent Complete
End Sub

Public Function Overflow() As Task
    m.Number = Overflow_
    m.Type = "Overflow"
    Set Overflow = Me
End Function

Public Function InvalidOperationException() As Task

End Function

Public Function OutOfRange() As Task
    m.Number = OutOfRange_
    m.Type = "OutOfRange"
    Set OutOfRange = Me
End Function

Public Function NotEqual(ByRef a As IObject, ByRef b As IObject) As Task
    m.Number = NotEqual_
    m.Message = "<" & Typename$(a) & "> !=  <" & Typename$(b) & ">"
    m.Type = "Not Equal"
    Set NotEqual = Me
End Function

Public Function Instantiation() As Task
    m.Number = PreDeclared_
    m.Type = "PreDeclared Instantating"
    m.Message = "Invalid use of 'New' Keyword "
    Set Instantiation = Me
End Function

Public Function InvalidInput() As Task
    m.Number = Invalid_
    m.Type = "Invalid Input"
    Set InvalidInput = Me
End Function

Public Sub Raise(Optional ByVal Message As String)
    Dim raiseVBA As Boolean
    raiseVBA = True
    
    m.Message = m.Message & Message
    
    RaiseEvent Failed(raiseVBA)
    
    If raiseGlobal Then _
        Call Err.Raise(m.Number, Typename$(m.Source), System.IObject(Me).ToString)
        
End Sub

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(m)
End Sub

Private Function IObject_Clone() As IObject

End Function

Private Function IObject_CompareTo(Other As IObject) As ComparerEnum

End Function

Private Sub IObject_Dispose()

End Sub

Private Function IObject_Equals(Other As IObject) As Boolean

End Function

Private Property Get IObject_HashValue() As Long

End Property

Private Function IObject_ToString() As String
    IObject_ToString = "Source :: " & Typename$(m.Source) _
                        & vbNewLine & "ErrorType :: " & m.Type _
                        & vbNewLine & "Errornumber :: (#" & m.Number & ")" _
                        & vbNewLine & "Procedure :: " & m.Method _
                        & vbNewLine & "Description :: " & m.Message
End Function

Private Property Get IObject_VirtualPointer() As LongLong
    IObject_VirtualPointer = Virtual.Pointer
End Property
