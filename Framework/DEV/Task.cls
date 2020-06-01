VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder "Base"
Implements IObject
Implements IPrototype

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
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get HasError() As Boolean
    HasError = myClass.HasError
End Property

Public Function Build(ByVal Source As Object, ByVal Method As String) As Task
    Set myClass.Source = Source
    myClass.Method = Method
    Set Build = Skynet.Create(Me)
End Function

Public Sub Complete()
    RaiseEvent Complete
End Sub

Public Function Overflow() As Task
    myClass.Number = Overflow_
    myClass.Type = "Overflow"
    Set Overflow = Me
End Function

Public Function InvalidOperationException() As Task

End Function

Public Function OutOfRange() As Task
    myClass.Number = OutOfRange_
    myClass.Type = "OutOfRange"
    Set OutOfRange = Me
End Function

Public Function NotEqual(ByRef A As IObject, ByRef B As IObject) As Task
    myClass.Number = NotEqual_
    myClass.Message = "<" & Typename$(A) & "> !=  <" & Typename$(B) & ">"
    myClass.Type = "Not Equal"
    Set NotEqual = Me
End Function

Public Function Instantiation() As Task
    myClass.Number = PreDeclared_
    myClass.Type = "Instantiation Error"
    myClass.Message = "Invalid use of 'New' Keyword "
    Set Instantiation = Me
End Function

Public Function InvalidInput() As Task
    myClass.Number = Invalid_
    myClass.Type = "Invalid Input"
    Set InvalidInput = Me
End Function

Public Sub Raise(Optional ByVal Message As String)
    Dim raiseVBA As Boolean
    raiseVBA = True
    
    myClass.Message = myClass.Message & Message
    
    RaiseEvent Failed(raiseVBA)
    
    If raiseVBA Then _
        Call Err.Raise(myClass.Number, Typename$(myClass.Source), Skynet.IObject(Me).ToString)
        
End Sub

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
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

Private Function IObject_IsRelatedTo(Other As IObject) As Boolean

End Function

Private Function IObject_ToString() As String
    IObject_ToString = "Source :: " & Typename$(myClass.Source) _
                        & vbNewLine & "ErrorType :: " & myClass.Type _
                        & vbNewLine & "Errornumber :: (#" & myClass.Number & ")" _
                        & vbNewLine & "Procedure :: " & myClass.Method _
                        & vbNewLine & "Description :: " & myClass.Message
End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New Task: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



