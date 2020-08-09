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
Implements IGeneric
Implements IPrototype

Public Event Failed(ByRef raiseGlobal As Boolean)
Public Event Assert()

Public Enum ErrorNumber
    InvalidInput_ = 1
    ArgOutOfRange_
    IndexOutofRange_
    Overflow_
    NotEqual_
    PreDeclared_
    InvalidOperation_
    NotImplemented_
    NotSupported_
    NoNegativ_
    NullPointer_
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

Public Property Get HasError() As Boolean: HasError = myClass.HasError: End Property
Public Property Get Source() As Object: Set Source = myClass.Source: End Property
Public Property Get ErrorNumber() As Long: ErrorNumber = myClass.Number: End Property
Public Property Get MethodName() As String: MethodName = myClass.Method: End Property

'@DefaultMember
Public Function Build(ByVal Source As Object, ByVal Method As String) As Task
Attribute Build.VB_UserMemId = 0
    
    With myClass
        Set myClass.Source = Source
        myClass.Method = Method
    End With
   
    Set Build = New Task
    Call Skynet.Generic(Task).Dispose
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    If Not Me Is Task Then
        With myClass
            Set .Source = Task.Source
            .Method = Task.MethodName
        End With
    End If
End Sub

Public Sub Complete()
    RaiseEvent Assert
End Sub

Public Sub RaiseIfNothing(ByRef o As Object)
    If (o Is Nothing) Then _
        Call Me.NullPointer.Raise

End Sub

Public Function Overflow() As Task
    myClass.Number = Overflow_
    myClass.Type = "Overflow"
    Set Overflow = Me
End Function

Public Function InvalidOperation() As Task
    myClass.Number = InvalidOperation_
    myClass.Type = "InvalidOperation"
    Set InvalidOperation = Me
End Function

Public Function NoNegativ() As Task
    myClass.Number = NoNegativ_
    myClass.Type = "NoNegativ"
    Set NoNegativ = Me
End Function

Public Function NullPointer() As Task
    myClass.Number = NullPointer_
    myClass.Type = "NullPointer"
    Set NullPointer = Me
End Function

Public Function NotSupported() As Task
    myClass.Number = NotSupported_
    myClass.Type = "NotSupported"
    Set NotSupported = Me
End Function

Public Function NotImplemented() As Task
    myClass.Number = NotImplemented_
    myClass.Type = "NotImplemented"
    Set NotImplemented = Me
End Function

Public Function IndexOutofRange() As Task
    myClass.Number = IndexOutofRange_
    myClass.Type = "IndexOutofRange"
    Set IndexOutofRange = Me
End Function

Public Function ArgOutOfRange() As Task
    myClass.Number = ArgOutOfRange_
    myClass.Type = "ArgOutOfRange"
    Set ArgOutOfRange = Me
End Function

Public Function NotEqual(ByRef a As IGeneric, ByRef b As IGeneric) As Task
    myClass.Number = NotEqual_
    myClass.Message = "<" & TypeName$(a) & "> !=  <" & TypeName$(b) & ">"
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
    myClass.Number = InvalidInput_
    myClass.Type = "Invalid Input"
    Set InvalidInput = Me
End Function

Public Sub Raise(Optional ByVal Message As String)
    Dim raiseVBA As Boolean
    raiseVBA = True
    
    myClass.Message = myClass.Message & Message
    
    RaiseEvent Failed(raiseVBA)
    
    If raiseVBA Then _
        Call Err.Raise(myClass.Number, TypeName$(myClass.Source), Skynet.Generic(Me).ToString)
        
End Sub

Private Sub Class_Terminate()
    On Error GoTo CleanFail
    If Err.Number <> 0 Then
        GoTo CleanFail
    End If
CleanExit:
    Exit Sub

CleanFail:
    Debug.Print Skynet.Generic(Me).ToString
End Sub

Private Function IGeneric_Clone() As IGeneric
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum

End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

End Function

Private Property Get IGeneric_HashValue() As Long

End Property

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean

End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = "Source :: " & TypeName$(myClass.Source) _
                        & vbNewLine & "ErrorType :: " & myClass.Type _
                        & vbNewLine & "Errornumber :: (#" & myClass.Number & ")" _
                        & vbNewLine & "Procedure :: " & myClass.Method _
                        & vbNewLine & "Description :: " & myClass.Message
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New Task: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = Task: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property
