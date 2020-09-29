VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Throwhelper"
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
    KeyIsNothing_
    DuplicateNotAllowed_
    NoSuchElement_
    InvalidIndex_
    
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
    size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get HasError() As Boolean: HasError = myClass.HasError: End Property
Public Property Get Source() As Object: Set Source = myClass.Source: End Property
Public Property Get ErrorNumber() As Long: ErrorNumber = myClass.Number: End Property
Public Property Get MethodName() As String: MethodName = myClass.Method: End Property

'@DefaultMember
Public Function Build(ByVal Source As Object, ByVal Method As String) As Throwhelper
    
    With myClass
        Set myClass.Source = Source
        myClass.Method = Method
    End With
   
    Set Build = Skynet.CreateInstance(New Throwhelper, Virtual.Pointer, Virtual.size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.size = LenB(myClass)
End Sub

Public Sub Complete()
    RaiseEvent Assert
End Sub

Public Function NoSuchElement(Optional ByVal Message As String = "Element not found in Collection") As Throwhelper
    myClass.Number = NoSuchElement_
    myClass.Type = "NoSuchElement"
    Call Me.Raise(Message)
    Set NoSuchElement = Me
End Function

Public Function DuplicateNotAllowed(Optional ByVal Element As IGeneric) As Throwhelper
    Dim Message As String
    
    myClass.Number = DuplicateNotAllowed_
    myClass.Type = "DuplicateNotAllowed"
    If Element Is Nothing Then
        Message = "Element is already in Collection"
    Else
        Message = Element.ToString & " is already in Collection"
    End If
    Call Me.Raise(Message)
    Set DuplicateNotAllowed = Me
End Function

Public Function KeyIsNothing(Optional ByVal Message As String = "Key is nothing") As Throwhelper
    myClass.Number = KeyIsNothing_
    myClass.Type = "Key is nothing"
    Call Me.Raise(Message)
    Set KeyIsNothing = Me
End Function

Public Function Overflow(Optional ByVal Message As String) As Throwhelper
    myClass.Number = Overflow_
    myClass.Type = "Overflow"
    Call Me.Raise(Message)
    Set Overflow = Me
End Function

Public Function InvalidOperation(Optional ByVal Message As String) As Throwhelper
    myClass.Number = InvalidOperation_
    myClass.Type = "InvalidOperation"
    Call Me.Raise(Message)
    Set InvalidOperation = Me
End Function

Public Function NoNegativ(Optional ByVal Message As String) As Throwhelper
    myClass.Number = NoNegativ_
    myClass.Type = "NoNegativ"
    Set NoNegativ = Me
End Function

Public Function NullPointer(Optional ByVal Message As String) As Throwhelper
    myClass.Number = NullPointer_
    myClass.Type = "NullPointer"
    Call Me.Raise(Message)
    Set NullPointer = Me
End Function

Public Function NotSupported(Optional ByVal Message As String) As Throwhelper
    myClass.Number = NotSupported_
    myClass.Type = "NotSupported"
    Call Me.Raise(Message)
    Set NotSupported = Me
End Function

Public Function NotImplemented(Optional ByVal Message As String) As Throwhelper
    myClass.Number = NotImplemented_
    myClass.Type = "NotImplemented"
    Call Me.Raise(Message)
    Set NotImplemented = Me
End Function

Public Function IndexOutofRange(Optional ByVal Message As String) As Throwhelper
    myClass.Number = IndexOutofRange_
    myClass.Type = "IndexOutofRange"
    Set IndexOutofRange = Me
End Function

Public Function ArgOutOfRange(Optional ByVal Message As String) As Throwhelper
    myClass.Number = ArgOutOfRange_
    myClass.Type = "ArgOutOfRange"
    Call Me.Raise(Message)
    Set ArgOutOfRange = Me
End Function

Public Function NotEqual(ByRef a As IGeneric, ByRef b As IGeneric) As Throwhelper
    myClass.Number = NotEqual_
    myClass.Message = "<" & TypeName$(a) & "> !=  <" & TypeName$(b) & ">"
    myClass.Type = "Not Equal"
    Set NotEqual = Me
End Function

Public Function Instantiation(Optional ByVal Message As String) As Throwhelper
    myClass.Number = PreDeclared_
    myClass.Type = "Instantiation Error"
    myClass.Message = "Invalid use of 'New' Keyword "
    Call Me.Raise(Message)
    Set Instantiation = Me
End Function

Public Function InvalidInput(Optional ByVal Message As String) As Throwhelper
    myClass.Number = InvalidInput_
    myClass.Type = "Invalid Input"
    Call Me.Raise(Message)
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
    
    Dim C As Member: C = myClass 'hardcopy
    Set IGeneric_Clone = Skynet.CreateInstance(New Throwhelper, VarPtr(C), Virtual.size)

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

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New Throwhelper: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = Throwhelper: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

