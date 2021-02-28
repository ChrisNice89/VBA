VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericThrowhelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder "Base"
Implements IGeneric

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
    IsNothing_
    DuplicateNotAllowed_
    NoSuchElement_
    InvalidIndex_
    InvalidIterator_
    BuildExeption_
End Enum

Private Type Member
    HasError As Boolean
    Source As Object
    Message As String
    Type As String
    Method As String
    Number As ErrorNumber
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get HasError() As Boolean: HasError = myclass.HasError: End Property
Public Property Get Source() As Object: Set Source = myclass.Source: End Property
Public Property Get ErrorNumber() As Long: ErrorNumber = myclass.Number: End Property
Public Property Get MethodName() As String: MethodName = myclass.Method: End Property

Public Function Build(ByVal Source As Object, ByVal Method As String) As GenericThrowhelper
    
    With myclass
        Set myclass.Source = Source
        myclass.Method = Method
    End With
   
    Set Build = System.CreateInstance(New GenericThrowhelper, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Sub Complete()
    RaiseEvent Assert
End Sub

Public Function BuildExeption(Optional ByVal Message As String = "Build called from already instantiated Object") As GenericThrowhelper
    myclass.Number = BuildExeption_
    myclass.Type = "BuildExeption"
    Call Me.Raise(Message)
    Set BuildExeption = Me
End Function

Public Function NoSuchElement(Optional ByVal Message As String = "Element not found in Collection") As GenericThrowhelper
    myclass.Number = NoSuchElement_
    myclass.Type = "NoSuchElement"
    Call Me.Raise(Message)
    Set NoSuchElement = Me
End Function

Public Function DuplicateNotAllowed(Optional ByVal Element As IGeneric) As GenericThrowhelper
    Dim Message As String
    
    myclass.Number = DuplicateNotAllowed_
    myclass.Type = "DuplicateNotAllowed"
    If Element Is Nothing Then
        Message = "Element is already in Collection"
    Else
        Message = Element.ToString & " is already in Collection"
    End If
    Call Me.Raise(Message)
    Set DuplicateNotAllowed = Me
End Function

Public Function IsNothing(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = IsNothing_
    myclass.Type = "IsNothing"
    Call Me.Raise(Message & " is nothing")
    Set IsNothing = Me
End Function

Public Function Overflow(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = Overflow_
    myclass.Type = "Overflow"
    Call Me.Raise(Message)
    Set Overflow = Me
End Function

Public Function InvalidIterator(Optional ByVal Message As String = "Iterator not loaded") As GenericThrowhelper
    myclass.Number = InvalidIterator_
    myclass.Type = "InvalidIterator"
    Call Me.Raise(Message)
    Set InvalidIterator = Me
End Function

Public Function InvalidOperation(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = InvalidOperation_
    myclass.Type = "InvalidOperation"
    Call Me.Raise(Message)
    Set InvalidOperation = Me
End Function

Public Function NoNegativ(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = NoNegativ_
    myclass.Type = "NoNegativ"
    Set NoNegativ = Me
End Function

Public Function NotSupported(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = NotSupported_
    myclass.Type = "NotSupported"
    Call Me.Raise(Message)
    Set NotSupported = Me
End Function

Public Function NotImplemented(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = NotImplemented_
    myclass.Type = "NotImplemented"
    Call Me.Raise(Message)
    Set NotImplemented = Me
End Function

Public Function IndexOutofRange(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = IndexOutofRange_
    myclass.Type = "IndexOutofRange"
    Call Me.Raise(Message)
    Set IndexOutofRange = Me
End Function

Public Function ArgOutOfRange(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = ArgOutOfRange_
    myclass.Type = "ArgOutOfRange"
    Call Me.Raise(Message)
    Set ArgOutOfRange = Me
End Function

Public Function NotEqual(ByRef A As IGeneric, ByRef B As IGeneric) As GenericThrowhelper
    myclass.Number = NotEqual_
    myclass.Message = "<" & TypeName$(A) & "> !=  <" & TypeName$(B) & ">"
    myclass.Type = "Not Equal"
    Set NotEqual = Me
End Function

Public Function Instantiation(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = PreDeclared_
    myclass.Type = "Instantiation Error"
    myclass.Message = "Invalid use of 'New' Keyword "
    Call Me.Raise(Message)
    Set Instantiation = Me
End Function

Public Function InvalidInput(Optional ByVal Message As String) As GenericThrowhelper
    myclass.Number = InvalidInput_
    myclass.Type = "Invalid Input"
    Call Me.Raise(Message)
    Set InvalidInput = Me
End Function

Public Sub Raise(Optional ByVal Message As String)
    Dim raiseVBA As Boolean
    raiseVBA = True
    
    myclass.Message = myclass.Message & Message
    
    RaiseEvent Failed(raiseVBA)
    
    If raiseVBA Then _
        Call Err.Raise(myclass.Number, TypeName$(myclass.Source), IGeneric_ToString)
        
End Sub

Private Sub Class_Terminate()
    On Error GoTo CleanFail
    If Err.Number <> 0 Then
        GoTo CleanFail
    End If
CleanExit:
    Exit Sub

CleanFail:
    Debug.Print IGeneric_ToString
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericThrowhelper: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericThrowhelper: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericThrowhelper):  End Property

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy
    Set IGeneric_Clone = System.CreateInstance(New GenericThrowhelper, VarPtr(c), Virtual.Size)

End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = "Source :: " & TypeName$(myclass.Source) _
                        & vbNewLine & "ErrorType :: " & myclass.Type _
                        & vbNewLine & "Errornumber :: (#" & myclass.Number & ")" _
                        & vbNewLine & "Procedure :: " & myclass.Method _
                        & vbNewLine & "Description :: " & myclass.Message
End Function
