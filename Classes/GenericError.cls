VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

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
    Source As Object
    Message As String
    Type As String
    Method As String
    Number As ErrorNumber
    Throw As Boolean
End Type
Private This As Member

Public Property Get Source() As Object: Set Source = This.Source: End Property
Public Property Get ErrorNumber() As Long: ErrorNumber = This.Number: End Property
Public Property Get MethodName() As String: MethodName = This.Method: End Property
Public Property Get ErrorType() As String: ErrorType = This.Type: End Property
Public Property Get Throw() As Boolean: Throw = This.Throw: End Property
Public Property Get Message() As String

    Message = "Source :: " & TypeName$(This.Source) _
                            & vbNewLine & "ErrorType :: " & This.Type _
                            & vbNewLine & "Errornumber :: (#" & This.Number & ")" _
                            & vbNewLine & "Procedure :: " & This.Method _
                            & vbNewLine & "Description :: " & This.Message

End Property

Public Function Build(ByVal Source As Object, ByVal Method As String, Optional ByVal Throw As Boolean = True) As GenericError
    
    With This
        Set This.Source = Source
        This.Method = Method
        This.Throw = Throw
    End With
   
    Set Build = New GenericError
    
End Function

Private Sub Class_Initialize()
    
    With GenericError
        Set This.Source = .Source
        This.Method = .MethodName
        This.Throw = .Throw
    End With
    
End Sub

Public Sub Complete()
    RaiseEvent Assert
End Sub

Public Function BuildExeption(Optional ByVal Message As String = "Build called from already instantiated Object") As GenericError
    This.Number = BuildExeption_
    This.Type = "BuildExeption"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set BuildExeption = Me
    
End Function

Public Function NoSuchElement(Optional ByVal Message As String = "Element not found in Collection") As GenericError
    This.Number = NoSuchElement_
    This.Type = "NoSuchElement"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set NoSuchElement = Me
End Function

Public Function DuplicateNotAllowed(Optional ByVal Element As Object) As GenericError
    Dim Message As String
    
    This.Number = DuplicateNotAllowed_
    This.Type = "DuplicateNotAllowed"
    Message = "Element is already in Collection"
    
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set DuplicateNotAllowed = Me
End Function

Public Function IsNothing(Optional ByVal Name As String) As GenericError
    This.Number = IsNothing_
    This.Type = "IsNothing"
    This.Message = Name & " is nothing"
    
    If This.Throw Then _
        Call Me.Raise
        
    Set IsNothing = Me
End Function

Public Function Overflow(Optional ByVal Message As String) As GenericError
    This.Number = Overflow_
    This.Type = "Overflow"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set Overflow = Me
End Function

Public Function InvalidIterator(Optional ByVal Message As String = "Iterator not loaded") As GenericError
    This.Number = InvalidIterator_
    This.Type = "InvalidIterator"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set InvalidIterator = Me
End Function

Public Function InvalidOperation(Optional ByVal Message As String) As GenericError
    This.Number = InvalidOperation_
    This.Type = "InvalidOperation"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set InvalidOperation = Me
End Function

Public Function NoNegativ(Optional ByVal Message As String) As GenericError
    This.Number = NoNegativ_
    This.Type = "NoNegativ"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set NoNegativ = Me
End Function

Public Function NotSupported(Optional ByVal Message As String) As GenericError
    This.Number = NotSupported_
    This.Type = "NotSupported"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set NotSupported = Me
End Function

Public Function NotImplemented(Optional ByVal Message As String = "Not implemented yet") As GenericError
    This.Number = NotImplemented_
    This.Type = "NotImplemented yet"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set NotImplemented = Me
End Function

Public Function IndexOutofRange(Optional ByVal Message As String) As GenericError
    This.Number = IndexOutofRange_
    This.Type = "IndexOutofRange"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set IndexOutofRange = Me
End Function

Public Function ArgOutOfRange(Optional ByVal Message As String) As GenericError
    This.Number = ArgOutOfRange_
    This.Type = "ArgOutOfRange"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set ArgOutOfRange = Me
End Function

Public Function NotEqual(ByRef A As Object, ByRef B As Object) As GenericError
    This.Number = NotEqual_
    This.Message = "<" & VBA.TypeName$(A) & "> !=  <" & VBA.TypeName$(B) & ">"
    This.Type = "Not Equal"
    
    If This.Throw Then _
        Call Me.Raise
        
    Set NotEqual = Me
End Function

Public Function Instantiation(Optional ByVal Message As String) As GenericError
    This.Number = PreDeclared_
    This.Type = "Instantiation Error"
    This.Message = "Invalid use of 'New' Keyword "
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set Instantiation = Me
End Function

Public Function InvalidInput(Optional ByVal Message As String) As GenericError
    This.Number = InvalidInput_
    This.Type = "Invalid Input"
    This.Message = Message
    
    If This.Throw Then _
        Call Me.Raise
        
    Set InvalidInput = Me
End Function

Public Sub Raise()
    
    Dim raiseVBA As Boolean
    raiseVBA = True
    
    RaiseEvent Failed(raiseVBA)
    
    If raiseVBA Then _
        Call Err.Raise(This.Number, TypeName$(This.Source), IGeneric_ToString)
        
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericError: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericError: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = This.Number: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericError):  End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim OtherError As GenericError
    
    If (Other Is Nothing) Then _
        Exit Function
    
    If (Other Is Me) Then _
        IGeneric_Equals = True: Exit Function
        
    If Not TypeOf Other Is GenericError Then _
        Exit Function
    Set OtherError = Other
    
    If Not (OtherError.ErrorNumber = This.Number) Then _
        Exit Function
    
    If Not (OtherError.Source Is This.Source) Then _
        Exit Function
    
    If Not (OtherError.Message = Me.Message) Then _
        Exit Function
        
    IGeneric_Equals = True
    
End Function

Private Function IGeneric_ToString() As String: IGeneric_ToString = Me.Message: End Function


