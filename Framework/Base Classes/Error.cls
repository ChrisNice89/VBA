VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Implements IObject

Private Enum ErrorNumber
    Invalid_
    OutOfRange_
    Overflow_
    NotEqual_
End Enum

Private Type Member
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

Public Function Build(ByVal Source As Object, ByVal Method As String, ByVal Message As String) As Error
    Set m.Source = Source
    m.Method = Method
    m.Message = Message
    Set Build = System.Create(New Error, Me, LenB(m))
End Function

Public Function Overflow() As Error
    m.Number = Overflow_
    m.Type = "Overflow"
    Set Overflow = Me
End Function

Public Function OutOfRange() As Error
    m.Number = OutOfRange_
    m.Type = "OutOfRange"
    Set OutOfRange = Me
End Function

Public Function NotEqual(ByRef a As IObject, ByRef b As IObject) As Error
    m.Number = NotEqual_
    m.Message = "<" & Typename$(a) & "> !=  <" & Typename$(b) & ">"
    m.Type = "Not Equal"
    Set NotEqual = Me
End Function

Public Function InvalidInput() As Error
    m.Number = Invalid_
    m.Type = "Invalid Input"
    Set InvalidInput = Me
End Function

Public Sub Raise()
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
