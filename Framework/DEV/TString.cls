VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "<T>Class"
Option Explicit

Implements IObject
Implements IPrototype

Private Type Member
    Value As String
    HashFunction As Hashing
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(ByVal Value As String, Optional ByVal HashFunction As Hashing = Hashing.Function1) As TString
    myClass.Value = Value
    myClass.HashFunction = HashFunction
    Set Build = Skynet.Create(TString)
End Function

Public Property Get Value() As String
    Value = myClass.Value
End Property

Public Function ToByteArray() As Byte()
    ToByteArray = StrConv(myClass.Value, vbFromUnicode)
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
    
    Dim s As TString

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Set s = Other
    
    IObject_Equals = (StrComp(s.Value, myClass.Value, vbBinaryCompare) = 0)
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim s As TString
    If Other.IsRelatedTo(Me) Then
        Set s = Other
        Select Case StrComp(myClass.Value, s.Value, vbBinaryCompare)
            Case 1: IObject_CompareTo = IsGreater '=  1     (String1 and String2 are not equal; String1 is greater than String2
            Case -1: IObject_CompareTo = IsLower '= -1     (String1 and String2 are not equal; String1 is less than String2)
            Case Else: IObject_CompareTo = IsEqual '=  0     (String1 and String2 are equal)
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = Skynet.HashValue(StrConv(Skynet.IObject(Me).ToString, vbFromUnicode), Function1): End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is TString): End Function

Private Function IObject_ToString() As String
    With myClass
        IObject_ToString = .Value
    End With
End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TString: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongLong: IPrototype_VirtualPointer = Virtual.Pointer: End Property

