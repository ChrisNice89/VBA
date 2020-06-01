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
    With myClass
        .Value = Value
        .HashFunction = HashFunction
    End With
    
    Set Build = Skynet.Create(TString)
End Function

Public Property Get Length() As String
    Length = Len(myClass.Value)
End Property

Public Property Get FirstChar() As String
    FirstChar = VBA.Mid$(myClass.Value, 1, 1)
End Property

Public Property Get LastChar() As String
    LastChar = VBA.Mid$(myClass.Value, Len(myClass.Value), 1)
End Property

Public Property Get Character(ByVal Index As Long) As String
    Character = VBA.Mid$(myClass.Value, Index, 1)
End Property

Public Property Get Value() As String
    Value = myClass.Value 'VBA.Mid$(.Ascii, 1, .Length)
End Property

Public Function ToCharArray() As String()
    Dim Chars() As String
    Dim i As Long

    With myClass
        ReDim Chars(1 To Len(myClass.Value))
        For i = LBound(Chars) To UBound(Chars)
            Chars(i) = VBA.Mid$(myClass.Value, i, 1)
        Next
    End With

    ToCharArray = Chars
End Function
Public Function ToByteArray() As Byte(): ToByteArray = StrConv(myClass.Value, vbFromUnicode): End Function

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

    Dim S As TString: Set S = Other
    If Not (S.Length = Me.Length) Then IObject_Equals = False: Exit Function

'    Dim i As Long
'    Dim strB() As Byte: strB = S.ToByteArray
'
'    For i = 0 To myClass.Length - 1
'        If Not myClass.Ascii(i) = strB(i) Then IObject_Equals = False: Exit Function
'    Next
'
'    IObject_Equals = True
'
    IObject_Equals = (StrComp(myClass.Value, S.Value, vbBinaryCompare) = 0)
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim S As TString
    Dim i As Long
    
    If Other.IsRelatedTo(Me) Then
        Set S = Other
'        Dim strB() As Byte: strB = S.ToByteArray
'
'        l = S.Length
'        If myClass.Length < l Then l = myClass.Length
'
'        For i = 0 To l - 1
'            Select Case CLng(myClass.Ascii(i)) - strB(i)
'                Case Is > 0: IObject_CompareTo = IsGreater: Exit Function
'                Case Is < 0: IObject_CompareTo = IsLower: Exit Function
'            End Select
'        Next
'        IObject_CompareTo = IsEqual
        Select Case StrComp(myClass.Value, S.Value, vbBinaryCompare)
            Case 1: IObject_CompareTo = IsGreater '=  1     (String1 and String2 are not equal; String1 is greater than String2
            Case -1: IObject_CompareTo = IsLower '= -1     (String1 and String2 are not equal; String1 is less than String2)
            Case Else: IObject_CompareTo = IsEqual '=  0     (String1 and String2 are equal)
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = Skynet.HashValue(StrConv(myClass.Value, vbFromUnicode), myClass.HashFunction): End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is TString): End Function
Private Function IObject_ToString() As String: IObject_ToString = myClass.Value: End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TString: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

