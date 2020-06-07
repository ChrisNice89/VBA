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
Implements IValue

Public Enum TStringFormat
    [_First]
    DefaultString
    Trim
    LTrim
    RTrim
    [_Last]
End Enum

Private Type Member
    Value As String
    Format As TStringFormat
    HashFunction As Hashing
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

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

Public Function Replace(ByVal Find As String, ReplaceWith As String) As TString
    Set Replace = TString(VBA.Replace$(myClass.Value, Find, ReplaceWith), myClass.HashFunction, myClass.Format)
End Function

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

Public Property Get Format() As TStringFormat
    Format = myClass.Format
End Property

Public Property Get Value() As String
    Value = myClass.Value 'VBA.Mid$(.Ascii, 1, .Length)
End Property

Public Function Build( _
                        ByVal Value As String, _
                        Optional ByVal HashFunction As Hashing = Hashing.Function1, _
                        Optional ByVal Format As TStringFormat = TStringFormat.DefaultString) As TString
Attribute Build.VB_UserMemId = 0
With myClass
        .Value = Value
        .Format = Format
        .HashFunction = HashFunction
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Function IValue_Build(ByVal Value As Variant, Optional ByVal Format As Byte) As IObject
    With myClass
        .Value = Value
        .Format = Format
    End With
    
    Set IValue_Build = Skynet.Create(Me)
End Function
Private Property Get IValue_Object() As IObject: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IObject_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Dim s As TString: Set s = Other
    If Not (s.Length = Me.Length) Then IObject_Equals = False: Exit Function

'    Dim i As Long
'    Dim strB() As Byte: strB = S.ToByteArray
'
'    For i = 0 To myClass.Length - 1
'        If Not myClass.Ascii(i) = strB(i) Then IObject_Equals = False: Exit Function
'    Next
'
'    IObject_Equals = True
'
    IObject_Equals = (StrComp(myClass.Value, s.Value, vbBinaryCompare) = 0)
     
End Function

Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim s As TString
    Dim i As Long
    
    If Other.IsRelatedTo(Me) Then
        Set s = Other
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
        Select Case StrComp(myClass.Value, s.Value, vbBinaryCompare)
            Case 1: IObject_CompareTo = IsGreater
            Case -1: IObject_CompareTo = IsLower
            Case Else: IObject_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = Skynet.HashValue(StrConv(myClass.Value, vbFromUnicode), myClass.HashFunction): End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is TString): End Function
Private Function IObject_ToString() As String

    Select Case myClass.Format
        Case TStringFormat.DefaultString: IObject_ToString = myClass.Value: Exit Function
        Case TStringFormat.Trim: IObject_ToString = VBA.Trim$(myClass.Value): Exit Function
        Case TStringFormat.LTrim: IObject_ToString = VBA.LTrim$(myClass.Value): Exit Function
        Case TStringFormat.RTrim: IObject_ToString = VBA.RTrim$(myClass.Value): Exit Function
    End Select
    
End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TString: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

