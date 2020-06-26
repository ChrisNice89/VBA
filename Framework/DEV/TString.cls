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
    Hashfunction As Hashing
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Hashfunction() As Hashing: Hashfunction = myClass.Hashfunction: End Property
Public Property Get Length() As Long: Length = Len(myClass.Value): End Property
Public Property Get FirstChar() As String: FirstChar = VBA.Mid$(myClass.Value, 1, 1): End Property
Public Property Get LastChar() As String: LastChar = VBA.Mid$(myClass.Value, Len(myClass.Value), 1): End Property
Public Property Get Character(ByVal Index As Long) As String: Character = VBA.Mid$(myClass.Value, Index, 1): End Property
Public Function Concat(ByVal Text As TString, Optional ByVal Delimiter = " ") As TString: Set Concat = TString(myClass.Value & Delimiter & Text.Value, myClass.Hashfunction, myClass.Format): End Function
Public Function Replace(ByVal Find As String, ReplaceWith As String) As TString: Set Replace = TString(VBA.Replace$(myClass.Value, Find, ReplaceWith), myClass.Hashfunction, myClass.Format): End Function
Public Property Get Format() As TStringFormat: Format = myClass.Format: End Property
Public Property Get Value() As String: Value = myClass.Value: End Property
Public Function ToByteArray() As Byte(): ToByteArray = StrConv(myClass.Value, vbFromUnicode): End Function

Public Function Join(ByRef Strings() As TString, Optional ByVal Delimiter) As TString
    
    
   
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

Public Function Build( _
                        ByVal Value As String, _
                        Optional ByVal Hashfunction As Hashing = Hashing.Function1, _
                        Optional ByVal Format As TStringFormat = TStringFormat.DefaultString) As TString
Attribute Build.VB_UserMemId = 0
    With myClass
        .Value = Value
        .Format = Format
        .Hashfunction = Hashfunction
    End With
    
    Set Build = New TString    'faster then Skynet
    'Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    If Not Me Is TString Then
         With myClass
            .Value = TString.Value
            .Format = TString.Format
            .Hashfunction = TString.Hashfunction
        End With
        Call Skynet.CastObject(TString).Dispose
    End If
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IObject: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType: IValue_DatabaseType = T_String: End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IObject_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not TypeOf Other Is TString Then IObject_Equals = False: Exit Function

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
'Private Property Get IObject_HashValue() As Long
'
'    Dim h As Double, i As Long
'    Dim Ascii() As Byte: Ascii = StrConv(myClass.Value, vbFromUnicode)
'
'    Select Case Hashfunction
'        Case myClass.Hashfunction
'
'            Const Shift As Long = 2& ^ 5
'            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
'            Dim h2 As Double: h2 = h1
'
'            For i = LBound(Ascii) To UBound(Ascii) - 1
'                h1 = Me.X0R((h1 * Shift) + h1, Ascii(i)): h2 = Me.X0R((h2 * Shift) + h2, Ascii(i + 1))
'            Next
'
'            h = h1 + (h2 * 1566083941#)
'
'        Case myClass.Hashfunction
'
'             Dim Length As Long: Length = UBound(Ascii) + 1
'             Dim n As Long
'
'             For n = (Length / 2) To 1 Step -1
'                h = h + Ascii(i)
'                h = Me.X0R(Me.X0R(LeftShift(h, 16), Me.LeftShift(Ascii(i + 1), 11)), h)
'                h = h + Me.RightShift(h, 11)
'                i = i + 2
'            Next
'
'            If ((Length Mod 2) = 1) Then
'                h = h + Ascii(i) + 1566083941
'                h = Me.X0R(h, LeftShift(h, 10))
'                h = h + Me.RightShift(h, 1)
'            End If
'
'            h = Me.X0R(h, Me.LeftShift(h, 3)): h = h + Me.RightShift(h, 5)
'            h = Me.X0R(h, Me.LeftShift(h, 4)): h = h + Me.RightShift(h, 17)
'            h = Me.X0R(h, Me.LeftShift(h, 25)): h = h + Me.RightShift(h, 6)
'    End Select
'
'    HashValue = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))
'
'End Property

Private Property Get IObject_HashValue() As Long: IObject_HashValue = Skynet.HashValue(StrConv(myClass.Value, vbFromUnicode), myClass.Hashfunction): End Property
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

