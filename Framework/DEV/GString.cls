VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "<T>Class"
Option Explicit

Implements IGeneric
Implements IGenericValue

#If Win64 Then
    Declare PtrSafe Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Declare PtrSafe Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare PtrSafe Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
'Groß- und Kleinschreibung spielt keine Rolle
    Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    'Groß- und Kleinschreibung wird beachtet
    Private Declare Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function CompareString Lib "kernel32.dll" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

' Retval = CompareString(0&, NORM_IGNORESYMBOLS Or _
'        SORT_STRINGSORT, .List(i), Len(.List(i)), .List(j), Len(.List(j)))
'        If Retval = CSTR_LESS_THAN Or Retval = CSTR_EQUAL Then
'
' CompareString Rückgabe-Konstanten
Private Const CSTR_LESS_THAN = 1 ' Kleiner als
Private Const CSTR_EQUAL = 2 ' Gleich
Private Const CSTR_GREATER_THAN = 3 ' Größer als
 
' CompareString Locale-Konstanten
Private Const LOCALE_SYSTEM_DEFAULT = &H400 ' Standard-Systemsprache
Private Const LOCALE_USER_DEFAULT = &H800 ' Standard-Benutzersprache
 
' CompareString dwCmpFlags-Konstanten
Private Const NORM_IGNORECASE = &H1 ' Groß und Kleinschreibung nicht beachten
Private Const NORM_IGNOREKANATYPE = &H10000 ' Für Japanischen Zeichensatz.
' Macht keine Unterschiede zwischen Hiragana und Katakana
Private Const NORM_IGNORENONSPACE = &H2 ' Ignoriert Leerzeichen
Private Const NORM_IGNORESYMBOLS = &H4 ' Ignoriert Sonderzeichen
Private Const NORM_IGNOREWIDTH = &H20000 ' Ignoriert den Unterschied
' zwischen Unicode und normalen Strings
Private Const SORT_STRINGSORT = &H1000 ' Benutzt eine String-Sort-Methode
' anstatt einer Wort Sortierung. In einer String-Sort-Sortierung kommen
' Sonderzeichen wie "?,#;*..." vor den Buchstaben

Private Const DEFAULT_DELIMITER As String = ","
Private Const MAXVALUE As Double = 2 ^ 31

Public Enum TStringFormat
    [_First] = -1
    DefaultString = 0
    Trim
    LTrim
    RTrim
    [_Last]
End Enum

Private Type Member
    Value As String
    Length As Long
    Format As TStringFormat
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get ElementAt(ByVal Index As Long) As GString: Set ElementAt = GString(VBA.Mid$(myclass.Value, Index, 1)): End Property
Public Property Get Length() As Long: Length = myclass.Length: End Property
Public Property Get FirstChar() As String: FirstChar = VBA.Mid$(myclass.Value, 1, 1): End Property
Public Property Get LastChar() As String: LastChar = VBA.Mid$(myclass.Value, Len(myclass.Value), 1): End Property
Public Property Get Character(ByVal Index As Long) As String: Character = VBA.Mid$(myclass.Value, Index, 1): End Property
Public Property Get Format() As TStringFormat: Format = myclass.Format: End Property

Public Function ToFormat() As String
    
    With myclass
        Select Case .Format
            Case TStringFormat.DefaultString: ToFormat = .Value: Exit Function
            Case TStringFormat.Trim: ToFormat = VBA.Trim$(.Value): Exit Function
            Case TStringFormat.LTrim: ToFormat = VBA.LTrim$(.Value): Exit Function
            Case TStringFormat.RTrim: ToFormat = VBA.RTrim$(.Value): Exit Function
        End Select
    End With
    
End Function
Public Property Get Value() As String: Value = myclass.Value: End Property

Public Function Replace(ByVal Find As String, ReplaceWith As String) As GString: Set Replace = GString(VBA.Replace$(myclass.Value, Find, ReplaceWith), myclass.Format): End Function
Public Function ToByteArray() As Byte(): ToByteArray = StrConv(myclass.Value, vbFromUnicode): End Function
Public Function IsLike(ByVal Pattern As String) As Boolean: IsLike = (myclass.Value Like Pattern): End Function
Public Function StartsWith(ByVal Text As String) As Boolean: StartsWith = (myclass.Value Like ("%" & Text)): End Function
Public Function EndsWith(ByVal Text As String) As Boolean: EndsWith = (myclass.Value Like (Text & "%")): End Function
Public Function Contains(ByVal Text As String) As Boolean: Contains = Not (InStr(1, myclass.Value, Text, vbTextCompare) = 0): End Function
Public Function CompareWithString(ByRef S As String) As ComparerEnum: CompareWithString = StrComp(myclass.Value, S, vbBinaryCompare): End Function
Public Function IsNullOrEmpty() As Boolean
    If myclass.Value = vbNullString Then _
        IsNullOrEmpty = True: Exit Function
    If Me Is GString Then _
        IsNullOrEmpty = True: Exit Function
End Function

Public Function CountChar(ByVal c As String) As Long
    
    Dim i As Long
    Dim ascii As Integer
    
    If Len(c) > 1 Then _
        Skynet.Throw(Me, "Private Function GetArrayStructure").NotSupported ("Len(C) > 1")

    ascii = AscW(c)
    
    For i = 1 To Me.Length
        If AscW(Mid$(myclass.Value, i, 1)) = ascii Then _
            CountChar = CountChar + 1
    Next
    
End Function

Public Function Concat(ByVal Item As IGeneric, Optional ByVal Delimiter = DEFAULT_DELIMITER) As GString
    
    Dim Other As GString
    
    If Item Is Nothing Then _
        Set Concat = Me: Exit Function
    
    If TypeOf Item Is GString Then _
        Set Other = Item
            
    If Not Me Is GString Then
        If Not Other Is Nothing Then
            Set Concat = GString(myclass.Value & Delimiter & Other.Value, myclass.Format)
        Else
            Set Concat = GString(myclass.Value & Delimiter & Item.ToString, myclass.Format)
        End If
    Else 'Static default call
        If Not Other Is Nothing Then
            Set Concat = Item.Clone ' must be a Gstring ;)
        Else
            Set Concat = GString(Item.ToString, myclass.Format)
        End If
    End If
    
End Function

Public Function Split(ByVal Delimiter As String) As GenericArray
    
    Dim chars As Variant
    chars = VBA.Split(myclass.Value, Delimiter)
    
    Dim i As Long
    ReDim Strings(1 To UBound(chars) + 1) As IGeneric
    
    For i = LBound(chars) To UBound(chars)
        Set Strings(i + 1) = GString(chars(i))
    Next
    
    Set Split = GenericArray.BuildFrom(Strings)
    
End Function

Public Function ToArray() As GenericArray
    
    Dim chars() As IGeneric
    Dim i As Long

    ReDim chars(1 To Len(myclass.Value))
    For i = LBound(chars) To UBound(chars)
        Set chars(i) = GString(VBA.Mid$(myclass.Value, i, 1))
    Next

    Set ToArray = GenericArray.BuildFrom(chars)
    
End Function

Public Function Join(ByVal Values As GenericArray, Optional ByVal Delimiter = DEFAULT_DELIMITER) As GString
    
    Dim S As IGeneric
    Dim i As Long
    
    If (Values Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Join").NullPointer("Values is nothing")
    
    ReDim Strings(Values.LowerBound To Values.Length) As String
    
    With Values.Iterator
        Do While .HasNext(S)
            i = i + 1
            If Not S Is Nothing Then _
                Strings(i) = S.ToString
        Loop
    End With
    
    Set Join = Me.Concat(GString(VBA.Join$(Strings, Delimiter), myclass.Format))
    
End Function

Public Function HashValueOf(ByRef Text As String) As Long

    Const Shift As Long = 2& ^ 5
    
    Dim H As Double
    Dim i As Long
    
    If LenB(Text) = 0 Then Exit Function
    
    Dim ascii() As Byte: ascii() = StrConv(Text, vbFromUnicode)
    Dim Length As Long: Length = UBound(ascii)
    
    Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
    Dim h2 As Double: h2 = h1
    
    Do Until i = Length: h1 = X0R((h1 * Shift) + h1, ascii(i)): h2 = X0R((h2 * Shift) + h2, ascii(InterlockedIncrement(i))): Loop
    
    H = h1 + (h2 * 1566083941#)
    
    HashValueOf = CLng(H - (Fix(H / MAXVALUE) * MAXVALUE))

End Function

Public Function Build( _
                        ByVal Value As Variant, _
                        Optional ByVal Format As TStringFormat = TStringFormat.DefaultString) As GString
Attribute Build.VB_UserMemId = 0
    With myclass
        If Not IsNull(Value) Then
            .Value = Value
            .Length = Len(Value)
        End If
        .Format = Format
    End With
    
    'Set Build = Skynet.Create(GString)
    Set Build = Skynet.CreateInstance(New GString, Virtual.Pointer, Virtual.Size)
 
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GString: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = Me.HashValueOf(myclass.Value): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GString):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GString, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    Dim S As GString: Set S = Other
    If (myclass.Length = S.Length) = False Then IGeneric_Equals = False: Exit Function

    IGeneric_Equals = (StrComp(myclass.Value, S.Value, vbBinaryCompare) = 0)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    Dim S As GString: Set S = Other
    IGeneric_CompareTo = StrComp(myclass.Value, S.Value, vbTextCompare)

End Function

Private Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

Private Function IGeneric_ToString() As String

    Select Case myclass.Format
        Case TStringFormat.DefaultString: IGeneric_ToString = myclass.Value: Exit Function
        Case TStringFormat.Trim: IGeneric_ToString = VBA.Trim$(myclass.Value): Exit Function
        Case TStringFormat.LTrim: IGeneric_ToString = VBA.LTrim$(myclass.Value): Exit Function
        Case TStringFormat.RTrim: IGeneric_ToString = VBA.RTrim$(myclass.Value): Exit Function
    End Select
    
End Function
'IGenericValue
Private Function IGenericValue_ToParameter( _
                                          Optional ByVal Name As GString, _
                                          Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As Variant
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Name = Name.Value
        .Attributes = adParamNullable '64
        .Type = ISqlDataType.VarChar_Type
        .Size = Len(myclass.Value)
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToParameter = P
    
End Function

'IGenericValue
Private Property Get IGenericValue_GetType() As VbVarType: IGenericValue_GetType = vbString: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Decimal_Type: End Property
Private Property Get IGenericValue_RawValue() As Variant: IGenericValue_RawValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property

Private Function IGenericValue_Create(ByVal Value As Variant, Optional ByVal Direction As ISqlParameterDirection = 1&) As IGenericValue
    Set IGenericValue_Create = Me.Build(Value)
End Function

Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As ComparerEnum
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim S As GString: Set S = Other
    IGenericValue_CompareTo = StrComp(myclass.Value, S.Value, vbTextCompare)
    
End Function

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Variant
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Name = Name
        .Attributes = adParamNullable '64
        .Type = ISqlDataType.VarChar_Type
        .Size = Len(myclass.Value)
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = P
    
End Function
