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
Implements IGenericReadOnlyList
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Declare PtrSafe Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Private Declare PtrSafe Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare PtrSafe Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
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

Private Const INTERNAL_POOL_SIZE As Long = 2 ^ 10
Private Const DEFAULT_DELIMITER As String = ","
Private Const MAXVALUE As Double = 2 ^ 31
Private Const LEFTSHIFT As Long = 2& ^ 5
Private Const RIGHTSHIFT As Long = 2& ^ 27
Private Const DEFAULT_VALUE As String = VBA.vbNullString

Public Enum StringFormat
    [_First] = -1
    DefaultString = 0
    Trim
    LTrim
    RTrim
    [_Last]
End Enum

Private Type Iterator
    current As GString
    Index As Long
End Type

Private Type Member
    Length As Long
    Value As String
    Ptr As LongPtr
    Hash As Long
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get Chars() As IGenericCollection: Set Chars = Me: End Property
Public Property Get Pointer() As LongPtr: Pointer = myclass.Ptr: End Property
Public Property Get ElementAt(ByVal Index As Long) As GString: Set ElementAt = GString.Build(VBA.Mid$(myclass.Value, Index + 1, 1)): End Property
Public Property Get IsDefault() As Boolean: IsDefault = Me Is GString: End Property
Public Property Get IsNumeric() As Boolean: IsNumeric = VBA.IsNumeric(myclass.Value): End Property
Public Property Get Length() As Long: Length = myclass.Length: End Property
Public Property Get LastChar() As GString: Set LastChar = Me.ElementAt(myclass.Length - 1): End Property
Public Property Get IsNullOrEmpty() As Boolean: myclass.Length = 0: End Property
Public Property Get Value() As String: Value = myclass.Value: End Property

Public Function Build(ByRef Value As String) As GString
Attribute Build.VB_UserMemId = 0
    
    If Not (Me Is GString) Then _
        Set Build = GString.Build(Value): Exit Function

    If Value = DEFAULT_VALUE Then _
        Set Build = GString: Exit Function

    myclass.Length = Len(Value)

    If myclass.Length = 1 Then

        Static Chars As GenericArray
        If Chars Is Nothing Then _
            Set Chars = GenericArray.Build(2 ^ 16)

        Dim Char As GString
        Dim Code As Long

        Code = AscW(Value)
        Set Char = Chars.ElementAt(Code)

        If Char Is Nothing Then
            myclass.Value = ChrW$(Code)
            Set Char = New GString
            Set Chars.ElementAt(Code) = Char
            myclass = EmptyStructure
        End If

        Set Build = Char
        Exit Function
    End If
        
    myclass.Value = Value
    Set Build = New GString
    myclass = EmptyStructure
  
End Function

Private Sub Class_Initialize()
    With GString
        myclass.Value = .Value: myclass.Length = .Length
    End With
    myclass.Ptr = StrPtr(myclass.Value)
End Sub
Public Function ToFormat(ByVal Format As StringFormat) As String
    
    Select Case Format
        Case StringFormat.DefaultString: ToFormat = myclass.Value: Exit Function
        Case StringFormat.Trim: ToFormat = VBA.Trim$(myclass.Value): Exit Function
        Case StringFormat.LTrim: ToFormat = VBA.LTrim$(myclass.Value): Exit Function
        Case StringFormat.RTrim: ToFormat = VBA.RTrim$(myclass.Value): Exit Function
    End Select
 
End Function
Public Function ToLower() As GString: Set ToLower = GString.Build(LCase(myclass.Value)): End Function
Public Function ToUpper() As GString: Set ToUpper = GString.Build(UCase(myclass.Value)): End Function
Public Function Replace(ByVal Find As String, ReplaceWith As String) As GString: Set Replace = GString.Build(VBA.Replace$(myclass.Value, Find, ReplaceWith)): End Function
Public Function IsLike(ByVal Pattern As String) As Boolean: IsLike = (myclass.Value Like Pattern): End Function
Public Function StartsWith(ByVal RegEx As String) As Boolean: StartsWith = (myclass.Value Like ("%" & RegEx)): End Function
Public Function EndsWith(ByVal RegEx As String) As Boolean: EndsWith = (myclass.Value Like (RegEx & "%")): End Function
Public Function Contains(ByRef Text As String) As Boolean: Contains = Not (Me.IndexOf(Text:=Text) = 0): End Function
Public Function IndexOf(ByRef Text As String) As Long: IndexOf = InStr(1, myclass.Value, Text, vbTextCompare): End Function

Public Function AsciiList() As GenericArray
    
    Dim Ascii As GenericArray
    Set Ascii = GenericArray.Build(255)
    
    Dim i As Long
    With Ascii
        For i = 0 To 255 - 1
            Set .ElementAt(i) = GString.Build(VBA.Chr$(i))
        Next
    End With
    
    Set AsciiList = Ascii
    
End Function
Public Function CountChar(ByVal c As String) As Long
    
    Dim i As Long
    Dim Ascii As Integer
    
    If Len(c) > 1 Then _
        System.Throw(Me, "Private Function GetArrayStructure").NotSupported ("Len(C) > 1")

    Ascii = AscW(c)
    
    For i = 1 To Me.Length
        If AscW(Mid$(myclass.Value, i, 1)) = Ascii Then _
            CountChar = CountChar + 1
    Next
    
End Function

Public Function Insert(ByVal Index As Long, Text As GString) As GString

    If Index < 1 Then _
        Index = 1
    
    If myclass.Length >= Index Then
        Set Insert = GString.Build(VBA.Mid$(myclass.Value, 1, Index - 1) _
                        & Text.Value _
                        & VBA.Mid$(myclass.Value, Index, myclass.Length - Index + 1))
    Else
        Set Insert = Me.Concat(Text)
    End If

End Function

Public Function Concat(ByVal Other As GString, Optional ByVal Delimiter = vbNullString) As GString
   
    If Other Is Nothing Then _
        Set Concat = Me: Exit Function
        
    Set Concat = GString.Build(myclass.Value & Delimiter & Other.Value)
    
End Function

Public Function Repeat(ByVal Number As Long, ByVal Text As String, Optional ByVal Delimiter As String = vbNullString) As GString
    
    Dim i As Long
        
    If Number <= 0 Then _
        Set Repeat = GString.Build(DEFAULT_VALUE): Exit Function
        
    ReDim Strings(1 To Number) As String
    
    For i = 1 To Number
        Strings(i) = Text
    Next
    
    Set Repeat = GString.Build(VBA.Join$(Strings, Delimiter))

End Function

Public Function Split(ByVal Delimiter As String) As GenericArray
    
    Dim i As Long
    Dim Parts As GenericArray
    Dim Strings As Variant
    
    Strings = VBA.Split(myclass.Value, Delimiter)
    Set Parts = GenericArray.Build(UBound(Strings) + 1)
    
    With Parts
        For i = LBound(Strings) To UBound(Strings)
            Set .ElementAt(i) = GString.Build(CStr(Strings(i)))
        Next
    End With

    Set Split = Parts
    
End Function

Public Function Join(ByVal Values As IGenericCollection, Optional ByVal Delimiter = DEFAULT_DELIMITER) As GString
    
    Dim s As IGeneric
    Dim i As Long
  
    If (Values Is Nothing) Then _
        Call System.Throw(Me, "Public Function Join").IsNothing("(Values Is Nothing)")
    
    If (Values.Count <= 0) Then _
        Call System.Throw(Me, "Public Function Join").IsNothing("(Values Is Nothing)")
    
    ReDim Strings(Values.Count - 1) As String
    
    With Values.Iterator
        Do While .HasNext(s)
            If Not s Is Nothing Then _
                Strings(i) = s.ToString
            i = i + 1
        Loop
    End With
    
    Set Join = GString(VBA.Join$(Strings, Delimiter))
    
End Function

Public Function HashValueOf(ByRef Text As String) As Long

    Dim Ascii() As Byte
    Dim numChars As Long
    Dim h As Double
    Dim i As Long
    Dim h1 As Double, h2 As Double
    
    If Text = vbNullString Then _
        Exit Function
    
    numChars = Len(Text)
    If numChars = 1 Then _
        HashValueOf = VBA.AscW(Text): Exit Function
    
    h1 = 352654597#
    h2 = h1 'LeftShift(5381,16))+5381
    
    Ascii() = StrConv(Text, vbFromUnicode)
    
    For i = 0 To numChars - 2 'For i = 1 To Len(Text) - 1 'For i = 0 To UBound(Ascii) - 1
       h1 = (h1 * LEFTSHIFT) + h1 + (h1 / RIGHTSHIFT): h1 = CLng(h1 - (Fix(h1 / MAXVALUE) * MAXVALUE)) Xor Ascii(i) 'AscW(Mid$(Text, i, 1))
       h2 = (h2 * LEFTSHIFT) + h2 + (h2 / RIGHTSHIFT): h2 = CLng(h2 - (Fix(h2 / MAXVALUE) * MAXVALUE)) Xor Ascii(i + 1) 'AscW(Mid$(Text, i + 1, 1))
    Next
    
    h = h1 + h2 'h1 + (h2 * 1566083941#)

    HashValueOf = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GString: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long
    
    If myclass.Hash = 0 Then _
        myclass.Hash = Me.HashValueOf(myclass.Value)
    
    IGeneric_HashValue = myclass.Hash
    
End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GString):  End Property

Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is GString Then _
        IGeneric_Equals = False: Exit Function
        
    Dim s As GString: Set s = Other

    If (myclass.Length = s.Length) Then _
        IGeneric_Equals = (StrComp(myclass.Value, s.Value, vbBinaryCompare) = 0)

'    IGeneric_Equals = (StrComp(myclass.Pointer, S.Pointer, vbBinaryCompare) = 0)
'    IGeneric_Equals = (StrComp(myclass.Bytes, S.Bytes, vbBinaryCompare) = 0)
'    IGeneric_Equals = CSTR_EQUAL = CompareString(0&, NORM_IGNORESYMBOLS Or SORT_STRINGSORT, myclass.Value, myclass.Length, S.Value, S.Length)

End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = myclass.Value: End Function

'IGenericValue
Private Function IGenericValue_ToParameter( _
                                          Optional ByVal Name As GString, _
                                          Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput) As Variant
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Name = Name.Value
        .Attributes = adParamNullable '64
        .Type = ISqlDataType.VarChar_Type
        .Size = Len(myclass.Value)
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToParameter = p
    
End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbString: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.VarWChar_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericComparer: End Property

Private Function IGenericValue_Convert(ByVal Value As Variant) As IGenericValue: Set IGenericValue_Convert = GString.Build(VBA.IIf(VBA.Isnull(Value), DEFAULT_VALUE, Value)): End Function
Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As CompareResult
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim s As GString: Set s = Other
    
    If s.IsNullOrEmpty Then _
        IGenericValue_CompareTo = IIf(myclass.Length = 0, CompareResult.IsEqual, CompareResult.IsGreater): Exit Function
    
    IGenericValue_CompareTo = StrComp(myclass.Value, s.Value, vbTextCompare)
'    IGenericValue_CompareTo = lstrcmpi(myclass.Value, S.Value)
    
End Function

Private Function IGenericValue_ToSqlParameter( _
                                                 Optional ByVal Direction As ISqlParameterDirection = 1&, _
                                                 Optional ByVal Name As String) As Object
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Name = Name
        .Attributes = adParamNullable '64
        .Type = IGenericValue_VarType
        .Size = myclass.Length
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = GenericArray.LowerBound: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = myclass.Length - GenericArray.LowerBound - 1: End Property

Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Text:=Element.ToString): End Function
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Length) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.Length)")

    Set IGenericReadOnlyList_ElementAt = Me.ElementAt(Index)
    
End Property

Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Private Sub IGenericReadOnlyList_CopyTo").IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call System.Throw(Me, "Private Sub IGenericReadOnlyList_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Length) Then _
        Call System.Throw(Me, "Private Sub IGenericReadOnlyList_CopyTo").InvalidInput("(GArray.Length - Index < String.Length)")
    
    Dim i As Long
    For i = IGenericReadOnlyList_First To Me.Length
        Set GArray(ArrayIndex) = Me.ElementAt(i)
        ArrayIndex = ArrayIndex + 1
    Next

End Sub

'IGenericCollection
Private Sub IGenericCollection_Clear()

End Sub

Private Function IGenericCollection_Contains(ByVal Element As IGeneric) As Boolean



End Function

Private Function IGenericCollection_Copy() As IGenericCollection: Set IGenericCollection_Copy = Me: End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IsNothing("(GArray Is Nothing)")
    
    With GArray
        If (ArrayIndex < .LowerBound) Or (ArrayIndex > .Length) Then _
            Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")
    
        If (.Length - ArrayIndex + .LowerBound < myclass.Length) Then _
            Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < String.Length)")
            
    End With
    
    Dim i As Long
    For i = IGenericReadOnlyList_First To Me.Length
        Set GArray(ArrayIndex) = Me.ElementAt(i)
        ArrayIndex = ArrayIndex + 1
    Next
    
End Sub

Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = myclass.Length: End Property

Private Function IGenericCollection_Iterator() As IGenericIterator
    
    Set IGenericCollection_Iterator = Me
    Call IGenericCollection_Iterator.Reset
    
End Function

Private Function IGenericCollection_ToArray() As GenericArray
    
    Dim Chars As GenericArray
    Dim i As Long
    
    Set Chars = GenericArray
    If Not myclass.Length = 0 Then
        Set Chars = GenericArray.Build(myclass.Length)
         
        With Chars
            For i = .LowerBound To .Length - 1
                Set .ElementAt(i) = Me.ElementAt(i)
            Next
        End With
    End If
    
    Set IGenericCollection_ToArray = Chars
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
        If .Index < Me.Chars.Count Then
            Set .current = Me.ElementAt(.Index)
            .Index = .Index + 1
            Set Out = .current
            IGenericIterator_HasNext = True
            Exit Function
        End If
    End With
    
    Set Out = Nothing
    Set myclass.Iterator.current = Nothing
            
End Function

Private Sub IGenericIterator_Reset()
   myclass.Iterator.Index = 0
End Sub

