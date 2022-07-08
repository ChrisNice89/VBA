VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@IgnoreModule UnassignedVariableUsage, VariableNotAssigned, SuspiciousPredeclaredInstanceAccess, ProcedureNotUsed, ConstantNotUsed, AssignedByValParameter
'@Folder "Primitives"
Option Explicit

Implements IGeneric
Implements IGenericValue
Implements IGenericSequence
Implements IGenericCollection

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef Dst As LongPtr, ByRef Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Private Declare PtrSafe Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare PtrSafe Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef Dst As Long, ByRef Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long 'Groß- und Kleinschreibung spielt keine Rolle
    Private Declare Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long 'Groß- und Kleinschreibung wird beachtet
    Private Declare Function CompareString Lib "kernel32.dll" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long  'Works
#End If

' Retval = CompareString(0&, NORM_IGNORESYMBOLS Or _
'        SORT_STRINGSORT, a, Len(a), b, Len(b))
'        If Retval = CSTR_LESS_THAN Or Retval = CSTR_EQUAL Then
'
' CompareString Rückgabe-Konstanten
'@Ignore ImplicitlyTypedConst
Private Const CSTR_LESS_THAN = 1 ' Kleiner als
'@Ignore ImplicitlyTypedConst
Private Const CSTR_EQUAL = 2 ' Gleich
'@Ignore ImplicitlyTypedConst
Private Const CSTR_GREATER_THAN = 3 ' Größer als
 
' CompareString Locale-Konstanten
'@Ignore ImplicitlyTypedConst
Private Const LOCALE_SYSTEM_DEFAULT = &H400 ' Standard-Systemsprache
'@Ignore ImplicitlyTypedConst
Private Const LOCALE_USER_DEFAULT = &H800 ' Standard-Benutzersprache
 
' CompareString dwCmpFlags-Konstanten
'@Ignore ImplicitlyTypedConst
Private Const NORM_IGNORECASE = &H1 ' Groß und Kleinschreibung nicht beachten
'@Ignore ImplicitlyTypedConst
Private Const NORM_IGNOREKANATYPE = &H10000 ' Für Japanischen Zeichensatz.
' Macht keine Unterschiede zwischen Hiragana und Katakana
'@Ignore ImplicitlyTypedConst
Private Const NORM_IGNORENONSPACE = &H2 ' Ignoriert Leerzeichen
'@Ignore ImplicitlyTypedConst
Private Const NORM_IGNORESYMBOLS = &H4 ' Ignoriert Sonderzeichen
'@Ignore ImplicitlyTypedConst
Private Const NORM_IGNOREWIDTH = &H20000 ' Ignoriert den Unterschied
' zwischen Unicode und normalen Strings
'@Ignore ImplicitlyTypedConst
Private Const SORT_STRINGSORT = &H1000 ' Benutzt eine String-Sort-Methode
' anstatt einer Wort Sortierung. In einer String-Sort-Sortierung kommen
' Sonderzeichen wie "?,#;*..." vor den Buchstaben

Private Const NOT_SPECIFIED As Long = -1
Private Const INITIAL_BUFFER_SIZE As Long = 512 'Number of chars
Private Const INTERNAL_CACHE_SIZE As Long = 2 ^ 16

Private Const DEFAULT_VALUE As String = VBA.vbNullString
'@Ignore IntegerDataType
Private Const DEFAULT_COMPARE_CULTURE As Integer = VBA.VbCompareMethod.vbTextCompare

Private Type Member
    Length As Long
    Hash As Long
    Value As String
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property

Public Property Get First() As Long: First = 0: End Property
Public Property Get Last() As Long: Last = This.Length - 1: End Property
Public Property Get Pointer() As LongPtr: Pointer = VarPtr(This.Value): End Property
Public Property Get IsDefault() As Boolean: IsDefault = Me Is gString: End Property
Public Property Get IsNumeric() As Boolean: IsNumeric = VBA.IsNumeric(This.Value): End Property
Public Property Get Length() As Long: Length = This.Length: End Property
Public Property Get FirstChar() As gString: Set FirstChar = gString.Of(This.Value, 0, 1): End Property
Public Property Get LastChar() As gString: Set LastChar = gString.Of(This.Value, This.Length - 1, 1): End Property
Public Property Get IsEmpty() As Boolean: IsEmpty = (This.Length = 0): End Property
'@DefaultMember
Public Property Get ToString(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As String
Attribute ToString.VB_UserMemId = 0

    If (BeginAt = NOT_SPECIFIED) Then
        If (Count = NOT_SPECIFIED) Then
            ToString = This.Value
        Else
            ToString = VBA.Left$(This.Value, Count)
        End If
    Else
        If (Count = NOT_SPECIFIED) Then
            ToString = VBA.mid$(This.Value, BeginAt + 1)
        Else
            ToString = VBA.mid$(This.Value, BeginAt + 1, Count)
        End If
    End If

End Property

'@Ignore MisleadingByRefParameter
Public Property Let ToString(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED, ByRef Value As String)
    
    If (BeginAt = NOT_SPECIFIED) Then
        If (Count = NOT_SPECIFIED) Then
            This.Length = VBA.Len(Value)
            This.Value = Value
        Else
            This.Length = Count
            This.Value = VBA.Left$(Value, This.Length)
        End If
    Else
        If (Count = NOT_SPECIFIED) Then
            This.Length = VBA.Len(Value) - BeginAt
            This.Value = VBA.mid$(Value, BeginAt, This.Length)
        Else
            This.Length = Count
            This.Value = VBA.mid$(Value, BeginAt + 1, This.Length)
        End If
    End If
    
    This.Hash = 0
    
End Property

Public Function Of( _
                        ByRef Value As String, _
                        Optional ByVal BeginAt As Long = NOT_SPECIFIED, _
                        Optional ByVal Count As Long = NOT_SPECIFIED) As gString
    
    Dim Char As gString
    Dim Code As Long, l As Long
    Static CharCache As GenericArray
  
    If Not (Me Is gString) Then _
        Set Of = gString.Of(Value): Exit Function
    
    If (Count = NOT_SPECIFIED) Then
        l = VBA.Len(Value)
    Else
        l = Count
    End If

    If (l = 1) Then

        If (CharCache Is Nothing) Then _
            Set CharCache = GenericArray.Build(INTERNAL_CACHE_SIZE)

        If Not (BeginAt = NOT_SPECIFIED) Then
            Code = VBA.AscW(VBA.mid$(Value, BeginAt + 1, 1&))
        Else
            Code = VBA.AscW(Value)
        End If

        Set Char = CharCache.ElementAt(Code)

        If Char Is Nothing Then
            Set Char = New gString
            Char.ToString(BeginAt, 1&) = Value
            Set CharCache.ElementAt(Code) = Char
        End If

        Set Of = Char
        Exit Function
    End If
    
    Set Of = New gString
    Of.ToString(BeginAt, Count) = Value

End Function

Public Function EqualsTo(ByRef Value As String) As Boolean: EqualsTo = (This.Value = Value): End Function

Public Function CompareTo(ByRef Value As String) As CompareResult
    
    If (This.Length = 0) Then _
        CompareTo = VBA.IIf(Value = vbNullString, CompareResult.IsEqual, CompareResult.IsLower): Exit Function
    
    CompareTo = Compare(This.Value, VBA.VbCompareMethod.vbBinaryCompare)
    
End Function

Public Function CompareIgnoreCaseTo(ByRef Value As String) As CompareResult
    
    If (This.Length = 0) Then _
        CompareIgnoreCaseTo = VBA.IIf(Value = vbNullString, CompareResult.IsEqual, CompareResult.IsLower): Exit Function
    
    CompareIgnoreCaseTo = Compare(This.Value, Value, VBA.VbCompareMethod.vbTextCompare)
    
End Function

Public Function Compare(ByRef A As String, ByRef B As String, Optional ByVal Culture As VBA.VbCompareMethod = DEFAULT_COMPARE_CULTURE) As CompareResult

    Select Case VBA.StrComp(A, B, Culture)
        Case -1: Compare = IsLower
        Case 1: Compare = isgreater
        Case Else: Compare = IsEqual
    End Select
    
End Function

Public Function IsBetween(ByRef A As String, ByRef B As String) As Boolean
    
    If (A > This.Value) Then _
        Exit Function
    
    If (B < This.Value) Then _
        Exit Function

    IsBetween = True
    
End Function

Public Function IsLessThan(ByRef Value As String, Optional ByVal Culture As VBA.VbCompareMethod = DEFAULT_COMPARE_CULTURE) As Boolean: IsLessThan = (VBA.StrComp(This.Value, Value, Culture) = -1): End Function
Public Function IsGreaterThan(ByRef Value As String, Optional ByVal Culture As VBA.VbCompareMethod = DEFAULT_COMPARE_CULTURE) As Boolean: IsGreaterThan = (VBA.StrComp(This.Value, Value, Culture) = 1): End Function

Public Function GetAt(ByVal Index As Long) As gString

     If (Index >= This.Length) Then _
        Call GenericError.Build(Me, "Public Function GetAt").IndexOutofRange("(Index >= This.Length)")

    Set GetAt = gString.Of(This.Value, Index, 1&)
    
End Function

Public Function GetCodeAt(ByVal Index As Long) As Long

     If (Index >= This.Length) Then _
        Call GenericError.Build(Me, "Public Function GetCodeAt").IndexOutofRange("(Index >= This.Length)")

    GetCodeAt = VBA.AscW(VBA.mid$(This.Length, Index + 1, 1))
    
End Function

Public Function ToLower() As gString: Set ToLower = gString.Of(VBA.LCase$(This.Value)): End Function
Public Function ToUpper() As gString: Set ToUpper = gString.Of(VBA.UCase$(This.Value)): End Function
Public Function Replace(ByRef Find As String, ByRef ReplaceWith As String) As gString: Set Replace = gString.Of(VBA.Replace$(This.Value, Find, ReplaceWith)): End Function
Public Function IsLike(ByVal Pattern As String) As Boolean: IsLike = (This.Value Like Pattern): End Function
Public Function IndexOf(ByRef Text As String, Optional ByVal Culture As VBA.VbCompareMethod = VBA.VbCompareMethod.vbBinaryCompare) As Long: IndexOf = VBA.InStr(1, This.Value, Text, Culture) - 1: End Function

Public Function StartsWith(ByRef Prefix As String, Optional ByVal IgnoreCase As Boolean = True, Optional ByVal BeginAt As Long = NOT_SPECIFIED) As Boolean
    
    '@Ignore IntegerDataType
    Const CASE_SHIFT As Integer = 32
    
    Dim prefixLength As Long
    Dim i As Long, j As Long, p As Long
    '@Ignore IntegerDataType
    '@Ignore IntegerDataType
    Dim A As Integer, B As Integer
    
    prefixLength = VBA.Len(Prefix)
    
    i = 1
    j = 1
    
    If (BeginAt > This.Length - prefixLength) Then _
        Exit Function
        
    If Not (BeginAt = NOT_SPECIFIED) Then _
        i = BeginAt + i
    
    If IgnoreCase Then
        For p = 1 To prefixLength
            A = VBA.AscW(VBA.mid$(This.Value, i, 1))
            B = VBA.AscW(VBA.mid$(Prefix, p, 1))
            If Not A = B Then _
                If Not A - CASE_SHIFT = B Then _
                    If Not A = B - CASE_SHIFT Then _
                        Exit Function
            i = i + 1
        Next
    Else
        Do While (prefixLength > 0)
            If Not (VBA.AscW(VBA.mid$(This.Value, i, 1))) = VBA.AscW(VBA.mid$(Prefix, j, 1)) Then _
                Exit Function
            i = i + 1
            j = j + 1
            prefixLength = prefixLength - 1
        Loop
    End If
    
    StartsWith = True
    
End Function

Public Function EndsWith(ByRef Suffix As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean: EndsWith = StartsWith(Suffix, IgnoreCase, This.Length - VBA.Len(Suffix)): End Function

Public Function IsEmptyOrWhiteSpace() As Boolean
    
    '@Ignore IntegerDataType
    Const UNICODE_WHITESPACE As Integer = 32
    Dim i As Long
    
    For i = 1 To This.Length
        If Not VBA.AscW(VBA.mid$(This.Value, i, 1)) = UNICODE_WHITESPACE Then _
            Exit Function
    Next
    
    IsEmptyOrWhiteSpace = True
    
End Function

Public Function Trim() As gString
    
    '@Ignore IntegerDataType
    Const UNICODE_WHITESPACE As Integer = 32
    Dim Count As Long
    Dim i As Long, j As Long
    
    For i = 1 To This.Length
        If Not VBA.AscW(VBA.mid$(This.Value, i, 1)) = UNICODE_WHITESPACE Then _
            Exit For
    Next

    For j = This.Length To i + 1 Step -1
        If Not VBA.AscW(VBA.mid$(This.Value, j, 1)) = UNICODE_WHITESPACE Then _
            Exit For
    Next
    
    Count = j - i + 1
    
    Select Case Count
        Case This.Length: Set Trim = Me
        Case 0: Set Trim = gString
        Case Else: Set Trim = gString.Of(This.Value, BeginAt:=i - 1, Count:=Count) 'i - 1 ' we re zero based
    End Select
    
End Function

Public Function Clean(Optional ByVal Chars As String = "!? @#$%^&()\/{}[]_.;:äüö*+-") As gString
    
    Dim i As Long
    Dim cleanString As String
    
    For i = 1 To VBA.Len(Chars): cleanString = VBA.Replace$(This.Value, VBA.mid$(Chars, i, 1&), vbNullString): Next i
    
    Set Clean = gString.Of(cleanString)

End Function

'@Ignore IntegerDataType
Public Function CountChar(ByVal ASCII As Integer) As Long
    
    Dim i As Long
    
    For i = 1 To Me.Length
        If VBA.AscW(VBA.mid$(This.Value, i, 1)) = ASCII Then _
            CountChar = CountChar + 1
    Next
    
End Function

Public Function Slice(ByVal BegintAt As Long, Optional ByVal Count As Long = NOT_SPECIFIED) As gString
  
    If (Count <= 0) Then _
        Count = This.Length - BegintAt
        
    If (This.Length < BegintAt + Count) Then _
        Call GenericError.Build(Me, "Public Function Slice").IndexOutofRange("(This.Length < BegintAt + Count)")
    
    Set Slice = Me
    
    If Not (Count = This.Length) Then _
        Set Slice = gString.Of(This.Value, BeginAt:=BegintAt, Count:=Count)
    
End Function

Public Function Repeat(ByVal Number As Long, ByRef Text As String, Optional ByVal Separator As String = vbNullString) As gString
    
    Dim i As Long
        
    If Number <= 0 Then _
        Set Repeat = gString.Of(DEFAULT_VALUE): Exit Function
        
    ReDim Strings(1 To Number) As String
    
    For i = 1 To Number: Strings(i) = Text: Next
    
    Set Repeat = gString.Of(VBA.Join$(Strings, Separator))

End Function

Public Function Split(ByVal Separator As String) As GenericArray
    
    Dim i As Long
    Dim parts As GenericArray
    Dim Strings As Variant
    
    If (Separator = vbNullString) Then _
        Set Split = Me.ToArray: Exit Function
    
    Set parts = GenericArray
    
    If (This.Length = 0) Then _
        Exit Function
      
    Strings = VBA.Split(This.Value, Separator)
    i = UBound(Strings)
    
    If (i > 0) Then
        Set parts = GenericArray.Build(i + 1)
        
        With parts
            For i = LBound(Strings) To UBound(Strings): Set .ElementAt(i) = gString.Of(CStr(Strings(i))): Next
        End With
    End If
    
    Set Split = parts
    
End Function

Public Function ToArray() As GenericArray
    
    Dim Chars As GenericArray
    
    Set Chars = GenericArray
    
    If Not (This.Length = 0) Then
        Set Chars = GenericArray.Build(This.Length)
        Call Me.CopyTo(Target:=Chars, Index:=Chars.LowerBound)
    End If
    
    Set ToArray = Chars
    
End Function

Public Sub CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    Dim i As Long
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").IsNothing("Target")
    
    With Target
    
        If (Index < .LowerBound) Or (Index >= .Length) Then _
            Call GenericError.Build(Me, "Public Sub CopyTo").IndexOutofRange("(Index < 0) Or (Index > Target.Length)")
    
        If (.Length - Index + .LowerBound < This.Length) Then _
            Call GenericError.Build(Me, "Public Sub CopyTo").InvalidInput("(Target.Length - Index < String.Length)")
        
        For i = 0 To This.Length - 1: Set .ElementAt(Index + i) = gString.Of(Value:=This.Value, BeginAt:=i, Count:=1): Next
        
    End With
    
End Sub

Public Function Concat(ByVal A As gString, ByVal B As gString) As gString

    Dim Buffer As String
    
    If (A.IsEmpty) Then _
        Set Concat = B: Exit Function

    If (B.IsEmpty) Then _
        Set Concat = A: Exit Function

    Buffer = VBA.Space$(A.Length + B.Length)
    
    Mid(Buffer, 1, A.Length) = A.ToString
    Mid(Buffer, A.Length + 1, B.Length) = B.ToString

    Set Concat = gString.Of(Buffer)

End Function

Public Function Insert(ByVal Index As Long, ByRef Text As String) As gString
    
    Dim Buffer As String
    Dim textLength As Long
    
    If (Index > This.Length) Then _
        Call GenericError.Build(Me, "Public Function Insert").IndexOutofRange("(Index > This.Length)")
            
    If (Index = This.Length) Then _
        Set Insert = Me.Append(Text): Exit Function
                        
    If (Text = vbNullString) Then _
        Set Insert = Me: Exit Function
    
    If (This.Length = 0) Then _
        Set Insert = gString.Of(Text): Exit Function
    
    textLength = VBA.Len(Text)
    
    Buffer = VBA.Space$(This.Length + textLength)
    
    Mid(Buffer, 1, Index) = VBA.Left$(This.Value, Index)
    Mid(Buffer, Index + 1, textLength) = Text
    Mid(Buffer, Index + textLength + 1, This.Length - Index) = VBA.Right$(This.Value, This.Length - Index)
    
    Set Insert = gString.Of(Buffer)
        
End Function

Public Function Append(ByRef Text As String) As gString
   
    Dim Buffer As String
    Dim textLength As Long
    
    If (Text = vbNullString) Then _
        Set Append = Me: Exit Function
        
    If (This.Length = 0) Then _
        Set Append = gString.Of(Text): Exit Function
    
    textLength = VBA.Len(Text)
    
    Buffer = VBA.Space$(This.Length + textLength)
    Mid(Buffer, 1, This.Length) = This.Value
    Mid(Buffer, This.Length + 1, textLength) = Text
        
    Set Append = gString.Of(Buffer)
  
End Function

Public Function Join(ByVal Sequence As IGenericSequence, Optional ByVal Separator As String = vbNullString) As gString
    
    Dim Buffer As String, tmp As String, Capacity As Long, OldSize As Long, NewBound As Long
    Dim Chars As String, CharSize As Long, SeparatorSize As Long
    Dim i As Long
    
    '@Ignore VariableNotAssigned
    Dim Element As IGeneric, FirstElement As IGeneric
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Join").IsNothing("Sequence")
    
    Set Join = gString
    
    With Sequence.Iterator
    
        If Not .HasNext(Element) Then _
            Exit Function
            
        Set FirstElement = Element
            
        If Not .HasNext(Element) Then
            If Not (FirstElement Is Nothing) Then
                If TypeOf FirstElement Is gString Then
                    Set Join = FirstElement
                Else
                    Set Join = gString.Of(FirstElement.ToString)
                End If
            End If
            Exit Function
        End If
        
        If Not (FirstElement Is Nothing) Then
            Chars = FirstElement.ToString
            CharSize = VBA.Len(Chars)
        End If
        
        SeparatorSize = VBA.Len(Separator)
        Capacity = INITIAL_BUFFER_SIZE + SeparatorSize + CharSize
        
        Buffer = VBA.Space$(Capacity)
        
        If (CharSize > 0) Then _
            Mid(Buffer, 1, CharSize) = Chars
        i = CharSize + 1
         
        If Not (SeparatorSize = 0) Then

            Do
                If Not (Element Is Nothing) Then
                    Chars = Element.ToString
                Else
                    Chars = "NULL"
                End If
                
                CharSize = VBA.Len(Chars)
                
                If (CharSize > 0) Then
                    
                    NewBound = i + CharSize + SeparatorSize
                    
                    If (NewBound > Capacity) Then
    
                        OldSize = Capacity
                        Capacity = Capacity * 2
                        
                        If (NewBound > Capacity) Then _
                            Capacity = NewBound * 2 + 1
                        
'                            Buffer = VBA.Space$(Capacity - OldSize) & Buffer
                        tmp = Buffer
                        Buffer = VBA.Space$(Capacity)
                        Mid(Buffer, 1, i) = tmp
                        
                    End If
                    
                    Mid(Buffer, i, SeparatorSize) = Separator
                    i = i + SeparatorSize
                    
                    Mid(Buffer, i, CharSize) = Chars
                    i = NewBound
                    
                End If
                
            Loop While .HasNext(Element)
            
        Else
            
            Do
                If Not (Element Is Nothing) Then
                
                    Chars = Element.ToString
                    CharSize = VBA.Len(Chars)
                    
                    If (CharSize > 0) Then
                    
                        NewBound = i + CharSize
                        
                        If (NewBound > Capacity) Then
                        
                            OldSize = Capacity
                            Capacity = Capacity * 2
                            
                            If (NewBound > Capacity) Then _
                                Capacity = NewBound * 2 + 1
                            
                            Buffer = (Buffer & VBA.Space$(Capacity - OldSize))
                            
                        End If
                        
                        Mid(Buffer, i, CharSize) = Chars
                        i = NewBound
                    
                    End If
                End If
            Loop While .HasNext(Element)
        
        End If
        
        i = i - 1
        If (i = Capacity) Then 'edge case
            Set Join = gString.Of(Buffer)
        Else
            Set Join = gString.Of(Value:=Buffer, BeginAt:=0, Count:=i)
        End If
        
    End With

End Function
'Copymemory based approach for Join
Public Function Concatenate(ByVal Sequence As IGenericSequence) As gString
    
    Dim Buffer() As Byte, Capacity As Long, NewBound As Long
    Dim Chars As String, CharSize As Long
    Dim i As Long
    
    Dim Element As IGeneric, FirstElement As IGeneric
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Concatenate").IsNothing("Sequence")
    
    Set Concatenate = gString

    With Sequence.Iterator
    
        If Not .HasNext(Element) Then _
            Exit Function
            
        Set FirstElement = Element
            
        If Not .HasNext(Element) Then _
            If Not (FirstElement Is Nothing) Then _
                Set Concatenate = gString.Of(FirstElement.ToString): Exit Function
        
        If Not (FirstElement Is Nothing) Then
            Chars = FirstElement.ToString
            CharSize = VBA.LenB(Chars)
        End If

        Capacity = (INITIAL_BUFFER_SIZE * 2) + CharSize
        
        ReDim Buffer(Capacity)
        
        If (CharSize > 0) Then _
            Call CopyMemory(Dst:=ByVal VarPtr(Buffer(0)), Src:=ByVal StrPtr(Chars), Length:=CharSize)
        i = CharSize
        
        Do
            If Not (Element Is Nothing) Then
            
                Chars = Element.ToString
                CharSize = VBA.LenB(Chars)
                
                If (CharSize > 0) Then
                
                    NewBound = i + CharSize
                    
                    If (NewBound > Capacity) Then
                        Capacity = Capacity * 2
                        If (NewBound > Capacity) Then _
                            Capacity = NewBound * 2 + 1
                        ReDim Preserve Buffer(Capacity)
                    End If
                   
                    Call CopyMemory(Dst:=ByVal VarPtr(Buffer(i)), Src:=ByVal StrPtr(Chars), Length:=CharSize)
                    i = NewBound
                
                End If
            End If
        Loop While .HasNext(Element)
        
        i = i \ 2
        If (i = Capacity) Then 'edge case
            Set Concatenate = gString.Of(CStr(Buffer))
        Else
            Set Concatenate = gString.Of(VBA.Left$(Buffer, i))
        End If
        
    End With

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = gString: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long
  
    If (This.Hash = 0) Then
        If (This.Length > 0) Then
            
            Dim i As Long
            Dim B() As Byte

            This.Hash = 5381 '7919
            B = VBA.StrConv(This.Value, vbFromUnicode)
            For i = 0 To This.Length - 1: This.Hash = (i + B(i) + This.Hash) * 37& And &H7FFFFF: Next
        End If
    End If
    
    IGeneric_HashCode = This.Hash
    
End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Property
        
    If Not (TypeOf Other Is gString) Then _
        Exit Property
        
    If Not (TypeOf Other Is IGenericValue) Then _
        Exit Property
        
    IGeneric_IsRelatedTo = True
    
End Property

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Function
        
    If TypeOf Other Is gString Then
        Dim OtherString As gString
        Set OtherString = Other
        IGeneric_Equals = OtherString.EqualsTo(This.Value)
    End If

'    IGeneric_Equals = CSTR_EQUAL = CompareString(0&, NORM_IGNORESYMBOLS Or SORT_STRINGSORT, This.Value, This.Length, S.Value, S.Length)

End Function

Private Function IGeneric_ToString() As String: IGeneric_ToString = This.Value: End Function

'IGenericValue
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_GetValueOrDefault(Optional ByVal DefaultValue As Variant = Null) As Variant: IGenericValue_GetValueOrDefault = VBA.IIf(This.Length = 0, DefaultValue, This.Value): End Property
Private Property Get IGenericValue_HasValue() As Boolean: IGenericValue_HasValue = Not (This.Length = 0): End Property
    
Private Property Get IGenericValue_VarType() As VBA.VbVarType: IGenericValue_VarType = VBA.VbVarType.vbString: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(This.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.VarWChar_Type: End Property
Private Property Get IGenericValue_GetValue() As Variant: IGenericValue_GetValue = This.Value: End Property
Private Property Get IGenericValue_ToString() As String: IGenericValue_ToString = This.Value: End Property

Private Function IGenericValue_HashCode() As Long

    If (This.Hash = 0) Then
        If (This.Length > 0) Then
            
            Dim i As Long
            Dim B() As Byte

            This.Hash = 5381 '7919
            B = VBA.StrConv(This.Value, vbFromUnicode)
            For i = 0 To This.Length - 1: This.Hash = (i + B(i) + This.Hash) * 37& And &H7FFFFF: Next
        End If
    End If
    
    IGenericValue_HashCode = This.Hash

End Function

Private Function IGenericValue_Of(ByRef Value As Variant) As IGenericValue: Set IGenericValue_Of = gString.Of(CStr(Value)): End Function

Private Function IGenericValue_Equals(ByVal Other As IGenericValue) As Boolean: IGenericValue_Equals = (This.Value = Other.GetValue): End Function
Private Function IGenericValue_IsLessThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsLessThan = (This.Value < Other.GetValue): End Function
Private Function IGenericValue_IsGreaterThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsGreaterThan = (This.Value > Other.GetValue): End Function

Private Function IGenericValue_IsBetween(ByVal A As IGenericValue, ByVal B As IGenericValue) As Boolean
    
    If (A Is Nothing) Then _
        Exit Function
    
    If (B Is Nothing) Then _
        Exit Function
    
    If (A.HasValue) Then _
        If (B.HasValue) Then _
            IGenericValue_IsBetween = Me.IsBetween(A.GetValue, B.GetValue)
    
End Function

Private Function IGenericValue_ToSqlParameter( _
                                                 Optional ByVal Direction As ISqlParameterDirection = 1&, _
                                                 Optional ByVal Name As String) As Object
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Name = Name
        .Attributes = 64 'adParamNullable '64
        .Type = IGenericValue_VarType
        
        If (This.Length > 0) Then
            .Size = This.Length
            .Value = This.Value
        Else
            .Size = 1
            .Value = Null
        End If
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function

Public Function Reverse() As IGenericIterator: Set Reverse = GenericIndexIterator.Build(Me, 0, This.Length, True): End Function

Public Function Range(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As IGenericIterator
    
    If (BeginAt = NOT_SPECIFIED) Then
        BeginAt = 0
    ElseIf (BeginAt < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("BeginAt < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Length - BeginAt
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("Length < 0")
    End If
    
    If (This.Length - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Range").InvalidInput("(This.Length - BeginAt < Count)")
        
    Set Range = GenericIndexIterator.Build(Me, BeginAt, Count)
    
End Function

'IGenericCollection
Private Property Get IGenericCollection_HasRandomAccess() As Boolean: IGenericCollection_HasRandomAccess = True: End Property
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = This.Length: End Property

Private Function IGenericCollection_GetAt(ByVal Index As Long) As Object

    If (Index >= This.Length) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= This.Length)")
    
    Set IGenericCollection_GetAt = gString.Of(This.Value, Index, 1&)

End Function

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean: IGenericCollection_Contains = False: End Function
Private Function IGenericCollection_ContainsAll(ByVal Sequence As IGenericSequence) As Boolean
 
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_ContainsAll").IsNothing("Sequence")
    
    If (This.Length = 0) Then _
        Exit Function
        
    IGenericCollection_ContainsAll = True

End Function

'ToArray returns a new Object array containing the contents of the String.
'This requires copying the List, which is an O(n) operation.
Private Function IGenericCollection_ToArray() As GenericArray
    
    Set IGenericCollection_ToArray = GenericArray 'avoid null
    
    If (This.Length = 0) Then _
        Exit Function
        
    Set IGenericCollection_ToArray = GenericArray.Build(This.Length)
    Call Me.CopyTo(Target:=IGenericCollection_ToArray, Index:=IGenericCollection_ToArray.LowerBound)

End Function

Private Sub IGenericCollection_CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    If (This.Length = 0) Then _
        Exit Sub
    
    Call Me.CopyTo(Target:=Target, Index:=Index)
    
End Sub

'@Ignore EmptyMethod
Private Sub IGenericCollection_Clear()
End Sub

Private Function IGenericCollection_Copy() As IGenericCollection: Set IGenericCollection_Copy = Me: End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range(0, This.Length): End Function

'IGenericSequence
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range(0, This.Length): End Function
