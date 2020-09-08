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
Implements IPrototype
Implements IValue

Private Const DEFAULT_DELIMITER As String = ","

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
Public Property Get Length() As Long: Length = myClass.Length: End Property
Public Property Get FirstChar() As String: FirstChar = VBA.Mid$(myClass.Value, 1, 1): End Property
Public Property Get LastChar() As String: LastChar = VBA.Mid$(myClass.Value, Len(myClass.Value), 1): End Property
Public Property Get Character(ByVal Index As Long) As String: Character = VBA.Mid$(myClass.Value, Index, 1): End Property
Public Property Get Format() As TStringFormat: Format = myClass.Format: End Property
Public Property Get Value() As String: Value = myClass.Value: End Property

Public Function Replace(ByVal Find As String, ReplaceWith As String) As GString: Set Replace = GString(VBA.Replace$(myClass.Value, Find, ReplaceWith), myClass.Hashfunction, myClass.Format): End Function
Public Function ToByteArray() As Byte(): ToByteArray = StrConv(myClass.Value, vbFromUnicode): End Function
Public Function IsLike(ByVal Pattern As String) As Boolean: IsLike = (myClass.Value Like Pattern): End Function
Public Function StartsWith(ByVal text As String) As Boolean: StartsWith = (myClass.Value Like ("%" & text)): End Function
Public Function EndsWith(ByVal text As String) As Boolean: EndsWith = (myClass.Value Like (text & "%")): End Function
Public Function Contains(ByVal text As String) As Boolean: Contains = Not (InStr(1, myClass.Value, text) = 0): End Function

Public Function CountChar(ByVal C As String) As Long
    
    Dim i As Long
    Dim ascii As Integer
    
    If Len(C) > 1 Then _
        Skynet.Throw(Me, "Private Function GetArrayStructure").NotSupported ("Len(C) > 1")

    ascii = AscW(C)
    
    For i = 1 To Me.Length
        If AscW(Mid$(myClass.Value, i, 1)) = ascii Then _
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
            Set Concat = GString(myClass.Value & Delimiter & Other.Value, myClass.Hashfunction, myClass.Format)
        Else
            Set Concat = GString(myClass.Value & Delimiter & Item.ToString, myClass.Hashfunction, myClass.Format)
        End If
    Else 'Static default call
        If Not Other Is Nothing Then
            Set Concat = Item.Clone ' must be a Gstring ;)
        Else
            Set Concat = GString(Item.ToString, myClass.Hashfunction, myClass.Format)
        End If
    End If
    
End Function

Public Function Split(ByVal Delimiter As String) As GenericArray
    
    Dim chars As Variant
    chars = VBA.Split(myClass.Value, Delimiter)
    
    Dim i As Long, lb As Long: lb = GenericArray.LowerBound
    ReDim Strings(lb To UBound(chars) + lb) As IGeneric
    
    For i = LBound(chars) To UBound(chars)
        Set Strings(i + lb) = GString(chars(i))
    Next
    
    Set Split = GenericArray.BuildFrom(Strings)
    
End Function

Public Function Join(ByVal Values As GenericArray, Optional ByVal Delimiter = DEFAULT_DELIMITER) As GString
    
    Dim s As IGeneric
    Dim i As Long
    
    If (Values Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Join").NullPointer("Values is nothing")
    
    ReDim Strings(Values.LowerBound To Values.Length) As String
    
    For i = Values.LowerBound To Values.Length
        Set s = Values(i) 'Should use enumerable.Select(IgnoreNulls:=true)
        If Not s Is Nothing Then _
            Strings(i) = s.ToString
    Next
    
    Set Join = Me.Concat(GString(VBA.Join$(Strings, Delimiter), myClass.Hashfunction, myClass.Format))
    
End Function

Public Function ToCharArray() As String()
    Dim chars() As String
    Dim i As Long

    With myClass
        ReDim chars(1 To Len(myClass.Value))
        For i = LBound(chars) To UBound(chars)
            chars(i) = VBA.Mid$(myClass.Value, i, 1)
        Next
    End With

    ToCharArray = chars
End Function

Public Function Build( _
                        ByVal Value As String, _
                        Optional ByVal Hashfunction As Hashing = Hashing.Function1, _
                        Optional ByVal Format As TStringFormat = TStringFormat.DefaultString) As GString
Attribute Build.VB_UserMemId = 0
    With myClass
        .Value = Value
        .Length = Len(Value)
        .Format = Format
        .Hashfunction = Hashfunction
    End With
    
    'Set Build = Skynet.Create(GString)
    Set Build = Skynet.CreateInstance(New GString, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
'    With GString
'        myClass.Value = .Value
'        myClass.Length = .Length
'        myClass.Format = .Format
'        myClass.Hashfunction = .Hashfunction
'    End With
    
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType: IValue_DatabaseType = T_String: End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GString, VarPtr(C), Virtual.Size)
   
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    Dim s As GString
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    Set s = Other
    If Not (myClass.Length = Me.Length) Then IGeneric_Equals = False: Exit Function

    IGeneric_Equals = (StrComp(myClass.Value, s.Value, vbBinaryCompare) = 0)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    Dim s As GString: Set s = Other
    IGeneric_CompareTo = StrComp(myClass.Value, s.Value, vbBinaryCompare)

End Function

Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = Skynet.HashValue(StrConv(myClass.Value, vbFromUnicode), myClass.Hashfunction): End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GString): End Function
Private Function IGeneric_ToString() As String

    Select Case myClass.Format
        Case TStringFormat.DefaultString: IGeneric_ToString = myClass.Value: Exit Function
        Case TStringFormat.Trim: IGeneric_ToString = VBA.Trim$(myClass.Value): Exit Function
        Case TStringFormat.LTrim: IGeneric_ToString = VBA.LTrim$(myClass.Value): Exit Function
        Case TStringFormat.RTrim: IGeneric_ToString = VBA.RTrim$(myClass.Value): Exit Function
    End Select
    
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GString: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GString: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

