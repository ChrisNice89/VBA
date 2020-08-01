VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "<T>Class"
Option Explicit

Implements IGeneric
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
Public Function Concat(ByVal text As GString, Optional ByVal Delimiter = " ") As GString: Set Concat = GString(myClass.Value & Delimiter & text.Value, myClass.Hashfunction, myClass.Format): End Function
Public Function Replace(ByVal Find As String, ReplaceWith As String) As GString: Set Replace = GString(VBA.Replace$(myClass.Value, Find, ReplaceWith), myClass.Hashfunction, myClass.Format): End Function
Public Property Get Format() As TStringFormat: Format = myClass.Format: End Property
Public Property Get Value() As String: Value = myClass.Value: End Property
Public Function ToByteArray() As Byte(): ToByteArray = StrConv(myClass.Value, vbFromUnicode): End Function

Public Function IsLike(ByVal Pattern As String) As Boolean: IsLike = (myClass.Value Like Pattern): End Function
Public Function StartsWith(ByVal text As String) As Boolean: StartsWith = (myClass.Value Like ("%" & text)): End Function
Public Function EndsWith(ByVal text As String) As Boolean: EndsWith = (myClass.Value Like (text & "%")): End Function
Public Function Contains(ByVal text As String) As Boolean: Contains = Not (InStr(1, myClass.Value, text) = 0): End Function

Public Function Split(ByVal Delimiter As String) As GString()

End Function

Public Function Join(ByRef Strings() As GString, Optional ByVal Delimiter) As GString
    
    
   
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
        .Format = Format
        .Hashfunction = Hashfunction
    End With
    
    Set Build = New GString    'faster then Skynet
    'Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    If Not Me Is GString Then
         With myClass
            .Value = GString.Value
            .Format = GString.Format
            .Hashfunction = GString.Hashfunction
        End With
        Call Skynet.Generic(GString).Dispose
    End If
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType: IValue_DatabaseType = T_String: End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not TypeOf Other Is GString Then IGeneric_Equals = False: Exit Function

    Dim S As GString: Set S = Other
    If Not (S.Length = Me.Length) Then IGeneric_Equals = False: Exit Function

    IGeneric_Equals = (StrComp(myClass.Value, S.Value, vbBinaryCompare) = 0)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Dim S As GString
    Dim i As Long
    
    If TypeOf Other Is GString Then
        Set S = Other
        Select Case StrComp(myClass.Value, S.Value, vbBinaryCompare)
            Case 1: IGeneric_CompareTo = IsGreater
            Case -1: IGeneric_CompareTo = IsLower
            Case Else: IGeneric_CompareTo = IsEqual
        End Select
    End If
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
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GString: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

