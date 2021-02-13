VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Class")
Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericValue

Private Const DEFAULT_FORMAT As String = "yyyy-MM-dd hh:mm:ss"
Private Const DEFAULT_DATE As Date = #1/1/1900#

Public Enum DateInterval
    [_First]
    Year
    Quarter
    month
    DayOfYear
    day
    Weekday
    Week
    Hour
    Minute
    second
    [_Last]
End Enum

Private Type Member
    Value As Date
   
End Type
Private myclass As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is GDateTime: End Property
Public Property Get IntervalName(ByVal i As DateInterval) As String

    Select Case i
        Case DateInterval.day: IntervalName = "d"
        Case DateInterval.DayOfYear: IntervalName = "y"
        Case DateInterval.Hour: IntervalName = "h"
        Case DateInterval.Minute: IntervalName = "n"
        Case DateInterval.month: IntervalName = "m"
        Case DateInterval.Quarter: IntervalName = "q"
        Case DateInterval.second: IntervalName = "s"
        Case DateInterval.Week: IntervalName = "ww"
        Case DateInterval.Weekday: IntervalName = "w"
        Case DateInterval.Year: IntervalName = "yyyy"
    End Select

End Property

Public Property Get Value() As Date: Value = myclass.Value: End Property

Public Function BuildSerial(ByVal day As Integer, ByVal month As Integer, ByVal Year As Integer) As GDateTime: Set BuildSerial = GDateTime(VBA.DateSerial(Year, month, day)): End Function

Public Function Build(ByVal Value As Date) As GDateTime
Attribute Build.VB_UserMemId = 0
    
    myclass.Value = Value
    Set Build = New GDateTime: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    myclass.Value = GDateTime.Value
End Sub

Public Function Convert(ByVal StringDate As String, Optional ByRef Format As String = DEFAULT_FORMAT) As Date
   
    Dim y As Long, m As Long, d As Long, h As Long, min As Long, s As Long
    Dim am As Boolean, pm As Boolean
    Dim pos As Long

    If VBA.Len(StringDate) <> VBA.Len(Format) Then _
        Exit Function
    
    pos = VBA.InStr(1, Format, "yyyy", vbTextCompare)
    If pos > 0 Then
        y = VBA.Val(Mid$(StringDate, pos, 4))
    Else: pos = VBA.InStr(1, Format, "yy", vbTextCompare)
        If pos > 0 Then
            y = VBA.Val(Mid$(StringDate, pos, 2))
            If y < 80 Then y = y + 2000 Else y = y + 1900
        End If
    End If

    pos = InStr(1, Format, "mmm", vbTextCompare)
    If pos > 0 Then
        m = VBA.month(DateValue("01 " & (Mid$(StringDate, pos, 3)) & " 2000"))
    Else: pos = InStr(1, Format, "MM", vbBinaryCompare)
        If pos > 0 Then m = VBA.Val(VBA.Mid$(StringDate, pos, 2))
    End If

    pos = InStr(1, Format, "dd", vbTextCompare)
    If pos > 0 Then d = Val(Mid$(StringDate, pos, 2))

    pos = InStr(1, Format, "hh", vbTextCompare)
    If pos > 0 Then h = Val(Mid$(StringDate, pos, 2))
    If VBA.InStr(1, StringDate, "am", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "a.m.", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "a. m.", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "pm", vbTextCompare) > 0 Then pm = True
    If VBA.InStr(1, StringDate, "p.m.", vbTextCompare) > 0 Then pm = True
    If VBA.InStr(1, StringDate, "p. m.", vbTextCompare) > 0 Then pm = True
    If am And h = 12 Then h = 0
    If pm And h <> 12 Then h = h + 12

    pos = VBA.InStr(1, Format, "mm", vbBinaryCompare)
    If pos > 0 Then min = VBA.Val(VBA.Mid$(StringDate, pos, 2))

    pos = VBA.InStr(1, Format, "ss", vbTextCompare)
    If pos > 0 Then s = VBA.Val(VBA.Mid$(StringDate, pos, 2))

    Convert = VBA.DateSerial(y, m, d) + VBA.TimeSerial(h, min, s)

End Function

Public Function Add(ByVal Number As Long, ByVal i As DateInterval) As GDateTime: Set Add = GDateTime(DateAdd(Me.IntervalName(i), Number, myclass.Value)): End Function
Public Function DifferenceTo(ByVal d As GDateTime, ByVal i As DateInterval) As Long: DifferenceTo = DateDiff(Me.IntervalName(i), myclass.Value, d.Value): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GDateTime: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GDateTime):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is GDateTime Then _
        IGeneric_Equals = False: Exit Function
        
    Dim d As GDateTime: Set d = Other
    IGeneric_Equals = (myclass.Value = d.Value)
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Format$(myclass.Value, DEFAULT_FORMAT): End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbDate: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.DBDate_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericComparer: End Property

Private Function IGenericValue_Convert(ByVal Value As Variant) As IGenericValue

    If Isnull(Value) Then _
        Set IGenericValue_Convert = IGeneric_Default: Exit Function
    
    If VarType(Value) = vbString Then 'try parse
        Set IGenericValue_Convert = GDateTime.Build(Convert(Value))
    Else
        Set IGenericValue_Convert = GDateTime.Build(Value)
    End If
    
End Function

Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As CompareResult
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim otherDate As GDateTime:  Set otherDate = Other
    Dim i As Date
    i = otherDate.Value
    
    Select Case True
        Case myclass.Value > i: IGenericValue_CompareTo = IsGreater
        Case myclass.Value < i: IGenericValue_CompareTo = IsLower
        Case Else: IGenericValue_CompareTo = IsEqual
    End Select
    
End Function

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = IGenericValue_VarType
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function

