VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GdateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IGeneric
Implements GenericValue

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
    Second
    [_Last]
End Enum

Private Type Member
    Value As Date
   
End Type
Private myclass As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is GdateTime: End Property
Public Property Get IntervalName(ByVal i As DateInterval) As String

    Select Case i
        Case DateInterval.day: IntervalName = "d"
        Case DateInterval.DayOfYear: IntervalName = "y"
        Case DateInterval.Hour: IntervalName = "h"
        Case DateInterval.Minute: IntervalName = "n"
        Case DateInterval.month: IntervalName = "m"
        Case DateInterval.Quarter: IntervalName = "q"
        Case DateInterval.Second: IntervalName = "s"
        Case DateInterval.Week: IntervalName = "ww"
        Case DateInterval.Weekday: IntervalName = "w"
        Case DateInterval.Year: IntervalName = "yyyy"
    End Select

End Property

Public Property Get Value() As Date: Value = myclass.Value: End Property

Public Function BuildSerial(ByVal day As Integer, ByVal month As Integer, ByVal Year As Integer) As GdateTime: Set BuildSerial = GdateTime(VBA.DateSerial(Year, month, day)): End Function

Public Function Of(ByVal Value As Date) As GdateTime
    
    myclass.Value = Value
    Set Of = New GdateTime: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    myclass.Value = GdateTime.Value
End Sub

Public Function Convert(ByVal StringDate As String, Optional ByRef Format As String = DEFAULT_FORMAT) As Date
   
    Dim Y As Long, m As Long, D As Long, H As Long, Min As Long, s As Long
    Dim am As Boolean, pm As Boolean
    Dim pos As Long

    If VBA.Len(StringDate) <> VBA.Len(Format) Then _
        Exit Function
    
    pos = VBA.InStr(1, Format, "yyyy", vbTextCompare)
    If pos > 0 Then
        Y = VBA.Val(mid$(StringDate, pos, 4))
    Else: pos = VBA.InStr(1, Format, "yy", vbTextCompare)
        If pos > 0 Then
            Y = VBA.Val(mid$(StringDate, pos, 2))
            If Y < 80 Then Y = Y + 2000 Else Y = Y + 1900
        End If
    End If

    pos = InStr(1, Format, "mmm", vbTextCompare)
    If pos > 0 Then
        m = VBA.month(DateValue("01 " & (mid$(StringDate, pos, 3)) & " 2000"))
    Else: pos = InStr(1, Format, "MM", vbBinaryCompare)
        If pos > 0 Then m = VBA.Val(VBA.mid$(StringDate, pos, 2))
    End If

    pos = InStr(1, Format, "dd", vbTextCompare)
    If pos > 0 Then D = Val(mid$(StringDate, pos, 2))

    pos = InStr(1, Format, "hh", vbTextCompare)
    If pos > 0 Then H = Val(mid$(StringDate, pos, 2))
    If VBA.InStr(1, StringDate, "am", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "a.m.", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "a. m.", vbTextCompare) > 0 Then am = True
    If VBA.InStr(1, StringDate, "pm", vbTextCompare) > 0 Then pm = True
    If VBA.InStr(1, StringDate, "p.m.", vbTextCompare) > 0 Then pm = True
    If VBA.InStr(1, StringDate, "p. m.", vbTextCompare) > 0 Then pm = True
    If am And H = 12 Then H = 0
    If pm And H <> 12 Then H = H + 12

    pos = VBA.InStr(1, Format, "mm", vbBinaryCompare)
    If pos > 0 Then Min = VBA.Val(VBA.mid$(StringDate, pos, 2))

    pos = VBA.InStr(1, Format, "ss", vbTextCompare)
    If pos > 0 Then s = VBA.Val(VBA.mid$(StringDate, pos, 2))

    Convert = VBA.DateSerial(Y, m, D) + VBA.TimeSerial(H, Min, s)

End Function

Public Function Add(ByVal Number As Long, ByVal i As DateInterval) As GdateTime: Set Add = GdateTime(DateAdd(Me.IntervalName(i), Number, myclass.Value)): End Function
Public Function DifferenceTo(ByVal D As GdateTime, ByVal i As DateInterval) As Long: DifferenceTo = DateDiff(Me.IntervalName(i), myclass.Value, D.Value): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GdateTime: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GdateTime):  End Property

Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is GdateTime Then _
        IGeneric_Equals = False: Exit Function
        
    Dim D As GdateTime: Set D = Other
    IGeneric_Equals = (myclass.Value = D.Value)
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Format$(myclass.Value, DEFAULT_FORMAT): End Function

'GenericValue
Private Property Get GenericValue_VarType() As VbVarType: GenericValue_VarType = vbDate: End Property
Private Property Get GenericValue_Size() As Byte: GenericValue_Size = LenB(myclass.Value): End Property
Private Property Get GenericValue_SqlType() As ISqlDataType: GenericValue_SqlType = ISqlDataType.DBDate_Type: End Property
Private Property Get GenericValue_ToValue() As Variant: GenericValue_ToValue = myclass.Value: End Property
Private Property Get GenericValue_Instance() As IGeneric: Set GenericValue_Instance = Me: End Property
Private Property Get GenericValue_Comparer() As IGenericComparer: Set GenericValue_Comparer = IGenericComparer: End Property

Private Function GenericValue_Convert(ByRef Value As Variant) As GenericValue

    If VBA.IsNull(Value) Then _
        Set GenericValue_Convert = IGeneric_Default: Exit Function
    
    If VarType(Value) = vbString Then 'try parse
        Set GenericValue_Convert = GdateTime.Of(Convert(Value))
    Else
        Set GenericValue_Convert = GdateTime.Of(Value)
    End If
    
End Function

Private Function GenericValue_CompareTo(ByVal Other As GenericValue) As CompareResult
    
    If Other Is Nothing Then _
        GenericValue_CompareTo = isgreater: Exit Function
    
    Dim otherDate As GdateTime:  Set otherDate = Other
    Dim i As Date
    i = otherDate.Value
    
    Select Case True
        Case myclass.Value > i: GenericValue_CompareTo = isgreater
        Case myclass.Value < i: GenericValue_CompareTo = islower
        Case Else: GenericValue_CompareTo = IsEqual
    End Select
    
End Function

Private Function GenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = GenericValue_VarType
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set GenericValue_ToSqlParameter = P
    
End Function



