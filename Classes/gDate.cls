VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Primitives"
Option Explicit

Implements IGeneric
Implements IGenericValue

Private Const MAX_LONG As Long = 2 ^ 31 - 1
Private Const DEFAULT_FORMAT As String = "yyyy-MM-dd hh:mm:ss"
Private Const DEFAULT_DATE As Date = #12:00:00 AM#

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
Private This As Member

Public Property Get IsDefault() As Boolean: IsDefault = Me Is gDate: End Property
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
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
'@DefaultMember
Public Property Get ToDate() As Date: ToDate = This.Value: End Property
Attribute ToDate.VB_UserMemId = 0
Public Function BuildSerial(ByVal day As Integer, ByVal month As Integer, ByVal Year As Integer) As gDate: Set BuildSerial = gDate.Of(VBA.DateSerial(Year, month, day)): End Function

Public Function Of(ByVal Value As Date) As gDate
    
    Select Case VBA.VarType(Value)
        Case VBA.VbVarType.vbNull, VBA.VbVarType.vbEmpty
            Set Of = gDate
        Case VBA.VbVarType.vbDate
            This.Value = Value
            Set Of = New gDate
        Case VBA.VbVarType.vbString
            This.Value = gDate.Convert(CStr(Value), DEFAULT_FORMAT)
            Set Of = New gDate
        Case Else
            This.Value = Value
            Set Of = New gDate
    End Select
    
    This.Value = DEFAULT_DATE
     
End Function

Private Sub Class_Initialize(): This.Value = gDate.ToDate: End Sub

Public Function Convert(ByRef StringDate As String, Optional ByRef Format As String = DEFAULT_FORMAT) As Date
   
    Dim y As Long, m As Long, d As Long, h As Long, Min As Long, s As Long
    Dim am As Boolean, pm As Boolean
    Dim pos As Long

    If VBA.Len(StringDate) <> VBA.Len(Format) Then _
        Exit Function
    
    pos = VBA.InStr(1, Format, "yyyy", vbTextCompare)
    If pos > 0 Then
        y = VBA.Val(VBA.mid$(StringDate, pos, 4))
    Else
        pos = VBA.InStr(1, Format, "yy", vbTextCompare)
        If pos > 0 Then
            y = VBA.Val(VBA.mid$(StringDate, pos, 2))
            If y < 80 Then
                y = y + 2000
            Else
                y = y + 1900
            End If
        End If
    End If

    pos = VBA.InStr(1, Format, "mmm", vbTextCompare)
    If pos > 0 Then
        m = VBA.month(VBA.DateValue("01 " & (VBA.mid$(StringDate, pos, 3)) & " 2000"))
    Else
        pos = VBA.InStr(1, Format, "MM", vbBinaryCompare)
        If pos > 0 Then _
            m = VBA.Val(VBA.mid$(StringDate, pos, 2))
    End If

    pos = VBA.InStr(1, Format, "dd", vbTextCompare)
    If pos > 0 Then _
        d = VBA.Val(VBA.mid$(StringDate, pos, 2))

    pos = VBA.InStr(1, Format, "hh", vbTextCompare)
    If pos > 0 Then _
        h = VBA.Val(VBA.mid$(StringDate, pos, 2))
    
    Select Case True
        Case VBA.InStr(1, StringDate, "am", vbTextCompare) > 0: am = True
        Case VBA.InStr(1, StringDate, "a.m.", vbTextCompare) > 0: am = True
        Case VBA.InStr(1, StringDate, "a. m.", vbTextCompare) > 0: am = True
        Case VBA.InStr(1, StringDate, "pm", vbTextCompare) > 0: pm = True
        Case VBA.InStr(1, StringDate, "p.m.", vbTextCompare) > 0: pm = True
        Case VBA.InStr(1, StringDate, "p. m.", vbTextCompare) > 0: pm = True
    End Select
    
    If am And h = 12 Then
        h = 0
    ElseIf pm And h <> 12 Then
        h = h + 12
    End If

    pos = VBA.InStr(1, Format, "mm", vbBinaryCompare)
    If pos > 0 Then _
        Min = VBA.Val(VBA.mid$(StringDate, pos, 2))

    pos = VBA.InStr(1, Format, "ss", vbTextCompare)
    If pos > 0 Then _
        s = VBA.Val(VBA.mid$(StringDate, pos, 2))

    Convert = VBA.DateSerial(y, m, d) + VBA.TimeSerial(h, Min, s)

End Function

Public Function IsBefore(ByVal A As Date) As Boolean: IsBefore = A < This.Value: End Function
Public Function IsAfter(ByVal A As Date) As Boolean: IsAfter = A > This.Value: End Function

Public Function MaxOf(ByVal A As Date, ByVal B As Date) As gDate: Set MaxOf = gDate.Of(VBA.IIf(A > B, A, B)): End Function
Public Function MinOf(ByVal A As Date, ByVal B As Date) As gDate: Set MinOf = gDate.Of(VBA.IIf(A < B, A, B)): End Function
Public Function EqualsTo(ByVal Value As Date) As Boolean: EqualsTo = This.Value = Value: End Function
Public Function CompareTo(ByVal Value As Date) As CompareResult: CompareTo = Compare(This.Value, Value): End Function

Public Function Compare(ByVal A As Date, ByVal B As Date) As CompareResult

    Select Case True
        Case A > B: Compare = isgreater
        Case A < B: Compare = IsLower
        Case Else: Compare = IsEqual
    End Select
    
End Function

Public Function IsBetween(ByVal A As Date, ByVal B As Date) As Boolean
    
    If (A > This.Value) Then _
        Exit Function
    
    If (B < This.Value) Then _
        Exit Function

    IsBetween = True
    
End Function

Public Function Add(ByVal Number As Long, ByVal i As DateInterval) As gDate: Set Add = gDate.Of(VBA.DateAdd(Me.IntervalName(i), Number, This.Value)): End Function
Public Function DifferenceTo(ByVal d As gDate, ByVal i As DateInterval) As Long: DifferenceTo = VBA.DateDiff(Me.IntervalName(i), This.Value, d.ToDate): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = gDate: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = (CLng(This.Value - (VBA.Fix(This.Value / MAX_LONG) * MAX_LONG)) Xor CLng((This.Value - VBA.Int(This.Value)) * 10000&)): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean
    
    If Not (Other Is Nothing) Then
        IGeneric_IsRelatedTo = (TypeOf Other Is gDate)
    Else
        IGeneric_IsRelatedTo = True
    End If
        
End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean
    
    If (Other Is Nothing) Then _
        Exit Function
        
    If TypeOf Other Is gDate Then
        Dim d As gDate: Set d = Other
        IGeneric_Equals = (This.Value = d.ToDate)
    End If
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = VBA.Format$(This.Value, DEFAULT_FORMAT): End Function

'IGenericValue
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_GetValueOrDefault(Optional ByVal DefaultValue As Variant = Null) As Variant: IGenericValue_GetValueOrDefault = VBA.IIf(Me Is gDate, DefaultValue, This.Value): End Property
Private Property Get IGenericValue_HasValue() As Boolean: IGenericValue_HasValue = Not (Me Is gDate): End Property
Private Property Get IGenericValue_VarType() As VBA.VbVarType: IGenericValue_VarType = VBA.VbVarType.vbDate: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(This.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.DBDate_Type: End Property
Private Property Get IGenericValue_GetValue() As Variant: IGenericValue_GetValue = This.Value: End Property
Private Property Get IGenericValue_ToString() As String: IGenericValue_ToString = This.Value: End Property
Private Function IGenericValue_HashCode() As Long: IGenericValue_HashCode = (CLng(This.Value - (VBA.Fix(This.Value / MAX_LONG) * MAX_LONG)) Xor CLng((This.Value - VBA.Int(This.Value)) * 10000&)): End Function
Private Function IGenericValue_Of(ByRef Value As Variant) As IGenericValue: Set IGenericValue_Of = gDate.Of(Value): End Function

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

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Type = IGenericValue_VarType
        .Name = Name
        .Value = This.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function
