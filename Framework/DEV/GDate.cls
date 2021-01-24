VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDate"
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

Public Enum DateFormat
    [_First]
    DEFAULTDATE
    DDMMYYYY
    MMDDYYYY
    [_Last]
End Enum

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
Private Interval() As String

Public Property Get IsDefault() As Boolean: IsDefault = Me Is GDate: End Property
Public Property Get IntervalName(ByVal i As DateInterval) As String: IntervalName = Interval(i): End Property
Public Property Get Value() As Date: Value = myclass.Value: End Property

Public Property Get ToDay() As String
    'ToDay = Format(DateSerial2016, 11, 4), "dddd")
End Property

Public Function BuildSerial(ByVal day As Integer, ByVal month As Integer, ByVal Year As Integer) As GDate: Set BuildSerial = GDate(VBA.DateSerial(Year, month, day)): End Function
Public Function BuildFromString(ByVal StringDate As String, ByVal Format As DateFormat) As GDate: Set BuildFromString = GDate(Me.Convert(StringDate, Format)): End Function

Public Function Build(ByVal Value As Date) As GDate
Attribute Build.VB_UserMemId = 0
    
    
    myclass.Value = Value
    Set Build = New GDate: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    
    myclass.Value = GDate.Value
    
    ReDim Interval(DateInterval.[_First] To DateInterval.[_Last])
    Interval(DateInterval.day) = "d"
    Interval(DateInterval.DayOfYear) = "y"
    Interval(DateInterval.Hour) = "h"
    Interval(DateInterval.Minute) = "n"
    Interval(DateInterval.month) = "m"
    Interval(DateInterval.Quarter) = "q"
    Interval(DateInterval.second) = "s"
    Interval(DateInterval.Week) = "ww"
    Interval(DateInterval.Weekday) = "w"
    Interval(DateInterval.Year) = "yyyy"

End Sub

Public Function Convert(ByVal d As String, ByVal Format As DateFormat) As Date
    
    Dim day As Integer, month As Integer, Year As Integer
    
    If Format = DateFormat.DDMMYYYY Then
        day = Int(VBA.Mid$(d, 1, 2))
        month = Int(VBA.Mid$(d, 4, 2))
        Year = Int(VBA.Mid$(d, 7, 4))
    ElseIf Format = DateFormat.DDMMYYYY Then
    
    Else
    
    End If
  
    Convert = VBA.DateSerial(Year, month, day)

End Function

Public Function Add(ByVal Number As Long, ByVal i As DateInterval) As GDate: Set Add = GDate(DateAdd(Me.IntervalName(i), Number, myclass.Value)): End Function
Public Function DifferenceTo(ByVal d As GDate, ByVal i As DateInterval) As Long: DifferenceTo = DateDiff(Me.IntervalName(i), myclass.Value, d.Value): End Function
'Todo
Private Function IsValid(ByVal day As Integer, ByVal month As Integer, ByVal Year As Integer) As Boolean

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GDate: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GDate):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is GDate Then _
        IGeneric_Equals = False: Exit Function
        
    Dim d As GDate: Set d = Other
    IGeneric_Equals = (myclass.Value = d.Value)
     
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Format$(myclass.Value, DEFAULT_FORMAT): End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbDate: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.DBTimeStamp_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericComparer: End Property

Private Function IGenericValue_Convert(ByVal Value As Variant) As IGenericValue: Set IGenericValue_Convert = Me.Build(Value): End Function

Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As CompareResult
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim otherDate As GDate:  Set otherDate = Other
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

