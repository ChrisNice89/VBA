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

Public Enum TDateFormat
    [_First]
    DefaultDate
    DDMMYYYY
    MMDDYYYY
    [_Last]
End Enum

Public Enum DateInterval
    [_First]
    year
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
    Format As TDateFormat
    
End Type
Private myclass As Member
Private Interval() As String
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get ToDay() As String
    'Day = Format(DateSerial2016, 11, 4), "dddd")
End Property

Public Function Add(ByVal Number As Long, ByVal i As DateInterval) As GDate
    Set Add = GDate(DateAdd(Me.IntervalName(i), Number, myclass.Value), myclass.Format)
End Function

Public Function DifferenceTo(ByVal d As GDate, ByVal i As DateInterval) As Long
    DifferenceTo = DateDiff(Me.IntervalName(i), myclass.Value, d.Value)
End Function

Public Property Get IntervalName(ByVal i As DateInterval) As String
    IntervalName = Interval(i)
End Property

Public Property Get Value() As Date
    Value = myclass.Value
End Property

Public Function Build( _
                        ByVal Value As Date, _
                        Optional ByVal Format As DateInterval = DateInterval.year) As GDate
Attribute Build.VB_UserMemId = 0
    With myclass
        .Value = Value
        .Format = Format
    End With
    
    Set Build = Skynet.CreateInstance(New GDate, Virtual.Pointer, Virtual.Size)
    
End Function

Public Function Build_2( _
                            ByVal day As Integer, _
                            ByVal month As Integer, _
                            ByVal year As Integer, _
                            Optional ByVal Format As TDateFormat = TDateFormat.DefaultDate) As GDate
   
    Set Build_2 = GDate(VBA.DateSerial(year, month, day), myclass.Format)
End Function

Public Function Build_3(ByVal StringDate As String, ByVal Format As TDateFormat) As GDate
   
    Set Build_3 = GDate(Me.Convert(StringDate, Format))
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
    ReDim Interval(DateInterval.[_First] To DateInterval.[_Last])
    Interval(DateInterval.day) = "d"
    Interval(DateInterval.DayOfYear) = "y"
    Interval(DateInterval.Hour) = "h"
    Interval(DateInterval.Minute) = "n"
    Interval(DateInterval.month) = "m"
    Interval(DateInterval.Quarter) = "q"
    Interval(DateInterval.Second) = "s"
    Interval(DateInterval.Week) = "ww"
    Interval(DateInterval.Weekday) = "w"
    Interval(DateInterval.year) = "yyyy"
    
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Public Function Convert(ByVal d As String, ByVal Format As TDateFormat) As Date
    Dim day As Integer, month As Integer, year As Integer
    
    If Format = TDateFormat.DDMMYYYY Then
        day = Int(VBA.Mid$(d, 1, 2))
        month = Int(VBA.Mid$(d, 4, 2))
        year = Int(VBA.Mid$(d, 7, 4))
    ElseIf Format = TDateFormat.DDMMYYYY Then
    
    Else
    
    End If
  
    Convert = VBA.DateSerial(year, month, day)

End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GDate: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GDate):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass

    Set IGeneric_Clone = Skynet.CreateInstance(New GDate, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IGeneric_Equals = False: Exit Function

    Dim d As GDate: Set d = Other
   
    IGeneric_Equals = (myclass.Value = d.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Dim d As GDate
    
    If Other.IsRelatedTo(Me) Then
        Set d = Other
        Select Case True
            Case myclass.Value > d.Value: IGeneric_CompareTo = IsGreater
            Case myclass.Value < d.Value: IGeneric_CompareTo = IsLower
            Case Else: IGeneric_CompareTo = IsEqual
        End Select
    End If
End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = Format$(myclass.Value, Interval(myclass.Format))
End Function
