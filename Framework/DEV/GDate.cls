VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IPrototype
Implements IValue

Public Enum TDateFormat
    [_First]
    DefaultDate
    
    [_Last]
End Enum

Public Enum DateInterval
    [_First]
    Year
    Quarter
    Month
    DayOfYear
    Day
    Weekday
    Week
    Hour
    Minute
    Second
    [_Last]
End Enum

Private Type Member
    Value As Date
    Format As TFloatFormat
   
End Type
Private myClass As Member
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
    Set Add = GDate(DateAdd(Me.IntervalName(i), Number, myClass.Value), myClass.Format)
End Function

Public Function DifferenceTo(ByVal d As GDate, ByVal i As DateInterval) As Long
    DifferenceTo = DateDiff(Me.IntervalName(i), myClass.Value, d.Value)
End Function

Public Property Get IntervalName(ByVal i As DateInterval) As String
    IntervalName = Interval(i)
End Property

Public Function Build2( _
                            ByVal Day As Integer, _
                            ByVal Month As Integer, _
                            ByVal Year As Integer, _
                            Optional ByVal Format As TDateFormat = TDateFormat.DefaultDate) As GDate
   
    Set Build2 = GDate(VBA.DateSerial(Year, Month, Day), myClass.Format)
End Function

Public Property Get Value() As Date
    Value = myClass.Value
End Property

Public Function Build( _
                        ByVal Value As Date, _
                        Optional ByVal Format As TDateFormat = TDateFormat.DefaultDate) As GDate
Attribute Build.VB_UserMemId = 0
With myClass
        .Value = Value
        .Format = Format
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    ReDim Interval(DateInterval.[_First] To DateInterval.[_Last])
    Interval(DateInterval.Day) = "d"
    Interval(DateInterval.DayOfYear) = "y"
    Interval(DateInterval.Hour) = "h"
    Interval(DateInterval.Minute) = "n"
    Interval(DateInterval.Month) = "m"
    Interval(DateInterval.Quarter) = "q"
    Interval(DateInterval.Second) = "s"
    Interval(DateInterval.Week) = "ww"
    Interval(DateInterval.Weekday) = "w"
    Interval(DateInterval.Year) = "yyyy"
    
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal other As IGeneric) As Boolean

    If other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not other.IsRelatedTo(Me) Then IGeneric_Equals = False: Exit Function

    Dim d As GDate: Set d = other
   
    IGeneric_Equals = (myClass.Value = d.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal other As IGeneric) As ComparerEnum
    Dim d As GDate
    
    If other.IsRelatedTo(Me) Then
        Set d = other
        Select Case True
            Case myClass.Value > d.Value: IGeneric_CompareTo = IsGreater
            Case myClass.Value < d.Value: IGeneric_CompareTo = IsLower
            Case Else: IGeneric_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = 0: End Property
Private Function IGeneric_IsRelatedTo(ByVal other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf other Is GDate): End Function
Private Function IGeneric_ToString() As String

    Select Case myClass.Format
        
        
    End Select
    
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GDate: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GDate: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

