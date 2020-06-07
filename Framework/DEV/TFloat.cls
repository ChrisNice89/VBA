VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Class")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IObject
Implements IPrototype
Implements IValue

Public Enum TFloatFormat
    [_First]
    DefaultFloat
   
    [_Last]
End Enum

Private Type Member
    Value As Double
    Format As TFloatFormat
   
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Value() As Double
    Value = myClass.Value
End Property

Public Function Build( _
                        ByVal Value As Double, _
                        Optional ByVal Format As TFloatFormat = TFloatFormat.DefaultFloat) As TFloat
Attribute Build.VB_UserMemId = 0
With myClass
        .Value = Value
        .Format = Format
        
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Function IValue_Build(ByVal Value As Variant, Optional ByVal Format As Byte) As IObject
    
    With myClass
        .Value = Value
        .Format = Format
    End With
    
    Set IValue_Build = Skynet.Create(Me)
End Function
Private Property Get IValue_Object() As IObject: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IObject_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Dim f As TFloat: Set f = Other
    
    IObject_Equals = (myClass.Value = f.Value)
     
End Function

Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim f As TFloat
    
    If Other.IsRelatedTo(Me) Then
        Set f = Other
        Select Case True
            Case myClass.Value > f.Value: IObject_CompareTo = IsGreater
            Case myClass.Value < f.Value: IObject_CompareTo = IsLower
            Case Else: IObject_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = 0: End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is TFloat): End Function
Private Function IObject_ToString() As String

    Select Case myClass.Format
        
        
    End Select
    
End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TFloat: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



