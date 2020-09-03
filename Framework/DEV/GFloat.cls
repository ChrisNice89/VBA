VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GFloat"
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
                        Optional ByVal Format As TFloatFormat = TFloatFormat.DefaultFloat) As GFloat
Attribute Build.VB_UserMemId = 0
With myClass
        .Value = Value
        .Format = Format
        
    End With
    
    'Set Build = Skynet.Create(GFloat)
    Set Build = Skynet.CreateInstance(New GFloat, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
    'ReDim Formats(TStringFormat.[_First] To TStringFormat.[_Last])
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GFloat, VarPtr(C), Virtual.Size)
   
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IGeneric_Equals = False: Exit Function

    Dim f As GFloat: Set f = Other
    
    IGeneric_Equals = (myClass.Value = f.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Dim f As GFloat
    
    If Other.IsRelatedTo(Me) Then
        Set f = Other
        Select Case True
            Case myClass.Value > f.Value: IGeneric_CompareTo = IsGreater
            Case myClass.Value < f.Value: IGeneric_CompareTo = IsLower
            Case Else: IGeneric_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = 0: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GFloat): End Function
Private Function IGeneric_ToString() As String

    Select Case myClass.Format
        
        
    End Select
    
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GFloat: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GFloat: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

