VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IGenericEnumerator
Implements IPrototype

Public Event Aggregate(ByRef Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
Public Event Contains(ByRef Contains As Boolean, ByVal Current As IGeneric)

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

Private Const DEFAULT_CAPACITY As Long = 16

Private Type Range
    Start As Long
    Count As Long
    i As Long
    Current As IGeneric
    Complete As Boolean
End Type

Private Type Member
    Sequence As IGenericEnumerator
    Range As Range
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Sequence() As IGenericEnumerator
    Set Sequence = myClass.Sequence
End Property

Public Function Build(ByVal Sequence As IGenericEnumerator) As GenericEnumerable
Attribute Build.VB_UserMemId = 0
    If Sequence Is Nothing Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Sequence is nothing")
      
    Set myClass.Sequence = Sequence
    Set Build = New GenericEnumerable
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
    Set myClass.Sequence = GenericEnumerable.Sequence
End Sub

Private Sub Check(ByVal Source As IPrototype, ByVal Caller As String)
    
    If (Source Is Nothing) Then _
        Set Source = Source.CreateDefault 'Call Skynet.Throw(Me, Caller).NullPointer("Key is nothing")

    If myClass.Sequence Is Nothing Then _
        Call Skynet.Throw(Me, Caller).NullPointer("Sequence is nothing")
    
    Call myClass.Sequence.Reset
    
End Sub

Public Function Aggregate(ByVal Result As IPrototype, Optional ByVal IgnoreNull As Boolean = True) As IGeneric
    
    Call Check(Result, "Public Sub Aggregate")
    
    With myClass.Sequence
        If IgnoreNull Then
            Do While .MoveNext
                If Not .Current Is Nothing Then _
                    RaiseEvent Aggregate(Result, .Current, IgnoreNull)
            Loop
        Else
            Do While .MoveNext
                RaiseEvent Aggregate(Result, .Current, IgnoreNull)
            Loop
        End If
        Call .Reset
    End With
    
    Set Aggregate = Result
    
End Function

Public Function Contains(ByVal Item As IGeneric, Optional ByVal Custom As Boolean = False) As Boolean
   
    Dim Check As Boolean
    
    With myClass.Sequence
        Call .Reset
        If Custom Then
            Do While .MoveNext
                RaiseEvent Contains(Check, .Current)
                If Check Then _
                    Exit Do
            Loop
        ElseIf Not Item Is Nothing Then
            Do While .MoveNext
                If Item.Equals(.Current) Then _
                   Check = True: Exit Do
            Loop
        Else
            Do While .MoveNext
                If .Current Is Nothing Then _
                   Check = True: Exit Do
            Loop
        End If
        Call .Reset
    End With
    
    Contains = Check
    
End Function

Public Function Range(ByVal Start As Long, ByVal Count As Long) As IGenericEnumerator
    
    Dim C As Member
    With C.Range
        .Start = Start
        .Count = Count
        .i = .Start
        .Complete = False
    End With
   
    Set Range = Skynet.CreateInstance(New GenericEnumerable, VarPtr(C), LenB(C))

End Function

Private Property Get IGenericEnumerator_Current() As IGeneric
    Set IGenericEnumerator_Current = myClass.Range.Current
End Property

Private Function IGenericEnumerator_MoveNext() As Boolean
    
    With myClass.Range
        If Not .Complete Then
            Set .Current = GNumeric(myClass.Range.i)
        Else
            Set .Current = Nothing: Exit Function
        End If
        
        IGenericEnumerator_MoveNext = Not .Complete
      
        .Complete = InterlockedIncrement(.i) > .Count
    End With
    
End Function

Private Sub IGenericEnumerator_Reset()
    myClass.Range.i = myClass.Range.Start
    myClass.Range.Complete = False
End Sub

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericEnumerable: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericEnumerable: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


