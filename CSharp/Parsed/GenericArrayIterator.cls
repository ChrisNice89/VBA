VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericArrayIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule AssignedByValParameter
'@PredeclaredId
'@Folder "Collections"
Option Explicit

Implements IGenericIterator

#If Win64 Then
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
#Else
    Private Declare Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As Long
#End If

Private Const S_OK As Long = &H0&

Private Const NOT_SPECIFIED As Long = -1 'Dont change!

Private Type Span
    '@Ignore IntegerDataType
    cDims        As Integer
    '@Ignore IntegerDataType
    fFeatures    As Integer
    cbElements   As Long
    cLocks       As Long
    #If Win64 Then
        pvData       As LongPtr
    #Else
        pvData       As Long
    #End If
    cElements    As Long
    lLbound      As Long
End Type

Private Type Member
    Source As GenericArray
    Index As Long
    Remaining As Long
    Current As Object
    Reverse As Boolean
    Span As Span
    #If Win64 Then
        Ptr As LongPtr
    #Else
        Ptr  As Long
    #End If
End Type
Private This As Member

Public Property Get Index() As Long: Index = This.Index: End Property
Public Property Get Remaining() As Long: Remaining = This.Remaining: End Property
Public Property Get Reverse() As Boolean: Reverse = This.Reverse: End Property
Public Property Get Source() As IGenericCollection: Set Source = This.Source: End Property

Public Function Build( _
                        ByVal Collection As GenericArray, _
                        Optional ByVal BeginAt As Long = NOT_SPECIFIED, _
                        Optional ByVal Count As Long = NOT_SPECIFIED, _
                        Optional ByVal Reverse As Boolean = False) As IGenericIterator
    
    If (Collection Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("(Length - BeginAt < Count)")
    
    If (BeginAt = NOT_SPECIFIED) Then
        BeginAt = 0
    ElseIf (BeginAt < 0) Then
        Call GenericError.Build(Me, "Public Function Build").ArgOutOfRange("BeginAt < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = Collection.Length - BeginAt
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Build").ArgOutOfRange("Length < 0")
    End If
    
    If (Collection.Length - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("(Length - BeginAt < Count)")
    
    If Not Reverse Then
        This.Reverse = False
        This.Index = BeginAt
    Else
        This.Reverse = True
        This.Index = Count - 1
    End If
    
    This.Remaining = Count
    Set This.Source = Collection
    
    Set Build = New GenericArrayIterator
    This = EmptyMember
    
End Function

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyMember() As Member: End Function

Private Sub Class_Initialize()
    
    If Me Is GenericArrayIterator Then _
        Exit Sub
        
    With GenericArrayIterator
        This.Index = .Index
        This.Remaining = .Remaining
        This.Reverse = .Reverse
        Set This.Source = .Source
    End With
    
    With This.Source
        This.Span.cDims = 1
        This.Span.cbElements = .SizeOf(1)
        This.Span.fFeatures = .Features
        This.Span.pvData = .AdressOf(0) + This.Index * This.Span.cbElements
        This.Span.cElements = This.Remaining
        This.Span.lLbound = .LowerBound
    End With
    
    This.Ptr = VarPtr(This.Span)
    
End Sub

Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Current: End Property

Private Function IGenericIterator_HasNext(Optional ByRef Out As Object) As Boolean
    
    If (This.Index < This.Span.cElements) Then

        '@Ignore ImplicitUnboundDefaultMemberAccess
        If Not (SafeArrayGetElement(ByVal This.Ptr, This.Index, ByVal This.Current) = S_OK) Then _
            Call GenericError.Build(Me, "Private Function IGenericIterator_HasNext").InvalidInput("Invalid Index")

        This.Index = This.Index + 1

        Set Out = This.Current
        IGenericIterator_HasNext = True
        Exit Function

    End If
    
    Set This.Current = Nothing
    Set This.Source = Nothing
    
End Function


