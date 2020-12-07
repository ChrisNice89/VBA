VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stringbuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Base"
Option Compare Text
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Dst As Long, ByVal Src As Long, ByVal Length As Long)

Private Const InitialCharCount As Long = 16
' Define members
Private mUBound As Long
Private mString() As Byte

Private Sub Class_Initialize()
    Me.Clear
End Sub

Public Sub Clear()
    
    mUBound = -1
    ReDim mString(0 To InitialCharCount * 2 - 1) 'Each unicode character is 2 bytes

End Sub

Public Property Get Self() As Stringbuilder
    Set Self = Me
End Property

Public Property Get Length() As Long
    Length = (mUBound + 1) / 2
End Property

Public Function Create() As Stringbuilder
        
    With New Stringbuilder
        Set Create = .Self
    End With

End Function

Public Function Append(ByVal Value As String) As Stringbuilder
    
    Dim NewUBound As Long
    Dim CapacityUBound As Long
    
    On Error GoTo Failed
    NewUBound = mUBound + LenB(Value)

    If NewUBound > UBound(mString) Then
        CapacityUBound = UBound(mString) * 2 + 1
        If NewUBound > CapacityUBound Then CapacityUBound = NewUBound * 2 + 1
        ReDim Preserve mString(0 To CapacityUBound)
    End If
    CopyMemory VarPtr(mString(mUBound + 1)), StrPtr(Value), LenB(Value)

    mUBound = NewUBound
    Set Append = Me
    Exit Function

Failed:
    Stop
    Resume
End Function

Public Function ToString() As String
    ToString = Mid(mString, 1, Length)
End Function

