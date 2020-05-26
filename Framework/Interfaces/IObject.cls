VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Property Get HashValue() As Long: End Property
Public Function Equals(ByRef Other As IObject) As Boolean: End Function
Public Function CompareTo(ByRef Other As IObject) As ComparerEnum: End Function
Public Function ToString() As String: End Function
Public Function Clone() As IObject: End Function
Public Sub Dispose(): End Sub
Public Property Get VirtualPointer() As LongPtr: End Property

Private Sub Class_Initialize()
    Err.Raise 5, TypeName$(Me), TypeName$(Me) & " (Run-time error '5': Invalid procedure call or argument) - Interface class must not be instantiated."
End Sub
