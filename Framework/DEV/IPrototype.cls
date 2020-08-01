VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrototype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "Base"
Public Property Get Default() As Object: End Property
Public Property Get VirtualPointer() As LongPtr: End Property
Public Property Get Size() As Long: End Property
Public Function CreateDefault() As IPrototype: End Function

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub
