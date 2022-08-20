VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataGridColumnDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "DataGrid"
Option Explicit

Implements GenericMapper

'@Ignore EncapsulatePublicField
Public Index As Long

Private Function GenericMapper_Build() As GenericMapper
    Set GenericMapper_Build = New DataGridColumnDelegate
End Function

Private Function GenericMapper_Map(ByVal Element As Object) As Object
    
    Dim Row As GenericArray
    Set Row = Element
    Set GenericMapper_Map = Row.ElementAt(Index)
    
End Function

