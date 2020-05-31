Attribute VB_Name = "Modul1"
'@Folder("Interface")
Option Compare Database

Sub test()

Dim s As TString
Dim i As Long

For i = 1 To 10000
    Set s = TString.Build("test" & i)
Next

End Sub
