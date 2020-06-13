VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iCAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Database")
Option Compare Database

Option Explicit
'@PredeclaredId


Private Type Member
    Library As PackageManager
    
End Type
Private myClass As Member

Private Sub Class_Initialize()
    If Me Is iCAT Then
        With New PackageManager
            Set myClass.Library = .Install("iCAT")
        End With
    Else
    
        'Error
    End If
End Sub

Public Function Factory() As Object 'iCAT.Factory
    Set Factory = myClass.Library.CreateInstance("Factory")
End Function

Public Function TString() As Object
    Set TString = myClass.Library.CreateInstance("TString")
End Function

Public Function Constants() As Object
    Set Constants = myClass.Library.CreateInstance("Constants")
End Function
