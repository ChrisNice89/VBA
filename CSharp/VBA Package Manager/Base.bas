Attribute VB_Name = "Base"
'@Folder("Database")
Option Compare Database

#If VBA7 Then
    Private Declare PtrSafe Function CreateDotNetObject Lib "C:\Users\cnitz\Documents\GitHub\VBA\CSharp\Skynet.Test\Libs\Skynet.Interopt\bin\Debug\SkynetInteropt.dll" (ByVal text As String) As Object
#Else
    Private Declare Function CreateDotNetObject Lib "C:\Users\cnitz\Documents\GitHub\VBA\CSharp\Skynet.Test\Libs\Skynet.Interopt\bin\Debug\SkynetInteropt.dll" (ByVal text As String) As Object
#End If
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function TestApi Lib "Skynet.Objects.dll" () As Long


Sub test()
    
    Dim Folder As String: Folder = "C:\Users\cnitz\Documents\GitHub\VBA\CSharp\Skynet.Test\Libs\"
    Dim Dest As String: Dest = "C:\Users\cnitz\Documents\GitHub\VBA\CSharp\lib"
    
    Dim Namespace As PackageManager
    Dim Namespace2 As PackageManager
     
    Call PackageManager.ImportLibrary(Folder & "Skynet.Interopt\bin\Debug", "Skynet.Interopt", "Skynet.Interopt", Dest)
    Call PackageManager.ImportLibrary(Folder & "Skynet.Objects\bin\Debug", "Skynet.Objects", "Skynet.Objects", Dest)

    Set Namespace = PackageManager.Install("Skynet.Objects", True)
    Set Namespace2 = PackageManager.Install("Skynet.Interopt", True)
    
    Debug.Print Namespace.Assembly
    Debug.Print Namespace2.Assembly
    
'    Call Namespace.RemoveReference
'    Call Namespace2.RemoveReference
    
    Dim f As Object
    Set f = Namespace.CreateInstance("Factory")
    Debug.Print TypeName(f)
    
    Dim s As SkynetObjects.TString
    Set s = f.TString("Test")
    Dim s_ As SkynetObjects.IValue
    Set s_ = s
    
End Sub

Sub API()

  Dim instance As Object

  Set instance = CreateDotNetObject("Test 1")
  Debug.Print instance.text

  Debug.Print instance.TestMethod

  instance.text = "abc 123" ' case insensitivity in VBA works as expected'

  Debug.Print instance.text
End Sub

Private Sub Foo()
    Dim lb As Long

    lb = LoadLibrary("C:\Users\cnitz\Documents\GitHub\VBA\CSharp\lib\Skynet.Interopt.dll")

    MsgBox CreateDotNetObject("Test")

    'I found I had to do repeated calls to FreeLibrary to force the reference count
    'to zero so the dll would be unloaded.
    Do Until FreeLibrary(lb) = 0
    Loop
End Sub
