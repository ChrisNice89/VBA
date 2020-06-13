VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PackageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Database")
Option Compare Database

Option Explicit
'@PredeclaredId

Private Declare PtrSafe Function CorBindToRuntimeEx Lib "mscoree" ( _
    ByVal pwszVersion As LongPtr, _
    ByVal pwszBuildFlavor As LongPtr, _
    ByVal startupFlags As Long, _
    ByRef rclsid As Long, _
    ByRef riid As Long, _
    ByRef ppvObject As mscoree.CorRuntimeHost) As Long

Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (Dest, Src) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal GetModuleHandle As String) As LongPtr

Private Const m_Table As String = "LibFiles"
Private Const m_Library  As String = "Library"
Private Const m_Namespace  As String = "Namespace"
Private Const m_dll  As String = "dll_file"
Private Const m_tlb  As String = "tlb_file"
Private Const m_Reference As String = "VBA_Reference"
Private Const m_Location As String = "LocationFolder"

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005

Private Type LibData
    Size As Long
    Path As String
    Data() As Byte
End Type

Private Type LibFile
    Exists As Boolean
    Name As String
    Namespace As String
    Reference As String
    LocationFolder As String
    dll As LibData
    tlb As LibData
End Type

Private Type Member
    Connection As ADODB.Connection
    Library As LibFile
    Domain As mscorlib.AppDomain
    Host As mscoree.CorRuntimeHost
    Earlybinding As Boolean
End Type
Dim myClass As Member

Public Property Get HasEarlybinding() As Boolean
    HasEarlybinding = myClass.Earlybinding
End Property

Public Property Get Name() As String
    Name = myClass.Library.Name
End Property

Public Property Get Folder() As String
    Folder = myClass.Library.LocationFolder
End Property

Public Property Get LibraryNamespace() As String
    LibraryNamespace = myClass.Library.Namespace
End Property

Public Property Get Assembly() As String
    Assembly = myClass.Library.dll.Path
End Property

Private Sub Class_Initialize()
    
    If Me Is PackageManager Then
        Select Case Application.Name
            Case "Microsoft Excel"
                
            Case "Microsoft Access"
                Debug.Print ReadRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Office\" & Application.Version & "\Excel\Security", "AccessVBOM", 64)
                'Debug.Print CheckRegKey("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM")
            Case Else
                MsgBox "Key nicht definiert"
        End Select
        Call activeVBE
        Set myClass.Connection = CurrentProject.Connection
        Call Me.CreateTable
    Else
        With myClass
            Set .Connection = CurrentProject.Connection
            .Library = LoadLib(PackageManager.Name)
            .Earlybinding = PackageManager.HasEarlybinding
            
            With myClass.Library
                If Not (.Exists) Then _
                    Err.Raise vbObjectError, "Install", "Lib doesnt exist"
               
                'Kill (.Assembly)
                Call ExportFile(.dll)
                    'Err.Raise vbObjectError, "Install", "Assembly couldnt export"
                   
                If myClass.Earlybinding Then
                    'Call RemoveReference(.tlb)
                    If ExportFile(.tlb) Then _
                        Call UpdateReference(.tlb)
               End If
           End With
        End With
    End If

End Sub

Public Sub Dispose()
    
End Sub

Private Sub Class_Terminate()
    With myClass
        If Not .Host Is Nothing Then
            .Host.Stop
            Set .Host = Nothing
            Set .Domain = Nothing
        End If
    End With
End Sub

Public Function TableExists() As Boolean
    
    Dim Schema As Integer
    Schema = ADODB.SchemaEnum.adSchemaTables
    
    If Not myClass.Connection Is Nothing Then
        If Not (myClass.Connection.State And 0) = 0 Then
            Call myClass.Connection.Open
        End If
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = myClass.Connection.OpenSchema(Schema, Array(Empty, Empty, m_Table, Empty))

    TableExists = Not (rst.BOF And rst.EOF)

End Function

Public Function DeleteTable() As Boolean
   
'    Const DeleteTableSql = "DROP TABLE " & m_Table
'    Call myClass.Command.Execute(DeleteTableSql, , 128&)

    DeleteTable = True

End Function


Public Function CreateTable() As Boolean
   
   Const CreateTableSql = _
      "CREATE TABLE " & m_Table & " ( " & _
                     m_Library & " VARCHAR(255) PRIMARY KEY, " & _
                     m_Namespace & " VARCHAR(255), " & _
                     m_Reference & " VARCHAR(255), " & _
                     m_Location & " VARCHAR(255), " & _
                     m_dll & " IMAGE, " & _
                     m_tlb & " IMAGE)"
    
    With New ADODB.Command
        If Not Me.TableExists Then
            Set .ActiveConnection = myClass.Connection
            .CommandText = CreateTableSql
            Call .Execute(, , 128&)
            Application.SetHiddenAttribute acTable, m_Table, True
        End If
    End With
    
    CreateTable = True

End Function

Public Function InstallFrom( _
                                ByVal Location As String, _
                                ByVal Library As String, _
                                Optional ByVal Namespace As String) As PackageManager

'    With myClass
'        .Library = Library
'        .Folder = CheckLocation(Location, True)
'        .Namespace = IIf(Namespace = vbNullString, .Library, Namespace)
'        '.ReferenceName = Name 'FileName(Assembly)
'        .dll = .Folder & .Library & ".dll"
'        .tlb = .Folder & .Library & ".tlb"
'        'Call UpdateReference(.ReferenceName, .tlb)
'    End With
    
    Set InstallFrom = Me
End Function

Public Function Install(ByVal Library As String, Optional ByVal Earlybinding As Boolean = False) As PackageManager
    
    With myClass.Library
        .Name = Library
    End With
    myClass.Earlybinding = Earlybinding
    
    Set Install = New PackageManager
End Function

Private Function ExportFile(ByRef File As LibData) As Boolean
    Dim FileNr As Integer
    
    With File
        If .Size = 0 Then Exit Function
        If .Path = vbNullString Then Exit Function
        If IsFileOpen(.Path) Then Exit Function
    
        FileNr = FreeFile
        Open .Path For Binary As #FileNr
        Put #FileNr, , .Data()
        Close #FileNr
        
        ExportFile = True
    End With
End Function

Private Function IsFileOpen(FileName As String) As Boolean

    Dim fileNum As Integer
    Dim errNum As Integer
    
    'Allow all errors to happen
    On Error Resume Next
    fileNum = FreeFile()
    
    'Try to open and close the file for input.
    'Errors mean the file is already open
    Open FileName For Input Lock Read As #fileNum
    Close fileNum
    
    'Get the error number
    errNum = Err
    
    'Do not allow errors to happen
    On Error GoTo 0
    
    'Check the Error Number
    Select Case errNum
    
        'errNum = 0 means no errors, therefore file closed
        Case 0
        IsFileOpen = False
     
        'errNum = 70 means the file is already open
        Case 70
        IsFileOpen = True
    
        'Something else went wrong
        Case Else
        IsFileOpen = False
    
    End Select

End Function

Private Function LoadLib(ByVal LibName As String) As LibFile
    
    Dim Lib As LibFile

    With AccessLibrary(LibName)
'        If (.BOF And .EOF) Then _
'            Err.Raise vbObjectError, "LoadLibrary", "Missing file data"
'
        If Not (.BOF And .EOF) Then
            Lib.Exists = True
            Lib.Name = .Fields(m_Library)
            Lib.Namespace = .Fields(m_Namespace)
            Lib.LocationFolder = CheckLocation(.Fields(m_Location), True)
            Lib.Reference = .Fields(m_Reference)
            
            With .Fields(m_dll)
                If .ActualSize > 0 Then
                    Lib.dll.Size = .ActualSize
                    Lib.dll.Path = Lib.LocationFolder & Lib.Name & ".dll"
                    Lib.dll.Data = .GetChunk(.ActualSize)
                End If
            End With
            With .Fields(m_tlb)
                If .ActualSize > 0 Then
                    Lib.tlb.Size = .ActualSize
                    Lib.tlb.Path = Lib.LocationFolder & Lib.Name & ".tlb"
                    Lib.tlb.Data = .GetChunk(.ActualSize)
                End If
            End With
        End If
        .Close
    End With
    
    LoadLib = Lib
    
End Function

Private Function AccessLibrary(ByVal LibName As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = myClass.Connection
        .CommandText = "SELECT * FROM " & m_Table & " WHERE " & m_Library & "=?"
        Call .Parameters.Append(.CreateParameter("Library", adVarChar, adParamInput, Len(LibName), LibName))
    End With
    
    Set AccessLibrary = GetRecordset(cmd)
End Function

Private Function GetRecordset(ByRef cmd As ADODB.Command) As ADODB.Recordset
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
   
    'Updatebatch
    With rst
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockBatchOptimistic
        Call .Open(cmd)
    End With
 
    Set GetRecordset = rst
    
End Function

Public Function ImportLibrary( _
                                ByVal Location As String, _
                                ByVal LibName As String, _
                                ByVal Namespace As String, _
                                ByVal Destination As String) As PackageManager
    
    ReDim Files(1 To 2) As String
    Dim dll As String
    Dim tlb As String
    
    
    Location = CheckLocation(Location)
    If VBA.Len(VBA.dir$(Location, vbDirectory)) = 0 Then
        Debug.Print Location: Err.Raise vbObjectError, "ImportLibrary", Location & " not exists"
    End If
    
    dll = Location & LibName & ".dll"
    tlb = Location & LibName & ".tlb"
    
    Files(1) = dll
    Files(2) = tlb
    
    Dim File
    For Each File In Files
        If VBA.Len(VBA.dir$(File)) = 0 Then
          Debug.Print File: Err.Raise vbObjectError, "ImportLibrary", File & " not found"
          Exit Function
        End If
    Next
    With AccessLibrary(LibName)
        If (.BOF And .EOF) Then
            .AddNew
            .Fields(m_Library) = LibName
            .Fields(m_Namespace) = Namespace
            .Fields(m_Location) = Destination
            .Fields(m_Reference) = Replace$(LibName, ".", "")
        End If
        .Fields(m_dll).AppendChunk ImportFile(dll)
        .Fields(m_tlb).AppendChunk ImportFile(tlb)
        .UpdateBatch
        .Close
    End With
   
'    With myClass.Command
'        .CommandText = "INSERT INTO " & m_Table & " VALUES (?,?,?,?,?,?)" ' Query
'        .CommandType = adCmdText
'
'        '---adding parameters
'        .Parameters.Append .CreateParameter("@Library", adVarChar, adParamInput, Len(LibName), LibName)
'        .Parameters.Append .CreateParameter("@Namespace", adVarChar, adParamInput, Len(Namespace), Namespace)
'        .Parameters.Append .CreateParameter("@VBA_Reference", adVarChar, adParamInput, Len(LibName), LibName)
'        .Parameters.Append .CreateParameter("@LocationFolder", adVarChar, adParamInput, Len(Destination), Destination)
'
'        .Parameters.Append .CreateParameter("@dll_file", adVarBinary, adParamInput, adoStream.Size, adoStream.Read)
'        '---
'    End With
    
    Set ImportLibrary = Me
    
End Function
    
Private Function ImportFile(ByVal File As String) As Byte()
    
    Dim FileNr As Integer
    Dim Size As Long
    Dim Data() As Byte
    
    FileNr = FreeFile
    Open File For Binary As #FileNr
        Size = LOF(FileNr) - 1
        If Size >= 0 Then
            ReDim Data(Size)
            Get #FileNr, , Data()
        End If
    Close #FileNr
    
    If Size < 0 Then
        'Err.Raise vbObjectError, "SaveAppFile", "Empty file (filesize = 0)"

    End If
    
    ImportFile = Data
End Function

Private Function CheckLocation( _
                                    ByVal Folder As String, _
                                    Optional ByVal CreateMissingDirectory As Boolean = False) As String
   
    

    Const APPDIR = "%APPDIR%"
    Const USER = "%USER%"
    Const DEFAULT = "%DEFAULT%"
    Const SYSTEMROOT = "%SystemRoot%"

    If VBA.LenB(Folder) = 0 Then _
        Err.Raise vbObjectError, "GetFullFilePath", "Missing directory info"

    Folder = Replace$(Folder, APPDIR, ApplicationLibPath)
    
    If VBA.Right$(Folder, 1) <> "\" Then _
       Folder = Folder & "\"
 
    If CreateMissingDirectory Then _
       Call CreateDirectory(Folder)
    
    CheckLocation = Folder

End Function

Private Sub activeVBE()
    Dim objVBE As Object
    
    On Error Resume Next
    Set objVBE = Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
    Set objVBE = Nothing
    On Error GoTo 0
End Sub

Private Function ReadRegistry( _
                                ByVal RootKey, _
                                ByVal Key As String, _
                                ByVal Value As String, _
                                Optional ByVal RegType As Integer = 32) As String
   
    ' Reads a REG_SZ value from the local computer's registry using WMI.
    ' Parameters:
    '   RootKey - The registry hive
    '   Key - The key that contains the desired value.
    '   Value - The value that you want to get.
    '   RegType - The registry bitness: 32 or 64.
'https://docs.microsoft.com/en-us/windows/win32/wmisdk/requesting-wmi-data-on-a-64-bit-platform

    Dim oCtx, oLocator, oReg, oInParams, oOutParams
    
    If KeyExists(RootKey, Key) Then
        Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
        oCtx.Add "__ProviderArchitecture", RegType
      
        Set oLocator = CreateObject("Wbemscripting.SWbemLocator")
        Set oReg = oLocator.ConnectServer("", "root\default", "", "", , , , oCtx).Get("StdRegProv")
    
        Set oInParams = oReg.Methods_("GetStringValue").Inparameters
        oInParams.Hdefkey = RootKey
        oInParams.Ssubkeyname = Key
        oInParams.Svaluename = Value
    
        Set oOutParams = oReg.ExecMethod_("GetStringValue", oInParams, , oCtx)
        If IsNull(oOutParams.sValue) Then
            ReadRegistry = vbNullString
        Else
            ReadRegistry = oOutParams.sValue
        End If
    End If
  
End Function

Private Function KeyExists(ByVal RootKey, KeyPath)
    Dim oReg: Set oReg = GetObject("winmgmts:!root/default:StdRegProv")
    KeyExists = (oReg.EnumKey(RootKey, KeyPath, "", "") = 0)
End Function

Private Sub UpdateReference(ByRef Lib As LibData)
    'On Error Resume Next
    'Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
    'Call RemoveReference(Lib)
    Call Application.VBE.ActiveVBProject.References.AddFromFile(Lib.Path)
    'On Error GoTo 0
End Sub

Public Function RemoveReference() As Boolean
    
'    Dim objVBE As Object
'    Set objVBE = Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
'
    On Error Resume Next
    If Not myClass.Library.Exists Then Exit Function
    
    Dim i As Long
    With Application.VBE.ActiveVBProject.References
        For i = 1 To .Count
            If StrComp(.Item(i).FullPath, myClass.Library.tlb.Path, vbBinaryCompare) = 0 Then
            
'                If Not .Item(i).IsBroken Then
'                    blnReturn = True
'                Else
'                    .Remove .Item(i)
'                End If
                .Remove .Item(i)
                myClass.Earlybinding = False
                RemoveReference = True
                Exit For
            End If
        Next i
    End With
End Function

Private Function FileName(ByVal Path As String) As String
    Dim FullName As String
    FullName = Right(Path, Len(Path) - InStrRev(Path, "\"))
    FileName = Left(FullName, (InStrRev(FullName, ".") - 1))
End Function

Private Sub ExportFileFromRecordField(FileDataField As DAO.Field, FullFilePath As String)

   Dim FieldSize As Long
   Dim BinFile() As Byte
   Dim FileNr As Integer

   FieldSize = FileDataField.FieldSize
   If FieldSize = 0 Then
      Err.Raise vbObjectError, "ExportFile", "Missing file data"
   End If

   ReDim BinFile(FieldSize - 1)

   BinFile = FileDataField.GetChunk(0, FieldSize)
   
   ' Datei erstellen
   FileNr = FreeFile
   Open FullFilePath For Binary As #FileNr
   Put #FileNr, , BinFile()
   Close #FileNr

End Sub

Function CreateInstance(ByVal ClassName As String) As Variant
    Const CLR$ = "v4.0.30319"
    
    'Static Domain As mscorlib.AppDomain
    If myClass.Domain Is Nothing Then
        'Dim Host As mscoree.CorRuntimeHost
        Dim hr&, T&(0 To 7)
        T(0) = &HCB2F6723: T(1) = &H11D2AB3A: T(2) = &HC000409C: T(3) = &H3E0AA34F
        T(4) = &HCB2F6722: T(5) = &H11D2AB3A: T(6) = &HC000409C: T(7) = &H3E0AA34F
        
        hr = CorBindToRuntimeEx(StrPtr(CLR), 0, 3, T(0), T(4), myClass.Host)
        If hr And -2 Then Err.Raise hr
        With myClass.Host
            .Start
            .GetDefaultDomain myClass.Domain
        End With
    End If
    
    With myClass.Library
        If Not .Exists Then
            Err.Raise 1#, "CreateInstance", "Library doesnt exist"
        End If
        
        VariantCopy CreateInstance, myClass.Domain.CreateInstanceFrom(.dll.Path, .Namespace & "." & ClassName).Unwrap
    End With
End Function
'v1
Public Function CreateInstance2(ByVal ClassName As String) As Object
    Set myClass.Host = New mscoree.CorRuntimeHost
    With myClass.Host
        .Start
        .GetDefaultDomain myClass.Domain
        Set CreateInstance2 = myClass.Domain.CreateInstanceFrom(myClass.Library.dll.Path, myClass.Library.Namespace & "." & ClassName).Unwrap
        .Stop
    End With
    Set myClass.Host = Nothing
    Set myClass.Domain = Nothing

End Function

Private Function GetFullFilePath( _
                                        ByVal FileName As String, _
                                        ByVal DestinationFolder As String, _
                                        Optional ByVal CreateMissingDirectory As Boolean = False) As String
   
    
    Const APPDIR = "[APPDIR]"
    
    If VBA.LenB(DestinationFolder) = 0 Then
        Err.Raise vbObjectError, "GetFullFilePath", "Missing directory info"
    End If
  
    DestinationFolder = Replace$(DestinationFolder, APPDIR, ApplicationLibPath)
    
    If VBA.Right$(DestinationFolder, 1) <> "\" Then
       DestinationFolder = DestinationFolder & "\"
    End If
    
    If CreateMissingDirectory Then
       Call CreateDirectory(DestinationFolder)
    End If
    
    GetFullFilePath = DestinationFolder & FileName

End Function

Private Sub CreateDirectory(ByVal Path As String)
   
    Dim PathBefore As String
    Dim TextPos As Long
    Dim BuildPath As String
    Dim Folder As Variant
    
    If VBA.LenB(VBA.dir$(Path, vbDirectory)) > 0 Then _
        Exit Sub
    
    If Right(Path, 1) = "\" Then _
        Path = Left(Path, Len(Path) - 1)

    TextPos = InStrRev(Path, "\")
    
    If TextPos > 3 Then
      PathBefore = VBA.Mid$(Path, 1, TextPos - 1)
      If VBA.Len(VBA.dir$(PathBefore, vbDirectory)) = 0 Then
         Call CreateDirectory(PathBefore)
      End If
    End If
    
    MkDir Path
     
'    For Each Folder In Split(Path, "\")
'        BuildPath = BuildPath & Folder & "\"
'        If Len(dir(BuildPath, vbDirectory)) = 0 Then _
'            MkDir BuildPath
'
'    Next
'
End Sub

Private Function ApplicationLibPath() As String
' MS.net-dlls laufen nicht in Netzwerkpfad

'   Dim p As String
'   p = CodeProject.Path
'   If p <> UNCPath(p) Then
'      'p = GetSpecFolder(CSIDL_LOCAL_APPDATA) & "\" & ApplicationName
'   End If
'   ApplicationLibPath = p
   
End Function
Private Function ApplicationName() As String
   Dim s As String
   s = Access.CodeProject.Name
   s = VBA.Left$(s, VBA.InStrRev(s, ".") - 1)
   ApplicationName = s
End Function

'Private Function UNCPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String
'
'  Dim UNC As String * 512
'
'  If VBA.Len(Path) = 1 Then Path = Path & ":"
'
'  If WNetGetConnection(Left$(Path, 2), UNC, Len(UNC)) Then
'
'    ' API-Routine gibt Fehler zurück:
'    If IgnoreErrors Then
'      UNCPath = Path
'    Else
'      Err.Raise 5 ' Invalid procedure call or argument
'    End If
'
'  Else
'
'    ' Ergebnis zurückgeben:
'    UNCPath = Left$(UNC, InStr(UNC, vbNullChar) - 1) _
'            & Mid$(Path, 3)
'
'  End If

'End Function

