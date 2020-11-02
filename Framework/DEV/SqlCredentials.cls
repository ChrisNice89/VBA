VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlCredentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Base")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IGeneric

'Provider Database
Public Enum ISqlConnectionProvider
    Default_Provider
    Excel_Provider
    Access97_2003_Provider
    Access2007_2013_Provider
    Sql_Provider
    DBMS_Provider
    CSV_File_Provider
End Enum

Private Type Member
    ProviderTyp As ISqlConnectionProvider
    Prompt As ISqlConnectionPrompt
    Mode As ISqlConnectionMode
    DataSource As String
    UserPassword As String
    UserID As String
    DatabasePassword As String
    HDR As Boolean
    ConnectionString As String
    SSPI As Boolean
    IntitialCatalog As String
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get ProviderTyp() As ISqlConnectionProvider
    ProviderTyp = myclass.ProviderTyp
End Property

Public Property Get SSPI() As Boolean
    SSPI = myclass.SSPI
End Property

Public Property Get DataSource() As String
    DataSource = myclass.DataSource
End Property

Public Property Get IntitialCatalog() As String
    IntitialCatalog = myclass.IntitialCatalog
End Property

Public Property Get Mode() As ISqlConnectionMode
    Mode = myclass.Mode
End Property

Public Property Get DatabasePassword() As String
    DatabasePassword = myclass.DatabasePassword
End Property

Public Property Get UserID() As String
    UserID = myclass.UserID
End Property

Public Property Get UserPassword() As String
    UserPassword = myclass.UserPassword
End Property

Public Property Get Prompt() As ISqlConnectionPrompt
    Prompt = myclass.Prompt
End Property


Public Function SqlConnection(ByVal Path As String, ByVal InitialCatalog As String, Optional ByVal SSPI As Boolean, Optional ByVal Password As String) As SqlCredentials
    
    With myclass
        .ProviderTyp = Sql_Provider
        .DataSource = Path
        .IntitialCatalog = InitialCatalog
        .SSPI = SSPI
        .DatabasePassword = Password
        .Prompt = PromptNever
    End With
   
    Set SqlConnection = Skynet.CreateInstance(New SqlCredentials, Virtual.Pointer, Virtual.Size)

End Function

Public Function AccessConnection(ByVal Path As String, Optional ByVal Password As String) As SqlCredentials
    
    With myclass
        .ProviderTyp = Access2007_2013_Provider
        .DataSource = Path
        .DatabasePassword = Password
        .Prompt = PromptNever
        .UserID = "Admin"
    End With
   
    Set AccessConnection = Skynet.CreateInstance(New SqlCredentials, Virtual.Pointer, Virtual.Size)

End Function

Public Function CsvConnection(ByVal File As String, Optional ByVal Password As String) As SqlCredentials
    
    With myclass
        .ProviderTyp = CSV_File_Provider
        .DataSource = VBA.Left$(File, InStrRev(File, "\") - 1)
        .DatabasePassword = Password
        .Prompt = PromptNever
    End With
    
    Set CsvConnection = Skynet.CreateInstance(New SqlCredentials, Virtual.Pointer, Virtual.Size)

End Function

Public Function ExcelConnection(ByVal File As String, Optional ByVal Password As String) As SqlCredentials
    
    With myclass
        .ProviderTyp = Excel_Provider
        .DataSource = File
        .DatabasePassword = Password
        .Prompt = PromptNever
    End With
    
    Set ExcelConnection = Skynet.CreateInstance(New SqlCredentials, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New SqlCredentials: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is SqlCredentials):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New SqlCredentials, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
End Function

Private Function IGeneric_ToString() As String

    
End Function

