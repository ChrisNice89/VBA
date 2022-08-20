Attribute VB_Name = "System"
'@Folder "Base"
'@IgnoreModule ProcedureNotUsed, ConstantNotUsed, ModuleWithoutFolder
Option Explicit

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal GetModuleHandle As String) As LongPtr
    Private Declare PtrSafe Sub CopyMemory_byVar Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef Dst As Any, ByRef Src As Any) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60" Alias "__vbaCopyBytesZero" (ByVal Length As Long, Dst As Any, Src As Any) As Long
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Size As LongPtr)
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32.dll" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32.dll" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal GetModuleHandle As String) As Long
    Private Declare Sub CopyMemory_byVar Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef Dst As Any, ByRef Src As Any) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60" Alias "__vbaCopyBytesZero" (ByVal Length As Long, Dst As Any, Src As Any) As Long
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Size As Long)
    Private Declare Function InterlockedIncrement Lib "kernel32.dll" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32.dll" (ByRef lpAddend As Long) As Long
#End If

Public Function Throw(ByVal Source As Object, ByVal Method As String) As GenericError
    Set Throw = GenericError.Build(Source, Method, True)
End Function
'Call Copymemory_byVar(ByVal Prototype.VirtualPointer, ByVal Source, Size): Call Zeromemory(ByVal Source, Size)
Public Sub Inject(ByVal Prototype As IGeneric, ByVal Source As LongPtr, ByVal Size As Long): Call CopyBytesZero(ByVal Size, ByVal Prototype.VirtualPointer, ByVal Source): End Sub

