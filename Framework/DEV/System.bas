Attribute VB_Name = "System"
Option Compare Database

Public Const DEFAULT_SORTORDER As Integer = 1 'SortOrder.Ascending
Public Const MAXVALUE As Double = 2 ^ 31
Public Const MAX_INT32 As Long = &H7FFFFFFF
Public Const NullPointer As Long = 0

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

Public Property Get DefaultComparer() As IGenericComparer: Set DefaultComparer = IGenericComparer: End Property

Public Function Throw(ByVal Source As Object, ByVal Method As String) As GenericThrowhelper
    Set Throw = GenericThrowhelper.Build(Source, Method)
End Function

Public Function CreateInstance(ByVal Prototype As IGeneric, ByVal Source As LongPtr, ByVal Size As Long) As IGeneric
        
    Call CopyBytesZero(ByVal Size, ByVal Prototype.VirtualPointer, ByVal Source) 'works
'    Call CopySystemmory_byVar(ByVal Prototype.VirtualPointer, ByVal Source, Size): Call ZeroSystemmory(ByVal Source, Size)
    Set CreateInstance = Prototype
End Function

Public Function HashValue(ByRef Ascii() As Byte) As Long
    
    Dim h As Double
    Dim i As Long
    
    Dim Length As Long: Length = UBound(Ascii) + 1
    Dim N As Long
    
    For N = (Length / 2) To 1 Step -1
        h = h + Ascii(i)
        h = System.X0R(System.X0R(LEFTSHIFT(h, 16), System.LEFTSHIFT(Ascii(i + 1), 11)), h)
        h = h + System.RIGHTSHIFT(h, 11)
        i = i + 2
    Next
    
    If ((Length Mod 2) = 1) Then
        h = h + Ascii(i) + 1566083941
        h = System.X0R(h, LEFTSHIFT(h, 10))
        h = h + System.RIGHTSHIFT(h, 1)
    End If
    
    h = System.X0R(h, System.LEFTSHIFT(h, 3)): h = h + System.RIGHTSHIFT(h, 5)
    h = System.X0R(h, System.LEFTSHIFT(h, 4)): h = h + System.RIGHTSHIFT(h, 17)
    h = System.X0R(h, System.LEFTSHIFT(h, 25)): h = h + System.RIGHTSHIFT(h, 6)
    
    HashValue = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))

End Function

Public Function ClassName(ByRef Instance As IGeneric) As String: ClassName = "<" & TypeName(Instance) & ">": End Function
Public Function Clone(ByVal Instance As IGeneric) As IGeneric: Set Clone = Instance.Clone: End Function
Public Function Generic(ByRef Instance As IGeneric) As IGeneric: Set Generic = Instance: End Function
Public Function Modulo(ByVal A As Double, ByVal m As Double) As Long: Modulo = (A - (Int(A / m) * m)): End Function
Public Function DecreSystemnt(ByRef i As Long) As Long: i = (Not -i): DecreSystemnt = i: End Function
Public Function IncreSystemnt(ByRef i As Long) As Long: i = (-(Not i)): IncreSystemnt = i: End Function
Public Function RIGHTSHIFT(ByVal Value As Long, Shift As Byte) As Double: RIGHTSHIFT = Value / (2& ^ Shift): End Function
Public Function LEFTSHIFT(ByVal Value As Long, Shift As Byte) As Double: LEFTSHIFT = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function
Public Function LOGn(ByVal Value, Optional ByVal Base As Byte = 2) As Long: LOGn = Log(Value) / Log(Base): End Function

