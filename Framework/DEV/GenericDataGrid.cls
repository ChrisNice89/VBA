VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericDataTable
'*
'* Purpose: The GenericList is a collection of IGeneric objects
'*          that can be accessed by index and having methods for sorting, searching, and modifying list.
'*          i.e. GenericList grows dynamically as the elements are added to it.
'*          If the Size of the current elements (including the new element to be added to the GenericList)

'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric

Public Event Update(ByVal Row As GenericArray)

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' VB equivilent of Screen.TwipsPerPixel in Access
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const LVM_FIRST As Long = &H1000
Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Const LVSCW_AUTOSIZE As Long = -1
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Const LVM_SCROLL As Long = &H1014
Const SOME_CHARS As String = "XXXXXX"
'
'customized settings
Public Enum DataviewerStyle
    Default_Style

End Enum

Public Enum DataViewerColumnView
    Hide
    Show
End Enum

Private Const DEFAULT_CAPACITY_ROWS As Long = 64
Private Const DEFAULT_CAPACITY_COLUMNS As Long = 32
Private Const FETCH_LIMIT As Long = 10

Private Type Member
    ViewControl As CustomControl
    ImageList As MSComctlLib.ImageList
    View As MSComctlLib.Listview
    RowMapper As IGenericIterator
    Rows As GenericList
    Columns As GenericOrderedMap
    SelectedColumn As Long
    HasFetchComplete As Boolean
    HasInitialized As Boolean
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get IndexOfColumn(ByVal Name As GString) As Long: IndexOfColumn = myclass.Columns.IndexOfKey(Name): End Property
Public Property Get Rows() As IGenericReadOnlyList: Set Rows = myclass.Rows: End Property
Public Property Get Columns() As IGenericReadOnlyList: Set Columns = myclass.Columns: End Property
Public Property Get SelectedColumn() As Long: SelectedColumn = myclass.SelectedColumn: End Property
Public Property Get Row(ByVal Index As Long) As GenericArray: Set Row = myclass.Rows(Index): End Property
Public Property Get Column(ByVal Index As Long) As GenericPair: Set Column = myclass.Columns.Entry(Index): End Property

Public Function Build( _
                            ByVal ListViewControl As CustomControl, _
                            Optional ByVal Setting As DataviewerStyle = DataviewerStyle.Default_Style, _
                            Optional ByVal ImageList As MSComctlLib.ImageList = Nothing) As GenericDataGrid
    
    If (Me Is GenericDataGrid) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
        
    If (ListViewControl Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").InvalidInput("(Records Is Nothing)")
    
    If Not TypeOf ListViewControl.Object Is MSComctlLib.Listview Then _
        Call Skynet.Throw(Me, "Public Function Build").InvalidInput("False Access.CustomControl: " & TypeName$(ListViewControl.Object))
        
    With myclass
        Set .ViewControl = ListViewControl
        Set .View = .ViewControl.Object
        Call Style(.View, Setting)
        Set .ImageList = ImageList
        Set .Rows = GenericList.Build(DEFAULT_CAPACITY_ROWS)
        Set .Columns = GenericOrderedMap.Build(DEFAULT_CAPACITY_COLUMNS)
    End With
    
    Set Build = Skynet.CreateInstance(New GenericDataGrid, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Public Sub Initialize(ByVal Records As ADODB.Recordset)

'    Dim Fields As ADODB.Fields
'    Dim Field As ADODB.Field
'    Dim i As Long
    
    Dim Field As IGeneric
'    If (Records Is Nothing) Then _
'        Call Skynet.Throw(Me, "Public Sub Initialize(").InvalidInput("(Records Is Nothing)")
'
'    Set myclass.Records = Records
    myclass.HasFetchComplete = False
    Set myclass.RowMapper = GenericSql.RowMapper(Records)
    
    Dim Row As GenericArray
    Set Row = myclass.RowMapper.Current
    With Row.Iterator
        Do While .HasNext(Field)
            Call Me.AddColumn(Field, vbString)
        Loop
    End With
  
'    With Records
'        If .State = ISqlCommandState.StateClosed Then _
'            .Open
'
'        For Each Field In .Fields
'            Call Me.AddColumn(GString(Field.Name), vbString)
'        Next
'
'        If (.EOF And .BOF) Then _
'            myclass.HasFetchComplete = True
'
'    End With
    
    Call Me.Fetch(Autosize:=True)
    myclass.HasInitialized = True
    
End Sub

Public Function Fetch(Optional ByVal Limit As Long = FETCH_LIMIT, Optional ByVal Autosize As Boolean) As Boolean
    
    Dim i As Long
    Dim Row As GenericArray
    
    If Not myclass.HasInitialized Then
        'Error
    End If
    
    If myclass.RowMapper Is Nothing Then
        Exit Function
    End If
    
    Call LockWindowUpdate(myclass.View.hwnd)
    
    Do
        If myclass.RowMapper.HasNext(Row) Then
            Call myclass.Rows.Add(Row)
            With myclass.View.ListItems.Add(Text:=myclass.View.ListItems.Count + 1)
                For i = Row.LowerBound To Row.Length
                    .SubItems(i) = Row.ElementAt(i).ToString
                Next
            End With
        Else
            Set myclass.RowMapper = Nothing
            myclass.HasFetchComplete = True
            Exit Do
        End If
        
    Loop While Not (InterlockedDecrement(Limit) = 0)
   
    If Autosize Then _
        Call Me.ColumnAutoSize

    With myclass.View
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(.ListItems.Count)
            .SelectedItem.Selected = True
            .SelectedItem.EnsureVisible
        End If
    End With
    
    Call LockWindowUpdate(0&)
    DoEvents
    Fetch = True
    
End Function

Public Sub Clear()
    myclass.SelectedColumn = -1
    
    If Not myclass.Rows Is Nothing Then _
        myclass.Rows.Clear
        
    If Not myclass.Columns Is Nothing Then _
        myclass.Columns.Clear
    
    If Not myclass.View Is Nothing Then
        With myclass.View
             .ListItems.Clear
             .ColumnHeaders.Clear
             .SortKey = 0
         End With
    End If
     
End Sub

Public Sub AddColumn(ByVal Name As GString, ByVal Datatype As VbVarType)
    
    Dim RightAlign As Boolean
    
    Select Case Datatype
        Case VbVarType.vbString: Call myclass.Columns.Add(Name, GString): RightAlign = True
        Case VbVarType.vbLong: Call myclass.Columns.Add(Name, GNumeric): RightAlign = True
        Case VbVarType.vbDate: Call myclass.Columns.Add(Name, GDate): RightAlign = True
        Case Else
    End Select
    
    If Not myclass.View Is Nothing Then
        With myclass.View.ColumnHeaders
            If .Count = 0 Then _
                Call .Add(Index:=1, Text:=vbNullString)
    
            Call .Add(Text:=Name.Value)
    
            If RightAlign Then _
                .Item(.Count).Alignment = lvwColumnRight
        
        End With
    End If
    
End Sub

Public Function AddRow() As GenericArray
    
 
End Function

Public Sub Sort(ByVal Index As Long)
    myclass.SelectedColumn = Index
    Call myclass.Rows.Sort(Order:=ascending)
End Sub

Public Sub Update(ByVal ID As Long)
    RaiseEvent Update(myclass.Rows(ID))
End Sub

Public Sub ColumnAutoSize(Optional ByVal Column As Variant)

    Dim i As Long
    Dim LstItm As MSComctlLib.ListItem
    
     With myclass.View
         If .ColumnHeaders.Count = 0 Then Exit Sub
         Set LstItm = .ListItems.Add(Text:=.ColumnHeaders.Item(1).Text & SOME_CHARS)

         If IsMissing(Column) Then
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, 0&, ByVal LVSCW_AUTOSIZE_USEHEADER)

            For i = 2 To .ColumnHeaders.Count
                LstItm.SubItems(i - 1) = .ColumnHeaders.Item(i).Text & SOME_CHARS
                Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, i - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
            Next i

         Else
            i = myclass.View.ColumnHeaders(Column).Index
            LstItm.SubItems(i - 1) = .ColumnHeaders.Item(i).Text & SOME_CHARS
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, i - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
         End If
         Call .ListItems.Remove(LstItm.Index)
     End With

     Set LstItm = Nothing

End Sub

Private Sub Style(ByVal Listview As MSComctlLib.Listview, ByVal Style As DataviewerStyle)
    
    If Listview Is Nothing Then _
        Call Skynet.Throw(Me, "Private Sub Style").NullPointer("Listview Is Nothing")
        
    With Listview
        Select Case Style
            Case DataviewerStyle.Default_Style
                .View = lvwReport
                'Spalten neu ordnen
                .AllowColumnReorder = True
                'ganze Zeile makieren
                .FullRowSelect = True
                .MultiSelect = False
                .LabelEdit = lvwManual 'lvwAutomatic
                'lininen in listbox
                .GridLines = True
                .Font.Size = 11
                .Font.Name = "Calibri" '"Arial Unicode MS"
                'Hintergrundbild ist das Bild "BackGround" in der Imagelist1
                'Set .Picture = ImageList1.ListImages("BackGround").Picture
                .FlatScrollBar = False
                .Appearance = cc3D
                .BorderStyle = ccFixedSingle
                .Checkboxes = False
                .BackColor = RGB(220, 230, 242)
                .Enabled = True 'Itemes columns etc veränderbar
                .HotTracking = False
                .MousePointer = ccArrow
                .HoverSelection = False
            Case Else
            
        End Select
    End With
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericDataGrid: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericDataGrid):  End Property

Private Function IGeneric_Clone() As IGeneric
     Call Skynet.Throw(Me, "Private Function IGeneric_Clone").NotImplemented
End Function

Private Sub IGeneric_Dispose()
    myclass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Rows: " & .Rows.Count & vbNewLine & _
                            "Columns: " & .Columns.Count
    
    End With
End Function

