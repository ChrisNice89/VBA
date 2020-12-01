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

Public Event RowSelected(ByVal RowID As Long, ByVal Row As GenericArray)
Public Event FetchComplete()

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' VB equivilent of Screen.TwipsPerPixel in Access
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Private Declare Function GetInputState Lib "user32" () As Long

Private Enum Keys
    VK_SPACE = &H20
    VK_LEFT = &H25
    VK_RIGHT = &H27
    VK_ESCAPE = &H1B
    VK_back = &H8
End Enum

Private Const FETCH_LIMIT As Long = 10
Private Const DEFAULT_CAPACITY_ROWS As Long = 128
Private Const DEFAULT_CAPACITY_COLUMNS As Long = 32

Private Const WM_SETREDRAW  As Long = &HB
Private Const SR_ON         As Long = 1
Private Const SR_OFF        As Long = 0

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Const SOME_CHARS As String = "XXXXXX"
'
'customized settings
Public Enum DataviewerStyle
    Default_Style

End Enum

Public Enum DataViewerColumnView
    Hide
    Show
End Enum

Private Type Member
    ViewControl As CustomControl
    ImageList As MSComctlLib.ImageList
    View As MSComctlLib.IListView
    RowMapper As IGenericIterator
    RowList As GenericList
    RowMap As IGenericDictionary
    Columns As GenericOrderedMap
    SelectedColumn As Long
    SortDirection As SortOrder
    HasFetchComplete As Boolean
    HasInitialized As Boolean
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Private WithEvents ViewControl As MSComctlLib.Listview
Attribute ViewControl.VB_VarHelpID = -1

Public Property Get IndexOfColumn(ByVal Name As GString) As Long

    Dim Column As IGenericValue
    Set Column = myclass.Columns.Item(Name)
    
    IndexOfColumn = Column.ToValue
    
End Property
Public Property Get Rows() As IGenericReadOnlyList: Set Rows = myclass.RowList: End Property
Public Property Get Columns() As IGenericReadOnlyList: Set Columns = myclass.Columns: End Property
Public Property Get SelectedColumn() As Long: SelectedColumn = myclass.SelectedColumn: End Property
Public Property Get Row(ByVal Index As Long) As GenericArray: Set Row = myclass.RowList.ElementAt(Index): End Property
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
        Call style(.View, Setting)
        Set .ImageList = ImageList
        Set .RowMap = GenericSortedSet.Build
        Set .RowList = GenericList.Build(DEFAULT_CAPACITY_ROWS)
        Set .Columns = GenericOrderedMap.Build(DEFAULT_CAPACITY_COLUMNS)
        .SortDirection = ascending
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

Private Sub ViewControl_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    RaiseEvent RowSelected(item.index, myclass.RowList(item.index))
End Sub

Private Sub ViewControl_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Me.Sort(ColumnHeader.Index - 1, myclass.SortDirection)
End Sub

Public Sub Update(ByVal RowID As Long)
    
    Dim i As Long
    Dim Row As GenericArray
    Set Row = myclass.RowList.ElementAt(RowID)
    
    Dim Element As IGeneric
    
    With myclass.View.ListItems(RowID)
        For i = Row.LowerBound To Row.Length
            Set Element = Row.ElementAt(i)
            If Element Is Nothing = False Then
                .SubItems(i) = Element.ToString
            Else
                .SubItems(i) = vbNullString
            End If
        Next
    End With
        
End Sub

Public Sub Initialize(ByVal Records As ADODB.Recordset)

    Dim Field As ADODB.Field
    
    myclass.HasFetchComplete = False
    Set ViewControl = myclass.View
    Set myclass.RowMapper = GenericSql.RowMapper(Records)
    
    With Records
        For Each Field In .Fields
            Call Me.AddColumn(GString(Field.Name), vbString)
        Next

        If (.EOF And .BOF) Then _
            myclass.HasFetchComplete = True

    End With

    Call Me.ColumnAutoSize
    Call Me.Fetch
    myclass.HasInitialized = True
   
End Sub

Public Function Fetch(Optional ByVal Limit As Long = FETCH_LIMIT) As Boolean

    Dim Row As IGenericReadOnlyList, i As Long
    
    If Not myclass.HasInitialized Then
        'Error
    End If
    
    If myclass.RowMapper Is Nothing Then _
        Exit Function
  
    With myclass

        Call SendMessage(.View.hWnd, WM_SETREDRAW, SR_OFF, ByVal 0&)

        Do
            .HasFetchComplete = Not .RowMapper.HasNext(Row)
            
            If .HasFetchComplete Then
                Set .RowMapper = Nothing
                .HasFetchComplete = True
                DoEvents
                Exit Do
            End If
            
            Call .RowList.Add(Row)
         
            With .View.ListItems.Add(TEXT:=.RowList.Count)
                For i = Row.BaseIndex To Row.Count
                    .SubItems(i) = Row.ElementAt(i).ToString
                Next
            End With
    
            If GetInputState Then
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) Then
                    .HasFetchComplete = True
                    DoEvents
                    Exit Do
                End If
            End If
            
        Loop While Not (InterlockedDecrement(Limit) = 0)
        
        If .RowList.Count > 0 Then
            With .View
                Set .SelectedItem = .ListItems(.ListItems.Count)
                .SelectedItem.Selected = True
                .SelectedItem.EnsureVisible
            End With
        End If
        
        Call SendMessage(.View.hWnd, WM_SETREDRAW, SR_ON, ByVal 0&)
        RaiseEvent FetchComplete
    End With

    Fetch = Not myclass.HasFetchComplete
    
End Function

Public Sub Clear()
    
    With myclass
        .SelectedColumn = -1
        
        If Not .RowMap Is Nothing Then _
            .RowMap.Clear
            
        If Not .RowList Is Nothing Then _
            .RowList.Clear
            
        If Not .Columns Is Nothing Then _
            .Columns.Clear
        
        If Not .View Is Nothing Then
            With .View
                 .ListItems.Clear
                 .ColumnHeaders.Clear
                 .SortKey = 0
             End With
        End If
     End With
     
End Sub

Public Sub AddColumn(ByVal Name As GString, ByVal Datatype As VbVarType)
    
    Dim RightAlign As Boolean
    
    Select Case Datatype
        Case VbVarType.vbString: RightAlign = True
        Case VbVarType.vbLong: RightAlign = True
        Case VbVarType.vbDate:  RightAlign = True
        Case Else
    End Select
    
    With myclass.Columns
        Call .Add(Name, GNumeric(.Count + 1)):
    End With
    
    If Not myclass.View Is Nothing Then
        With myclass.View.ColumnHeaders
            If .Count = 0 Then _
                Call .Add(Index:=1, TEXT:=vbNullString)
    
            Call .Add(TEXT:=Name.Value)
    
            If RightAlign Then _
                .Item(.Count).Alignment = lvwColumnRight
        
        End With
    End If
    
End Sub

Public Sub Sort(ByVal Index As Long, ByVal SortDirection As SortOrder)
    
    Dim i As Long
    
    With myclass
        If Not .SelectedColumn = Index Then
            Call .RowList.Sort(SortDirection, GenericRowComparer.Build(Index))
        Else
            Call .RowList.Reverse
        End If
        
        .SortDirection = IIf(SortDirection = ascending, SortOrder.descending, SortOrder.ascending)
        .SelectedColumn = Index
    
        Call SendMessage(.View.hWnd, WM_SETREDRAW, SR_OFF, ByVal 0&)
        
        For i = .RowList.BaseIndex To .RowList.Count
           Call Me.Update(i)
        Next
        
        Call SendMessage(.View.hWnd, WM_SETREDRAW, SR_ON, ByVal 0&)
   
   End With
    
End Sub

Public Sub ColumnAutoSize(Optional ByVal Column As Variant)

    Dim i As Long
    Dim LstItm As MSComctlLib.ListItem
    
    With myclass.View
    
        If .ColumnHeaders.Count = 0 Then _
            Exit Sub
        
        Call SendMessage(.hWnd, WM_SETREDRAW, SR_OFF, ByVal 0&)
        
        Set LstItm = .ListItems.Add(TEXT:=.ColumnHeaders.Item(1).TEXT & SOME_CHARS)
        
        If IsMissing(Column) Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, 0&, ByVal LVSCW_AUTOSIZE_USEHEADER)
            
            For i = 2 To .ColumnHeaders.Count
                LstItm.SubItems(i - 1) = .ColumnHeaders.Item(i).TEXT & SOME_CHARS
                Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, i - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
            Next i
        
        Else
            i = myclass.View.ColumnHeaders(Column).Index
            LstItm.SubItems(i) = .ColumnHeaders.Item(i).TEXT & SOME_CHARS
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER)
        End If
        
        Call .ListItems.Remove(LstItm.Index): Set LstItm = Nothing
        Call SendMessage(.hWnd, WM_SETREDRAW, SR_ON, ByVal 0&)
    End With
    
End Sub

Private Sub style(ByVal Listview As MSComctlLib.Listview, ByVal style As DataviewerStyle)
    
    If Listview Is Nothing Then _
        Call Skynet.Throw(Me, "Private Sub Style").NullPointer("Listview Is Nothing")
        
    With Listview
        Select Case style
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
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Rows: " & .RowList.Count & vbNewLine & _
                            "Columns: " & .Columns.Count
    
    End With
End Function
