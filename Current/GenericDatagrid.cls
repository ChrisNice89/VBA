VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDatagrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, ParameterNotUsed, UnassignedVariableUsage, AssignedByValParameter
'@Folder "DataGrid"
Option Explicit

Implements IGeneric

Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hWnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare PtrSafe Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, _
    ByVal lprcUpdate As Long, _
    ByVal hrgnUpdate As Long, _
    ByVal fuRedraw As Long) As Long

Private Declare PtrSafe Function LockWindowUpdate Lib "user32" ( _
    ByVal hwndLock As Long) As Long

Public Event Clear()
Public Event RowClick(ByVal Index As Long) ', ByVal Row As GenericSequence)
Public Event RowSelected(ByVal Index As Long) ', ByVal Row As GenericSequence)
Public Event ColumnSelected(ByVal Index As Long, ByVal Column As GenericSequence)
Public Event ColumnClick(ByVal Index As Long, ByRef Cancel As Boolean)
Public Event PageSelected(ByVal Index As Long, ByVal PageSize As Long, ByVal HasPreviousPage As Boolean, HasNextPage As Boolean)
Public Event Process(ByVal Element As Object, ByRef Row() As IGenericValue)

Private Const DEFAULT_PAGE_SIZE As Long = 25
Private Const DEFAULT_SORT_ORDER As Long = SortOrder.Ascending
Private Const INITIAL_ROW_SIZE  As Long = 1024
Private Const FONT_NAME As String = "Tahoma"
'@Ignore ConstantNotUsed
Private Const FONT_SIZE As Long = 9

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
'@Ignore ConstantNotUsed
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
'@Ignore ConstantNotUsed
Private Const LVM_SCROLL As Long = &H1014

Private Const WM_SETREDRAW As Long = &HB
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ALLCHILDREN As Long = &H80
'@Ignore ConstantNotUsed
Private Const RDW_UPDATENOW As Long = &H100
'@Ignore ConstantNotUsed
Private Const RDW_ERASE As Long = &H4

Private Type Class
    StatusBar As StatusBar
    Rows As GenericOrderedList
    View As GenericOrderedList
    Columns As GenericHashMap
    SelectedRow As GenericArray
    ListView As MSComctlLib.ListView
    ListViewParent As CustomControl
    CurrentRow As Long
    CurrentColumn As Long
    SortedColumn As Long
    Direction As SortOrder
    VirtualMode As Boolean
    PageSize As Long
    PageIndex As Long
End Type
Private This As Class

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ControlEvent As MSComctlLib.ListView
Attribute ControlEvent.VB_VarHelpID = -1

Public Property Set ElementAt(ByVal i As Long, j As Long, Element As IGenericValue)
    
    If Not (i = This.CurrentRow) Then _
        Set This.SelectedRow = This.Rows.ElementAt(i)
    This.CurrentRow = i
    
    Set This.SelectedRow.ElementAt(j) = Element
    
End Property

Public Property Get ElementAt(ByVal i As Long, j As Long) As IGenericValue
    
    If Not (i = This.CurrentRow) Then _
        Set This.SelectedRow = This.Rows.ElementAt(i)
    This.CurrentRow = i
    
    Set ElementAt = This.SelectedRow.ElementAt(j)
    
End Property

Public Property Get IsVirtualMode() As Boolean: IsVirtualMode = This.VirtualMode: End Property

Public Property Get SelectedRow() As Long: SelectedRow = This.CurrentRow: End Property
Public Property Get HasRowSelected() As Boolean: HasRowSelected = This.CurrentRow >= 0: End Property

Public Property Get SelectedColumn() As Long: SelectedColumn = This.CurrentColumn: End Property

Public Property Get Header(ByVal Index As Long) As gString: Set Header = This.Columns.GetKeyAt(Index): End Property
Public Property Get IndexOf(ByVal Name As gString) As Long: IndexOf = This.Columns.IndexOf(Name): End Property

Public Property Get SelectedPage() As Long: SelectedPage = This.PageIndex: End Property
Public Property Get CountRows() As Long: CountRows = This.Rows.Elements.Count: End Property
Public Property Get CountColumns() As Long: CountColumns = This.Columns.Elements.Count: End Property

Public Property Get PageCount() As Long
    
    Dim c As Long, p As Long
    
    c = Me.CountRows
    If (c = 0) Then _
        Exit Property
    
    p = Me.PageSize
    If (p = 0) Then _
        Exit Property
        
    PageCount = VBA.Fix((c / p))
            
    If Not ((c Mod p) = 0) Then _
        PageCount = PageCount + 1
        
End Property

Public Property Get DisplayedRows() As Long: DisplayedRows = This.ListView.ListItems.Count: End Property
Public Property Let PageSize(ByVal Size As Long): This.PageSize = Size: End Property
Public Property Get PageSize() As Long: PageSize = This.PageSize: End Property
Public Property Get HasNextPage() As Boolean: HasNextPage = This.PageIndex < Me.PageCount: End Property
Public Property Get HasPreviousPage() As Boolean: HasPreviousPage = This.PageIndex > 1: End Property
Public Function GetColumns() As GenericArray: Set GetColumns = This.Columns.GetKeys: End Function

Public Function Build(ByVal ListView As CustomControl, Optional ByVal StatusBar As CustomControl) As GenericDatagrid
    
    Dim newGrid As Class
    
    With newGrid
        Set .ListView = ListView.Object
        Set .ListViewParent = ListView
        Call .ListView.ListItems.Clear
        Call .ListView.ColumnHeaders.Clear
        
        If Not StatusBar Is Nothing Then
            Set .StatusBar = StatusBar.Object
            With .StatusBar
                .Enabled = True
                .Font.Name = FONT_NAME
                .Font.Size = 8.25
                .MousePointer = ccArrow
                .Style = sbrNormal
                
                With newGrid.StatusBar.Panels
                    
                    Do While .Count > 0: Call .Remove(.Count): Loop
        
                    With .Add
                        .Visible = False
                        .Bevel = sbrInset
                        .Alignment = sbrLeft
                        .AutoSize = sbrContents
                        .Style = sbrText
                        .ToolTipText = "Maus ist auf Panel Text"
                    End With
                    
                    With .Add
                        .Visible = False
                        .Bevel = sbrInset
                        .Alignment = sbrRight
                        .AutoSize = sbrContents
                        .Style = sbrText
                        .ToolTipText = "Maus ist auf Panel Text"
                    End With
                    
                    With .Add
                        .Visible = False
                        .Bevel = sbrInset
                        .Alignment = sbrRight
                        .AutoSize = sbrContents
                        .Style = sbrText
                        .ToolTipText = "Maus ist auf Panel Text"
                    End With
                    
                    With .Add
                        .Visible = False
                        .Bevel = sbrInset
                        .Alignment = sbrRight
                        .AutoSize = sbrSpring
                        .Style = sbrText
                        .ToolTipText = "Maus ist auf Panel Text"
                    End With
               
                End With
                
                Call .Refresh
                
            End With
        End If
        
        Set .Rows = GenericOrderedList.Build()
        Set .View = GenericOrderedList.Build()
        Set .Columns = GenericHashMap.Build(32)
        .Direction = SortOrder.Descending
        .CurrentRow = -1
        .CurrentColumn = -1
        
        With .ListView
            .ColumnHeaders.Clear
            .ListItems.Clear
            .SortKey = 0
            .Font.Name = FONT_NAME
            .Font.Size = 8.25
            .FlatScrollBar = False 'BUG!
            .Checkboxes = False
            .BorderStyle = ccFixedSingle
            .LabelEdit = lvwManual
            .LabelWrap = False
            .ForeColor = -2147483630
            .BackColor = VBA.RGB(220, 230, 242) ' -2147483643 '
            .Appearance = ccFlat
            .BorderStyle = ccFixedSingle
            .Enabled = True
            .HotTracking = False
            .MousePointer = ccArrow
            .HoverSelection = False
            .Sorted = False
            .AllowColumnReorder = True
            .FullRowSelect = True
            .Multiselect = True
            .HideSelection = False 'Returns or sets a Boolean that specifies whether selected text remains highlighted when a control does not have the focus. Read/write.
            .HideColumnHeaders = False
            .HotTracking = True
            .GridLines = True
            .View = lvwReport
        End With
        
    End With
    
    Set Build = New GenericDatagrid
    Call System.Inject(Build, VarPtr(newGrid), LenB(newGrid))
    
End Function

Public Sub EnsureCapacity(ByVal Count As Long)
    
    If (This.Rows Is Nothing) Then _
        Exit Sub
        
    If (Count < 0) Then _
        Exit Sub
        
    Call This.Rows.EnsureCapacity(Count)
    
End Sub

Public Sub AddColumns(ParamArray Elements() As Variant)

    Dim i As Long
    For i = 0 To UBound(Elements)
       Call Me.AddColumn(Elements(i))
    Next

End Sub

Public Sub AddColumn(ByVal Name As IGenericValue, Optional ByVal RightAlign As Boolean = False)
    
    Dim Header As MSComctlLib.ColumnHeader
    Dim Delegate As DataGridColumnDelegate
    
    If (This.ListView.ColumnHeaders.Count = 0) Then
        Call This.ListView.ColumnHeaders.Add(, , vbNullString)
        Set ControlEvent = This.ListView
    End If
    
    If (Name Is Nothing) Then _
        Set Name = gString.Of(This.Columns.Elements.Count + 1)
        
    If This.Columns.TryAdd(Name, Nothing) Then
        
        Set Delegate = New DataGridColumnDelegate
        Delegate.Index = This.Columns.Elements.Count - 1
        
        Set This.Columns.ElementAt(This.Columns.Elements.Count - 1) = Delegate
        Set Header = This.ListView.ColumnHeaders.Add(, , Name.ToString)
    
        If RightAlign Then _
            Header.Alignment = lvwColumnRight
        
    End If
    
End Sub

Public Sub RightAlign(ByVal i As Long)
    This.ListView.ColumnHeaders.Item(i + 1).Alignment = lvwColumnRight
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    Call This.ListView.ListItems.Remove(Index)
    Call This.View.RemoveAt(Index - 1)
End Sub

Public Sub Stream(ByVal Reader As GenericSqlStream, Optional ByVal PageSize As Long = DEFAULT_PAGE_SIZE)
    
    Dim i As Long
    Dim FieldType As VBA.VbVarType
    
    Call Me.BeginUpdate
    Call Me.Clear
    
    With Reader
        If .HasLoaded Then
            
            This.PageSize = PageSize
            
            For i = 0 To .FieldCount - 1
                FieldType = .FieldType(i)
                Call Me.AddColumn(.GetName(i), Not (FieldType = vbString Or FieldType = vbDate))
            Next
          
            Call Fetch(Reader)
            Call Me.SelectPage(1)
            
        End If
    End With
    
    Call Me.EndUpdate
    
End Sub

Public Sub AddRow(ParamArray Elements() As Variant)

    Dim Element As IGenericValue
    Dim LstItm As MSComctlLib.ListItem
    Dim i As Long
    
    Call Me.LastPage
    
    Set This.SelectedRow = GenericArray.Build(This.Columns.Elements.Count)
    
    Call This.Rows.Add(This.SelectedRow)
    Call This.View.Add(This.SelectedRow)
    
    Call Me.BeginUpdate
    Set LstItm = This.ListView.ListItems.Add(, , This.View.Elements.Count)

    With This.SelectedRow
        For i = 0 To UBound(Elements)
            Set Element = Elements(i)
            If Not Element Is Nothing Then
                Set .ElementAt(i) = Element
                LstItm.SubItems(i + 1) = Element.ToString
            End If
        Next
    End With

    Call Me.SelectRow(This.View.Elements.Count - 1)
    Call Me.EndUpdate
    
End Sub
'TODO
'@Ignore EmptyMethod
Public Sub DeleteRow(ByVal Index As Long)
    
End Sub

'@Ignore EmptyMethod
Public Sub DeleteSelectedRow()
    
End Sub

Public Sub AddAll(ByVal Collection As IGenericCollection, Optional ByVal PageSize As Long = DEFAULT_PAGE_SIZE)
    
    Dim Row As GenericArray
    Dim Cache() As IGenericValue
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim i As Long, j As Long
    
    If Collection Is Nothing Then _
        Exit Sub
        
    If Collection.Count = 0 Then _
        Exit Sub
    
    Call This.Rows.EnsureCapacity(Collection.Count)
    j = This.Columns.Elements.Count
    ReDim Cache(j)
    
    With Collection.Iterator
        Do While .HasNext(Element)
            RaiseEvent Process(Element, Cache)
            Set Row = GenericArray.Build(j)
            For i = 0 To j - 1: Set Row.ElementAt(i) = Cache(i): Next
            Call This.Rows.Add(Row)
        Loop
    End With
  
    This.PageSize = PageSize
    Call Me.SelectPage(1)
    
End Sub

Public Sub BeginUpdate()
    
    Dim pl As MSComctlLib.Panel

    If Not This.VirtualMode Then
     
        '@Ignore ValueRequired
        Call SendMessage(This.ListView.hWnd, ByVal WM_SETREDRAW, False, 0&)
        '@Ignore ValueRequired
        Call LockWindowUpdate(This.ListView.hWnd)
        
        If Not This.StatusBar Is Nothing Then
            For Each pl In This.StatusBar.Panels
                pl.Visible = False
                pl.Text = vbNullString
            Next

            With This.StatusBar.Panels.Item(1)
                .Bevel = sbrInset
                .Alignment = sbrLeft
                .AutoSize = sbrSpring
                .Style = sbrText
                .Text = " Updating, please wait... "
                .Visible = True
            End With

            Call This.StatusBar.Refresh
        End If
        This.VirtualMode = True
    End If
    
End Sub

Public Sub EndUpdate()
    
    Dim pl As MSComctlLib.Panel
    
    If This.VirtualMode Then
    
        Call Me.AutoSize
        
        If This.ListView.ListItems.Count > 0 Then _
            This.ListView.ListItems.Item(1).Selected = False
        
        '@Ignore ValueRequired
        Call SendMessage(This.ListView.hWnd, ByVal WM_SETREDRAW, True, 0&)
        '@Ignore ValueRequired
        Call RedrawWindow(This.ListView.hWnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN)
        Call LockWindowUpdate(0&)
 
        If Not This.StatusBar Is Nothing Then
                
            For Each pl In This.StatusBar.Panels
                pl.Visible = False
            Next
            
             With This.StatusBar.Panels
             
                With .Item(1)
                    .Bevel = sbrInset
                    .Alignment = sbrLeft
                    .AutoSize = sbrContents
                    .Style = sbrText
                    .Text = " Total records: " & VBA.Format$(Me.CountRows, "#,##0")
                    .ToolTipText = "Maus ist auf Panel Text"
                End With
                
                With .Item(2)
                    .Bevel = sbrInset
                    .Alignment = sbrRight
                    .AutoSize = sbrContents
                    .Style = sbrText
                    .Text = " Records per page: " & VBA.Format$(Me.PageSize, "#,##0")
                    .ToolTipText = "Maus ist auf Panel Text"
                End With
                
                With .Item(3)
                    .Bevel = sbrInset
                    .Alignment = sbrRight
                    .AutoSize = sbrContents
                    .Style = sbrText
                    .Text = " Page " & Me.SelectedPage & " of " & Me.PageCount
                    .ToolTipText = "Maus ist auf Panel Text"
                End With
                
                With .Item(4)
                    .Bevel = sbrInset
                    .Alignment = sbrRight
                    .AutoSize = sbrSpring
                    .Style = sbrText
'                    .Text = " Displaying records from " & This.ListView.ListItems.Item(1).Text & " to " & This.ListView.ListItems.Item(This.ListView.ListItems.Count).Text
                    .Text = " Displaying records: " & Me.DisplayedRows
                    .ToolTipText = "Maus ist auf Panel Text"
                End With
               
            End With
            For Each pl In This.StatusBar.Panels
                pl.Visible = True
            Next
            Call This.StatusBar.Refresh
        End If
     
        This.VirtualMode = False
    End If
    
End Sub

Public Sub Clear()
    
    If (This.ListView Is Nothing) Then _
        Exit Sub
        
    This.CurrentRow = -1
    This.CurrentColumn = -1
    This.PageIndex = 0
    This.PageSize = 0
    This.SortedColumn = 0
    
    Set This.SelectedRow = Nothing
    
    Call This.ListView.ColumnHeaders.Clear
    Call This.ListView.ListItems.Clear
  
    Call This.View.Elements.Clear
    Call This.Rows.Elements.Clear
    Call This.Columns.Elements.Clear
    
    RaiseEvent Clear
    
End Sub

Public Function GetRow(ByVal i As Long) As GenericSequence
    Set GetRow = GenericSequence.Stream(This.Rows.GetAt(i))
End Function

Public Function FindPage(ByVal Row As Long) As Long
    
    Dim i As Long
    
    i = This.PageIndex
    
    If (Row <= Me.CountRows) Then
        Row = VBA.Abs(Row)
        i = VBA.Fix(Row / This.PageSize)
        If Not ((Row Mod This.PageSize) = 0) Then _
            i = i + 1
    End If
    
    FindPage = i
    
End Function

Public Sub Shuffle()
    Call This.Rows.Shuffle
    Call Me.Refresh
End Sub

Private Sub Class_Terminate()
    Call Me.Clear
End Sub

Public Sub AutoSize(Optional ByVal Index As Long = -1)

    Dim i As Long
    Dim VirtualMode As Boolean
    
    With This.ListView
        If .ColumnHeaders.Count = 0 Then _
            Exit Sub
        
        VirtualMode = This.VirtualMode
        
        If Not VirtualMode Then _
            Call Me.BeginUpdate
      
        If Index = -1 Then
        
            '@Ignore ValueRequired
            If Not (.ColumnHeaders.Item(1).Width = 0) Then _
                Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, 0&, ByVal LVSCW_AUTOSIZE_USEHEADER)
            
            For i = 1 To .ColumnHeaders.Count - 1
                '@Ignore ValueRequired
                Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER)
            Next i
        
        Else
    
            '@Ignore ValueRequired
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, i - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
            
        End If
       
        If Not VirtualMode Then _
            Call Me.EndUpdate
            
    End With
 
End Sub

Public Sub ChangePageSize(ByVal Size As Long)
    
    Dim i As Long
    
    If (Me.CountRows = 0) Then _
        Exit Sub
    
    If (Me.DisplayedRows = 0) Then _
        This.PageSize = 0
        
    If (Size = This.PageSize) Then _
        Exit Sub
    
    Call Me.BeginUpdate
    
    If (Size < This.PageSize) Then
        
        With This.ListView.ListItems
            For i = .Count To Size + 1 Step -1
                Call .Remove(i)
                Call This.View.RemoveAt(i - 1)
            Next
        End With
        
    Else
        
        If Size > Me.CountRows Then
            Call This.View.EnsureCapacity(Me.CountRows)
        Else
            Call This.View.EnsureCapacity(Size)
        End If
        
        If (This.SortedColumn > 0) Then
            Call LoadView(This.Rows.Stream.OrderBy(This.Columns.GetItemAt(This.SortedColumn), This.Direction).Skip(This.PageIndex * This.PageSize).Take(Size - This.PageSize))
        Else
            Call LoadView(This.Rows.Stream.Skip(This.PageIndex * This.PageSize).Take(Size - This.PageSize))
        End If
       
    End If
    
    This.PageSize = Size
    This.PageIndex = 1
    
    Call Me.EndUpdate
    
End Sub

Public Sub SortColumn(ByVal Index As Long, Optional ByVal Direction As SortOrder = DEFAULT_SORT_ORDER)
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 0 Or Index >= Me.CountColumns Then
        'error
    End If
    
    If (This.SortedColumn = Index) Then _
        If (This.Direction = Direction) Then _
            Exit Sub
            
    Me.BeginUpdate
    
    This.SortedColumn = Index
    This.Direction = Direction
    
    Call UpdateView(This.Rows.Stream.OrderBy(This.Columns.GetItemAt(Index), Direction).Partition(This.PageIndex - 1, This.PageSize))
    
    Me.EndUpdate
    
End Sub

Public Sub FilterTop(ByVal Top As Long, ByVal Index As Long, Optional ByVal Direction As SortOrder = DEFAULT_SORT_ORDER)
    
    Dim i As Long
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 1 Or Index > Me.CountColumns Then
        'error
    End If
    
    If (This.SortedColumn = Index) Then _
        If (This.Direction = Direction) Then _
            If (This.PageSize = Top) Then _
                Exit Sub
            
    Call Me.BeginUpdate
    
    This.SortedColumn = Index
    This.Direction = Direction
    This.CurrentColumn = Index
    This.PageIndex = 1
    
    i = Me.CountRows
    If (Top > i) Then _
        Top = i
        
    This.PageSize = Top
    
    If (This.PageSize > i) Then
        Call This.View.EnsureCapacity(i)
    Else
        Call This.View.EnsureCapacity(This.PageSize)
    End If
    
    If (Me.DisplayedRows > 0) Then
        Call UpdateView(This.Rows.Stream.OrderBy(This.Columns.GetItemAt(Index), Direction).Take(Top))
    Else
        Call LoadView(This.Rows.Stream.OrderBy(This.Columns.GetItemAt(Index), Direction).Take(Top))
    End If
    
    Call Me.EndUpdate
    
    RaiseEvent PageSelected(This.PageIndex, Me.DisplayedRows, Me.HasPreviousPage, Me.HasNextPage)
    
End Sub

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function SelectPage(ByVal Index As Long) As GenericDatagrid
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 1 Or Index > Me.PageCount Then
        'error
    End If
    
    Set SelectPage = Me
    
    If Not (This.PageIndex = Index) Then
        
        Me.BeginUpdate
        
        This.SortedColumn = 0
        This.PageIndex = Index
        
        If This.PageSize > Me.CountRows Then
            Call This.View.EnsureCapacity(Me.CountRows)
        Else
            Call This.View.EnsureCapacity(This.PageSize)
        End If
        
        If (Me.DisplayedRows > 0) Then
            Call UpdateView(This.Rows.Stream.Partition(This.PageIndex - 1, This.PageSize))
        Else
            Call LoadView(This.Rows.Stream.Partition(This.PageIndex - 1, This.PageSize))
        End If
        
        Me.EndUpdate
        
        RaiseEvent PageSelected(This.PageIndex, Me.DisplayedRows, Me.HasPreviousPage, Me.HasNextPage)
        
    End If
    
End Function

Public Function SelectRecord(ByVal Index As Long) As GenericDatagrid
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 1 Or Index > Me.CountRows Then
        'error
    End If
    
    Call Me.SelectPage(Me.FindPage(Index))
    Call Me.SelectRow((Index - 1) - ((This.PageIndex - 1) * This.PageSize))
    
    Set SelectRecord = Me
    
End Function

Public Sub UnselectCurrentRow()

     With This.ListView
        If Not (.SelectedItem Is Nothing) Then
            .SelectedItem.Selected = False
            Set .SelectedItem = Nothing
        End If
    End With
        
    This.CurrentRow = -1
    Set This.SelectedRow = Nothing

End Sub

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function SelectRow(ByVal Index As Long) As GenericDatagrid
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 0 Or Index >= Me.CountRows Then
        'error
    End If
    
    Set SelectRow = Me
    
    If Not (Index = This.CurrentRow) Then
        
        With This.ListView
           
            If Not (.SelectedItem Is Nothing) Then
                .SelectedItem.Selected = False
                Set .SelectedItem = Nothing
            End If
            
            Set .SelectedItem = .ListItems(Index + 1)
            .SelectedItem.Selected = True
            Call .SelectedItem.EnsureVisible
            
        End With
        
        This.CurrentRow = Index
        Set This.SelectedRow = This.View.GetAt(This.CurrentRow)
        
        Call This.ListViewParent.SetFocus
        RaiseEvent RowSelected(This.CurrentRow)
    
    End If
    
End Function

Public Function SelectColumn(ByVal Index As Long) As GenericSequence
    'TODO
    '@Ignore EmptyIfBlock
    If Index < 0 Or Index >= Me.CountColumns Then
        'error
    End If
    
    This.CurrentColumn = Index
    Set SelectColumn = This.Rows.Stream.Apply(This.Columns.GetItemAt(This.CurrentColumn))
    RaiseEvent ColumnSelected(Me.SelectedColumn, SelectColumn)
    
End Function

Public Sub RefreshView()
    
    If (Me.PageCount = 0) Then _
        Exit Sub
 
    Me.BeginUpdate
    Call UpdateView(This.View)
    Me.EndUpdate
    
End Sub

Public Sub Refresh()
    
    If (Me.PageCount = 0) Then _
        Exit Sub
 
    Me.BeginUpdate
    
    This.PageIndex = 1
    
    If Me.DisplayedRows = 0 Then
        Call LoadView(This.Rows.Stream.Take(This.PageSize))
    Else
        Call UpdateView(This.Rows.Stream.Take(This.PageSize))
    End If
    
    Me.EndUpdate
    
    RaiseEvent PageSelected(This.PageIndex, Me.DisplayedRows, Me.HasPreviousPage, Me.HasNextPage)
 
End Sub

Public Sub FirstPage(): Call SelectPage(1): End Sub
Public Sub LastPage(): Call SelectPage(Me.PageCount): End Sub

Public Function NextPage() As Boolean
    
    If Me.HasNextPage Then _
        Call SelectPage(This.PageIndex + 1)
     
    NextPage = Me.HasNextPage
    
End Function

Public Function PreviousPage() As Boolean

    If Me.HasPreviousPage Then _
        Call SelectPage(This.PageIndex - 1)
     
    PreviousPage = Me.HasPreviousPage
    
End Function

Private Sub Fetch(ByVal Reader As GenericSqlStream)
    
    Dim CacheSize As Long
    
    With Reader
        If (.HasLoaded And .HasRow) Then
            
            CacheSize = .CacheSize
            If (CacheSize = 0) Then _
                CacheSize = INITIAL_ROW_SIZE
                
            Call This.Rows.EnsureCapacity(CacheSize)
            
            Do While .Read: Call This.Rows.Add(.ToArray): Loop
            
        End If
    End With
    
End Sub

Public Sub ReverseView()
   
    '@Ignore VariableNotAssigned
    Dim Row As GenericArray
    Dim Element As IGenericValue
    
    Dim LstItms As MSComctlLib.ListItems
    Dim LstItm As MSComctlLib.ListItem
    
    Dim i As Long, RowCount As Long
    
    Set LstItms = This.ListView.ListItems
    
    Call Me.BeginUpdate
    
    With This.View.Reverse.Range
        Do While .HasNext(Row)
            With Row
            
                RowCount = RowCount + 1
                Set LstItm = LstItms.Item(RowCount)
                
                LstItm.Text = RowCount
                LstItm.Selected = False
           
                For i = 1 To .Length - 1
                    Set Element = .ElementAt(i - 1)
                    If Not Element Is Nothing Then
                        LstItm.SubItems(i) = Element.ToString
                    Else
                        LstItm.SubItems(i) = vbNullString
                    End If
                Next
                
            End With
        Loop
    End With
    
    Call Me.UnselectCurrentRow
    Call Me.EndUpdate
    
End Sub

Private Sub LoadView(ByVal Rows As IGenericSequence)
    
    '@Ignore VariableNotAssigned
    Dim Row As GenericArray
    Dim Element As IGenericValue
    
    Dim LstItms As MSComctlLib.ListItems
    Dim LstItm As MSComctlLib.ListItem
    
    Dim i As Long, RowCount As Long
    
    Set LstItms = This.ListView.ListItems
    RowCount = LstItms.Count
    
    With Rows.Iterator
        Do While .HasNext(Row)
            With Row
                
                RowCount = RowCount + 1
                Set LstItm = LstItms.Add(, , RowCount)
                Call This.View.Add(Row)
                
                For i = 0 To .Length - 1
                    Set Element = .ElementAt(i)
                    If Not Element Is Nothing Then _
                        LstItm.SubItems(i + 1) = Element.ToString
                Next
            End With
        Loop
    End With
   
End Sub

Private Sub UpdateView(ByVal Rows As IGenericSequence)

    '@Ignore VariableNotAssigned
    Dim Row As GenericArray
    Dim Element As IGenericValue
    
    Dim LstItms As MSComctlLib.ListItems
    Dim LstItm As MSComctlLib.ListItem
    
    Dim i As Long, RowCount As Long
    
    Set LstItms = This.ListView.ListItems
    
    With Rows.Iterator
        Do While .HasNext(Row)
            With Row
            
                RowCount = RowCount + 1
                If (RowCount <= LstItms.Count) Then
                    Set LstItm = LstItms.Item(RowCount)
                    Set This.View.ElementAt(RowCount - 1) = Row
                Else
                    Set LstItm = LstItms.Add()
                    Call This.View.Add(Row)
                End If

                LstItm.Text = RowCount
                LstItm.Selected = False
           
                For i = 0 To .Length - 1
                    Set Element = .ElementAt(i)
                    If Not Element Is Nothing Then
                        LstItm.SubItems(i + 1) = Element.ToString
                    Else
                        LstItm.SubItems(i + 1) = vbNullString
                    End If
                Next
                
            End With
                  
        Loop
    End With
    
    If (LstItms.Count > RowCount) Then
        With LstItms
            For i = .Count To RowCount + 1 Step -1
                Call .Remove(i)
                Call This.View.RemoveAt(i - 1)
            Next
        End With
    End If
    
End Sub

Private Sub ControlEvent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    '@Ignore VariableNotAssigned
    Dim i As Long, Cancel As Boolean
    i = ColumnHeader.Index
    
    If (i = 1) Then _
        Exit Sub
        
    Call Me.SelectColumn(i - 2)
    RaiseEvent ColumnClick(i - 2, Cancel)
    
    If Cancel Then _
        Exit Sub
        
    Call Me.SortColumn(i - 2, VBA.IIf(This.Direction = Ascending, SortOrder.Descending, SortOrder.Ascending))
    Call Me.UnselectCurrentRow
  
End Sub

Private Sub ControlEvent_DblClick()
    RaiseEvent RowClick(This.CurrentRow)
End Sub

Private Sub ControlEvent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call Me.SelectRow(Item.Index - 1)
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericDatagrid: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericDatagrid: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericDatagrid):  End Property

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Me Is Other: End Function

 Private Function IGeneric_ToString() As String
    With This
        IGeneric_ToString = _
                            "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                            "Rows: " & .Rows.Elements.Count & vbNewLine & _
                            "Columns: " & gString.Join(This.Columns, vbNewLine).ToString
    
    End With
End Function


