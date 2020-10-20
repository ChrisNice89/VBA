VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDataviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")
Option Explicit
'@PredeclaredId


Implements IGeneric
Implements IPrototype

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

'customized settings
Public Enum DataviewerStyle
    Default_Style
   
End Enum
'
Public Enum DataViewerColumnView
    Hide
    Show
End Enum
'
Public Enum DataViewerColumnType
    Numeric
    text
    Date
End Enum
'
Private Type Member
    Listview As MSComctlLib.Listview
    ListViewControl As CustomControl
    ImageList As MSComctlLib.ImageList
    Columns As GenericOrderedMap
    Rows As GenericList
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Columns() As IGenericReadOnlyList
    Set Columns = myclass.Columns
End Property

Public Property Get HasRows() As Boolean
    If myclass.Rows.Count > 0 Then
        HasRows = True
    Else
        HasRows = False
    End If
End Property

Public Function Build( _
                         ByVal ListViewControl As CustomControl, _
                         Optional ByVal Setting As DataviewerStyle = DataviewerStyle.Default_Style, _
                         Optional ByVal ImageList As MSComctlLib.ImageList = Nothing) As GenericDataviewer
    
    If Not TypeOf ListViewControl.Object Is MSComctlLib.Listview Then _
        Err.Raise 1, , "False Access.CustomControl: " & TypeName$(ListViewControl.Object)
          
    With myclass
        Set .ListViewControl = ListViewControl
        Set .Listview = ListViewControl.Object
        Set .ImageList = ImageList
        Call Style(.Listview, Setting)
    End With

    Set Build = Skynet.CreateInstance(New GenericDataviewer, Virtual.Pointer, Virtual.Size)
 
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Sub Clear()

     With myclass.Listview
         .ListItems.Clear
         .ColumnHeaders.Clear
         .SortKey = 0
     End With
     
End Sub
Public Sub AddColumn( _
                        ByVal Name As String, _
                        ByVal ColumnType As DataViewerColumnType, _
                        Optional ByVal RightAlign As Boolean = False)
    
    Call myclass.Columns.Add(GString(Name), GNumeric(ColumnType))
    
    With myclass.Listview.ColumnHeaders
        If .Count = 0 Then _
            Call .Add(1, vbNullString)
        
        Call .Add(Index:=.Count, Key:=Name)
        
        If RightAlign Then _
            .Item(.Count).Alignment = lvwColumnRight
         
    End With

End Sub

Public Function FillFromRecordset( _
                                    ByVal Data As ADODB.Recordset, _
                                    Optional ByVal AutoSize As Boolean = True, _
                                    Optional ByVal NumericRightAlign As Boolean = True) As Boolean
                                                               
    Dim Fields As ADODB.Fields
    Dim Field As ADODB.Field
    Dim Records As Long
   
    If Data Is Nothing Then _
        Call Err.Raise(1)

    Call Me.Clear

    With Data
        If .State = ISqlState.StateClosed Then _
           .Open
        
        If (.EOF And .BOF) Then _
           Err.Raise 1
        
        Set myclass.Columns = GenericOrderedMap.Build(.Fields.Count)
        
        If Data.RecordCount > 0 Then
            Set myclass.Rows = GenericList.Build(.RecordCount)
        Else
            Set myclass.Rows = GenericList.Build
        End If
        
        Set Fields = .Fields
        For Each Field In Fields
    
            Call Me.AddColumn(Field.Name, Date)
        Next
        
        Do While Not .EOF
            With myclass.Listview.ListItems.Add(text:=.AbsolutePosition)
                For Each Field In Fields
                   Call .ListSubItems.Add(text:=Field.Value)
                Next
            End With
            .MoveNext
        Loop
        
    End With

End Function

Public Sub ColumnAutoSize(Optional ByVal Column As Variant)

    Dim i As Long
    Dim LstItm As MSComctlLib.ListItem
    
     With myclass.Listview
         If .ColumnHeaders.Count = 0 Then Exit Sub
         Set LstItm = .ListItems.Add(text:=.ColumnHeaders.Item(1).text & SOME_CHARS)

         If IsMissing(Column) Then
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, 0&, ByVal LVSCW_AUTOSIZE_USEHEADER)

            For i = 2 To .ColumnHeaders.Count
                LstItm.SubItems(i - 1) = .ColumnHeaders.Item(i).text & SOME_CHARS
                Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, i - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
            Next i

         Else
            i = myclass.Listview.ColumnHeaders(Column).Index
            LstItm.SubItems(i - 1) = .ColumnHeaders.Item(i).text & SOME_CHARS
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
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Sub IGeneric_Dispose(): myclass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericDataviewer, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If TypeOf Other Is GenericDataviewer Then _
        IGeneric_Equals = (Me Is Other)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function

End Function

Private Property Get IGeneric_HashValue() As Long: ObjPtr (Me): End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericDataviewer): End Function
Private Function IGeneric_ToString() As String

End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericDataviewer: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericDataviewer: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


