VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSqlStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSqlStream
'*
'* Purpose:
'*
'*
'*
'* Notes:
'*
'*
'*
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericSequence
Implements IGenericIterator

Private Const DEBUG_MODE  As Boolean = False

Private Enum ISqlObjectState
    IsClosed_State = 0        'Indicates that the object is closed.
    IsOpen_State = 1          'Indicates that the object is open.
    IsConnecting_State = 2    'Indicates that the object is connecting.
    IsExecuting_State = 4     'Indicates that the object is executing a command.
    IsFetching_State = 8      'Indicates that the rows of the object are being retrieved.
End Enum

Private Type Member
    Recordset As ADODB.Recordset
    Field() As ADODB.Field
    Converter() As IGenericValue
    CacheLine() As IGenericValue
    FieldCount As Long
    HasRow As Boolean
    FetchedRows As Long
    Lookup As GenericHashMap
    HasLoaded As Boolean
End Type

Private This As Member
Attribute This.VB_VarHelpID = -1

Public Property Get Stream() As GenericSequence: Set Stream = GenericSequence.Stream(Me): End Property
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property

Public Function Of(ByVal Records As ADODB.Recordset) As IGenericSequence
  
    Dim newReader As Member
    
    If (Records Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Records")
    
    Set newReader.Recordset = Records
    Set newReader.Lookup = GenericHashMap.Build()
    
    Set Of = New GenericSqlStream
    Call System.Inject(Of, VarPtr(newReader), LenB(newReader))
    
End Function

Public Function Read() As Boolean
    
    Dim i As Long
    Dim Field As ADODB.Field
    
    If This.HasRow Then
            
        For i = 0 To This.FieldCount - 1
            Set Field = This.Field(i)
            If (Field.ActualSize > 0) Then
                Set This.CacheLine(i) = This.Converter(i).Of(Field.Value)
            Else
                Set This.CacheLine(i) = Nothing
            End If
            
        Next
 
        This.FetchedRows = This.FetchedRows + 1
        Call This.Recordset.MoveNext
        This.HasRow = Not This.Recordset.EOF
    
        Read = True
        Exit Function

    End If

    If (This.HasLoaded = False) Then _
        Call GenericError.Build(Me, "Private Function Read").InvalidOperation
    
    Erase This.Field ' release
    Erase This.Converter ' release
    Erase This.CacheLine ' release

End Function

Public Property Get FieldCount() As Long: FieldCount = This.FieldCount: End Property
Public Property Get FetchedRows() As Long: FetchedRows = This.FetchedRows: End Property

Public Property Get CacheSize() As Long

    If Not This.HasLoaded Then _
        Call StartFetch
    CacheSize = (This.Recordset.RecordCount - This.FetchedRows)
    
End Property

Public Property Get HasLoaded() As Boolean

    If Not This.HasLoaded Then _
        Call StartFetch
    HasLoaded = This.HasLoaded
        
End Property

Public Property Get HasRow() As Boolean
    
    If Not This.HasLoaded Then _
        Call StartFetch
    HasRow = This.HasRow

End Property

Public Property Get IndexOf(ByVal Field As gString) As Long: IndexOf = This.Lookup.IndexOf(Field): End Property
Public Property Get GetValue(ByVal Index As Long) As IGenericValue: Set GetValue = This.CacheLine(Index): End Property
Public Property Get GetName(ByVal Index As Long) As gString: Set GetName = This.Lookup.GetKeyAt(Index): End Property
Public Property Get FieldType(ByVal Index As Long) As VBA.VbVarType: FieldType = This.Converter(Index).VarType: End Property
Public Property Get Item(ByVal Field As gString) As IGenericValue: Set Item = This.CacheLine(This.Lookup.IndexOf(Field)): End Property

Public Function GetNames() As GenericArray

    If Not This.HasLoaded Then _
        Call StartFetch
    Set GetNames = This.Lookup.GetKeys

End Function

Public Sub CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    Dim i As Long
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").IsNothing("Target")
    
    With Target
        
        If (Index < .LowerBound) Or (Index >= .Length) Then _
            Call GenericError.Build(Me, "Public Sub CopyTo").IndexOutofRange("(Index < Target.LowerBound) Or (Index >= Target.Length)")
        
        If (.Length - Index < This.FieldCount) Then _
            Call GenericError.Build(Me, "Public Sub CopyTo").IndexOutofRange("(Index - .Length < This.FieldCount)")
        
        For i = 0 To This.FieldCount - 1: Set .ElementAt(Index + i) = This.CacheLine(i): Next

    End With
    
End Sub

Public Function ToArray() As GenericArray
    
    Dim i As Long
    
    Set ToArray = GenericArray.Build(This.FieldCount)
   
    With ToArray
        For i = 0 To This.FieldCount - 1: Set .ElementAt(i) = This.CacheLine(i): Next
    End With

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSqlQuery: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericSqlQuery: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSqlStream):  End Property

Private Function IGeneric_ToString() As String

        IGeneric_ToString = _
                        "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                        "Fetched rows: " & This.FetchedRows
    
End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Other Is Me: End Function

Private Function IGenericSequence_Iterator() As IGenericIterator

    Call StartFetch
    Set IGenericSequence_Iterator = Me
    
End Function

Private Sub StartFetch()
    
    Dim Field As ADODB.Field
    Dim Converter As IGenericValue
    Dim i As Long
    
    If (This.Recordset Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("This.Recordset")
    
    If (This.Lookup Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("This.Lookup")
    
'    Set Stream = This.Recordset
    
    With This.Recordset
        
        If (.State = IsClosed_State) Then _
            Call GenericError.Build(Me, "Public Function Of").InvalidInput("(Records.State = IsClosed)")
            
        Do While (.State And ISqlObjectState.IsExecuting_State) = ISqlObjectState.IsExecuting_State
            If DEBUG_MODE Then _
                Debug.Print "is executing..."
        Loop

        Do While (.State And ISqlObjectState.IsFetching_State) = ISqlObjectState.IsFetching_State
            If DEBUG_MODE Then _
                Debug.Print "is fetching..."
        Loop

        If Not (.State = ISqlObjectState.IsOpen_State) Then _
            Call GenericError.Build(Me, "Public Function Fetch").InvalidOperation

        If Not (.BOF Or .EOF) Then

            This.HasRow = True
            This.Recordset.CacheSize = 1024
            This.FieldCount = .Fields.Count
            Call This.Lookup.EnsureCapacity(This.FieldCount)
            ReDim This.CacheLine(This.FieldCount - 1)
            ReDim This.Converter(This.FieldCount - 1)
            ReDim This.Field(This.FieldCount - 1)
            
            For i = 0 To .Fields.Count - 1
                
                Set Field = .Fields(i)
                Select Case Field.Type
                    Case ISqlDataType.Boolean_Type: Set Converter = gBool
                    Case ISqlDataType.Char_Type: Set Converter = gString
                    Case ISqlDataType.Numeric_Type, ISqlDataType.SmallInt_Type: Set Converter = gNumeric
                    Case ISqlDataType.Currency_Type: Set Converter = gNumeric
                    Case ISqlDataType.Single_Type, ISqlDataType.Double_Type: Set Converter = gString
                    Case ISqlDataType.Decimal_Type: Set Converter = gNumeric
                    Case ISqlDataType.Integer_Type: Set Converter = gInt
                    Case ISqlDataType.BigInt_Type: Set Converter = gNumeric
                    Case ISqlDataType.UnsignedTinyInt_Type: Set Converter = gInt
                    Case ISqlDataType.Date_Type: Set Converter = gDate
                    Case ISqlDataType.VarWChar_Type: Set Converter = gString
                    Case ISqlDataType.LongVarChar_Type: Set Converter = gString
                    Case ISqlDataType.LongVarWChar_Type: Set Converter = gString  '203 Gibt eine lange Null-terminierte Unicode-Zeichenfolgenwert an.
                    Case Else
                        If DEBUG_MODE Then _
                            Debug.Print This.Field(i).Name & " :: " & This.Field(i).Type & " // Unknown Sql Field - Mapped to IGenericValue"
                        Set Converter = IGenericValue
                End Select
                
                Set This.Field(i) = Field
                Set This.Converter(i) = Converter
                Call This.Lookup.TryAdd(gString.Of(Field.Name), Field)
                
    '                Debug.Print .Fields(i).Name & " :: " & .Fields(i).Type & " :: " & This.Fields(i).SqlType.Instance.ClassName
    '                Debug.Print .Fields(i).Name & " AS Gstring"
    '                Debug.Print "Public " & .Fields(i).Name & " AS " & This.Fields(i).SqlType.Instance.ClassName
    '                Debug.Print "Set ." & .Fields(i).Name & " = Row.Item(Table.Sanierung." & .Fields(i).Name & ")"
    '                Debug.Print "Set ." & .Fields(i).Name & " = GString.Of(""" & .Fields(i).Name & """)"
            Next

        End If
    End With
    
    This.FetchedRows = 0
    This.HasLoaded = True
    
End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = Me: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean

    If Me.Read Then
        Set Out = Me
        IGenericIterator_HasNext = True
   End If
   
End Function

Private Sub Stream_EndOfRecordset(fMoreData As Boolean, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    This.HasRow = False
End Sub
