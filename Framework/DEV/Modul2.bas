Attribute VB_Name = "Modul2"
Public Function HeapSort(Keys)
   Dim Base As Long: Base = LBound(Keys)                    ' array index base
   Dim n As Long: n = UBound(Keys) - LBound(Keys) + 1       ' array Size
   ReDim Index(Base To Base + n - 1) As Long                ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(Base + i) = Base + i: Next     ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                           ' generate ordered heap
      Heapify Keys, Index, i, n
      Next
   For m = n To 2 Step -1
      Exchange Index, 0, m - 1                              ' move highest element to top
      Heapify Keys, Index, 0, m - 1
      Next
   HeapSort = Index
   End Function

Private Sub Heapify(Keys, Index() As Long, ByVal i1 As Long, ByVal n As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim Base As Long: Base = LBound(Index)
   Dim nDiv2 As Long: nDiv2 = n \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim K As Long: K = 2 * i + 1
      If K + 1 < n Then
         If Keys(Index(Base + K)) < Keys(Index(Base + K + 1)) Then K = K + 1
         End If
      If Keys(Index(Base + i)) >= Keys(Index(Base + K)) Then Exit Do
      Exchange Index, i, K
      i = K
      Loop
   End Sub

Private Sub Exchange(A() As Long, ByVal i As Long, ByVal j As Long)
   Dim Base As Long: Base = LBound(A)
   Dim temp As Long: temp = A(Base + i)
   A(Base + i) = A(Base + j)
   A(Base + j) = temp
   End Sub

Public Sub TestHeapSort()
   Debug.Print "Start"
  
    Dim Keys: Keys = GenerateArrayWithRandomValues()
    Dim Index: Index = HeapSort(Keys)
    VerifyIndexIsSorted Keys, Index
 
   Debug.Print "OK"
   End Sub

Private Function GenerateArrayWithRandomValues()
   Dim n As Long: n = 100
   ReDim A(0 To n - 1) As String
   Dim i As Long
   A(0) = "c"
    A(0) = ""
    A(0) = "B"
    A(0) = "a"
   GenerateArrayWithRandomValues = A
   End Function

Private Sub VerifyIndexIsSorted(Keys, Index)
   Dim i As Long
   For i = LBound(Index) To UBound(Index) - 1
    Debug.Print Keys(Index(i))
      If Keys(Index(i)) > Keys(Index(i + 1)) Then
         Err.Raise vbObjectError, , "Index array is not sorted!"
         End If
      Next
   End Sub
