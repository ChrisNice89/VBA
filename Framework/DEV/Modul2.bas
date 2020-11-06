Attribute VB_Name = "Modul2"
'@Folder "Entwicklung"
Public Function HeapSort(Keys)
   Dim Base As Long: Base = LBound(Keys)                    ' array index base
   Dim N As Long: N = UBound(Keys) - LBound(Keys) + 1       ' array Size
   ReDim index(Base To Base + N - 1) As Long                ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To N - 1: index(Base + i) = Base + i: Next     ' fill index array
   For i = N \ 2 - 1 To 0 Step -1                           ' generate ordered heap
      Heapify Keys, index, i, N
      Next
   For m = N To 2 Step -1
      Exchange index, 0, m - 1                              ' move highest element to top
      Heapify Keys, index, 0, m - 1
      Next
   HeapSort = index
   End Function

Private Sub Heapify(Keys, index() As Long, ByVal i1 As Long, ByVal N As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim Base As Long: Base = LBound(index)
   Dim nDiv2 As Long: nDiv2 = N \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim K As Long: K = 2 * i + 1
      If K + 1 < N Then
         If Keys(index(Base + K)) < Keys(index(Base + K + 1)) Then K = K + 1
         End If
      If Keys(index(Base + i)) >= Keys(index(Base + K)) Then Exit Do
      Exchange index, i, K
      i = K
      Loop
   End Sub

Private Sub Exchange(a() As Long, ByVal i As Long, ByVal j As Long)
   Dim Base As Long: Base = LBound(a)
   Dim temp As Long: temp = a(Base + i)
   a(Base + i) = a(Base + j)
   a(Base + j) = temp
   End Sub

Public Sub TestHeapSort()
   Debug.Print "Start"
  
    Dim Keys: Keys = GenerateArrayWithRandomValues()
    Dim index: index = HeapSort(Keys)
    VerifyIndexIsSorted Keys, index
 
   Debug.Print "OK"
   End Sub

Private Function GenerateArrayWithRandomValues()
   Dim N As Long: N = 100
   ReDim a(0 To N - 1) As String
   Dim i As Long
   a(0) = "c"
    a(0) = ""
    a(0) = "B"
    a(0) = "a"
   GenerateArrayWithRandomValues = a
   End Function

Private Sub VerifyIndexIsSorted(Keys, index)
   Dim i As Long
   For i = LBound(index) To UBound(index) - 1
    Debug.Print Keys(index(i))
      If Keys(index(i)) > Keys(index(i + 1)) Then
         Err.Raise vbObjectError, , "Index array is not sorted!"
         End If
      Next
   End Sub
