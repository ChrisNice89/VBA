Attribute VB_Name = "Modul1"

Sub Test()
    
    Dim ht As HashTable
    Set ht = New HashTable
    Call ht.Build(Capacity:=10, HashFunction:=Function1)
    
    Dim i As Long
    For i = 1 To 15
        Call ht.Add("Key" & i, "Value" & i)
        If Not ht.Contains("Key" & i) Then
            Debug.Print "Error Key" & i
        End If
    Next

    Dim ht2 As HashTable
    Set ht2 = ht.Copy
    
    For i = 1 To ht2.Count
        Debug.Print ht2.GetKeys()(i)
    Next
    
    Dim v As Variant
    For Each v In ht2.GetValues
        Debug.Print v
    Next
    
    Call ht.StartIterator
    Do While ht.EntryLoaded
        Debug.Print ht.CurrentType
        Debug.Print ht.CurrentKey
        Debug.Print ht.CurrentItem
    Loop

    Debug.Print ht.ToString
    
    Call ht.RemoveAll
    Call ht2.RemoveAll

End Sub
