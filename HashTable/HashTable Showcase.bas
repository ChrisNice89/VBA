Attribute VB_Name = "Showcase"

Sub Test()
    
'    The Hashtable class represents a dictionary of associated keys and values with constant lookup time.
    Dim ht As HashTable
    Set ht = New HashTable
    
'    The capacity argument serves as an indication of
'    the number of entries the hashtable will contain. When this number (or
'    an approximation) is known, specifying it in the constructor can
'    eliminate a number of resizing operations that would otherwise be
'    performed when elements are added to the hashtable
'    As entries are added to a hashtable, the hashtable's actual capacity increases
'    Smaller load factors cause faster average lookup times at the cost of increased memory consumption
'    This Hashtable uses double hashing, two different hash functions are implemented
    Call ht.Build(Capacity:=10, LoadFactor:=0.75, HashFunction:=Function1)
    
    Dim i As Long
    For i = 1 To 20
        'Objects can also be inserted
        'existing key value pairs are overwritten
        Call ht.Add("Key" & i, "Value" & i)
        
        If ht.Contains("Key" & i) Then
            'Fast item access, triggers an error if 'Contains' failed
            Debug.Print ht.LastAccess
            'Slowly because 'Contains' method is implicitly called again
            'Debug.Print ht.Item("Key" & i)
        End If
    Next

    'Ensures that the iterator records all current entries
    'Newly added entries are not recorded after starting an iterator
    Call ht.StartIterator
    'Entries that are loaded by the iterator can only be called up once
    'The second time the entry is retrieved, vbempty is returned!
    Do While ht.EntryLoaded
        Debug.Print ht.CurrentType
        Debug.Print ht.CurrentKey
        Debug.Print ht.CurrentItem
    Loop

    Dim ht_Copy As HashTable
    Set ht_Copy = ht.Copy
    
    'Snapshot' generate a snapshot of the data internally
    'Call ht_Copy.Snapshot
    'GetKeys or GetValues uses a current snapshot
    
    'Entries do not correspond to the insert order and can differ from the order of the original table
    For i = 1 To ht_Copy.Count
        Debug.Print ht_Copy.GetKeys()(i)
    Next
    
    Dim v As Variant
    For Each v In ht_Copy.GetValues
        Debug.Print v
    Next
    
    'deletes the current snapshot
    Call ht_Copy.ResetIterator
    
    'Generates an overview
    Debug.Print ht.ToString
    
    'All entries are deleted
    Call ht.RemoveAll
    Call ht_Copy.RemoveAll

End Sub
