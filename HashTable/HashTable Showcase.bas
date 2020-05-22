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
    Call ht.Build(Capacity:=7, LoadFactor:=0.7, HashFunction:=Function1)
    
    Dim i As Long
    For i = 1 To 10
        'Objects can also be inserted
        'existing key value pairs are overwritten
        Call ht.Add("Key" & i, "Value" & i)
        
        If ht.Contains("Key" & i) Then
            'Fast item access, triggers an error if 'Contains' failed
            'Debug.Print ht.LastAccess
            'Slowly because 'Contains' method is implicitly called again
            'Debug.Print ht.Item("Key" & i)
        End If
    Next

    'Ensures that the Cache records all current entries
    'Newly added entries are not recorded after starting caching

    Dim k As String
    Dim v As Variant
    
    Call ht.CachePrepare
    Do While ht.Cached(k, v)
        Debug.Print k
        Debug.Print v
    Loop
    
    'deletes the current snapshot / is implicitly called when cached fails
    'Call htClone.CacheClear
    
    Dim htClone As HashTable
    Set htClone = New HashTable
    Call htClone.CloneBy(ht)
    
    'Entries do not correspond to the insert order and can differ from the order of the original table
    For Each v In htClone.GetValues
        Debug.Print v
    Next
    
    For Each v In htClone.GetKeys
        Debug.Print v
    Next
    
    'Generates an overview
    Debug.Print ht.ToString
    Debug.Print htClone.ToString
    
    'All entries are deleted
    Call ht.RemoveAll
    Call htClone.RemoveAll

End Sub










