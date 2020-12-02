Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private Sql As GenericSql
Private RandomList As GenericList
Sub TestRange()
    

    Dim i As Long
    Dim t As CTimer
    Set t = New CTimer
    
    Dim N As Long
    N = 1000
    ReDim g(N) As IGeneric
    ReDim iunk(N) As Object
    
    Dim oNumber As Object
    Dim INumber As IGeneric
    Dim Number As GNumeric
    Set Number = GNumeric
    Set INumber = Number
    Set oNumber = Number
    
    t.StartCounter
    For i = LBound(iunk) To UBound(iunk)
        Set iunk(i) = Number
        Set INumber = iunk(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    For i = LBound(g) To UBound(g)
        Set g(i) = Number
        Set INumber = g(i)
    Next
    Debug.Print t.TimeElapsed

End Sub

Sub TestCreaet()
    
    Dim Result As Boolean
    Dim t As CTimer
    Set t = New CTimer
    Dim i As Long
    Dim g As IGeneric
    Set g = GenericPair
    t.StartCounter

    For i = 1 To 10000
        Set g = GenericPair(IGenericValue, IGenericValue)
        Result = g.IsRelatedTo(g)
    Next

    Debug.Print t.TimeElapsed
End Sub
Sub TestListEquals()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim SortedList As GenericSortedList
    Set SortedList = GenericSortedList.Build
    
    Dim tree As GenericSortedSet
    Set tree = GenericSortedSet.Build
    
    Dim S As GString
    
    Dim i As Long
    t.StartCounter
    For i = 1 To 10
        Set S = GString("Key: " & i)
        Call tree.Add(S)
        Call SortedList.Add(S, S)
'        Call GString.HashValueOf("asfethzrnasx fvfcsc" & i)
    Next
    Debug.Print t.TimeElapsed
'    Call List.Sort(ascending)
    
    Debug.Print GenericList.IsEqual(SortedList, tree)
    Debug.Print tree.IndexOf(GString("Key: " & i - 1))
    Debug.Print SortedList.IndexOfKey(GString("Key: " & i - 1))
    
    
    Debug.Print GString.Join(tree, ";").Value
End Sub

Sub CreateTables()

    Dim Stammdaten As Stringbuilder: Set Stammdaten = New Stringbuilder
    Dim Portfolio As Stringbuilder: Set Portfolio = New Stringbuilder
    Dim Normal As Stringbuilder: Set Normal = New Stringbuilder
    Dim Intensiv As Stringbuilder: Set Intensiv = New Stringbuilder
    Dim Sanierng As Stringbuilder: Set Sanierng = New Stringbuilder
    Dim Abwicklung As Stringbuilder: Set Abwicklung = New Stringbuilder
    Dim Paar As Stringbuilder: Set Paar = New Stringbuilder
    
    With Portfolio
        .Append "CREATE TABLE LDRS_PORTFOLIO ( "
        .Append "[ID] AUTOINCREMENT,"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[NUMMER] TEXT,"
        .Append "[NAME] TEXT,"
        .Append "[PR�FER] TEXT,"
        .Append "[TRANCHE] TEXT,"
        .Append "[PR�FUNGSSCHWERPUNKT] BYTE,"
        .Append "[AUSWAHLGRUND] TEXT,"
        .Append "[DATUM] DATE,"
        .Append "[KUNDENNUMMER] TEXT,"
        .Append "[RATINGVERFAHREN] TEXT,"
        .Append "[RATINGNOTE] TEXT,"
        .Append "[RATINGDATUM] DATE,"
        .Append "[RISIKOVOLUMEN] CURRENCY,"
        .Append "[INANSPRUCHNAHME] CURRENCY,"
        .Append "[BLANKOVOLUMEN] CURRENCY,"
        .Append "[EWB] CURRENCY,"
        .Append "[KONTONUMMER] TEXT,"
        .Append "[PRODUKTGRUPPE] TEXT,"
        .Append "[PRODUKTTYP] TEXT,"
        .Append "[SOLLZINS] SINGLE,"
        .Append "[LIMIT (EXTERN)] CURRENCY,"
        .Append "[LIMIT (INTERN)] CURRENCY,"
        .Append "[INANSPRUCHNAHME] CURRENCY,"
        .Append "[�BERZIEHUNGSDAUER] BYTE,"
        .Append "[GEBER-NUMMER] TEXT,"
        .Append "[GEBER-NAME] TEXT,"
        .Append "[NUMMER] BYTE,"
        .Append "[SICHERHEITENART] TEXT,"
        .Append "[IMMOBILIEN-NUMMER] BYTE,"
        .Append "[OBJEKTART] TEXT,"
        .Append "[BLW-AUSLAUF] CURRENCY,"
        .Append "[ANRECHNUNG] CURRENCY,"
        .Append "[VERF�GBAR] CURRENCY"
        .Append " )"
    End With

    Call TestSql(Portfolio.ToString)
    
    With Normal
        .Append "CREATE TABLE LDRS_NORMAL ( "
        .Append "[ID] AUTOINCREMENT,"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[KREDITENTSCHEIDUNG/KREDITPROTOKOLL_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[KREDITENTSCHEIDUNG/KREDITPROTOKOLL_DOKUMENTATION] MEMO,"
        .Append "[KDF (INKL OFFENLEGUNG)_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[KDF (INKL OFFENLEGUNG)_DOKUMENTATION] MEMO,"
        .Append "[RKV_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[RKV_DOKUMENTATION] MEMO,"
        .Append "[SICHERHEITEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[SICHERHEITEN_DOKUMENTATION] MEMO,"
        .Append "[LAUFENDE �BERWACHUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[LAUFENDE �BERWACHUNG_DOKUMENTATION] MEMO,"
        .Append "[ZUORDNUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ZUORDNUNG_DOKUMENTATION] MEMO,"
        .Append "[RISIKOVORSORGE_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[RISIKOVORSORGE_DOKUMENTATION] MEMO,"
        .Append "[VOTIERUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VOTIERUNG_DOKUMENTATION] MEMO,"
        .Append "[GENEHMIGUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[GENEHMIGUNG_DOKUMENTATION] MEMO,"
        .Append "[ADRESSENAUSFALLRISIKO_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ADRESSENAUSFALLRISIKO_DOKUMENTATION] MEMO,"
        .Append "[STRATEGIEKONFORMIT�T_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[STRATEGIEKONFORMIT�T_DOKUMENTATION] MEMO,"
        .Append "[BERICHTSWESEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[BERICHTSWESEN_DOKUMENTATION] MEMO,"
        .Append "[VERTRAGSERSTELLUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VERTRAGSERSTELLUNG_DOKUMENTATION] MEMO,"
        .Append "[AUSZAHLUNGSKONTROLLE / MITTELVERWENDUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[AUSZAHLUNGSKONTROLLE / MITTELVERWENDUNG_DOKUMENTATION] MEMO,"
        .Append "[FORBEARANCE_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[FORBEARANCE_DOKUMENTATION] MEMO,"
        .Append "[�BERWACHUNG DER OFFENLEGUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[�BERWACHUNG DER OFFENLEGUNG_DOKUMENTATION] MEMO,"
        .Append "[VOLLST�NDIGKEIT ERFORDERLICHE UNTERLAGEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VOLLST�NDIGKEIT ERFORDERLICHE UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[EINREICHUNGSFRIST_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[EINREICHUNGSFRIST_DOKUMENTATION] MEMO,"
        .Append "[AUSWERTUNGSFRIST_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[AUSWERTUNGSFRIST_DOKUMENTATION] MEMO,"
        .Append "[MAHNVERFAHREN F�R AUSSTEHENDE UNTERLAGEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[MAHNVERFAHREN F�R AUSSTEHENDE UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[VOLLST�NDIGKEIT KDF PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VOLLST�NDIGKEIT KDF DOKUMENTATION] MEMO,"
        .Append "[NACHVOLLZIEHBARE BERECHNUNG DER KDF PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[NACHVOLLZIEHBARE BERECHNUNG DER KDF DOKUMENTATION] MEMO,"
        .Append "[NACHHALTIGE KDF GEGEBEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[NACHHALTIGE KDF GEGEBEN_DOKUMENTATION] MEMO,"
        .Append "[VERFAHRENSAUSWAHL_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VERFAHRENSAUSWAHL_DOKUMENTATION] MEMO,"
        .Append "[RISIKOFAKTOREN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[RISIKOFAKTOREN_DOKUMENTATION] MEMO,"
        .Append "[�BERSCHREIBUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[�BERSCHREIBUNG_DOKUMENTATION] MEMO,"
        .Append "[TURNUSPR�FUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[TURNUSPR�FUNG_DOKUMENTATION] MEMO,"
        .Append "[ANLASSPR�FUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ANLASSPR�FUNG_DOKUMENTATION] MEMO,"
        .Append "[AUSFALLERKENNUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[AUSFALLERKENNUNG_DOKUMENTATION] MEMO,"
        .Append "[SICHERHEITENVERTR�GE_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[SICHERHEITENVERTR�GE_DOKUMENTATION] MEMO,"
        .Append "[RECHTLICHE DURCHSETZBARKEIT_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[RECHTLICHE DURCHSETZBARKEIT_DOKUMENTATION] MEMO,"
        .Append "[PLAUSIBILISIERUNG DER WERTERMITTLUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[PLAUSIBILISIERUNG DER WERTERMITTLUNG_DOKUMENTATION] MEMO,"
        .Append "[TURNUSPR�FUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[TURNUSPR�FUNG_DOKUMENTATION] MEMO,"
        .Append "[ANLASSPR�FUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ANLASSPR�FUNG_DOKUMENTATION] MEMO,"
        .Append "[VERWALTUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VERWALTUNG_DOKUMENTATION] MEMO,"
        .Append "[VOLLST�NDIGKEIT DER NOTWENDIGEN UNTERLAGEN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[VOLLST�NDIGKEIT DER NOTWENDIGEN UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[�BERZIEHUNGSBEARBEITUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[�BERZIEHUNGSBEARBEITUNG_DOKUMENTATION] MEMO,"
        .Append "[AUFLAGEN / COVENANTS_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[AUFLAGEN / COVENANTS_DOKUMENTATION] MEMO,"
        .Append "[FR�HWARNSYSTEM - SYSTEMATISCHE INDIKATOREN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[FR�HWARNSYSTEM - SYSTEMATISCHE INDIKATOREN_DOKUMENTATION] MEMO,"
        .Append "[ANLASSBEZOGENE INDIKATOREN_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ANLASSBEZOGENE INDIKATOREN_DOKUMENTATION] MEMO,"
        .Append "[BESTANDSAUFNAHME_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[BESTANDSAUFNAHME_DOKUMENTATION] MEMO,"
        .Append "[�BEREINSTIMMUNG SOLL-ZUORDNUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[�BEREINSTIMMUNG SOLL-ZUORDNUNG_DOKUMENTATION] MEMO,"
        .Append "[ERMITTLUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[ERMITTLUNG_DOKUMENTATION] MEMO,"
        .Append "[BESCHLUSSFASSUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[BESCHLUSSFASSUNG_DOKUMENTATION] MEMO,"
        .Append "[BERICHTERSTATTUNG_PR�FUNGSINTENSIT�T] BYTE,"
        .Append "[BERICHTERSTATTUNG_DOKUMENTATION] MEMO"
        .Append " )"
    End With
        
    Call TestSql(Normal.ToString)

    With Stammdaten
        .Append "CREATE TABLE LDRS_STAMMDATEN ( "
        .Append "[ID] AUTOINCREMENT,"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[KUNDENSTATUS] BYTE,"
        .Append "[BETREUUNGSSTATUS] BYTE,"
        .Append "[RISIKOVOLUMEN (PORTFOLIOABZUG)] CURRENCY,"
        .Append "[BLANKOVOLUMEN (PORTFOLIOABZUG)] CURRENCY,"
        .Append "[INANSPRUCHNAHME (PORTFOLIOABZUG)] CURRENCY,"
        .Append "[EWB (PORTFOLIOABZUG)] CURRENCY,"
        .Append "[RISIKOVOLUMEN (PR�FUNGSZEITPUNKT)] CURRENCY,"
        .Append "[BLANKOVOLUMEN (PR�FUNGSZEITPUNKT)] CURRENCY,"
        .Append "[INANSPRUCHNAHME (PR�FUNGSZEITPUNKT)] CURRENCY,"
        .Append "[EWB (PR�FUNGSZEITPUNKT)] CURRENCY,"
        .Append "[RISIKOVOLUMEN (NACH PAAR)] CURRENCY,"
        .Append "[BLANKOVOLUMEN (NACH PAAR)] CURRENCY,"
        .Append "[INANSPRUCHNAHME (NACH PAAR)] CURRENCY,"
        .Append "[EWB (NACH PAAR)] CURRENCY,"
        .Append "[WERTHALTIG ANGESETZTE SICHERHEITEN] BYTE,"
        .Append "[KURZBESCHREIBUNG DES ENGAGEMENTS] MEMO,"
        .Append "[GESAMTERGEBNIS] MEMO,"
        .Append "[NOTIZEN / BEMERKUNGEN] MEMO"
        .Append " )"
    End With
        
'    Call TestSql(Stammdaten.ToString)
    
End Sub

Public Sub TestSql(ByVal CMDText As String)
    
    Dim Path As String
    Path = "C:\Users\cnitz\Desktop\iCAT Neu\Backend\Vers. 2.5\2020-02-24 iCAT-Backend Vers. 2.5.accdb"
    Dim PW As String
    PW = "OpenSesame"
    
    Set Sql = GenericSql.Build(SqlCredentials.AccessConnection(Path, PW), CMDText, 0)
    Call Sql.Execute

End Sub

Sub TestMultiDimArray()

    Dim ga As GenericArray
    Set ga = GenericArray.Build(3, 4)
    
    Call ga.SetValue(GNumeric(1), 1, 3)
    Call ga.SetValue(GNumeric(2), 2, 3)
    Call ga.SetValue(GNumeric(3), 3, 3)
    
    Dim Column As GenericArray
    Set Column = ga.SlizeColumn(3)
    
    Dim Element As GNumeric
    With Column.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call GenericArray.Copy(Column, Column.LowerBound, ga, ga.LowerBound, Column.Length)
    Debug.Print ga.GetValue(1, 1).Equals(ga.GetValue(1, 3))
    
    Call ga.Transpose
    Debug.Print ga.GetValue(3, 1).Equals(Column(1))

End Sub
Sub testArrayConstructor()

    Dim List As GenericList
    Set List = GenericList.BuildWith(GNumeric(VBA.Now), GString("   now: " & VBA.Now & "!   ", Trim), GDate(VBA.Now, year))
    
    Dim Element As IGeneric
    With List.Iterator
        Do While .HasNext(Element)
            Debug.Print Element
        Loop
    End With
    
End Sub
Sub TestArrayIterator2()
    
    Dim char As IGeneric
    Dim S As GString
    Set S = GString("Ich bin ein Fuchs")
    
    With S.ToArray.Iterator
        Do While .HasNext(char)
            Debug.Print char
        Loop
    End With
    
    With S.Split(" ").Iterator
        Do While .HasNext(char)
            Debug.Print char
        Loop
    End With
    
    Debug.Print S.ElementAt(1).Contains("i")
    
End Sub
Sub TestArrayGetter()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim List As GenericArray
    Dim Element As IGeneric
   
    Dim i As Long, N As Long
    N = 1000
        
    Set List = GenericArray.Build(N)
    ReDim x(1 To N) As IGeneric
    
    t.StartCounter
    For i = 1 To N
        Set List(i) = GNumeric(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    For i = 1 To N
        Set x(i) = GNumeric(i)
    Next
    Debug.Print t.TimeElapsed
    
    
'
'    With List.Iterator
'        t.StartCounter
'        Do While .HasNext(Element)
'
'        Loop
'        Debug.Print t.TimeElapsed
'   End With

End Sub

Sub TestArraySort()
    Dim t As CTimer
    Set t = New CTimer
    
    Dim i As Long, N As Long
    N = 40000
    
    Dim List As GenericList
    Set List = GenericList.Build(N)
   
    For i = 1 To N
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(GNumeric(i))
    Next
    Call List.Sort(Random)
    
    t.StartCounter
    Call List.Sort(descending)
    Debug.Print t.TimeElapsed
    
    Dim Item As IGeneric
    With List.Iterator
        Do While .HasNext(Item)
          
        Loop
    End With
    
End Sub

Sub TestEquals()

    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim S1 As IGeneric
    Dim S2 As IGeneric
    
    Set S1 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    Set S2 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, N As Long
    
    t.StartCounter
    N = 10000
    For i = 1 To N
        S1.Equals S2
    Next
    Debug.Print t.TimeElapsed
     
End Sub
Sub TestArray()
    
    Dim t As CTimer
    
    Dim ga As GenericArray
    Dim ga2 As GenericArray
    Dim ga3 As GenericArray
    Dim x() As IGeneric
    Dim i As Long, N As Long
   
    N = 1000
    Set ga = GenericArray.Build(N)
    Set ga2 = GenericArray.Build(N)
    ReDim x(1 To N)
    
    For i = 1 To N
        Call ga.SetValue(GNumeric(i), i)
        Call ga2.SetValue(GString("Value: " & i), i)
        Set x(i) = GString("Value: " & i)
    Next
    
    Set ga3 = GenericArray.BuildFrom(x)
    
    Dim c As IGeneric
    Set c = Skynet.Clone(ga3)
    
    Set t = New CTimer
    t.StartCounter
    Debug.Print c.Equals(ga3)
    Debug.Print t.TimeElapsed
    
End Sub

Sub TestOrderedMap()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim Map As GenericOrderedMap
    Set Map = GenericOrderedMap.Build
    
    Dim Imap As IGenericDictionary
    Set Imap = GenericSortedSet.Build
    
    Dim i As Long, N As Long
    
    N = 10000
    t.StartCounter
    For i = 1 To N
        Call Imap.Add(GNumeric(i), GNumeric(i))
    Next
    
    t.StartCounter
    Call Map.AddAll(Imap)
    Debug.Print t.TimeElapsed
    
    Dim c As GenericOrderedMap
    Set c = Skynet.Clone(Map)
    
    Dim Item As GenericPair
    t.StartCounter
    With c.IteratorOf(PairData)
        Do While .HasNext(Item)
'            Debug.Print Item.Key
        Loop
    End With
    Debug.Print t.TimeElapsed
    
End Sub

Sub TestListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim L As GenericList
    Set L = GenericList.Build

    Dim i As Long, N As Long
    
    N = 5000
    For i = 1 To N
        Call L.Add(GenericPair(GNumeric(i), GNumeric(i)))
    Next
    
    Dim c As GenericList
    Set c = Skynet.Clone(L)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Iterator
        Do While .HasNext(Item)
'           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestSortedListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build

    Dim i As Long, N As Long
    
    N = 50
    For i = 1 To N
        Call sl.Add(GNumeric(i), GNumeric(i))
    Next
    
    Dim c As GenericSortedList
    Set c = Skynet.Clone(sl)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.IteratorOf(t:=KeyData)
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim t As CTimer
    
    Dim Map As IGenericDictionary
    Set Map = GenericOrderedMap.Build 'GenericSortedList.Build()
    
    Dim i As Long, N As Long, j As Long
    N = 30
    
    If RandomList Is Nothing Then
        Set RandomList = GenericList.Build
        For i = 1 To N
            Call RandomList.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Next
        Call RandomList.Sort(Random)
    End If
 
    Dim P As GenericPair
    Dim Item As IGeneric
    
    Set t = New CTimer
    t.StartCounter
    For i = 1 To N
        Set P = RandomList(i)
        Call Map.Add(P.Key, P.Value)
    Next
    Debug.Print N & " :: "; t.TimeElapsed

'
'    Dim tree As GenericSortedSet
'    Set tree = Map
'    Dim Node As GenericNode
'    Set Node = tree.ElementAt(N)
'
'    Do While Node Is Nothing = False
'        Debug.Print Node.Key
'        Set Node = Node.InOrderPrevious
'    Loop
'    Exit Sub
'
    For i = 1 To N
        Set P = RandomList(i)
        Set Item = Map.Item(P.Key)
    Next
  
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Map.Count)
    Call Map.CopyOf(PairData, ga, ga.LowerBound)
    
    Dim GenericPairComparer As IGenericComparer
    Set GenericPairComparer = GenericPair
    
    For i = ga.LowerBound + 1 To ga.Length
        Debug.Print ga.ElementAt(i)
    Next
'
    t.StartCounter
    With Map.IteratorOf(PairData)
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
'
End Sub

Sub TestSortedList()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build()
    
    Dim i As Long, N As Long
    
    N = 100
    For i = N To 1 Step -1
        Call sl.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print t.TimeElapsed
    
    Dim c As GenericSortedList
    Set c = Skynet.Clone(sl)
    
    Dim Item As IGeneric

    With c.IteratorOf(PairData)
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestTree()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim tree As GenericSortedSet
    Set tree = GenericSortedSet.Build
    
    Dim i As Long
    
    For i = 100 To 1 Step -1
        Call tree.Add(GNumeric(i))
    Next
    
    Dim N As IGeneric
    Set N = tree.ElementAt(1)
    
    Debug.Print N.ToString
    Dim c As IGenericReadOnlyList
    Set c = Skynet.Clone(tree)
    
    Debug.Print c.IndexOf(GNumeric(99))
    Debug.Print c.IndexOf(GNumeric(1))
    Dim Item As IGeneric
    
    t.StartCounter
    
    With c.Iterator
        Do While .HasNext(Item)
'            Debug.Print Item.ToString
        Loop
    End With
    
    Debug.Print t.TimeElapsed
'
'    Dim collection As IGenericCollection
'    Set collection = C
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(collection.Count)
'    Call collection.CopyTo(ga, ga.LowerBound)

End Sub

Sub TestGenericCollection()
    
    Dim c As GenericOrderedMap
    Set c = GenericOrderedMap.Build
    
    Dim i As Long
    For i = 1 To 10
        Call c.Add(GString("Key: " & i), GString("Value: " & i))
    Next

    Dim List As GenericList
    Set List = GenericList.Build
    
    Call List.AddAll(c.IteratorOf(PairData)) 'size is unknown
    'Call List.AddAll(c)' faster because size is known
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = Skynet.Clone(List.AsReadOnly)
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Count)
    
    Call Clone.CopyTo(ga, ga.LowerBound)
    Call Skynet.Dispose(Clone)
       
    For i = ga.LowerBound To ga.Length
        Debug.Print ga(i)
    Next

    Debug.Print ga.IndexOf(List(10))

End Sub

Sub TestArray2()
    
    Dim i As Long
    Dim ga As GenericArray
    Set ga = GenericArray.Build(100)
    Dim t As CTimer
    Set t = New CTimer
    With ga
        Call .SetValue(GString("b"), 13)
        Call .SetValue(GString("c"), 14)
        Call .SetValue(GString("a"), 15)
        Call .SetValue(GString("h"), 16)
        Call .SetValue(GString("s"), 17)
        Call .SetValue(GString("d"), 18)
        Call .SetValue(GString("zz"), 19)
        Call .SetValue(GString("c"), 20)
        Call .SetValue(GString("x"), 21)
        Call .SetValue(GString("e"), 22)
        Call .SetValue(GString("t"), 23)
        Call .SetValue(GString("a"), 24)
    
        Call .SetValue(GString("a"), 50)
        Call .SetValue(GString("c"), 51)
        Call .SetValue(GString("a"), 52)
        Call .SetValue(GString("j"), 53)
        Call .SetValue(GString("s"), 54)
        Call .SetValue(GString("�"), 55)
        Call .SetValue(GString("q"), 56)
        Call .SetValue(GString("k"), 57)
        Call .SetValue(GString("x"), 58)
        Call .SetValue(GString("h"), 59)
        Call .SetValue(GString("t"), 60)
        Call .SetValue(GString("a"), 61)
    
        Call .SetValue(GString("z"), 70)
        Call .SetValue(GString("h"), 71)
        Call .SetValue(GString("t"), 72)
        Call .SetValue(GString("�"), 73)
    
        Call .SetValue(GString("c"), 80)
        Call .SetValue(GString(""), 81)
        Call .SetValue(GString("e"), 82)
        Call .SetValue(GString("f"), 83)
        Call .SetValue(GString("d"), 84)
        Call .SetValue(GString("zz"), 85)
        Call .SetValue(GString("c"), 86)
        Call .SetValue(GString("x"), 87)
        Call .SetValue(GString("e"), 88)
        Call .SetValue(GString("f"), 89)
        Call .SetValue(GString("a"), 90)
        Call .SetValue(GString("a"), 100)
        t.StartCounter
        Call .Sort(descending)
        Debug.Print t.TimeElapsed
        
        For i = 1 To .Length
            If Not ga(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga(i)
        Next

        Debug.Print .BinarySearch(GString("zzz"), 1, .Length, descending, IGenericValue.Comparer)
        Call .Reverse
        Call .Clear
    End With
    
End Sub

Public Sub ListTest()

    Dim i As Long
    Dim List As GenericList
    Set List = GenericList.Build()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    For i = 1 To 1000
        Call List.Add(GString("test" & i))
    Next
    Debug.Print t.TimeElapsed
'
    Debug.Print List.IndexOf(GString("test" & 999), 1, 999)
    Call List.Insert(500, GString("eingef�gt an 500"))
    Debug.Print List.IndexOf(GString("eingef�gt an 500"))
'
    Dim List2 As GenericList
    Set List2 = Skynet.Clone(List)
    Debug.Print List2.Count
    Call List.Insert(500, GString("eingef�gt an 500"))
    Debug.Print List(500)
    Debug.Print List2(500)

    Dim List3 As GenericList
    Set List3 = List.GetRange(500, 502)
    
    Dim readOnly As IGenericReadOnlyList
    Set readOnly = List3.AsReadOnly
    Debug.Print readOnly(1)
    Debug.Print readOnly(10)
    Debug.Print readOnly.Count
    Set List = Nothing

End Sub

Sub testMap()

    Dim i As Long
 
    Dim hm As IGenericDictionary
    Set hm = GenericMap.Build()
    Dim t As CTimer
    Set t = New CTimer
   
    For i = 1 To 35
        Call hm.Add(GNumeric(i), GNumeric(i))
    Next
    
    Dim Clone As IGenericDictionary
    Set Clone = Skynet.Clone(hm)
    Set hm = Nothing
    
    Dim Item As IGeneric
    t.StartCounter
    With Clone.IteratorOf(PairData)
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
    
    With GenericSortedList.BuildFrom(Clone, IGenericValue.Comparer).IteratorOf(PairData)
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With

End Sub

