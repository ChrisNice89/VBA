Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private Sql As GenericSql

Sub CreateTables()

    Dim Stammdaten As Stringbuilder: Set Stammdaten = New Stringbuilder
    Dim Portfolio As Stringbuilder: Set Portfolio = New Stringbuilder
    Dim Normal As Stringbuilder: Set Normal = New Stringbuilder
    Dim Intensiv As Stringbuilder: Set Intensiv = New Stringbuilder
    Dim Sanierng As Stringbuilder: Set Sanierng = New Stringbuilder
    Dim Abwicklung As Stringbuilder: Set Abwicklung = New Stringbuilder
    Dim Paar As Stringbuilder: Set Paar = New Stringbuilder
    
'    With Portfolio
'        .Append "CREATE TABLE LDRS_PORTFOLIO ( "
'        .Append "[ID] AUTOINCREMENT,"
'        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
'        .Append "[NUMMER] TEXT,"
'        .Append "[NAME] TEXT,"
'        .Append "[PRÜFER] TEXT,"
'        .Append "[TRANCHE] TEXT,"
'        .Append "[PRÜFUNGSSCHWERPUNKT] BYTE,"
'        .Append "[AUSWAHLGRUND] TEXT,"
'        .Append "[DATUM] DATE,"
'        .Append "[KUNDENNUMMER] TEXT,"
'        .Append "[RATINGVERFAHREN] TEXT,"
'        .Append "[RATINGNOTE] TEXT,"
'        .Append "[RATINGDATUM] DATE,"
'        .Append "[RISIKOVOLUMEN] CURRENCY,"
'        .Append "[INANSPRUCHNAHME] CURRENCY,"
'        .Append "[BLANKOVOLUMEN] CURRENCY,"
'        .Append "[EWB] CURRENCY,"
'        .Append "[KONTONUMMER] TEXT,"
'        .Append "[PRODUKTGRUPPE] TEXT,"
'        .Append "[PRODUKTTYP] TEXT,"
'        .Append "[SOLLZINS] SINGLE,"
'        .Append "[LIMIT (EXTERN)] CURRENCY,"
'        .Append "[LIMIT (INTERN)] CURRENCY,"
'        .Append "[INANSPRUCHNAHME] CURRENCY,"
'        .Append "[ÜBERZIEHUNGSDAUER] BYTE,"
'        .Append "[GEBER-NUMMER] TEXT,"
'        .Append "[GEBER-NAME] TEXT,"
'        .Append "[NUMMER] BYTE,"
'        .Append "[SICHERHEITENART] TEXT,"
'        .Append "[IMMOBILIEN-NUMMER] BYTE,"
'        .Append "[OBJEKTART] TEXT,"
'        .Append "[BLW-AUSLAUF] CURRENCY,"
'        .Append "[ANRECHNUNG] CURRENCY,"
'        .Append "[VERFÜGBAR] CURRENCY"
'        .Append " )"
'    End With
'
'    Call TestSql(Portfolio.ToString)
    
    With Normal
        .Append "CREATE TABLE LDRS_NORMAL ( "
        .Append "[ID] AUTOINCREMENT,"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[KREDITENTSCHEIDUNG/KREDITPROTOKOLL_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[KREDITENTSCHEIDUNG/KREDITPROTOKOLL_DOKUMENTATION] MEMO,"
        .Append "[KDF (INKL OFFENLEGUNG)_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[KDF (INKL OFFENLEGUNG)_DOKUMENTATION] MEMO,"
        .Append "[RKV_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[RKV_DOKUMENTATION] MEMO,"
        .Append "[SICHERHEITEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[SICHERHEITEN_DOKUMENTATION] MEMO,"
        .Append "[LAUFENDE ÜBERWACHUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[LAUFENDE ÜBERWACHUNG_DOKUMENTATION] MEMO,"
        .Append "[ZUORDNUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ZUORDNUNG_DOKUMENTATION] MEMO,"
        .Append "[RISIKOVORSORGE_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[RISIKOVORSORGE_DOKUMENTATION] MEMO,"
        .Append "[VOTIERUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VOTIERUNG_DOKUMENTATION] MEMO,"
        .Append "[GENEHMIGUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[GENEHMIGUNG_DOKUMENTATION] MEMO,"
        .Append "[ADRESSENAUSFALLRISIKO_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ADRESSENAUSFALLRISIKO_DOKUMENTATION] MEMO,"
        .Append "[STRATEGIEKONFORMITÄT_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[STRATEGIEKONFORMITÄT_DOKUMENTATION] MEMO,"
        .Append "[BERICHTSWESEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[BERICHTSWESEN_DOKUMENTATION] MEMO,"
        .Append "[VERTRAGSERSTELLUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VERTRAGSERSTELLUNG_DOKUMENTATION] MEMO,"
        .Append "[AUSZAHLUNGSKONTROLLE / MITTELVERWENDUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[AUSZAHLUNGSKONTROLLE / MITTELVERWENDUNG_DOKUMENTATION] MEMO,"
        .Append "[FORBEARANCE_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[FORBEARANCE_DOKUMENTATION] MEMO,"
        .Append "[ÜBERWACHUNG DER OFFENLEGUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ÜBERWACHUNG DER OFFENLEGUNG_DOKUMENTATION] MEMO,"
        .Append "[VOLLSTÄNDIGKEIT ERFORDERLICHE UNTERLAGEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VOLLSTÄNDIGKEIT ERFORDERLICHE UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[EINREICHUNGSFRIST_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[EINREICHUNGSFRIST_DOKUMENTATION] MEMO,"
        .Append "[AUSWERTUNGSFRIST_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[AUSWERTUNGSFRIST_DOKUMENTATION] MEMO,"
        .Append "[MAHNVERFAHREN FÜR AUSSTEHENDE UNTERLAGEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[MAHNVERFAHREN FÜR AUSSTEHENDE UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[VOLLSTÄNDIGKEIT KDF PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VOLLSTÄNDIGKEIT KDF DOKUMENTATION] MEMO,"
        .Append "[NACHVOLLZIEHBARE BERECHNUNG DER KDF PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[NACHVOLLZIEHBARE BERECHNUNG DER KDF DOKUMENTATION] MEMO,"
        .Append "[NACHHALTIGE KDF GEGEBEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[NACHHALTIGE KDF GEGEBEN_DOKUMENTATION] MEMO,"
        .Append "[VERFAHRENSAUSWAHL_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VERFAHRENSAUSWAHL_DOKUMENTATION] MEMO,"
        .Append "[RISIKOFAKTOREN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[RISIKOFAKTOREN_DOKUMENTATION] MEMO,"
        .Append "[ÜBERSCHREIBUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ÜBERSCHREIBUNG_DOKUMENTATION] MEMO,"
        .Append "[TURNUSPRÜFUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[TURNUSPRÜFUNG_DOKUMENTATION] MEMO,"
        .Append "[ANLASSPRÜFUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ANLASSPRÜFUNG_DOKUMENTATION] MEMO,"
        .Append "[AUSFALLERKENNUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[AUSFALLERKENNUNG_DOKUMENTATION] MEMO,"
        .Append "[SICHERHEITENVERTRÄGE_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[SICHERHEITENVERTRÄGE_DOKUMENTATION] MEMO,"
        .Append "[RECHTLICHE DURCHSETZBARKEIT_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[RECHTLICHE DURCHSETZBARKEIT_DOKUMENTATION] MEMO,"
        .Append "[PLAUSIBILISIERUNG DER WERTERMITTLUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[PLAUSIBILISIERUNG DER WERTERMITTLUNG_DOKUMENTATION] MEMO,"
        .Append "[TURNUSPRÜFUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[TURNUSPRÜFUNG_DOKUMENTATION] MEMO,"
        .Append "[ANLASSPRÜFUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ANLASSPRÜFUNG_DOKUMENTATION] MEMO,"
        .Append "[VERWALTUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VERWALTUNG_DOKUMENTATION] MEMO,"
        .Append "[VOLLSTÄNDIGKEIT DER NOTWENDIGEN UNTERLAGEN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[VOLLSTÄNDIGKEIT DER NOTWENDIGEN UNTERLAGEN_DOKUMENTATION] MEMO,"
        .Append "[ÜBERZIEHUNGSBEARBEITUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ÜBERZIEHUNGSBEARBEITUNG_DOKUMENTATION] MEMO,"
        .Append "[AUFLAGEN / COVENANTS_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[AUFLAGEN / COVENANTS_DOKUMENTATION] MEMO,"
        .Append "[FRÜHWARNSYSTEM - SYSTEMATISCHE INDIKATOREN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[FRÜHWARNSYSTEM - SYSTEMATISCHE INDIKATOREN_DOKUMENTATION] MEMO,"
        .Append "[ANLASSBEZOGENE INDIKATOREN_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ANLASSBEZOGENE INDIKATOREN_DOKUMENTATION] MEMO,"
        .Append "[BESTANDSAUFNAHME_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[BESTANDSAUFNAHME_DOKUMENTATION] MEMO,"
        .Append "[ÜBEREINSTIMMUNG SOLL-ZUORDNUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ÜBEREINSTIMMUNG SOLL-ZUORDNUNG_DOKUMENTATION] MEMO,"
        .Append "[ERMITTLUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[ERMITTLUNG_DOKUMENTATION] MEMO,"
        .Append "[BESCHLUSSFASSUNG_PRÜFUNGSINTENSITÄT] BYTE,"
        .Append "[BESCHLUSSFASSUNG_DOKUMENTATION] MEMO,"
        .Append "[BERICHTERSTATTUNG_PRÜFUNGSINTENSITÄT] BYTE,"
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
        .Append "[RISIKOVOLUMEN (PRÜFUNGSZEITPUNKT)] CURRENCY,"
        .Append "[BLANKOVOLUMEN (PRÜFUNGSZEITPUNKT)] CURRENCY,"
        .Append "[INANSPRUCHNAHME (PRÜFUNGSZEITPUNKT)] CURRENCY,"
        .Append "[EWB (PRÜFUNGSZEITPUNKT)] CURRENCY,"
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

    Dim GA As GenericArray
    Set GA = GenericArray.Build(3, 4)
    
    Call GA.SetValue(GNumeric(1), 1, 3)
    Call GA.SetValue(GNumeric(2), 2, 3)
    Call GA.SetValue(GNumeric(3), 3, 3)
    
    Dim Column As GenericArray
    Set Column = GA.SlizeColumn(3)
    
    Dim Element As GNumeric
    With Column.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call GenericArray.Copy(Column, Column.LowerBound, GA, GA.LowerBound, Column.Length)
    Debug.Print GA.GetValue(1, 1).Equals(GA.GetValue(1, 3))
    
    Call GA.Transpose
    Debug.Print GA.GetValue(3, 1).Equals(Column(1))

End Sub
Sub testArrayConstructor()

    Dim GA As GenericArray
    Set GA = GenericArray.BuildWith(GNumeric(VBA.Now), GString("   now: " & VBA.Now & "!   ", Trim), GDate(VBA.Now, year))
    
    Dim Element As IGeneric
    With GA.Iterator
        Do While .HasNext(Element)
            Debug.Print Element
        Loop
    End With
    
End Sub
Sub TestArrayIterator2()
    
    Dim Char As IGeneric
    Dim S As GString
    Set S = GString("Ich bin ein Fuchs")
    
    With S.ToArray.Iterator
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    With S.Split(" ").Iterator
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    Debug.Print S.ElementAt(1).Contains("i")
    
End Sub
Sub TestArrayIterator()
   
    Dim i As Long, N As Long
    N = 10000
    
    Dim x() As IGeneric
    ReDim x(1 To N)
    
    For i = 1 To N
        Set x(i) = GNumeric(i)
    Next
    
    Dim Number As GNumeric
    Dim GA As GenericArray
    Set GA = GenericArray.BuildFrom(x)
     
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    For i = GA.LowerBound To GA.Length
        Set Number = GA(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    With GA.Iterator
        Do While .HasNext(Number) ' Fast
'            Set Number = .Current
        Loop
    End With
    Debug.Print t.TimeElapsed
    
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
    Call List.Sort(random)
    
    t.StartCounter
    Call List.Sort(Descending)
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
    
    Dim s1 As IGeneric
    Dim s2 As IGeneric
    
    Set s1 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    Set s2 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, N As Long
    
    t.StartCounter
    N = 10000
    For i = 1 To N
        s1.Equals s2
    Next
    Debug.Print t.TimeElapsed
     
End Sub
Sub TestArray()
    
    Dim t As CTimer
    
    Dim GA As GenericArray
    Dim ga2 As GenericArray
    Dim ga3 As GenericArray
    Dim x() As IGeneric
    Dim i As Long, N As Long
   
    N = 1000
    Set GA = GenericArray.Build(N)
    Set ga2 = GenericArray.Build(N)
    ReDim x(1 To N)
    
    For i = 1 To N
        Call GA.SetValue(GNumeric(i), i)
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
    Set Imap = GenericTree.Build
    
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
    With c.Iterator(Pairs_)
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
    With c.Iterator(t:=Keys_)
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim t As CTimer
    
    Dim Map As IGenericDictionary
    Set Map = GenericTree.Build ' GenericOrderedMap.Build 'GenericSortedList.Build() 'GenericTree.Build '
    
    Dim i As Long, N As Long, j As Long
    N = 1000
    
    Dim List As GenericList
    If List Is Nothing Then
        Set List = GenericList.Build
        For i = 1 To N
            Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Next
        Call List.Sort(random)
    End If
 
    Dim P As GenericPair
    Dim Item As IGeneric
    
    Set t = New CTimer
    t.StartCounter
    For i = 1 To N
        Set P = List(i)
        Call Map.Add(P.Key, P.Value)
    Next
    Debug.Print N & " :: "; t.TimeElapsed
  
    For i = 1 To N
        Set P = List(i)
        Set Item = Map.Item(P.Key)
    Next
    Dim Tree As GenericTree
    Set Tree = Map
    
    If Tree.Count = N = False Then
        Debug.Print "Tree.Count = n = False"
    Else
        Debug.Print Tree.ElementAt(N - 1)
    End If
    
    Set List = Nothing
    Set Item = Nothing
    Set Map = Nothing
    Set Tree = Nothing
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(Map.Count)
'    Call Map.CopyOf(Pairs_, ga, ga.LowerBound)
'
'    For i = ga.LowerBound + 1 To ga.Length
'        If ga(i - 1).CompareTo(ga(i)) = IsGreater Then
'            Debug.Print "Error"
'        End If
'    Next
''
'    t.StartCounter
'    With Map.Iterator(Pairs_)
'        Do While .HasNext(Item)
'            Debug.Print Item
'        Loop
'    End With
'    Debug.Print t.TimeElapsed
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

    With c.Iterator(Pairs_)
        Do While .HasNext(Item)
           
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestTree()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim Tree As GenericTree
    Set Tree = GenericTree.Build
    
    Dim i As Long
    
    For i = 1 To 100
        Call Tree.Add(GNumeric(i), GNumeric(i))
    Next
  
    Dim c As IGenericReadOnlyList
    Set c = Skynet.Clone(Tree)
    
    Debug.Print c.IndexOf(GNumeric(99))
    Debug.Print c.IndexOf(GNumeric(1))
    Dim Item As IGeneric
    
    t.StartCounter
    
    With c.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
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
    
    Call List.AddAll(c.Iterator(Pairs_)) 'size is unknown
    'Call List.AddAll(C)' faster because size is known
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = Skynet.Clone(List.AsReadOnly)
        
    Dim GA As GenericArray
    Set GA = GenericArray.Build(Clone.Count)
    
    Call Clone.CopyTo(GA, GA.LowerBound)
    Call Skynet.Dispose(Clone)
       
    For i = GA.LowerBound To GA.Length
        Debug.Print GA(i)
    Next

    Debug.Print GA.IndexOf(List(10))

End Sub

Sub TestArray2()
    
    Dim i As Long
    Dim GA As GenericArray
    Set GA = GenericArray.Build(100)
    
    With GA
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
        Call .SetValue(GString("ö"), 55)
        Call .SetValue(GString("q"), 56)
        Call .SetValue(GString("k"), 57)
        Call .SetValue(GString("x"), 58)
        Call .SetValue(GString("h"), 59)
        Call .SetValue(GString("t"), 60)
        Call .SetValue(GString("a"), 61)
    
        Call .SetValue(GString("z"), 70)
        Call .SetValue(GString("h"), 71)
        Call .SetValue(GString("t"), 72)
        Call .SetValue(GString("ä"), 73)
    
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
        Call .Sort(Descending, GA.LowerBound, GA.Length)
    
        For i = 1 To .Length
            If Not GA(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & GA(i)
        Next
        
        Debug.Print .BinarySearch(GString("a"), 1, .Length, Descending)
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
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List.IndexOf(GString("eingefügt an 500"))
'
    Dim List2 As GenericList
    Set List2 = Skynet.Clone(List)
    Debug.Print List2.Count
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List(500)
    Debug.Print List2(500)

    Dim List3 As GenericList
    Set List3 = List.GetRange(500, 503)
    
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
   
    For i = 1 To 50000
        Call hm.Add(GString("Key" & i), GNumeric(i))
    Next
    
    Dim Clone As IGenericDictionary
    Set Clone = Skynet.Clone(hm)
    Set hm = Nothing
    
    Dim Item As IGeneric
    t.StartCounter
    With Clone.Iterator(t:=Pairs_)
        Do While .HasNext(Item)
'            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
    
    With GenericSortedList.BuildFrom(Dictionary:=Clone).Iterator(Pairs_)
        Do While .HasNext(Item)
'            Debug.Print Item
        Loop
    End With

End Sub

'Public Sub TestString()
'
'    Dim sentence As GString
'    Set sentence = GString("the quick brown fox jumps over the lazy dog")
'
'    Debug.Print "Before: " & sentence.Value
'
'    Set WordSequence = GenericEnumerable(sentence.Split(" ").Iterator(1, 9))
'
'    Debug.Print "After: " & WordSequence.Aggregate(GString, IgnoreNull:=True)
'
'End Sub
'
'Public Sub TestInteger()
'
'    Dim ints As GenericArray
'    Set ints = GenericArray.BuildWithIntegers(4, 8, 8, 3, 9, 0, 7, 8, 2)
'
'    Set IntegerSequence = GenericEnumerable(ints.Iterator(1, 9))
'    Debug.Print "Integers: " & GString.Join(ints, ";").Value
'
'    Debug.Print "The number of even integers is: " & IntegerSequence.Aggregate(GNumeric, IgnoreNull:=True)
'
'End Sub
'
'Private Sub IntegerSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
'    If (Cast.ToNumeric(Current).IsEven) Then Set Result = Cast.ToNumeric(Result).Add(1)
'End Sub
'
'Private Sub WordSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
'    Set Result = Cast.ToString(Current).Concat(Result, " ")
'End Sub


'Sub Cmdtest()
'
'    Dim cmd As SqlCommand
'    Dim Sql As String
'    Sql = "SomeSql"
'    Set cmd = SqlCommand.Build(Sql, SqlConnection)
'    Debug.Print cmd.Sql.Replace("Some", "Somee").Value
'End Sub
'
'Sub ParameterTest()
'
'Dim cmd As SqlCommand
'Set cmd = SqlCommand.Build("SomeSql", SqlConnection)
'
'Call cmd.CreateParameter(GString("Christian"), "Name").AddValue(GString("Christoph"))
'Debug.Print cmd.Parameter("Name").CurrentValue.Value
'
'Debug.Print cmd.Parameter("Name").UseValue(2).Object.ToString
'
'Dim Christian As IGeneric
'Dim Christoph As IGeneric
'
'Set Christian = cmd.Parameter("Name").Value(1)
'Set Christoph = cmd.Parameter("Name").Value(2)
'
'Debug.Print Christoph.CompareTo(Christian) = IsGreater
'
'Dim p1 As IGeneric
'Dim p2 As IGeneric
'
'Set p1 = cmd.Parameter(1)
'Set p2 = p1
'
'Debug.Print p1.Equals(p2)
'
'Dim p3 As SqlParameter
'Set p3 = SqlParameter(GString("Christian"), "Name")
'Debug.Print Christian.Equals(p3.CurrentValue)
'
'
'Dim p4 As IGeneric
''Set p4 = cmd.CreateParameter(TDate(#4/4/2020#), "Datum").AddValue(TDate(#1/1/2021#))
'
'Debug.Print p4.Equals(p1)
'
'
'End Sub
'
'


