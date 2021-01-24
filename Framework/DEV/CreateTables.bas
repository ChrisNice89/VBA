Attribute VB_Name = "CreateTables"
Option Compare Database

Public Function Test() As String

     With New Stringbuilder
        .Append "CREATE TABLE TEST ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] INT ," 'PRIMARY KEY
        .Append "[DATUM] DATETIME,"
        .Append "[KATEGORIE] CHAR(20),"
        .Append "[BETRAG] NUMERIC(18,4),"
        .Append "[BESCHREIBUNG] VARCHAR(MAX),"
        .Append "[STATUS] BIT "
        .Append " )"
        
        Test = .ToString
    
    End With
    
End Function

Public Function �berblick() As String

     With New Stringbuilder
        .Append "CREATE TABLE UEBERBLICK ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[STATUS] CHAR(20),"
        .Append "[BETREUUNG] CHAR(20),"
        .Append "[EKPA_RV] NUMERIC(18,4),"
        .Append "[EKPA_BV] NUMERIC(18,4),"
        .Append "[EKPA_IA] NUMERIC(18,4),"
        .Append "[EKPA_EWB] NUMERIC(18,4),"
        .Append "[STICH_RV] NUMERIC(18,4),"
        .Append "[STICH_BV] NUMERIC(18,4),"
        .Append "[STICH_IA] NUMERIC(18,4),"
        .Append "[STICH_EWB] NUMERIC(18,4),"
        .Append "[PAAR_RV] NUMERIC(18,4),"
        .Append "[PAAR_BV] NUMERIC(18,4),"
        .Append "[PAAR_IA] NUMERIC(18,4),"
        .Append "[PAAR_EWB] NUMERIC(18,4),"
        .Append "[WERTHALTIG] CHAR(20),"
        .Append "[BESCHREIBUNG] VARCHAR(MAX),"
        .Append "[ERGEBNIS] VARCHAR(MAX),"
        .Append "[NOTIZEN] VARCHAR(MAX)"

        .Append " )"
        
        �berblick = .ToString
    
    End With
    
End Function

Public Function Normal() As String

     With New Stringbuilder
        .Append "CREATE TABLE NORMAL ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[E1_BP_ENTSCHEIDUNG_INT] BIT ,"
        .Append "[E1_BP_ENTSCHEIDUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_KDF_INT] BIT ,"
        .Append "[E1_BP_KDF_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_RKV_INT] BIT ,"
        .Append "[E1_BP_RKV_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_SICHERHEITEN_INT] BIT ,"
        .Append "[E1_BP_SICHERHEITEN_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_UEBERWACHUNG_INT] BIT ,"
        .Append "[E1_BP_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_ZUORDNUNG_INT] BIT ,"
        .Append "[E1_BP_ZUORDNUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_VORSORGE_INT] BIT ,"
        .Append "[E1_BP_VORSORGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VOTIERUNG_INT] BIT ,"
        .Append "[E2_TP1_VOTIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_GENEHMIGUNG_INT] BIT ,"
        .Append "[E2_TP1_GENEHMIGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AAR_INT] BIT ,"
        .Append "[E2_TP1_AAR_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_STRATEGIE_INT] BIT ,"
        .Append "[E2_TP1_STRATEGIE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_BERICHT_INT] BIT ,"
        .Append "[E2_TP1_BERICHT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VERTRAG_INT] BIT ,"
        .Append "[E2_TP1_VERTRAG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AUSZAHLUNG_INT] BIT ,"
        .Append "[E2_TP1_AUSZAHLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_FORBEARANCE_INT] BIT ,"
        .Append "[E2_TP1_FORBEARANCE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_OFFENLEGUNG_INT] BIT ,"
        .Append "[E2_TP2_OFFENLEGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_UNTERLAGEN_INT] BIT ,"
        .Append "[E2_TP2_UNTERLAGEN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_EINREICHUNG_INT] BIT ,"
        .Append "[E2_TP2_EINREICHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_AUSWERTUNG_INT] BIT ,"
        .Append "[E2_TP2_AUSWERTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_MAHNVERFAHREN_INT] BIT ,"
        .Append "[E2_TP2_MAHNVERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BEURTEILUNG_INT] BIT ,"
        .Append "[E2_TP2_BEURTEILUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BERECHNUNG_INT] BIT ,"
        .Append "[E2_TP2_BERECHNUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_NACHHALTIG_INT] BIT ,"
        .Append "[E2_TP2_NACHHALTIG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_VERFAHREN_INT] BIT ,"
        .Append "[E2_TP3_VERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_FAKTOREN_INT] BIT ,"
        .Append "[E2_TP3_FAKTOREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_UEBERSCHREIBUNG_INT] BIT ,"
        .Append "[E2_TP3_UEBERSCHREIBUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_TURNUS_INT] BIT ,"
        .Append "[E2_TP3_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_ANLASS_INT] BIT ,"
        .Append "[E2_TP3_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_AUSFALL_INT] BIT ,"
        .Append "[E2_TP3_AUSFALL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERTRAEGE_INT] BIT ,"
        .Append "[E2_TP4_VERTRAEGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_DURCHSETZBARKEIT_INT] BIT ,"
        .Append "[E2_TP4_DURCHSETZBARKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_INT] BIT ,"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_TURNUS_INT] BIT ,"
        .Append "[E2_TP4_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_ANLASS_INT] BIT ,"
        .Append "[E2_TP4_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERWALTUNG_INT] BIT ,"
        .Append "[E2_TP4_VERWALTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_INT] BIT ,"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERZIEHUNG_INT] BIT ,"
        .Append "[E2_TP5_UEBERZIEHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_COVENANTS_INT] BIT ,"
        .Append "[E2_TP5_COVENANTS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SYSTEMATISCH_INT] BIT ,"
        .Append "[E2_TP6_SYSTEMATISCH_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_ANLASS_INT] BIT ,"
        .Append "[E2_TP6_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_BESTANDSAUFNAHME_INT] BIT ,"
        .Append "[E2_TP6_BESTANDSAUFNAHME_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SOLL_INT] BIT ,"
        .Append "[E2_TP6_SOLL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_ERMITTLUNG_INT] BIT ,"
        .Append "[E2_TP7_ERMITTLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BESCHLUSS_INT] BIT ,"
        .Append "[E2_TP7_BESCHLUSS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BERICHT_INT] BIT ,"
        .Append "[E2_TP7_BERICHT_DOK] VARCHAR(MAX)"

        .Append " )"
        
        Normal = .ToString
    
    End With

End Function

Public Function Intensiv() As String

     With New Stringbuilder
        .Append "CREATE TABLE INTENSIV ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[E1_BP_ENTSCHEIDUNG_INT] BIT,"
        .Append "[E1_BP_ENTSCHEIDUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_KDF_INT] BIT,"
        .Append "[E1_BP_KDF_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_RKV_INT] BIT,"
        .Append "[E1_BP_RKV_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_SICHERHEITEN_INT] BIT,"
        .Append "[E1_BP_SICHERHEITEN_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_UEBERWACHUNG_INT] BIT,"
        .Append "[E1_BP_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_ZUORDNUNG_INT] BIT,"
        .Append "[E1_BP_ZUORDNUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_VORSORGE_INT] BIT,"
        .Append "[E1_BP_VORSORGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VOTIERUNG_INT] BIT,"
        .Append "[E2_TP1_VOTIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_GENEHMIGUNG_INT] BIT,"
        .Append "[E2_TP1_GENEHMIGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AAR_INT] BIT,"
        .Append "[E2_TP1_AAR_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_STRATEGIE_INT] BIT,"
        .Append "[E2_TP1_STRATEGIE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_BERICHT_INT] BIT,"
        .Append "[E2_TP1_BERICHT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VERTRAG_INT] BIT,"
        .Append "[E2_TP1_VERTRAG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AUSZAHLUNG_INT] BIT,"
        .Append "[E2_TP1_AUSZAHLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_MA�NAHMEN_INT] BIT,"
        .Append "[E2_TP1_MA�NAHMEN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_FORBEARANCE_INT] BIT,"
        .Append "[E2_TP1_FORBEARANCE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_OFFENLEGUNG_INT] BIT,"
        .Append "[E2_TP2_OFFENLEGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_UNTERLAGEN_INT] BIT,"
        .Append "[E2_TP2_UNTERLAGEN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_EINREICHUNG_INT] BIT,"
        .Append "[E2_TP2_EINREICHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_AUSWERTUNG_INT] BIT,"
        .Append "[E2_TP2_AUSWERTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_MAHNVERFAHREN_INT] BIT,"
        .Append "[E2_TP2_MAHNVERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BEURTEILUNG_INT] BIT,"
        .Append "[E2_TP2_BEURTEILUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BERECHNUNG_INT] BIT,"
        .Append "[E2_TP2_BERECHNUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_NACHHALTIG_INT] BIT,"
        .Append "[E2_TP2_NACHHALTIG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_VERFAHREN_INT] BIT,"
        .Append "[E2_TP3_VERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_FAKTOREN_INT] BIT,"
        .Append "[E2_TP3_FAKTOREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_UEBERSCHREIBUNG_INT] BIT,"
        .Append "[E2_TP3_UEBERSCHREIBUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_TURNUS_INT] BIT,"
        .Append "[E2_TP3_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_ANLASS_INT] BIT,"
        .Append "[E2_TP3_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_AUSFALL_INT] BIT,"
        .Append "[E2_TP3_AUSFALL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERTRAEGE_INT] BIT,"
        .Append "[E2_TP4_VERTRAEGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_DURCHSETZBARKEIT_INT] BIT,"
        .Append "[E2_TP4_DURCHSETZBARKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_INT] BIT,"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_TURNUS_INT] BIT,"
        .Append "[E2_TP4_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_ANLASS_INT] BIT,"
        .Append "[E2_TP4_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERWALTUNG_INT] BIT,"
        .Append "[E2_TP4_VERWALTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_INT] BIT,"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERZIEHUNG_INT] BIT,"
        .Append "[E2_TP5_UEBERZIEHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_COVENANTS_INT] BIT,"
        .Append "[E2_TP5_COVENANTS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERWACHUNG_INT] BIT,"
        .Append "[E2_TP5_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SYSTEMATISCH_INT] BIT,"
        .Append "[E2_TP6_SYSTEMATISCH_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_ANLASS_INT] BIT,"
        .Append "[E2_TP6_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_BESTANDSAUFNAHME_INT] BIT,"
        .Append "[E2_TP6_BESTANDSAUFNAHME_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SOLL_INT] BIT,"
        .Append "[E2_TP6_SOLL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_ERMITTLUNG_INT] BIT,"
        .Append "[E2_TP7_ERMITTLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BESCHLUSS_INT] BIT,"
        .Append "[E2_TP7_BESCHLUSS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BERICHT_INT] BIT,"
        .Append "[E2_TP7_BERICHT_DOK] VARCHAR(MAX)"

        .Append " )"
        
        Intensiv = .ToString
    
    End With
    
End Function

Public Function Sanierung() As String

     With New Stringbuilder
        .Append "CREATE TABLE SANIERUNG ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[E1_BP_ENTSCHEIDUNG_INT] BIT,"
        .Append "[E1_BP_ENTSCHEIDUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_KDF_INT] BIT,"
        .Append "[E1_BP_KDF_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_RKV_INT] BIT,"
        .Append "[E1_BP_RKV_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_SICHERHEITEN_INT] BIT,"
        .Append "[E1_BP_SICHERHEITEN_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_UEBERWACHUNG_INT] BIT,"
        .Append "[E1_BP_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_ZUORDNUNG_INT] BIT,"
        .Append "[E1_BP_ZUORDNUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_VORSORGE_INT] BIT,"
        .Append "[E1_BP_VORSORGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VOTIERUNG_INT] BIT,"
        .Append "[E2_TP1_VOTIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_GENEHMIGUNG_INT] BIT,"
        .Append "[E2_TP1_GENEHMIGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AAR_INT] BIT,"
        .Append "[E2_TP1_AAR_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_STRATEGIE_INT] BIT,"
        .Append "[E2_TP1_STRATEGIE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_BERICHT_INT] BIT,"
        .Append "[E2_TP1_BERICHT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_VERTRAG_INT] BIT,"
        .Append "[E2_TP1_VERTRAG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AUSZAHLUNG_INT] BIT,"
        .Append "[E2_TP1_AUSZAHLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_FAEHIGKEIT_INT] BIT,"
        .Append "[E2_TP1_FAEHIGKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_KONZEPT_INT] BIT,"
        .Append "[E2_TP1_KONZEPT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_FORBEARANCE_INT] BIT,"
        .Append "[E2_TP1_FORBEARANCE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_OFFENLEGUNG_INT] BIT,"
        .Append "[E2_TP2_OFFENLEGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_UNTERLAGEN_INT] BIT,"
        .Append "[E2_TP2_UNTERLAGEN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_EINREICHUNG_INT] BIT,"
        .Append "[E2_TP2_EINREICHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_AUSWERTUNG_INT] BIT,"
        .Append "[E2_TP2_AUSWERTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_MAHNVERFAHREN_INT] BIT,"
        .Append "[E2_TP2_MAHNVERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BEURTEILUNG_INT] BIT,"
        .Append "[E2_TP2_BEURTEILUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_BERECHNUNG_INT] BIT,"
        .Append "[E2_TP2_BERECHNUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_NACHHALTIG_INT] BIT,"
        .Append "[E2_TP2_NACHHALTIG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_VERFAHREN_INT] BIT,"
        .Append "[E2_TP3_VERFAHREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_FAKTOREN_INT] BIT,"
        .Append "[E2_TP3_FAKTOREN_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_UEBERSCHREIBUNG_INT] BIT,"
        .Append "[E2_TP3_UEBERSCHREIBUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_TURNUS_INT] BIT,"
        .Append "[E2_TP3_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_ANLASS_INT] BIT,"
        .Append "[E2_TP3_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_AUSFALL_INT] BIT,"
        .Append "[E2_TP3_AUSFALL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERTRAEGE_INT] BIT,"
        .Append "[E2_TP4_VERTRAEGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_DURCHSETZBARKEIT_INT] BIT,"
        .Append "[E2_TP4_DURCHSETZBARKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_INT] BIT,"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_TURNUS_INT] BIT,"
        .Append "[E2_TP4_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_ANLASS_INT] BIT,"
        .Append "[E2_TP4_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERWALTUNG_INT] BIT,"
        .Append "[E2_TP4_VERWALTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_INT] BIT,"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERZIEHUNG_INT] BIT,"
        .Append "[E2_TP5_UEBERZIEHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_COVENANTS_INT] BIT,"
        .Append "[E2_TP5_COVENANTS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERWACHUNG_INT] BIT,"
        .Append "[E2_TP5_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_STILL_INT] BIT,"
        .Append "[E2_TP5_STILL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SYSTEMATISCH_INT] BIT,"
        .Append "[E2_TP6_SYSTEMATISCH_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_ANLASS_INT] BIT,"
        .Append "[E2_TP6_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_BESTANDSAUFNAHME_INT] BIT,"
        .Append "[E2_TP6_BESTANDSAUFNAHME_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SOLL_INT] BIT,"
        .Append "[E2_TP6_SOLL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_ERMITTLUNG_INT] BIT,"
        .Append "[E2_TP7_ERMITTLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BESCHLUSS_INT] BIT,"
        .Append "[E2_TP7_BESCHLUSS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BERICHT_INT] BIT,"
        .Append "[E2_TP7_BERICHT_DOK] VARCHAR(MAX)"

        .Append " )"
        
        Sanierung = .ToString
    
    End With
    
End Function

Public Function Abwicklung() As String

     With New Stringbuilder
        .Append "CREATE TABLE ABWICKLUNG ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[E1_BP_ENTSCHEIDUNG_INT] BIT,"
        .Append "[E1_BP_ENTSCHEIDUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_KDF_INT] BIT,"
        .Append "[E1_BP_KDF_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_RKV_INT] BIT,"
        .Append "[E1_BP_RKV_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_SICHERHEITEN_INT] BIT,"
        .Append "[E1_BP_SICHERHEITEN_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_UEBERWACHUNG_INT] BIT,"
        .Append "[E1_BP_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_ZUORDNUNG_INT] BIT,"
        .Append "[E1_BP_ZUORDNUNG_DOK] VARCHAR(MAX),"
        .Append "[E1_BP_VORSORGE_INT] BIT,"
        .Append "[E1_BP_VORSORGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_GENEHMIGUNG_INT] BIT,"
        .Append "[E2_TP1_GENEHMIGUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_AAR_INT] BIT,"
        .Append "[E2_TP1_AAR_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_BERICHT_INT] BIT,"
        .Append "[E2_TP1_BERICHT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_KONZEPT_INT] BIT,"
        .Append "[E2_TP1_KONZEPT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_OUTSOURCING_INT] BIT,"
        .Append "[E2_TP1_OUTSOURCING_DOK] VARCHAR(MAX),"
        .Append "[E2_TP1_FORBEARANCE_INT] BIT,"
        .Append "[E2_TP1_FORBEARANCE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP2_RUMPFKDF_INT] BIT,"
        .Append "[E2_TP2_RUMPFKDF_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_ANLASS_INT] BIT,"
        .Append "[E2_TP3_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP3_AUSFALL_INT] BIT,"
        .Append "[E2_TP3_AUSFALL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERTRAEGE_INT] BIT,"
        .Append "[E2_TP4_VERTRAEGE_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_DURCHSETZBARKEIT_INT] BIT,"
        .Append "[E2_TP4_DURCHSETZBARKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_INT] BIT,"
        .Append "[E2_TP4_PLAUSIBILISIERUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_TURNUS_INT] BIT,"
        .Append "[E2_TP4_TURNUS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_ANLASS_INT] BIT,"
        .Append "[E2_TP4_ANLASS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERWALTUNG_INT] BIT,"
        .Append "[E2_TP4_VERWALTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VERWERTUNG_INT] BIT,"
        .Append "[E2_TP4_VERWERTUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_INT] BIT,"
        .Append "[E2_TP4_VOLLSTAENDIGKEIT_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_UEBERWACHUNG_INT] BIT,"
        .Append "[E2_TP5_UEBERWACHUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP5_STILL_INT] BIT,"
        .Append "[E2_TP5_STILL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP6_SOLL_INT] BIT,"
        .Append "[E2_TP6_SOLL_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_ERMITTLUNG_INT] BIT,"
        .Append "[E2_TP7_ERMITTLUNG_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BESCHLUSS_INT] BIT,"
        .Append "[E2_TP7_BESCHLUSS_DOK] VARCHAR(MAX),"
        .Append "[E2_TP7_BERICHT_INT] BIT,"
        .Append "[E2_TP7_BERICHT_DOK] VARCHAR(MAX)"

        .Append " )"
        
        Abwicklung = .ToString
    
    End With
    
End Function

Public Function Paar() As String

     With New Stringbuilder
        .Append "CREATE TABLE PAAR ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[DARSTELLUNG] VARCHAR(MAX),"
        .Append "[VERHAELTNISSE] VARCHAR(MAX),"
        .Append "[SICHERHEITEN] VARCHAR(MAX),"
        .Append "[EWB] VARCHAR(MAX),"
        .Append "[FESTSTELLUNGEN] VARCHAR(MAX)"

        .Append " )"
        
        Paar = .ToString
    
    End With
    
End Function

Public Function Abstimmung() As String

     With New Stringbuilder
        .Append "CREATE TABLE ABSTIMMUNG ( "
        .Append "[ID] INT IDENTITY(1,1),"
        .Append "[KNE] VARCHAR(60) PRIMARY KEY,"
        .Append "[DATUM] DATETIME,"
        .Append "[INSTITUT] VARCHAR(MAX),"
        .Append "[BBK] VARCHAR(MAX),"
        .Append "[FRAGE] VARCHAR(MAX),"
        .Append "[REAKTION] VARCHAR(MAX),"
        .Append "[FAZIT] VARCHAR(MAX)"
        
        .Append " )"
        
        Abstimmung = .ToString
    
    End With
    
End Function
