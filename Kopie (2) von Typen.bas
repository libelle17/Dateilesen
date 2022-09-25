Dim sql$, T1!, T2!

Public type namen
 Pat_ID as long 'Pat_ID 3 '3000
 lfdnr as long 'lfdnr 3 'laufende Patientennummer
 NVorsatz as string 'NVorsatz 129 '3100
 Nachname as string 'Nachname 129 '3101
 Vorname as string 'Vorname 129 '3102
 GebDat as date 'GebDat 135 '3103
 Straße as string 'Straße 129 '3107
 KVKStatus as string 'KVKStatus 129 '3108
 Geschlecht as string 'Geschlecht 129 '3110
 Plz as string 'Plz 129 '3112
 Ort as string 'Ort 129 '3113
 Weggeldzone as string 'Weggeldzone 129 '3631
 AufnDat as date 'AufnDat 135 '3610
 LANR as string 'LANR 129 '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr as string 'BStNr 129 '3536 Betriebsstättennummer
 Titel as string 'Titel 129 '3104
 Versichertennummer as string 'Versichertennummer 129 '3105
 PrivatTel as string 'PrivatTel 129 '3629
 KVNr as string 'KVNr 129 '3630
 PrivatTel_2 as string 'PrivatTel_2 129 '3629
 PrivatFax as string 'PrivatFax 129 '3629
 DienstTel as string 'DienstTel 129 '3629
 PrivatMobil as string 'PrivatMobil 129 '3629
 Email as string 'Email 129 'Email
 Arbeitgeber as string 'Arbeitgeber 129 '3625
 AnAllgda as integer 'AnAllgda 11 'Anamnese allgemein da
 An1da as integer 'An1da 11 'Anamnese S.1 da
 An2da as integer 'An2da 11 'Anamnese S.2 da
 Checkda as integer 'Checkda 11 'Checkliste da
 DMTypaD as string 'DMTypaD 129 'aus Diagnosen
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Cave as string 'Cave 129 '3654
 Notiz as string 'Notiz 129 '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>
 zubenach as string 'zubenach 129 '3633
 Verwandt as string 'Verwandt 129 '3632
 Sprache as string 'Sprache 129 '3628
end type

Public type faelle
 FID as long 'FID 3 'Primärschlüssel
 Pat_ID as long 'Pat_ID 3 '3000
 Quartal as string 'Quartal 129 '4101
 Nachname as string 'Nachname 129 '3101
 Vorname as string 'Vorname 129 '3102
 lfdnr as long 'lfdnr 3 'laufende Fallnummer
 TMFNr as string 'TMFNr 129 '4144 Fallnummer in Turbomed
 VKNr as string 'VKNr 129 '4104
 BhFB as date 'BhFB 135 '4150
 BhFE1 as date 'BhFE1 135 '4151
 BhFE2 as date 'BhFE2 135 '4152
 f4202 as string 'f4202 129 '4202
 ausgst as date 'ausgst 135 '4102 ('ausgestellt am')
 KtrAbrB as string 'KtrAbrB 129 '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr as string 'AbrAr 129 '4107, Abrechnungsart (1 = Primärkassen)
 lVorl as date 'lVorl 135 '4109, letzte Vorlage
 IK as string 'IK 129 '4111 (auch patientenbezogen)
 KVKs as string 'KVKs 129 '4112
 KVKserg as string 'KVKserg 129 '4113
 GebOr as string 'GebOr 129 '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb as string 'AbrGb 129 '4122, Abrechnungsgebiet (07 = Diabetes)
 f4206 as string 'f4206 129 '4206, ?
 ÜwText as string 'ÜwText 129 '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 as string 'f4210 129 '4210, unbekannt
 statNuller as string 'statNuller 129 '4216, nu bei Musterfrau 16 Nuller
 ÜbwV as string 'ÜbwV 129 '4218, überwiesen von
 AndÜw as string 'AndÜw 129 '4219, anderer Überweiser
 Übw as string 'Übw 129 '4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR as string 'ÜbwLANR 129 '4242 LANR des Überweisers
 ÜWZiel as string 'ÜWZiel 129 '4220
 ÜWNNr as string 'ÜWNNr 129 '4231(4): KV-Nummer des Überweisers
 ÜWNaN as string 'ÜWNaN 129 '4231(3): Nachname des Überweisers
 ÜWTit as string 'ÜWTit 129 '4231(3): Titel des Überweisers
 ÜWVor as string 'ÜWVor 129 '4231(2): Vorname des Überweisers
 ÜWVsw as string 'ÜWVsw 129 '4231(2b): Vorsatzwort des Überweisers
 statKlasse as string 'statKlasse 129 '4236
 f4237 as string 'f4237 129 '4237, ? (nur bei Musterw)
 statBehTage as long 'statBehTage 3 '4238
 SchGr as string 'SchGr 129 '4239, Scheingruppe
 Weiterbeh as string 'Weiterbeh 129 '4243, Weiterbehandelnder
 PGeb as string 'PGeb 129 '4401, Praxisgebühr
 PGebErg as string 'PGebErg 129 '4402, ?
 Mahnfrist as string 'Mahnfrist 129 '4403, Mahnfrist bis
 GOÄKatNr as string 'GOÄKatNr 129 '4580 (1): Katalog-Nummer
 GOÄKatName as string 'GOÄKatName 129 '4580 (2): Privat-Abrechnungskatalog
 abrArzt as string 'abrArzt 129 '4585 abrechnender Arzt
 privVers as string 'privVers 129 '4586 private Versicherung
 AdNam as string 'AdNam 129 '4602(1) Name bei Musterfrau
 AdStr as string 'AdStr 129 '4602(2) Straße bei Musterfrau
 AdPlz as string 'AdPlz 129 '4602(3) PLZ bei Musterfrau
 AdOrt as string 'AdOrt 129 '4602(4) Ort bei Musterfrau
 BhFE as date 'BhFE 135 '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 as string 's8000 129 '8000, ???
 s8100 as string 's8100 129 '8100, ???
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 Fanf as date 'Fanf 135 'Fallanfang
 altQuart as string 'altQuart 129 '
 QAnf as date 'QAnf 135 'Quartalsanfang
 QEnd as date 'QEnd 135 'Quartalsende
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 TherArt as long 'TherArt 3 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
end type

Public type au
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '6200 + 6201
 Beginn as string 'Beginn 129 '6285 1. Hälfte
 Ende as string 'Ende 129 '6285 2. Hälfte
 ICDs as string 'ICDs 129 '6286
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type briefe
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '
 ZeitPunkt as date 'ZeitPunkt 135 '
 Pfad as string 'Pfad 129 '
 Art as string 'Art 129 '
 Name as string 'Name 129 '
 Typ as string 'Typ 129 '
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 DokGroe as long 'DokGroe 3 'Größe der Datei
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type diagnosen
 ID1 as long 'ID1 3 '
 FID as long 'FID 3 'Fall-Bezug
 Pat_id as long 'Pat_id 3 'Bezug auf Anamneseblattt
 GesName as string 'GesName 129 '
 DiagDatum as date 'DiagDatum 135 '
 DiagSicherheit as string 'DiagSicherheit 129 '6003
 DiagText as string 'DiagText 129 '
 DiagSeite as string 'DiagSeite 129 '6004
 DiagAttr as string 'DiagAttr 129 '6006
 ICD as string 'ICD 129 '
 obDauer as byte 'obDauer 16 'ob Dauerdiagnose
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type dokumente
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '
 ZeitPunkt as date 'ZeitPunkt 135 '
 DokPfad as string 'DokPfad 129 '
 DokArt as string 'DokArt 129 '
 DokName as string 'DokName 129 '
 Quelldatum as date 'Quelldatum 135 'Datum, auf das sich das Dokument bezieht
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 DokGroe as long 'DokGroe 3 'Dokument-Größe
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type eintraege
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '
 Art as string 'Art 129 '6330
 Inhalt as string 'Inhalt 129 '8480
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type forminhaltform_abk
 Form_AbkVW as long 'Form_AbkVW 3 '
 Form_Abk as string 'Form_Abk 129 '
end type

Public type formulare
 FormID as long 'FormID 3 '
 Form_Abk as string 'Form_Abk 129 '
 FormBez as string 'FormBez 129 '
 FormVorl as string 'FormVorl 129 '
 AktZeit as date 'AktZeit 135 'Zeitpunkt der Aktualisierung
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type forminhkopf
 FoID as long 'FoID 3 '
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '
 Form_ID as long 'Form_ID 3 '
 ZeitPunkt as date 'ZeitPunkt 135 '
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Satzart as string 'Satzart 129 '8000
 Satzlänge as string 'Satzlänge 129 '8100
end type

Public type forminhfeld
 FoID as long 'FoID 3 '
 Nr as integer 'Nr 2 '
 FeldNr as integer 'FeldNr 2 '
 FeldVW as long 'FeldVW 3 '
 Feld as string
 FeldInhVW as long 'FeldInhVW 3 '
 FeldInh as string
end type

Public type kheinweis
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '
 Ziel as string 'Ziel 129 '6291
 Diagnose as string 'Diagnose 129 '6230
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type lbanforderungen
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '6200 + 6201
 AnfText as string 'AnfText 129 '6280
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Statusbyte
end type

Public type laborneu
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '
 FertigStGrad as string 'FertigStGrad 129 '8401
 Abkü as string 'Abkü 129 '8410
 LangtextVW as long 'LangtextVW 3 '8411
 Langtext as string
 Wert as string 'Wert 129 '8420
 Einheit as string 'Einheit 129 '8421
 KommentarVW as long 'KommentarVW 3 '8480
 Kommentar as string
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 Refnr as long 'Refnr 3 'Bezug auf LaborXUS
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type leistungen
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '5000 + 6201
 Leistung as string 'Leistung 129 '5001
 f5002 as string 'f5002 129 '5002
 f5005 as string 'f5005 129 '5005
 f5006 as string 'f5006 129 '5006
 f5009 as string 'f5009 129 '5009
 Med as string 'Med 129 '5010
 f5015 as string 'f5015 129 '5015
 f5016 as string 'f5016 129 '5016
 f5021 as string 'f5021 129 '5021
 f5026 as string 'f5026 129 '5026
 Faktor as string 'Faktor 129 '5062 Multiplikator für GOÄ-Rechnung
 f5098 as string 'f5098 129 '5098 0000000000
 LANR as string 'LANR 129 '5099 LANR
 letzVorg as date 'letzVorg 135 '5101 letzter Vorgang
 Ausn as string 'Ausn 129 '3677 Ausnahme/Begründung für abweichendes Geschlecht
 Beme as string 'Beme 129 '         Bemerkung
 absPos as long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type medplan
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 MPNr as long 'MPNr 3 'Ordnungsziffer für Medikamentenplan
 ZeitPunkt as date 'ZeitPunkt 135 'Zeitpunkt, der Speicherung im Turbomed
 Datum as date 'Datum 135 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament as string 'Medikament 129 '
 MedAnfang as string 'MedAnfang 129 '
 FeldNr as integer 'FeldNr 2 '
 mo as string 'mo 129 '
 mi as string 'mi 129 '
 nm as string 'nm 129 '
 ab as string 'ab 129 '
 zn as string 'zn 129 '
 bBed as integer 'bBed 11 '
 Bemerkung as string 'Bemerkung 129 '
 AbsPos as long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type rezepteintraege
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '6200 + 6201
 Rezept as string 'Rezept 129 '6210, 3652(1), 6218(1)
 Rezeptklasse as string 'Rezeptklasse 129 '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Medikament as string 'Medikament 129 '3652(2), 6218(4)
 PZN as string 'PZN 129 '6210(2), 6218(3)
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 QS as string 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT as string 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte as long 'StByte 3 'Statusbyte
end type

Public type rr
 FID as long 'FID 3 'Fall-Bezug
 Pat_ID as long 'Pat_ID 3 '3000
 ZeitPunkt as date 'ZeitPunkt 135 '6200 + 6201
 RR as string 'RR 129 '6230
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type kvnrue
 lfdnr as long 'lfdnr 3 '
 Pat_ID as long 'Pat_ID 3 '
 KVNr as string 'KVNr 129 '
 absPos as long 'absPos 3 'Zeile in BDT-Datei
 AktZeit as date 'AktZeit 135 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type unbekannte_kennungen
 Kennung as string 'Kennung 129 '
 absPos as long 'absPos 3 '
 StByte as long 'StByte 3 '
 Pat_id as long 'Pat_id 3 'zugehöriger Patient für spätere Ermittlungen
 Inhalt as string 'Inhalt 129 'Inhalt Zeile zum Wiederauffinden
end type

Public type dmpreihe
 Abk as string 'Abk 129 'Abkürzung der DMP-Art
 Art as string 'Art 129 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum as date 'KarteiDatum 133 'Datum des Karteikarteneintrags der Dokumentation
 exportiert as date 'exportiert 135 'Datum des Exports
 DokuDatum as date 'DokuDatum 135 'Datum der Dokumentation
 obvoll as integer 'obvoll 11 'ob vollständig
 NachName as string 'NachName 129 '
 VorName as string 'VorName 129 '
 GebDat as date 'GebDat 133 '
 Pat_id as long 'Pat_id 3 '
 StByte as long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AktZeit as date 'AktZeit 135 'Aktualisierungzeit
end type

Public type laborxsaetze
 SatzID as long 'SatzID 3 'zum Bezug für LaborUS
 DatID as long 'DatID 3 'Bezug zu LaborEingelesen
 Satzart as string 'Satzart 129 '8000 Satzart (Turbomed)
 Satzlänge as string 'Satzlänge 129 '8100 Satzlänge (Turbomed)
 SatzlängeSchluss as string 'SatzlängeSchluss 129 '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb as string 'VersionSatzb 129 '9212 Version der Satzbeschreibung (Turbomed)
 Arztnr as string 'Arztnr 129 '201 Arztnummer (Turbomed)
 Arztname as string 'Arztname 129 '203 Arztname (Turbomed)
 StraßePraxis as string 'StraßePraxis 129 '205 Straße der Praxis (Turbomed)
 PLZPraxis as string 'PLZPraxis 129 '215 PLZ der Praxis (Turbomed)
 OrtPraxis as string 'OrtPraxis 129 '216 Ort der Praxis (Turbomed)
 Labor as string 'Labor 129 '8320 Labor
 StraßeLabor as string 'StraßeLabor 129 '8321 Straße der Laboradresse (Turbomed)
 PLZLabor as string 'PLZLabor 129 '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor as string 'OrtLabor 129 '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr as string 'KBVPrüfnr 129 '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz as string 'Zeichensatz 129 '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr as string 'Kundenarztnr 129 '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum as string 'Erstellungsdatum 129 '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge as string 'Gesamtlänge 129 '9202 Gesamtlänge des Datenpaketes (Turbomed)
end type

Public type laborxeingel
 DatID as long 'DatID 3 'Bezug auf LaborEingelesen
 Pfad as string 'Pfad 129 'Pfadname
 Name as string 'Name 129 'Name der eingelesenen Labordatei ohne Endung
 Zp as date 'Zp 135 'Einlesezeitpunkt
 fertig as integer 'fertig 11 'ob Einlesen fertig
end type

Public type laborxus
 RefNr as long 'RefNr 3 'Bezug auf LaborWert
 DatID as long 'DatID 3 'Bezug auf LaborEingelesen
 SatzID as long 'SatzID 3 'Bezug auf LaborXSätze
 Satzart as string 'Satzart 129 '8000 Satzart (Turbomed)
 Satzlänge as string 'Satzlänge 129 '8100 Satzlänge (Turbomed)
 Auftragsnummer as string 'Auftragsnummer 129 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel as string 'Auftragsschlüssel 129 '8311 Anforderungsnr d Labors (Turbomed)
 Eingang as date 'Eingang 135 '8301 Eingangsdatum in Datumsform
 Berichtsdatum as string 'Berichtsdatum 129 '8302 Berichtsdatum
 Pat_id as long 'Pat_id 3 '
 Nachname as string 'Nachname 129 '3101
 Vorname as string 'Vorname 129 '3102
 GebDat as string 'GebDat 129 '3103
 Titel as string 'Titel 129 '3104
 NVorsatz as string 'NVorsatz 129 '3100
 BefArt as string 'BefArt 129 '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Abrechnungstyp as string 'Abrechnungstyp 129 '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd as string 'GebüOrd 129 '8403 Gebührenordnung (Turbomed)
 Patienteninformation as string 'Patienteninformation 129 '8405 Patienteninformation (Turbomed)
 Geschlecht as string 'Geschlecht 129 '8407 Geschlecht (Turbomed)
 AuftrHinw as string 'AuftrHinw 129 '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp as string 'Pat_idUrsp 129 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG as string 'Pat_idErwVNG 129 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN as string 'Pat_idErwVN 129 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG as string 'Pat_idErwG 129 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB as string 'Pat_idErwGB 129 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL as string 'Pat_idErwGL 129 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu as string 'Pat_idLaborNeu 129 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu as date 'ZeitpunktLaborneu 135 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP as integer 'ZdüP 2 'Zahl der verglichenen Parameter
 ZdiP as long 'ZdiP 3 'Zahl der infragekommenden Patienten
 LWerte as string 'LWerte 129 'Laborwerte, die zur Zuordnung geführt haben
 verglichen as date 'verglichen 135 'Datum, zu dem Datensatz zuletzt verglichen wurde
 AfN as integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
end type

Public type laborxbakt
 RefNr as long 'RefNr 3 '
 Verf as string 'Verf 129 '
 KuQu as string 'KuQu 129 '8428 Probenmaterial-Ident (Turbomed)
 Quelle as string 'Quelle 129 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez as string 'QSpez 129 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat as date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Kommentar as string 'Kommentar 129 '8480 Ergebnistest (Turbomed)
 Erklärung as string 'Erklärung 129 '
 Keimzahl as string 'Keimzahl 129 '
end type

Public type laborxwert
 RefNr as long 'RefNr 3 'Bezug auf LaborUS
 Abkü as string 'Abkü 129 '8410 Test-Ident  (Turbomed)
 Langname as string 'Langname 129 '8411 Testbezeichnung (Turbomed)
 Quelle as string 'Quelle 129 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez as string 'QSpez 129 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat as date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Wert as string 'Wert 129 '8420 Ergebniswert (Turbomed)
 Einheit as string 'Einheit 129 '8421 Einheit (Turbomed)
 Grenzwerti as string 'Grenzwerti 129 '8422 Grenzwertindikator (Turbomed)
 Kommentar as string 'Kommentar 129 '8480 Ergebnistext (Turbomed)
 Teststatus as string 'Teststatus 129 '8418 Teststatus (Turbomed)
 Erklärung as string 'Erklärung 129 '8470 Testbezogene Hinweise (Turbomed)
 Normbereich as string 'Normbereich 129 '8460 Normalwert-Text (Turbomed)
 NormU as string 'NormU 129 '8461 Normuntergrenze
 NormO as string 'NormO 129 '8461 Normobergrenze
 AuftrHinw as string 'AuftrHinw 129 '8490 Auftragsbezogene Hinweise (Turbomed)
end type

Public type laborxleist
 RefNr as long 'RefNr 3 'Bezug auf LaborUS
 Abkü as string 'Abkü 129 '8410 Test-Ident (Turbomed)
 Verf as string 'Verf 129 '8434
 EBM as string 'EBM 129 '5001 GNR (Turbomed)
 goä as string 'goä 129 '8406
 Anzahl as string 'Anzahl 129 '5005
end type

Public type anamnesebogen
 Pat_id as long 'Pat_id 3 '
 Nachname as string 'Nachname 129 '-
 Vorname as string 'Vorname 129 '
 NVorsatz as string 'NVorsatz 129 '
 Titel as string 'Titel 129 '
 Anrede as string 'Anrede 129 '
 GebDat as date 'GebDat 135 ', geb.
 Tkz as byte 'Tkz 16 'Tod-Kennzeichen
 Versicherungsart as string 'Versicherungsart 129 '
 Diabetestyp as string 'Diabetestyp 129 '^Diabetes Typ
 Diabetes_seit as string 'Diabetes_seit 129 '<seit
 Tabletten_seit as string 'Tabletten_seit 129 ', Tabletten seit
 Insulin_seit as string 'Insulin_seit 129 ', Insulin seit
 Grund_für_Vorstellung as string 'Grund_für_Vorstellung 129 '^:
 Familienanamnese as string 'Familienanamnese 129 '^:
 Größe as double 'Größe 5 '^:
 Gewicht as double 'Gewicht 5 ',:
 Tendenz as string 'Tendenz 129 '<, Tendenz
 DiabetesMedikament_1 as string 'DiabetesMedikament_1 129 '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge as string 'DiabetesMedikament_1_Menge 129 '<
 DiabetesMedikament_2 as string 'DiabetesMedikament_2 129 '<,
 DiabetesMedikament_2_Menge as string 'DiabetesMedikament_2_Menge 129 '<
 DiabetesMedikament_3 as string 'DiabetesMedikament_3 129 '<,
 DiabetesMedikament_3_Menge as string 'DiabetesMedikament_3_Menge 129 '<
 DiabetesMedikament_4 as string 'DiabetesMedikament_4 129 '<,
 DiabetesMedikament_4_Menge as string 'DiabetesMedikament_4_Menge 129 '<,
 Insulinpumpe as byte 'Insulinpumpe 16 '^:
 Insulinpumpe_seit as string 'Insulinpumpe_seit 129 '<seit
 Insulinpumpe_Marke as string 'Insulinpumpe_Marke 129 '<, Marke:
 Broteinheiten_gesamt as string 'Broteinheiten_gesamt 129 '^Broteinheiten:gesamt
 Broteinheiten_früh as string 'Broteinheiten_früh 129 '<, früh
 Broteinheiten_ZM_früh as string 'Broteinheiten_ZM_früh 129 '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags as string 'Broteinheiten_mittags 129 '<, mittags
 Broteinheiten_nachmittags as string 'Broteinheiten_nachmittags 129 '<, nachmittags
 Broteinheiten_abends as string 'Broteinheiten_abends 129 '<, abends
 Broteinheiten_nachts as string 'Broteinheiten_nachts 129 '<, nachts
 Essenszeit_früh as string 'Essenszeit_früh 129 '^Essenszeiten:früh
 Essenszeit_vormittags as string 'Essenszeit_vormittags 129 '<, vormittags
 Essenszeit_mittags as string 'Essenszeit_mittags 129 '<, mittags
 Essenszeit_nachmittags as string 'Essenszeit_nachmittags 129 '<, nachmittags
 Essenszeit_abends as string 'Essenszeit_abends 129 '<, abends
 Essenszeit_spät as string 'Essenszeit_spät 129 '<, spät
 Spritz_Eß_Abstand_früh as string 'Spritz_Eß_Abstand_früh 129 '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags as string 'Spritz_Eß_Abstand_mittags 129 '<, mittags
 Spritz_Eß_Abstand_abends as string 'Spritz_Eß_Abstand_abends 129 '<, abends
 Spritzstelle_früh as string 'Spritzstelle_früh 129 '^Spritzstellen:früh
 Spritzstelle_mittags as string 'Spritzstelle_mittags 129 '<, mittags
 Spritzstelle_abends as string 'Spritzstelle_abends 129 '<, abends
 Spritzstelle_nachts as string 'Spritzstelle_nachts 129 '<, nachts
 Jahr_letzte_Diabetesschulung as string 'Jahr_letzte_Diabetesschulung 129 '^Letzte Diabetesschulung:
 Ort_Schulung as string 'Ort_Schulung 129 '<in
 letztes_HbA1c as string 'letztes_HbA1c 129 '^Letztes HbA1c:
 gemessen_am as date 'gemessen_am 135 '<, gemessen
 vorherige_Werte as string 'vorherige_Werte 129 '<, vorher:
 BZMessungen_selbst as string 'BZMessungen_selbst 129 '^Blutzuckermessung:Selbstmessung?
 Gerät as string 'Gerät 129 '<:
 BZMessungen_pW as string 'BZMessungen_pW 129 '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE as string 'BZMessungen_pW_ndE 129 '<, davon nach dem Essen:
 BZMessungen_p_W_nachts as string 'BZMessungen_p_W_nachts 129 '<, nachts:
 Aufschreiben as string 'Aufschreiben 129 '<, Dokumentation:
 BZWerte_v_d_Essen as string 'BZWerte_v_d_Essen 129 '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen as string 'BZWerte_n_d_Essen 129 '<, nach dem Essen:
 UZ_Tageszeit as string 'UZ_Tageszeit 129 '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM as string 'Unterzucker_pM 129 '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig as string 'UZ_rechtzeitig 129 '<, rechtzeitig bemerkt:
 Fremde_Hilfe_pa as string 'Fremde_Hilfe_pa 129 '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa as string 'Bewußtlos_pa 129 '<, bewußtlos deshalb:
 Keto_pa as string 'Keto_pa 129 '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM as string 'BZgr300_pM 129 ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck as string 'Bluthochdruck 129 '^Bluthochdruck:
 BHD_seit as string 'BHD_seit 129 '<seit:
 BHD_beh_mit as string 'BHD_beh_mit 129 '<, behandelt mit:
 Blutdruckwerte as string 'Blutdruckwerte 129 '^Blutdruckwerte:
 BDselbst as string 'BDselbst 129 '^Blutdruckselbstmessung:
 Schwanger as string 'Schwanger 129 '^Aktuelle Schwangerschaft:
 Schwanger_seit as string 'Schwanger_seit 129 '<, seit:
 Augensp_zuletzt as string 'Augensp_zuletzt 129 '^Letzte Augenspiegelung:
 Augensp_Befund as string 'Augensp_Befund 129 '<, Befund:
 Netzhaut_gelasert as string 'Netzhaut_gelasert 129 ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar as string 'Sehminderung_unbehebbar 129 ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden as string 'Diabet_Nierenschaden 129 '^Diabetischer Nierenschaden:
 Albumin_zuletzt as string 'Albumin_zuletzt 129 ', letztes Albumin:
 erhöht as string 'erhöht 129 '<, Befund:
 Dialyse as byte 'Dialyse 16 ',:
 Dialyse_seit as string 'Dialyse_seit 129 '<seit
 andere_Nierenerkrankung as string 'andere_Nierenerkrankung 129 ', andere Nierenerkrankung:
 Herzkrankheit as string 'Herzkrankheit 129 '^Herzkrankheit:
 Angina_pectoris as string 'Angina_pectoris 129 ',:
 Herzinfarkt as string 'Herzinfarkt 129 ',:
 Herzinfarkt_wann as string 'Herzinfarkt_wann 129 '<, wann:
 PTCA_oder_Stent as string 'PTCA_oder_Stent 129 ',:
 Bypass_kardial as byte 'Bypass_kardial 16 ',:
 Bypass_wann as string 'Bypass_wann 129 '<, wann:
 Herzschwäche as string 'Herzschwäche 129 ',:
 Herzkrankheit_Beschreibung as string 'Herzkrankheit_Beschreibung 129 ', Beschreibung:
 Hirndurchblutungsstörung as string 'Hirndurchblutungsstörung 129 '^:
 Schlaganfall as string 'Schlaganfall 129 ',:
 Beindurchblutungsstörung as string 'Beindurchblutungsstörung 129 '^:
 Schaufensterkrankheit as string 'Schaufensterkrankheit 129 ',:
 Bypaß_peripher as byte 'Bypaß_peripher 16 ',:
 Geschwür as string 'Geschwür 129 ',:
 Amputation as string 'Amputation 129 ',:
 pAVK_Beschreibung as string 'pAVK_Beschreibung 129 ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen as string 'Ameisenlaufen 129 '^:
 Ameisen_Ausmaß as string 'Ameisen_Ausmaß 129 '<, Ausmaß:
 Druckstellen as string 'Druckstellen 129 ',:
 Verformungen as string 'Verformungen 129 ',:
 Verformungen_Beschreibung as string 'Verformungen_Beschreibung 129 '<Beschreibung:
 Fußpflege as string 'Fußpflege 129 '^:
 Podologie as string 'Podologie 129 ',:
 Einlagen as string 'Einlagen 129 ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen as string 'Neue_Fußkomplikationen 129 '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen as string 'Entleerungsstörungen_Magen 129 '^:
 Entleerungsstörungen_Harnblase as string 'Entleerungsstörungen_Harnblase 129 ',:
 Schwindel_Aufstehen as string 'Schwindel_Aufstehen 129 ',:
 Folgeerkrankungen_Haut as string 'Folgeerkrankungen_Haut 129 '^:
 Bewegungseinschränkungen as string 'Bewegungseinschränkungen 129 ',:
 Sexualstörung as string 'Sexualstörung 129 '^:
 Sexualstörung_seit as string 'Sexualstörung_seit 129 '<seit
 Weitere_Anamnese as string 'Weitere_Anamnese 129 '^:
 Alkohol as string 'Alkohol 129 '^:
 Tabak as string 'Tabak 129 ',:
 Weitere_Medikation as string 'Weitere_Medikation 129 '^:
 Liphypertrophien_Abdomen as string 'Liphypertrophien_Abdomen 129 '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine as string 'Liphypertrophien_Beine 129 '<, Beine:
 Liphypertrophien_Arme as string 'Liphypertrophien_Arme 129 '<, Arme:
 Beinbefund as string 'Beinbefund 129 '^:
 Hyperkeratosen as string 'Hyperkeratosen 129 ',:
 Ulcera as string 'Ulcera 129 ',:
 Kraft_Zehenheber as string 'Kraft_Zehenheber 129 '^Kraft:Zehenheber
 Kraft_Zehenbeuger as string 'Kraft_Zehenbeuger 129 '<, Zehenbeuger:
 Kraft_Knie as string 'Kraft_Knie 129 '<, Knie:
 ASR as string 'ASR 129 ',:
 PSR as string 'PSR 129 ',:
 Oberflächensensibilität as string 'Oberflächensensibilität 129 '^:
 Monofilamenttest as string 'Monofilamenttest 129 ',:
 Kalt_Warm as string 'Kalt_Warm 129 ', Kalt-Warm-Diskrimination:
 Vibration_IK as string 'Vibration_IK 129 ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe as string 'Vibration_Großzehe 129 '<, Großzehe:
 Puls_Leiste as string 'Puls_Leiste 129 '^Pulse:Leiste
 Puls_Kniekehle as string 'Puls_Kniekehle 129 '<,Kniekehle:
 Puls_Atp as string 'Puls_Atp 129 '<,Innenknöchel:
 Puls_Adp as string 'Puls_Adp 129 '<,Fußrücken:
 RR as string 'RR 129 '^Blutdruck:
 RRTurboMed as string 'RRTurboMed 129 '
 Herz as string 'Herz 129 '^:
 Lunge as string 'Lunge 129 ',:
 Bauch as string 'Bauch 129 ', Abdomen:
 WS as string 'WS 129 ', Wirbelsäule:
 NL as string 'NL 129 ', Nierenlager:
 SD as string 'SD 129 ', Schilddrüse:
 Carotiden as string 'Carotiden 129 ', Halsschlagadern:
 NNH as string 'NNH 129 ', Nasennebenhöhlen:
 Zähne as string 'Zähne 129 ',:
 Mundhöhle as string 'Mundhöhle 129 ',:
 LK as string 'LK 129 ', Lymphknoten:
 BeinödVen as string 'BeinödVen 129 ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst as string 'Neuro_sonst 129 '^Sonstige neurologische Befunde:
 Weitere_Befunde as string 'Weitere_Befunde 129 ', weitere Befunde:
 Schulung as string 'Schulung 129 'ob Schulungsbedarf
 DMP as string 'DMP 129 'ob Pat. bei HA im DMP
 DMSchulz as integer 'DMSchulz 2 'Zahl der DMP-Schulungen hier
 DMSchL as integer 'DMSchL 2 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz as integer 'RRSchulz 2 'Zahl der Hypertonie-Schulungen hier
 DMPhier as date 'DMPhier 135 'ob Pat hier im DMP
 HANr as string 'HANr 129 'mit "/"
 HANr2 as string 'HANr2 129 'mit "/"
 letzte_Änderung as date 'letzte_Änderung 135 'Datum der letzten Änderung
 Diagnosen as string 'Diagnosen 129 '
 Vorgestellt as date 'Vorgestellt 135 'Erstvorstellung
 Versicherung as string 'Versicherung 129 '
 AktZeit as date 'AktZeit 135 'Aktualisierungszeit
 Ther1 as string 'Ther1 129 'Diät, OAD, CT, ICT, CSII
 TherAkt as string 'TherAkt 129 'Diät, OAD, CT, ICT, CSII
 Prim as long 'Prim 3 'Primärschlüssel
 obAn1eing as byte 'obAn1eing 16 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing as byte 'obAn2eing 16 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing as byte 'obAnAeing 16 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck as byte 'obCheck 16 'ob Checkliste vorliegt
 obBZausgew as byte 'obBZausgew 16 'ob Blutzuckergerät ausgewechselt
 obOSaufgek as byte 'obOSaufgek 16 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek as byte 'obPodAufgek 16 'ob über Podologie aufgeklärt
 obMBlAusgeh as byte 'obMBlAusgeh 16 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek as string 'obSchulaufgek 129 'ob über Podologie aufgeklärt
 obDMPaufgekl as string 'obDMPaufgekl 129 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz as byte 'obMedNetz 16 'ob von Med. Netz geschickt
 Hausarzt as string 'Hausarzt 129 'Hausarzt laut Anamnesebogen
 ob as byte 'ob 16 'für verschiedene Aktionen
 QS as string 'QS 129 'Quartal sortiert von vorgestellt
 QT as string 'QT 129 'Quartal sortiert von vorgestellt
end type

Public rNa() as namen
Public rFa() as faelle
Public rAu() as au
Public rBr() as briefe
Public rDi() as diagnosen
Public rDo() as dokumente
Public rEi() as eintraege
Public rFi() as forminhaltform_abk
Public rFo() as formulare
Public rFr() as forminhkopf
Public rFm() as forminhfeld
Public rKh() as kheinweis
Public rLb() as lbanforderungen
Public rLa() as laborneu
Public rLe() as leistungen
Public rMe() as medplan
Public rRe() as rezepteintraege
Public rRr() as rr
Public rKv() as kvnrue
Public rUn() as unbekannte_kennungen
Public rDm() as dmpreihe
Public rLs() as laborxsaetze
Public rLg() as laborxeingel
Public rLu() as laborxus
Public rLo() as laborxbakt
Public rLw() as laborxwert
Public rLL() as laborxleist
Public rAna() as anamnesebogen

Public Function Tinit()
 static wdh%
 ReDim rAna(0)
 ReDim rNa(0)
 ReDim rFa(0)
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 if wdh = 0 then reDim rFi(0)
 if wdh = 0 then reDim rFo(0)
 ReDim rFr(0)
 ReDim rFm(0)
 ReDim rKh(0)
 ReDim rLb(0)
 ReDim rLa(0)
 ReDim rLe(0)
 ReDim rMe(0)
 ReDim rRe(0)
 ReDim rRr(0)
 ReDim rKv(0)
 ReDim rUn(0)
 ReDim rDm(0)
 wdh = -1
End Function ' Tinit

Public Function LabInit()
 static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 wdh = -1
End Function ' Tinit

Public Function doLösch(frm as lese, Tbl$)
 Set rs = DBCn.Execute("select count(*) as ct from " + kla + Tbl + klz)
 frm.Ausgabe = "Lösche: " & kla & Tbl & klz & " (" & rs!ct & " Datensätze)" & vbCrLf & altAusgabe
 altausgabe = frm.Ausgabe
 sql = sqldeletefrom & kla & Tbl & klz
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doLösch

Public Function AllesLösch(frm as lese)
 Dim ct&, rs as new ADODB.recordset
 on error goto fehler
 call ForeignNo0
 call ForeignNo1
 call doLösch(frm, "dmpreihe")
 call doLösch(frm, "unbekannte kennungen")
 call doLösch(frm, "kvnrue")
 call doLösch(frm, "rr")
 call doLösch(frm, "rezepteintraege")
 call doLösch(frm, "medplan")
 call doLösch(frm, "leistungen")
 call doLösch(frm, "laborneu")
 call doLösch(frm, "lbanforderungen")
 call doLösch(frm, "kheinweis")
 call doLösch(frm, "forminhfeld")
 call doLösch(frm, "forminhkopf")
 call doLösch(frm, "formulare")
 call doLösch(frm, "forminhaltform_abk")
 call doLösch(frm, "eintraege")
 call doLösch(frm, "dokumente")
 call doLösch(frm, "diagnosen")
 call doLösch(frm, "briefe")
 call doLösch(frm, "au")
 call doLösch(frm, "faelle")
 call doLösch(frm, "namen")
 call ForeignYes0
 call ForeignYes1
 exit function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' AllesLösch

Public Function LabLösch(frm as lese)
 Dim ct&, rs as new ADODB.recordset
 on error goto fehler
 call ForeignNo0
 call ForeignNo1
 call doLösch(frm, "laborxleist")
 call doLösch(frm, "laborxwert")
 call doLösch(frm, "laborxbakt")
 call doLösch(frm, "laborxus")
 call doLösch(frm, "laborxeingel")
 call doLösch(frm, "laborxsaetze")
 call ForeignYes0
 call ForeignYes1
 exit function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' LabLösch


Public Function namenSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere namen"
 if not Allepat then
  sql = "delete from " & kla & "namen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "namen" & klz & " (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,AufnDat," & _
     "LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "Cave,Notiz,zubenach,Verwandt,Sprache) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "namen" & klz & " (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,AufnDat," & _
     "LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "Cave,Notiz,zubenach,Verwandt,Sprache) values"))
 For i = 0 to ubound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rNa(i).Pat_ID & "," & rNa(i).lfdnr & ",'" & rNa(i).NVorsatz & "','" & rNa(i).Nachname & "','" & rNa(i).Vorname & "'," & datform(rNa(i).GebDat) & ",'" &  _
   rNa(i).Straße & "','" & rNa(i).KVKStatus & "','" & rNa(i).Geschlecht & "','" & rNa(i).Plz & "','" & rNa(i).Ort & "','" &  _
   rNa(i).Weggeldzone & "'," & datform(rNa(i).AufnDat) & ",'" & rNa(i).LANR & "','" & rNa(i).BStNr & "','" & rNa(i).Titel & "','" & rNa(i).Versichertennummer & "','" &  _
   rNa(i).PrivatTel & "','" & rNa(i).KVNr & "','" & rNa(i).PrivatTel_2 & "','" & rNa(i).PrivatFax & "','" & rNa(i).DienstTel & "','" &  _
   rNa(i).PrivatMobil & "','" & rNa(i).Email & "','" & rNa(i).Arbeitgeber & "'," & cstr(cint(rNa(i).AnAllgda)) & "," & cstr(cint(rNa(i).An1da)) & "," & cstr(cint( _
   rNa(i).An2da)) & "," & cstr(cint(rNa(i).Checkda)) & ",'" & rNa(i).DMTypaD & "'," & datform( 0 ) & "," & rNa(i).absPos & "," &  _
   rNa(i).StByte & ",'" & rNa(i).Cave & "','" & rNa(i).Notiz & "','" & rNa(i).zubenach & "','" & rNa(i).Verwandt & "','" & rNa(i).Sprache & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rNa(i).Pat_ID , "," , rNa(i).lfdnr , ",'" , rNa(i).NVorsatz , "','" , rNa(i).Nachname , "','" , rNa(i).Vorname , "'," , datform( _
   rNa(i).GebDat) , ",'" , rNa(i).Straße , "','" , rNa(i).KVKStatus , "','" , rNa(i).Geschlecht , "','" , rNa(i).Plz , "','" , rNa(i).Ort , "','" ,  _
   rNa(i).Weggeldzone , "'," , datform(rNa(i).AufnDat) , ",'" , rNa(i).LANR , "','" , rNa(i).BStNr , "','" , rNa(i).Titel , "','" , rNa(i).Versichertennummer , "','" ,  _
   rNa(i).PrivatTel , "','" , rNa(i).KVNr , "','" , rNa(i).PrivatTel_2 , "','" , rNa(i).PrivatFax , "','" , rNa(i).DienstTel , "','" ,  _
   rNa(i).PrivatMobil , "','" , rNa(i).Email , "','" , rNa(i).Arbeitgeber , "'," , cstr(cint(rNa(i).AnAllgda)) , "," , cstr(cint(rNa(i).An1da)) , "," , cstr(cint( _
   rNa(i).An2da)) , "," , cstr(cint(rNa(i).Checkda)) , ",'" , rNa(i).DMTypaD , "'," , datform( 0 ) , "," , rNa(i).absPos , "," ,  _
   rNa(i).StByte , ",'" , rNa(i).Cave , "','" , rNa(i).Notiz , "','" , rNa(i).zubenach , "','" , rNa(i).Verwandt , "','" , rNa(i).Sprache , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rNa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "NVorsatz":  If SpMod(rNa(i).NVorsatz, "namen", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rNa(i).Nachname, "namen", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rNa(i).Vorname, "namen", rsc) Then Exit Do
   Case "Straße":  If SpMod(rNa(i).Straße, "namen", rsc) Then Exit Do
   Case "KVKStatus":  If SpMod(rNa(i).KVKStatus, "namen", rsc) Then Exit Do
   Case "Geschlecht":  If SpMod(rNa(i).Geschlecht, "namen", rsc) Then Exit Do
   Case "Plz":  If SpMod(rNa(i).Plz, "namen", rsc) Then Exit Do
   Case "Ort":  If SpMod(rNa(i).Ort, "namen", rsc) Then Exit Do
   Case "Weggeldzone":  If SpMod(rNa(i).Weggeldzone, "namen", rsc) Then Exit Do
   Case "LANR":  If SpMod(rNa(i).LANR, "namen", rsc) Then Exit Do
   Case "BStNr":  If SpMod(rNa(i).BStNr, "namen", rsc) Then Exit Do
   Case "Titel":  If SpMod(rNa(i).Titel, "namen", rsc) Then Exit Do
   Case "Versichertennummer":  If SpMod(rNa(i).Versichertennummer, "namen", rsc) Then Exit Do
   Case "PrivatTel":  If SpMod(rNa(i).PrivatTel, "namen", rsc) Then Exit Do
   Case "KVNr":  If SpMod(rNa(i).KVNr, "namen", rsc) Then Exit Do
   Case "PrivatTel_2":  If SpMod(rNa(i).PrivatTel_2, "namen", rsc) Then Exit Do
   Case "PrivatFax":  If SpMod(rNa(i).PrivatFax, "namen", rsc) Then Exit Do
   Case "DienstTel":  If SpMod(rNa(i).DienstTel, "namen", rsc) Then Exit Do
   Case "PrivatMobil":  If SpMod(rNa(i).PrivatMobil, "namen", rsc) Then Exit Do
   Case "Email":  If SpMod(rNa(i).Email, "namen", rsc) Then Exit Do
   Case "Arbeitgeber":  If SpMod(rNa(i).Arbeitgeber, "namen", rsc) Then Exit Do
   Case "DMTypaD":  If SpMod(rNa(i).DMTypaD, "namen", rsc) Then Exit Do
   Case "Cave":  If SpMod(rNa(i).Cave, "namen", rsc) Then Exit Do
   Case "Notiz":  If SpMod(rNa(i).Notiz, "namen", rsc) Then Exit Do
   Case "zubenach":  If SpMod(rNa(i).zubenach, "namen", rsc) Then Exit Do
   Case "Verwandt":  If SpMod(rNa(i).Verwandt, "namen", rsc) Then Exit Do
   Case "Sprache":  If SpMod(rNa(i).Sprache, "namen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
ElseIf Err.Number = -2147217871 or Err.Number = -2147217859 or err.Number = -2147467259 Then
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere faelle"
 if not Allepat then
  sql = "delete from " & kla & "faelle" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "faelle" & klz & " (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,f4206,ÜwText,f4210," & _
     "statNuller,ÜbwV,AndÜw,Übw,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor," & _
     "ÜWVsw,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr," & _
     "GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,BhFE,s8000,s8100," & _
     "AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,TherArt,StByte,absPos) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "faelle" & klz & " (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,f4206,ÜwText,f4210," & _
     "statNuller,ÜbwV,AndÜw,Übw,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor," & _
     "ÜWVsw,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr," & _
     "GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,BhFE,s8000,s8100," & _
     "AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,TherArt,StByte,absPos) values"))
 For i = 1 to ubound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFa(i).Pat_ID & ",'" & rFa(i).Quartal & "','" & rFa(i).Nachname & "','" & rFa(i).Vorname & "'," & rFa(i).lfdnr & ",'" & rFa(i).TMFNr & "','" &  _
   rFa(i).VKNr & "'," & datform(rFa(i).BhFB) & "," & datform(rFa(i).BhFE1) & "," & datform(rFa(i).BhFE2) & ",'" & rFa(i).f4202 & "'," & datform( _
   rFa(i).ausgst) & ",'" & rFa(i).KtrAbrB & "','" & rFa(i).AbrAr & "'," & datform(rFa(i).lVorl) & ",'" & rFa(i).IK & "','" & rFa(i).KVKs & "','" &  _
   rFa(i).KVKserg & "','" & rFa(i).GebOr & "','" & rFa(i).AbrGb & "','" & rFa(i).f4206 & "','" & rFa(i).ÜwText & "','" & rFa(i).f4210 & "','" &  _
   rFa(i).statNuller & "','" & rFa(i).ÜbwV & "','" & rFa(i).AndÜw & "','" & rFa(i).Übw & "','" & rFa(i).ÜbwLANR & "','" & rFa(i).ÜWZiel & "','" &  _
   rFa(i).ÜWNNr & "','" & rFa(i).ÜWNaN & "','" & rFa(i).ÜWTit & "','" & rFa(i).ÜWVor & "','" & rFa(i).ÜWVsw & "','" & rFa(i).statKlasse & "','" &  _
   rFa(i).f4237 & "'," & rFa(i).statBehTage & ",'" & rFa(i).SchGr & "','" & rFa(i).Weiterbeh & "','" & rFa(i).PGeb & "','" & rFa(i).PGebErg & "','" &  _
   rFa(i).Mahnfrist & "','" & rFa(i).GOÄKatNr & "','" & rFa(i).GOÄKatName & "','" & rFa(i).abrArzt & "','" & rFa(i).privVers & "','" &  _
   rFa(i).AdNam & "','" & rFa(i).AdStr & "','" & rFa(i).AdPlz & "','" & rFa(i).AdOrt & "'," & datform(rFa(i).BhFE) & ",'" & rFa(i).s8000 & "','" &  _
   rFa(i).s8100 & "'," & datform(rFa(i).AktZeit) & "," & datform(rFa(i).Fanf) & ",'" & rFa(i).altQuart & "'," & datform(rFa(i).QAnf) & "," & datform( _
   rFa(i).QEnd) & ",'" & rFa(i).QS & "','" & rFa(i).QT & "'," & rFa(i).TherArt & "," & rFa(i).StByte & "," & rFa(i).absPos & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFa(i).Pat_ID , ",'" , rFa(i).Quartal , "','" , rFa(i).Nachname , "','" , rFa(i).Vorname , "'," , rFa(i).lfdnr , ",'" , rFa(i).TMFNr , "','" ,  _
   rFa(i).VKNr , "'," , datform(rFa(i).BhFB) , "," , datform(rFa(i).BhFE1) , "," , datform(rFa(i).BhFE2) , ",'" , rFa(i).f4202 , "'," , datform( _
   rFa(i).ausgst) , ",'" , rFa(i).KtrAbrB , "','" , rFa(i).AbrAr , "'," , datform(rFa(i).lVorl) , ",'" , rFa(i).IK , "','" , rFa(i).KVKs , "','" ,  _
   rFa(i).KVKserg , "','" , rFa(i).GebOr , "','" , rFa(i).AbrGb , "','" , rFa(i).f4206 , "','" , rFa(i).ÜwText , "','" , rFa(i).f4210 , "','" ,  _
   rFa(i).statNuller , "','" , rFa(i).ÜbwV , "','" , rFa(i).AndÜw , "','" , rFa(i).Übw , "','" , rFa(i).ÜbwLANR , "','" , rFa(i).ÜWZiel , "','" ,  _
   rFa(i).ÜWNNr , "','" , rFa(i).ÜWNaN , "','" , rFa(i).ÜWTit , "','" , rFa(i).ÜWVor , "','" , rFa(i).ÜWVsw , "','" , rFa(i).statKlasse , "','" ,  _
   rFa(i).f4237 , "'," , rFa(i).statBehTage , ",'" , rFa(i).SchGr , "','" , rFa(i).Weiterbeh , "','" , rFa(i).PGeb , "','" , rFa(i).PGebErg , "','" ,  _
   rFa(i).Mahnfrist , "','" , rFa(i).GOÄKatNr , "','" , rFa(i).GOÄKatName , "','" , rFa(i).abrArzt , "','" , rFa(i).privVers , "','" ,  _
   rFa(i).AdNam , "','" , rFa(i).AdStr , "','" , rFa(i).AdPlz , "','" , rFa(i).AdOrt , "'," , datform(rFa(i).BhFE) , ",'" , rFa(i).s8000 , "','" ,  _
   rFa(i).s8100 , "'," , datform(rFa(i).AktZeit) , "," , datform(rFa(i).Fanf) , ",'" , rFa(i).altQuart , "'," , datform(rFa(i).QAnf) , "," , datform( _
   rFa(i).QEnd) , ",'" , rFa(i).QS , "','" , rFa(i).QT , "'," , rFa(i).TherArt , "," , rFa(i).StByte , "," , rFa(i).absPos , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_ID & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & datform(rFa(i).BhFB) & " and bhfe1 = " & datform(rFa(i).BhFE1) & " and ausgst = " & datform(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgabe = "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & rs!FID & vbCrLf & Lese.Ausgabe
    altAusgabe = Lese.ausgabe
    Dim jjj&
    For jjj = 1 To UBound(rEi)
     If rEi(jjj).FID = rFa(i).FID Then
      rEi(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLa)
     If rLa(jjj).FID = rFa(i).FID Then
      rLa(jjj).FID = rs!FID
     End If
    Next jjj
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Quartal":  If SpMod(rFa(i).Quartal, "faelle", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rFa(i).Nachname, "faelle", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rFa(i).Vorname, "faelle", rsc) Then Exit Do
   Case "TMFNr":  If SpMod(rFa(i).TMFNr, "faelle", rsc) Then Exit Do
   Case "VKNr":  If SpMod(rFa(i).VKNr, "faelle", rsc) Then Exit Do
   Case "f4202":  If SpMod(rFa(i).f4202, "faelle", rsc) Then Exit Do
   Case "KtrAbrB":  If SpMod(rFa(i).KtrAbrB, "faelle", rsc) Then Exit Do
   Case "AbrAr":  If SpMod(rFa(i).AbrAr, "faelle", rsc) Then Exit Do
   Case "IK":  If SpMod(rFa(i).IK, "faelle", rsc) Then Exit Do
   Case "KVKs":  If SpMod(rFa(i).KVKs, "faelle", rsc) Then Exit Do
   Case "KVKserg":  If SpMod(rFa(i).KVKserg, "faelle", rsc) Then Exit Do
   Case "GebOr":  If SpMod(rFa(i).GebOr, "faelle", rsc) Then Exit Do
   Case "AbrGb":  If SpMod(rFa(i).AbrGb, "faelle", rsc) Then Exit Do
   Case "f4206":  If SpMod(rFa(i).f4206, "faelle", rsc) Then Exit Do
   Case "ÜwText":  If SpMod(rFa(i).ÜwText, "faelle", rsc) Then Exit Do
   Case "f4210":  If SpMod(rFa(i).f4210, "faelle", rsc) Then Exit Do
   Case "statNuller":  If SpMod(rFa(i).statNuller, "faelle", rsc) Then Exit Do
   Case "ÜbwV":  If SpMod(rFa(i).ÜbwV, "faelle", rsc) Then Exit Do
   Case "AndÜw":  If SpMod(rFa(i).AndÜw, "faelle", rsc) Then Exit Do
   Case "Übw":  If SpMod(rFa(i).Übw, "faelle", rsc) Then Exit Do
   Case "ÜbwLANR":  If SpMod(rFa(i).ÜbwLANR, "faelle", rsc) Then Exit Do
   Case "ÜWZiel":  If SpMod(rFa(i).ÜWZiel, "faelle", rsc) Then Exit Do
   Case "ÜWNNr":  If SpMod(rFa(i).ÜWNNr, "faelle", rsc) Then Exit Do
   Case "ÜWNaN":  If SpMod(rFa(i).ÜWNaN, "faelle", rsc) Then Exit Do
   Case "ÜWTit":  If SpMod(rFa(i).ÜWTit, "faelle", rsc) Then Exit Do
   Case "ÜWVor":  If SpMod(rFa(i).ÜWVor, "faelle", rsc) Then Exit Do
   Case "ÜWVsw":  If SpMod(rFa(i).ÜWVsw, "faelle", rsc) Then Exit Do
   Case "statKlasse":  If SpMod(rFa(i).statKlasse, "faelle", rsc) Then Exit Do
   Case "f4237":  If SpMod(rFa(i).f4237, "faelle", rsc) Then Exit Do
   Case "SchGr":  If SpMod(rFa(i).SchGr, "faelle", rsc) Then Exit Do
   Case "Weiterbeh":  If SpMod(rFa(i).Weiterbeh, "faelle", rsc) Then Exit Do
   Case "PGeb":  If SpMod(rFa(i).PGeb, "faelle", rsc) Then Exit Do
   Case "PGebErg":  If SpMod(rFa(i).PGebErg, "faelle", rsc) Then Exit Do
   Case "Mahnfrist":  If SpMod(rFa(i).Mahnfrist, "faelle", rsc) Then Exit Do
   Case "GOÄKatNr":  If SpMod(rFa(i).GOÄKatNr, "faelle", rsc) Then Exit Do
   Case "GOÄKatName":  If SpMod(rFa(i).GOÄKatName, "faelle", rsc) Then Exit Do
   Case "abrArzt":  If SpMod(rFa(i).abrArzt, "faelle", rsc) Then Exit Do
   Case "privVers":  If SpMod(rFa(i).privVers, "faelle", rsc) Then Exit Do
   Case "AdNam":  If SpMod(rFa(i).AdNam, "faelle", rsc) Then Exit Do
   Case "AdStr":  If SpMod(rFa(i).AdStr, "faelle", rsc) Then Exit Do
   Case "AdPlz":  If SpMod(rFa(i).AdPlz, "faelle", rsc) Then Exit Do
   Case "AdOrt":  If SpMod(rFa(i).AdOrt, "faelle", rsc) Then Exit Do
   Case "s8000":  If SpMod(rFa(i).s8000, "faelle", rsc) Then Exit Do
   Case "s8100":  If SpMod(rFa(i).s8100, "faelle", rsc) Then Exit Do
   Case "altQuart":  If SpMod(rFa(i).altQuart, "faelle", rsc) Then Exit Do
   Case "QS":  If SpMod(rFa(i).QS, "faelle", rsc) Then Exit Do
   Case "QT":  If SpMod(rFa(i).QT, "faelle", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "insert into " & kla & "kassenliste" & klz & "(" & kla & "GO" & klz & "," & kla & "VK" & klz & "," & kla & "IK" & klz & ") values (" & "'" & rFa(I).GOÄKatName & "', '" & rFa(I).VKNr & "', '" & rFa(I).VKNr & "')"
 Call DBCn.Execute(sqlquer, rAf)
 Resume
end if ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere au"
 if not Allepat then
  sql = "delete from " & kla & "au" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "au" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "au" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rAu(i).FID & "," & rAu(i).Pat_ID & "," & datform(rAu(i).ZeitPunkt) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "','" & rAu(i).ICDs & "'," &  _
   rAu(i).absPos & "," & datform(rAu(i).AktZeit) & "," & rAu(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rAu(i).FID , "," , rAu(i).Pat_ID , "," , datform(rAu(i).ZeitPunkt) , ",'" , rAu(i).Beginn , "','" , rAu(i).Ende , "','" , rAu(i).ICDs , "'," ,  _
   rAu(i).absPos , "," , datform(rAu(i).AktZeit) , "," , rAu(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rAu) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Beginn":  If SpMod(rAu(i).Beginn, "au", rsc) Then Exit Do
   Case "Ende":  If SpMod(rAu(i).Ende, "au", rsc) Then Exit Do
   Case "ICDs":  If SpMod(rAu(i).ICDs, "au", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere briefe"
 if not Allepat then
  sql = "delete from " & kla & "briefe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "briefe" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "briefe" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte) values"))
 For i = 1 to ubound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rBr(i).FID & "," & rBr(i).Pat_ID & "," & datform(rBr(i).ZeitPunkt) & ",'" & rBr(i).Pfad & "','" & rBr(i).Art & "','" & rBr(i).Name & "','" &  _
   rBr(i).Typ & "'," & datform(rBr(i).AktZeit) & "," & rBr(i).DokGroe & ",'" & rBr(i).QS & "','" & rBr(i).QT & "'," & rBr(i).absPos & "," &  _
   rBr(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rBr(i).FID , "," , rBr(i).Pat_ID , "," , datform(rBr(i).ZeitPunkt) , ",'" , rBr(i).Pfad , "','" , rBr(i).Art , "','" , rBr(i).Name , "','" ,  _
   rBr(i).Typ , "'," , datform(rBr(i).AktZeit) , "," , rBr(i).DokGroe , ",'" , rBr(i).QS , "','" , rBr(i).QT , "'," , rBr(i).absPos , "," ,  _
   rBr(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rBr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Pfad":  If SpMod(rBr(i).Pfad, "briefe", rsc) Then Exit Do
   Case "Art":  If SpMod(rBr(i).Art, "briefe", rsc) Then Exit Do
   Case "Name":  If SpMod(rBr(i).Name, "briefe", rsc) Then Exit Do
   Case "Typ":  If SpMod(rBr(i).Typ, "briefe", rsc) Then Exit Do
   Case "QS":  If SpMod(rBr(i).QS, "briefe", rsc) Then Exit Do
   Case "QT":  If SpMod(rBr(i).QT, "briefe", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere diagnosen"
 if not Allepat then
  sql = "delete from " & kla & "diagnosen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "diagnosen" & klz & " (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "diagnosen" & klz & " (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,absPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rDi(i).FID & "," & rDi(i).Pat_id & ",'" & rDi(i).GesName & "'," & datform(rDi(i).DiagDatum) & ",'" & rDi(i).DiagSicherheit & "','" &  _
   rDi(i).DiagText & "','" & rDi(i).DiagSeite & "','" & rDi(i).DiagAttr & "','" & rDi(i).ICD & "'," & rDi(i).obDauer & "," & rDi(i).absPos & "," & datform( _
   rDi(i).AktZeit) & "," & rDi(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDi(i).FID , "," , rDi(i).Pat_id , ",'" , rDi(i).GesName , "'," , datform(rDi(i).DiagDatum) , ",'" , rDi(i).DiagSicherheit , "','" ,  _
   rDi(i).DiagText , "','" , rDi(i).DiagSeite , "','" , rDi(i).DiagAttr , "','" , rDi(i).ICD , "'," , rDi(i).obDauer , "," , rDi(i).absPos , "," , datform( _
   rDi(i).AktZeit) , "," , rDi(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "GesName":  If SpMod(rDi(i).GesName, "diagnosen", rsc) Then Exit Do
   Case "DiagSicherheit":  If SpMod(rDi(i).DiagSicherheit, "diagnosen", rsc) Then Exit Do
   Case "DiagText":  If SpMod(rDi(i).DiagText, "diagnosen", rsc) Then Exit Do
   Case "DiagSeite":  If SpMod(rDi(i).DiagSeite, "diagnosen", rsc) Then Exit Do
   Case "DiagAttr":  If SpMod(rDi(i).DiagAttr, "diagnosen", rsc) Then Exit Do
   Case "ICD":  If SpMod(rDi(i).ICD, "diagnosen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere dokumente"
 if not Allepat then
  sql = "delete from " & kla & "dokumente" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "dokumente" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "dokumente" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte) values"))
 For i = 1 to ubound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rDo(i).FID & "," & rDo(i).Pat_ID & "," & datform(rDo(i).ZeitPunkt) & ",'" & rDo(i).DokPfad & "','" & rDo(i).DokArt & "','" & rDo(i).DokName & "'," & datform( _
   rDo(i).Quelldatum) & "," & rDo(i).absPos & "," & datform(rDo(i).AktZeit) & "," & rDo(i).DokGroe & ",'" & rDo(i).QS & "','" &  _
   rDo(i).QT & "'," & rDo(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDo(i).FID , "," , rDo(i).Pat_ID , "," , datform(rDo(i).ZeitPunkt) , ",'" , rDo(i).DokPfad , "','" , rDo(i).DokArt , "','" ,  _
   rDo(i).DokName , "'," , datform(rDo(i).Quelldatum) , "," , rDo(i).absPos , "," , datform(rDo(i).AktZeit) , "," , rDo(i).DokGroe , ",'" , rDo(i).QS , "','" ,  _
   rDo(i).QT , "'," , rDo(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "DokPfad":  If SpMod(rDo(i).DokPfad, "dokumente", rsc) Then Exit Do
   Case "DokArt":  If SpMod(rDo(i).DokArt, "dokumente", rsc) Then Exit Do
   Case "DokName":  If SpMod(rDo(i).DokName, "dokumente", rsc) Then Exit Do
   Case "QS":  If SpMod(rDo(i).QS, "dokumente", rsc) Then Exit Do
   Case "QT":  If SpMod(rDo(i).QT, "dokumente", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere eintraege"
 if not Allepat then
  sql = "delete from " & kla & "eintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "eintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "eintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte) values"))
 For i = 1 to ubound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rEi(i).FID & "," & rEi(i).Pat_ID & "," & datform(rEi(i).ZeitPunkt) & ",'" & rEi(i).Art & "','" & rEi(i).Inhalt & "'," & rEi(i).absPos & "," & datform( _
   rEi(i).AktZeit) & ",'" & rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rEi(i).FID , "," , rEi(i).Pat_ID , "," , datform(rEi(i).ZeitPunkt) , ",'" , rEi(i).Art , "','" , rEi(i).Inhalt , "'," , rEi(i).absPos , "," , datform( _
   rEi(i).AktZeit) , ",'" , rEi(i).QS , "','" , rEi(i).QT , "'," , rEi(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rEi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rEi(i).Art, "eintraege", rsc) Then Exit Do
   Case "Inhalt":  If SpMod(rEi(i).Inhalt, "eintraege", rsc) Then Exit Do
   Case "QS":  If SpMod(rEi(i).QS, "eintraege", rsc) Then Exit Do
   Case "QT":  If SpMod(rEi(i).QT, "eintraege", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere forminhaltform_abk"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "forminhaltform_abk" & klz & " (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "forminhaltform_abk" & klz & " (Form_Abk) values"))
 For i = rFi1 + 1 to ubound(rFi)
'  rFi(i).AktZeit = now()
'   sql = sql0 & "('" & rFi(i).Form_Abk & "')"
  If Lese.dlg.SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rFi(i).Form_Abk , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rFi1 = ubound(rFi)
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFi(i).Form_Abk, "forminhaltform_abk", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere formulare"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "formulare" & klz & " (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "formulare" & klz & " (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values"))
 For i = rFo1 + 1 to ubound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFo(i).FormID & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "','" & rFo(i).FormVorl & "'," & datform(rFo(i).AktZeit) & "," &  _
   rFo(i).absPos & "," & rFo(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFo(i).FormID , ",'" , rFo(i).Form_Abk , "','" , rFo(i).FormBez , "','" , rFo(i).FormVorl , "'," , datform(rFo(i).AktZeit) , "," ,  _
   rFo(i).absPos , "," , rFo(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rFo1 = ubound(rFo)
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFo(i).Form_Abk, "formulare", rsc) Then Exit Do
   Case "FormBez":  If SpMod(rFo(i).FormBez, "formulare", rsc) Then Exit Do
   Case "FormVorl":  If SpMod(rFo(i).FormVorl, "formulare", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere forminhkopf"
 if not Allepat then
'  sql = "delete FROM " & kla & "forminhfeld" & klz & " where foid in (select foid from " & kla & "forminhkopf" & klz & " where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "delete from " & kla & "forminhkopf" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "forminhkopf" & klz & " (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "forminhkopf" & klz & " (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge) values"))
 For i = 1 to ubound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFr(i).FoID & "," & rFr(i).FID & "," & rFr(i).Pat_ID & "," & rFr(i).Form_ID & "," & datform(rFr(i).ZeitPunkt) & "," & rFr(i).AbsPos & "," & datform( _
   rFr(i).AktZeit) & "," & rFr(i).StByte & ",'" & rFr(i).Satzart & "','" & rFr(i).Satzlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFr(i).FoID , "," , rFr(i).FID , "," , rFr(i).Pat_ID , "," , rFr(i).Form_ID , "," , datform(rFr(i).ZeitPunkt) , "," , rFr(i).AbsPos , "," , datform( _
   rFr(i).AktZeit) , "," , rFr(i).StByte , ",'" , rFr(i).Satzart , "','" , rFr(i).Satzlänge , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rFr(i).Satzart, "forminhkopf", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rFr(i).Satzlänge, "forminhkopf", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere forminhfeld"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "forminhfeld" & klz & " (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "forminhfeld" & klz & " (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values"))
 For i = 1 to ubound(rFm)
'  rFm(i).AktZeit = now()
'   sql = sql0 & "(" & rFm(i).FoID & "," & rFm(i).Nr & "," & rFm(i).FeldNr & "," & rFm(i).FeldVW & "," & rFm(i).FeldInhVW & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFm(i).FoID , "," , rFm(i).Nr , "," , rFm(i).FeldNr , "," , rFm(i).FeldVW , "," , rFm(i).FeldInhVW , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFm) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere kheinweis"
 if not Allepat then
  sql = "delete from " & kla & "kheinweis" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "kheinweis" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "kheinweis" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKh(i).FID & "," & rKh(i).Pat_ID & "," & datform(rKh(i).ZeitPunkt) & ",'" & rKh(i).Ziel & "','" & rKh(i).Diagnose & "'," & rKh(i).absPos & "," & datform( _
   rKh(i).AktZeit) & "," & rKh(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKh(i).FID , "," , rKh(i).Pat_ID , "," , datform(rKh(i).ZeitPunkt) , ",'" , rKh(i).Ziel , "','" , rKh(i).Diagnose , "'," , rKh(i).absPos , "," , datform( _
   rKh(i).AktZeit) , "," , rKh(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKh) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Ziel":  If SpMod(rKh(i).Ziel, "kheinweis", rsc) Then Exit Do
   Case "Diagnose":  If SpMod(rKh(i).Diagnose, "kheinweis", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere lbanforderungen"
 if not Allepat then
  sql = "delete from " & kla & "lbanforderungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "lbanforderungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "lbanforderungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLb(i).FID & "," & rLb(i).Pat_ID & "," & datform(rLb(i).ZeitPunkt) & ",'" & rLb(i).AnfText & "'," & rLb(i).absPos & "," & datform( _
   rLb(i).AktZeit) & "," & rLb(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLb(i).FID , "," , rLb(i).Pat_ID , "," , datform(rLb(i).ZeitPunkt) , ",'" , rLb(i).AnfText , "'," , rLb(i).absPos , "," , datform( _
   rLb(i).AktZeit) , "," , rLb(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLb) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AnfText":  If SpMod(rLb(i).AnfText, "lbanforderungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborneu"
 if not Allepat then
  sql = "delete from " & kla & "laborneu" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "laborneu" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborneu" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values"))
 For i = 1 to ubound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLa(i).FID & "," & rLa(i).Pat_ID & "," & datform(rLa(i).ZeitPunkt) & ",'" & rLa(i).FertigStGrad & "','" & rLa(i).Abkü & "'," & rLa(i).LangtextVW & ",'" &  _
   rLa(i).Wert & "','" & rLa(i).Einheit & "'," & rLa(i).KommentarVW & "," & rLa(i).AbsPos & "," & datform(rLa(i).AktZeit) & "," &  _
   rLa(i).Refnr & "," & rLa(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLa(i).FID , "," , rLa(i).Pat_ID , "," , datform(rLa(i).ZeitPunkt) , ",'" , rLa(i).FertigStGrad , "','" , rLa(i).Abkü , "'," ,  _
   rLa(i).LangtextVW , ",'" , rLa(i).Wert , "','" , rLa(i).Einheit , "'," , rLa(i).KommentarVW , "," , rLa(i).AbsPos , "," , datform(rLa(i).AktZeit) , "," ,  _
   rLa(i).Refnr , "," , rLa(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "FertigStGrad":  If SpMod(rLa(i).FertigStGrad, "laborneu", rsc) Then Exit Do
   Case "Abkü":  If SpMod(rLa(i).Abkü, "laborneu", rsc) Then Exit Do
   Case "Wert":  If SpMod(rLa(i).Wert, "laborneu", rsc) Then Exit Do
   Case "Einheit":  If SpMod(rLa(i).Einheit, "laborneu", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere leistungen"
 if not Allepat then
  sql = "delete from " & kla & "leistungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "leistungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "leistungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte) values"))
 For i = 1 to ubound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLe(i).FID & "," & rLe(i).Pat_ID & "," & datform(rLe(i).ZeitPunkt) & ",'" & rLe(i).Leistung & "','" & rLe(i).f5002 & "','" & rLe(i).f5005 & "','" &  _
   rLe(i).f5006 & "','" & rLe(i).f5009 & "','" & rLe(i).Med & "','" & rLe(i).f5015 & "','" & rLe(i).f5016 & "','" & rLe(i).f5021 & "','" &  _
   rLe(i).f5026 & "','" & rLe(i).Faktor & "','" & rLe(i).f5098 & "','" & rLe(i).LANR & "'," & datform(rLe(i).letzVorg) & ",'" & rLe(i).Ausn & "','" &  _
   rLe(i).Beme & "'," & rLe(i).absPos & "," & datform(rLe(i).AktZeit) & ",'" & rLe(i).QS & "','" & rLe(i).QT & "'," & rLe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLe(i).FID , "," , rLe(i).Pat_ID , "," , datform(rLe(i).ZeitPunkt) , ",'" , rLe(i).Leistung , "','" , rLe(i).f5002 , "','" ,  _
   rLe(i).f5005 , "','" , rLe(i).f5006 , "','" , rLe(i).f5009 , "','" , rLe(i).Med , "','" , rLe(i).f5015 , "','" , rLe(i).f5016 , "','" , rLe(i).f5021 , "','" ,  _
   rLe(i).f5026 , "','" , rLe(i).Faktor , "','" , rLe(i).f5098 , "','" , rLe(i).LANR , "'," , datform(rLe(i).letzVorg) , ",'" , rLe(i).Ausn , "','" ,  _
   rLe(i).Beme , "'," , rLe(i).absPos , "," , datform(rLe(i).AktZeit) , ",'" , rLe(i).QS , "','" , rLe(i).QT , "'," , rLe(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Leistung":  If SpMod(rLe(i).Leistung, "leistungen", rsc) Then Exit Do
   Case "f5002":  If SpMod(rLe(i).f5002, "leistungen", rsc) Then Exit Do
   Case "f5005":  If SpMod(rLe(i).f5005, "leistungen", rsc) Then Exit Do
   Case "f5006":  If SpMod(rLe(i).f5006, "leistungen", rsc) Then Exit Do
   Case "f5009":  If SpMod(rLe(i).f5009, "leistungen", rsc) Then Exit Do
   Case "Med":  If SpMod(rLe(i).Med, "leistungen", rsc) Then Exit Do
   Case "f5015":  If SpMod(rLe(i).f5015, "leistungen", rsc) Then Exit Do
   Case "f5016":  If SpMod(rLe(i).f5016, "leistungen", rsc) Then Exit Do
   Case "f5021":  If SpMod(rLe(i).f5021, "leistungen", rsc) Then Exit Do
   Case "f5026":  If SpMod(rLe(i).f5026, "leistungen", rsc) Then Exit Do
   Case "Faktor":  If SpMod(rLe(i).Faktor, "leistungen", rsc) Then Exit Do
   Case "f5098":  If SpMod(rLe(i).f5098, "leistungen", rsc) Then Exit Do
   Case "LANR":  If SpMod(rLe(i).LANR, "leistungen", rsc) Then Exit Do
   Case "Ausn":  If SpMod(rLe(i).Ausn, "leistungen", rsc) Then Exit Do
   Case "Beme":  If SpMod(rLe(i).Beme, "leistungen", rsc) Then Exit Do
   Case "QS":  If SpMod(rLe(i).QS, "leistungen", rsc) Then Exit Do
   Case "QT":  If SpMod(rLe(i).QT, "leistungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere medplan"
 if not Allepat then
  sql = "delete from " & kla & "medplan" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "medplan" & klz & " (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "medplan" & klz & " (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rMe(i).FID & "," & rMe(i).Pat_ID & "," & rMe(i).MPNr & "," & datform(rMe(i).ZeitPunkt) & "," & datform(rMe(i).Datum) & ",'" & rMe(i).Medikament & "','" &  _
   rMe(i).MedAnfang & "'," & rMe(i).FeldNr & ",'" & rMe(i).mo & "','" & rMe(i).mi & "','" & rMe(i).nm & "','" & rMe(i).ab & "','" &  _
   rMe(i).zn & "'," & cstr(cint(rMe(i).bBed)) & ",'" & rMe(i).Bemerkung & "'," & rMe(i).AbsPos & "," & datform(rMe(i).AktZeit) & "," & rMe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rMe(i).FID , "," , rMe(i).Pat_ID , "," , rMe(i).MPNr , "," , datform(rMe(i).ZeitPunkt) , "," , datform(rMe(i).Datum) , ",'" ,  _
   rMe(i).Medikament , "','" , rMe(i).MedAnfang , "'," , rMe(i).FeldNr , ",'" , rMe(i).mo , "','" , rMe(i).mi , "','" , rMe(i).nm , "','" , rMe(i).ab , "','" ,  _
   rMe(i).zn , "'," , cstr(cint(rMe(i).bBed)) , ",'" , rMe(i).Bemerkung , "'," , rMe(i).AbsPos , "," , datform(rMe(i).AktZeit) , "," ,  _
   rMe(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rMe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Medikament":  If SpMod(rMe(i).Medikament, "medplan", rsc) Then Exit Do
   Case "MedAnfang":  If SpMod(rMe(i).MedAnfang, "medplan", rsc) Then Exit Do
   Case "mo":  If SpMod(rMe(i).mo, "medplan", rsc) Then Exit Do
   Case "mi":  If SpMod(rMe(i).mi, "medplan", rsc) Then Exit Do
   Case "nm":  If SpMod(rMe(i).nm, "medplan", rsc) Then Exit Do
   Case "ab":  If SpMod(rMe(i).ab, "medplan", rsc) Then Exit Do
   Case "zn":  If SpMod(rMe(i).zn, "medplan", rsc) Then Exit Do
   Case "Bemerkung":  If SpMod(rMe(i).Bemerkung, "medplan", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere rezepteintraege"
 if not Allepat then
  sql = "delete from " & kla & "rezepteintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "rezepteintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "rezepteintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte) values"))
 For i = 1 to ubound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRe(i).FID & "," & rRe(i).Pat_ID & "," & datform(rRe(i).ZeitPunkt) & ",'" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "','" &  _
   rRe(i).Medikament & "','" & rRe(i).PZN & "'," & rRe(i).absPos & "," & datform(rRe(i).AktZeit) & ",'" & rRe(i).QS & "','" & rRe(i).QT & "'," & rRe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRe(i).FID , "," , rRe(i).Pat_ID , "," , datform(rRe(i).ZeitPunkt) , ",'" , rRe(i).Rezept , "','" , rRe(i).Rezeptklasse , "','" ,  _
   rRe(i).Medikament , "','" , rRe(i).PZN , "'," , rRe(i).absPos , "," , datform(rRe(i).AktZeit) , ",'" , rRe(i).QS , "','" , rRe(i).QT , "'," ,  _
   rRe(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Rezept":  If SpMod(rRe(i).Rezept, "rezepteintraege", rsc) Then Exit Do
   Case "Rezeptklasse":  If SpMod(rRe(i).Rezeptklasse, "rezepteintraege", rsc) Then Exit Do
   Case "Medikament":  If SpMod(rRe(i).Medikament, "rezepteintraege", rsc) Then Exit Do
   Case "PZN":  If SpMod(rRe(i).PZN, "rezepteintraege", rsc) Then Exit Do
   Case "QS":  If SpMod(rRe(i).QS, "rezepteintraege", rsc) Then Exit Do
   Case "QT":  If SpMod(rRe(i).QT, "rezepteintraege", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere rr"
 if not Allepat then
  sql = "delete from " & kla & "rr" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "rr" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "rr" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte) values"))
 For i = 1 to ubound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRr(i).FID & "," & rRr(i).Pat_ID & "," & datform(rRr(i).ZeitPunkt) & ",'" & rRr(i).RR & "'," & rRr(i).absPos & "," & datform(rRr(i).AktZeit) & "," &  _
   rRr(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRr(i).FID , "," , rRr(i).Pat_ID , "," , datform(rRr(i).ZeitPunkt) , ",'" , rRr(i).RR , "'," , rRr(i).absPos , "," , datform( _
   rRr(i).AktZeit) , "," , rRr(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "RR":  If SpMod(rRr(i).RR, "rr", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere kvnrue"
 if not Allepat then
  sql = "delete from " & kla & "kvnrue" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "kvnrue" & klz & " (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "kvnrue" & klz & " (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values"))
 For i = 1 to ubound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKv(i).Pat_ID & ",'" & rKv(i).KVNr & "'," & rKv(i).absPos & "," & datform(rKv(i).AktZeit) & "," & rKv(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKv(i).Pat_ID , ",'" , rKv(i).KVNr , "'," , rKv(i).absPos , "," , datform(rKv(i).AktZeit) , "," , rKv(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKv) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "KVNr":  If SpMod(rKv(i).KVNr, "kvnrue", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error resume next
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere unbekannte_kennungen"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "unbekannte kennungen" & klz & " (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "unbekannte kennungen" & klz & " (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values"))
 For i = rUn1 + 1 to ubound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
'   sql = sql0 & "('" & rUn(i).Kennung & "'," & rUn(i).absPos & "," & rUn(i).StByte & "," & rUn(i).Pat_id & ",'" & rUn(i).Inhalt & "')"
  If Lese.dlg.SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rUn(i).Kennung , "'," , rUn(i).absPos , "," , rUn(i).StByte , "," , rUn(i).Pat_id , ",'" , rUn(i).Inhalt , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rUn) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rUn) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rUn1 = ubound(rUn)
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Kennung":  If SpMod(rUn(i).Kennung, "unbekannte kennungen", rsc) Then Exit Do
   Case "Inhalt":  If SpMod(rUn(i).Inhalt, "unbekannte kennungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere dmpreihe"
 if not Allepat then
  sql = "delete from " & kla & "dmpreihe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_ID)
  Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into " & kla & "dmpreihe" & klz & " (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "dmpreihe" & klz & " (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit) values"))
 For i = 1 to ubound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
'   sql = sql0 & "('" & rDm(i).Abk & "','" & rDm(i).Art & "'," & datform(rDm(i).KarteiDatum) & "," & datform(rDm(i).exportiert) & "," & datform(rDm(i).DokuDatum) & "," & cstr(cint( _
   rDm(i).obvoll)) & ",'" & rDm(i).NachName & "','" & rDm(i).VorName & "'," & datform(rDm(i).GebDat) & "," & rDm(i).Pat_id & "," &  _
   rDm(i).StByte & "," & datform(rDm(i).AktZeit) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rDm(i).Abk , "','" , rDm(i).Art , "'," , datform(rDm(i).KarteiDatum) , "," , datform(rDm(i).exportiert) , "," , datform(rDm(i).DokuDatum) , "," , cstr(cint( _
   rDm(i).obvoll)) , ",'" , rDm(i).NachName , "','" , rDm(i).VorName , "'," , datform(rDm(i).GebDat) , "," , rDm(i).Pat_id , "," ,  _
   rDm(i).StByte , "," , datform(rDm(i).AktZeit) , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDm) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 DBCn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abk":  If SpMod(rDm(i).Abk, "dmpreihe", rsc) Then Exit Do
   Case "Art":  If SpMod(rDm(i).Art, "dmpreihe", rsc) Then Exit Do
   Case "NachName":  If SpMod(rDm(i).NachName, "dmpreihe", rsc) Then Exit Do
   Case "VorName":  If SpMod(rDm(i).VorName, "dmpreihe", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 DBCn.BeginTrans
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function laborxsaetzeSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxsaetze" & klz & " (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,PLZPraxis,OrtPraxis,Labor,StraßeLabor,PLZLabor," & _
     "OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxsaetze" & klz & " (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,PLZPraxis,OrtPraxis,Labor,StraßeLabor,PLZLabor," & _
     "OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values"))
 For i = 0 to ubound(rLs)
'  rLs(i).AktZeit = now()
'   sql = sql0 & "(" & rLs(i).DatID & ",'" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).VersionSatzb & "','" &  _
   rLs(i).Arztnr & "','" & rLs(i).Arztname & "','" & rLs(i).StraßePraxis & "','" & rLs(i).PLZPraxis & "','" & rLs(i).OrtPraxis & "','" & rLs(i).Labor & "','" &  _
   rLs(i).StraßeLabor & "','" & rLs(i).PLZLabor & "','" & rLs(i).OrtLabor & "','" & rLs(i).KBVPrüfnr & "','" & rLs(i).Zeichensatz & "','" &  _
   rLs(i).Kundenarztnr & "','" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLs(i).DatID , ",'" , rLs(i).Satzart , "','" , rLs(i).Satzlänge , "','" , rLs(i).SatzlängeSchluss , "','" , rLs(i).VersionSatzb , "','" ,  _
   rLs(i).Arztnr , "','" , rLs(i).Arztname , "','" , rLs(i).StraßePraxis , "','" , rLs(i).PLZPraxis , "','" , rLs(i).OrtPraxis , "','" ,  _
   rLs(i).Labor , "','" , rLs(i).StraßeLabor , "','" , rLs(i).PLZLabor , "','" , rLs(i).OrtLabor , "','" , rLs(i).KBVPrüfnr , "','" , rLs(i).Zeichensatz , "','" ,  _
   rLs(i).Kundenarztnr , "','" , rLs(i).Erstellungsdatum , "','" , rLs(i).Gesamtlänge , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLs) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rLs(i).Satzart, "laborxsaetze", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLs(i).Satzlänge, "laborxsaetze", rsc) Then Exit Do
   Case "SatzlängeSchluss":  If SpMod(rLs(i).SatzlängeSchluss, "laborxsaetze", rsc) Then Exit Do
   Case "VersionSatzb":  If SpMod(rLs(i).VersionSatzb, "laborxsaetze", rsc) Then Exit Do
   Case "Arztnr":  If SpMod(rLs(i).Arztnr, "laborxsaetze", rsc) Then Exit Do
   Case "Arztname":  If SpMod(rLs(i).Arztname, "laborxsaetze", rsc) Then Exit Do
   Case "StraßePraxis":  If SpMod(rLs(i).StraßePraxis, "laborxsaetze", rsc) Then Exit Do
   Case "PLZPraxis":  If SpMod(rLs(i).PLZPraxis, "laborxsaetze", rsc) Then Exit Do
   Case "OrtPraxis":  If SpMod(rLs(i).OrtPraxis, "laborxsaetze", rsc) Then Exit Do
   Case "Labor":  If SpMod(rLs(i).Labor, "laborxsaetze", rsc) Then Exit Do
   Case "StraßeLabor":  If SpMod(rLs(i).StraßeLabor, "laborxsaetze", rsc) Then Exit Do
   Case "PLZLabor":  If SpMod(rLs(i).PLZLabor, "laborxsaetze", rsc) Then Exit Do
   Case "OrtLabor":  If SpMod(rLs(i).OrtLabor, "laborxsaetze", rsc) Then Exit Do
   Case "KBVPrüfnr":  If SpMod(rLs(i).KBVPrüfnr, "laborxsaetze", rsc) Then Exit Do
   Case "Zeichensatz":  If SpMod(rLs(i).Zeichensatz, "laborxsaetze", rsc) Then Exit Do
   Case "Kundenarztnr":  If SpMod(rLs(i).Kundenarztnr, "laborxsaetze", rsc) Then Exit Do
   Case "Erstellungsdatum":  If SpMod(rLs(i).Erstellungsdatum, "laborxsaetze", rsc) Then Exit Do
   Case "Gesamtlänge":  If SpMod(rLs(i).Gesamtlänge, "laborxsaetze", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxeingel" & klz & " (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxeingel" & klz & " (Pfad,Name,Zp," & _
     "fertig) values"))
 For i = 1 to ubound(rLg)
'  rLg(i).AktZeit = now()
'   sql = sql0 & "('" & rLg(i).Pfad & "','" & rLg(i).Name & "'," & datform(rLg(i).Zp) & "," & cstr(cint(rLg(i).fertig)) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rLg(i).Pfad , "','" , rLg(i).Name , "'," , datform(rLg(i).Zp) , "," , cstr(cint(rLg(i).fertig)) , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLg) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Pfad":  If SpMod(rLg(i).Pfad, "laborxeingel", rsc) Then Exit Do
   Case "Name":  If SpMod(rLg(i).Name, "laborxeingel", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxus"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxus" & klz & " (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxus" & klz & " (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN) values"))
 For i = j to j
'  rLu(i).AktZeit = now()
'   sql = sql0 & "(" & rLu(i).DatID & "," & rLu(i).SatzID & ",'" & rLu(i).Satzart & "','" & rLu(i).Satzlänge & "','" & rLu(i).Auftragsnummer & "','" & rLu(i).Auftragsschlüssel & "'," & datform( _
   rLu(i).Eingang) & ",'" & rLu(i).Berichtsdatum & "'," & rLu(i).Pat_id & ",'" & rLu(i).Nachname & "','" &  _
   rLu(i).Vorname & "','" & rLu(i).GebDat & "','" & rLu(i).Titel & "','" & rLu(i).NVorsatz & "','" & rLu(i).BefArt & "','" & rLu(i).Abrechnungstyp & "','" &  _
   rLu(i).GebüOrd & "','" & rLu(i).Patienteninformation & "','" & rLu(i).Geschlecht & "','" & rLu(i).AuftrHinw & "','" & rLu(i).Pat_idUrsp & "','" &  _
   rLu(i).Pat_idErwVNG & "','" & rLu(i).Pat_idErwVN & "','" & rLu(i).Pat_idErwG & "','" & rLu(i).Pat_idErwGB & "','" & rLu(i).Pat_idErwGL & "','" &  _
   rLu(i).Pat_idLaborNeu & "'," & datform(rLu(i).ZeitpunktLaborneu) & "," & rLu(i).ZdüP & "," & rLu(i).ZdiP & ",'" & rLu(i).LWerte & "'," & datform( _
   rLu(i).verglichen) & "," & rLu(i).AfN & ")"
  If Lese.dlg.SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLu(i).DatID , "," , rLu(i).SatzID , ",'" , rLu(i).Satzart , "','" , rLu(i).Satzlänge , "','" , rLu(i).Auftragsnummer , "','" ,  _
   rLu(i).Auftragsschlüssel , "'," , datform(rLu(i).Eingang) , ",'" , rLu(i).Berichtsdatum , "'," , rLu(i).Pat_id , ",'" , rLu(i).Nachname , "','" ,  _
   rLu(i).Vorname , "','" , rLu(i).GebDat , "','" , rLu(i).Titel , "','" , rLu(i).NVorsatz , "','" , rLu(i).BefArt , "','" , rLu(i).Abrechnungstyp , "','" ,  _
   rLu(i).GebüOrd , "','" , rLu(i).Patienteninformation , "','" , rLu(i).Geschlecht , "','" , rLu(i).AuftrHinw , "','" , rLu(i).Pat_idUrsp , "','" ,  _
   rLu(i).Pat_idErwVNG , "','" , rLu(i).Pat_idErwVN , "','" , rLu(i).Pat_idErwG , "','" , rLu(i).Pat_idErwGB , "','" , rLu(i).Pat_idErwGL , "','" ,  _
   rLu(i).Pat_idLaborNeu , "'," , datform(rLu(i).ZeitpunktLaborneu) , "," , rLu(i).ZdüP , "," , rLu(i).ZdiP , ",'" , rLu(i).LWerte , "'," , datform( _
   rLu(i).verglichen) , "," , rLu(i).AfN , ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLu) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rLu(i).Satzart, "laborxus", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLu(i).Satzlänge, "laborxus", rsc) Then Exit Do
   Case "Auftragsnummer":  If SpMod(rLu(i).Auftragsnummer, "laborxus", rsc) Then Exit Do
   Case "Auftragsschlüssel":  If SpMod(rLu(i).Auftragsschlüssel, "laborxus", rsc) Then Exit Do
   Case "Berichtsdatum":  If SpMod(rLu(i).Berichtsdatum, "laborxus", rsc) Then Exit Do
   Case "Nachname":  If SpMod(rLu(i).Nachname, "laborxus", rsc) Then Exit Do
   Case "Vorname":  If SpMod(rLu(i).Vorname, "laborxus", rsc) Then Exit Do
   Case "GebDat":  If SpMod(rLu(i).GebDat, "laborxus", rsc) Then Exit Do
   Case "Titel":  If SpMod(rLu(i).Titel, "laborxus", rsc) Then Exit Do
   Case "NVorsatz":  If SpMod(rLu(i).NVorsatz, "laborxus", rsc) Then Exit Do
   Case "BefArt":  If SpMod(rLu(i).BefArt, "laborxus", rsc) Then Exit Do
   Case "Abrechnungstyp":  If SpMod(rLu(i).Abrechnungstyp, "laborxus", rsc) Then Exit Do
   Case "GebüOrd":  If SpMod(rLu(i).GebüOrd, "laborxus", rsc) Then Exit Do
   Case "Patienteninformation":  If SpMod(rLu(i).Patienteninformation, "laborxus", rsc) Then Exit Do
   Case "Geschlecht":  If SpMod(rLu(i).Geschlecht, "laborxus", rsc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLu(i).AuftrHinw, "laborxus", rsc) Then Exit Do
   Case "Pat_idUrsp":  If SpMod(rLu(i).Pat_idUrsp, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwVNG":  If SpMod(rLu(i).Pat_idErwVNG, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwVN":  If SpMod(rLu(i).Pat_idErwVN, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwG":  If SpMod(rLu(i).Pat_idErwG, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwGB":  If SpMod(rLu(i).Pat_idErwGB, "laborxus", rsc) Then Exit Do
   Case "Pat_idErwGL":  If SpMod(rLu(i).Pat_idErwGL, "laborxus", rsc) Then Exit Do
   Case "Pat_idLaborNeu":  If SpMod(rLu(i).Pat_idLaborNeu, "laborxus", rsc) Then Exit Do
   Case "LWerte":  If SpMod(rLu(i).LWerte, "laborxus", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxbakt" & klz & " (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxbakt" & klz & " (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values"))
 For i = 1 to ubound(rLo)
'  rLo(i).AktZeit = now()
'   sql = sql0 & "(" & rLo(i).RefNr & ",'" & rLo(i).Verf & "','" & rLo(i).KuQu & "','" & rLo(i).Quelle & "','" & rLo(i).QSpez & "'," & datform(rLo(i).AbnDat) & ",'" &  _
   rLo(i).Kommentar & "','" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLo(i).RefNr , ",'" , rLo(i).Verf , "','" , rLo(i).KuQu , "','" , rLo(i).Quelle , "','" , rLo(i).QSpez , "'," , datform(rLo(i).AbnDat) , ",'" ,  _
   rLo(i).Kommentar , "','" , rLo(i).Erklärung , "','" , rLo(i).Keimzahl , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Verf":  If SpMod(rLo(i).Verf, "laborxbakt", rsc) Then Exit Do
   Case "KuQu":  If SpMod(rLo(i).KuQu, "laborxbakt", rsc) Then Exit Do
   Case "Quelle":  If SpMod(rLo(i).Quelle, "laborxbakt", rsc) Then Exit Do
   Case "QSpez":  If SpMod(rLo(i).QSpez, "laborxbakt", rsc) Then Exit Do
   Case "Kommentar":  If SpMod(rLo(i).Kommentar, "laborxbakt", rsc) Then Exit Do
   Case "Erklärung":  If SpMod(rLo(i).Erklärung, "laborxbakt", rsc) Then Exit Do
   Case "Keimzahl":  If SpMod(rLo(i).Keimzahl, "laborxbakt", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxwert" & klz & " (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxwert" & klz & " (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw) values"))
 For i = 1 to ubound(rLw)
'  rLw(i).AktZeit = now()
'   sql = sql0 & "(" & rLw(i).RefNr & ",'" & rLw(i).Abkü & "','" & rLw(i).Langname & "','" & rLw(i).Quelle & "','" & rLw(i).QSpez & "'," & datform(rLw(i).AbnDat) & ",'" &  _
   rLw(i).Wert & "','" & rLw(i).Einheit & "','" & rLw(i).Grenzwerti & "','" & rLw(i).Kommentar & "','" & rLw(i).Teststatus & "','" &  _
   rLw(i).Erklärung & "','" & rLw(i).Normbereich & "','" & rLw(i).NormU & "','" & rLw(i).NormO & "','" & rLw(i).AuftrHinw & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLw(i).RefNr , ",'" , rLw(i).Abkü , "','" , rLw(i).Langname , "','" , rLw(i).Quelle , "','" , rLw(i).QSpez , "'," , datform(rLw(i).AbnDat) , ",'" ,  _
   rLw(i).Wert , "','" , rLw(i).Einheit , "','" , rLw(i).Grenzwerti , "','" , rLw(i).Kommentar , "','" , rLw(i).Teststatus , "','" ,  _
   rLw(i).Erklärung , "','" , rLw(i).Normbereich , "','" , rLw(i).NormU , "','" , rLw(i).NormO , "','" , rLw(i).AuftrHinw , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLw) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLw(i).Abkü, "laborxwert", rsc) Then Exit Do
   Case "Langname":  If SpMod(rLw(i).Langname, "laborxwert", rsc) Then Exit Do
   Case "Quelle":  If SpMod(rLw(i).Quelle, "laborxwert", rsc) Then Exit Do
   Case "QSpez":  If SpMod(rLw(i).QSpez, "laborxwert", rsc) Then Exit Do
   Case "Wert":  If SpMod(rLw(i).Wert, "laborxwert", rsc) Then Exit Do
   Case "Einheit":  If SpMod(rLw(i).Einheit, "laborxwert", rsc) Then Exit Do
   Case "Grenzwerti":  If SpMod(rLw(i).Grenzwerti, "laborxwert", rsc) Then Exit Do
   Case "Kommentar":  If SpMod(rLw(i).Kommentar, "laborxwert", rsc) Then Exit Do
   Case "Teststatus":  If SpMod(rLw(i).Teststatus, "laborxwert", rsc) Then Exit Do
   Case "Erklärung":  If SpMod(rLw(i).Erklärung, "laborxwert", rsc) Then Exit Do
   Case "Normbereich":  If SpMod(rLw(i).Normbereich, "laborxwert", rsc) Then Exit Do
   Case "NormU":  If SpMod(rLw(i).NormU, "laborxwert", rsc) Then Exit Do
   Case "NormO":  If SpMod(rLw(i).NormO, "laborxwert", rsc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLw(i).AuftrHinw, "laborxwert", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern
 Dim i%, rAF& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 on error goto fehler
 syscmd 4, CStr(rNa(0).Pat_ID) & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into " & kla & "laborxleist" & klz & " (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into ", kla, "laborxleist" & klz & " (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl) values"))
 For i = 1 to ubound(rLL)
'  rLL(i).AktZeit = now()
'   sql = sql0 & "(" & rLL(i).RefNr & ",'" & rLL(i).Abkü & "','" & rLL(i).Verf & "','" & rLL(i).EBM & "','" & rLL(i).goä & "','" & rLL(i).Anzahl & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLL(i).RefNr , ",'" , rLL(i).Abkü , "','" , rLL(i).Verf , "','" , rLL(i).EBM , "','" , rLL(i).goä , "','" , rLL(i).Anzahl , "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLL) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLL) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  if lese.dlg.sammelinsert = 0 then csql.m_len = 0
  If lies.obmysql and obmitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If not rs.BOF() then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 exit function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abkü":  If SpMod(rLL(i).Abkü, "laborxleist", rsc) Then Exit Do
   Case "Verf":  If SpMod(rLL(i).Verf, "laborxleist", rsc) Then Exit Do
   Case "EBM":  If SpMod(rLL(i).EBM, "laborxleist", rsc) Then Exit Do
   Case "goä":  If SpMod(rLL(i).goä, "laborxleist", rsc) Then Exit Do
   Case "Anzahl":  If SpMod(rLL(i).Anzahl, "laborxleist", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function doSpeichern(frm As Lese)
 Dim rAf&
 on error goto fehler
 call namenSpeichern
 call faelleSpeichern
   if not lies.obmysql then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   end if
 call auSpeichern
 call briefeSpeichern
 call diagnosenSpeichern
 call dokumenteSpeichern
 call eintraegeSpeichern
 call forminhaltform_abkSpeichern
   if not lies.obmysql then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   end if
 call formulareSpeichern
 call forminhkopfSpeichern
 call forminhfeldSpeichern
 call kheinweisSpeichern
 call lbanforderungenSpeichern
 call laborneuSpeichern
 call leistungenSpeichern
 call medplanSpeichern
 call rezepteintraegeSpeichern
 call rrSpeichern
 call kvnrueSpeichern
 call unbekannte_kennungenSpeichern
 call dmpreiheSpeichern
 Call DBCn.Execute("update namen set aktZeit = " & datform(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_ID,rAf)
 If rAf <> 1 Then 
  frm.Ausgabe = "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_ID & " " & rNa(0).Nachname & " " & rNa(0).Vorname & vbcrfl & altausgabe
  altausgabe = frm.ausgabe
 end if
 exit function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

