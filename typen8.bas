Attribute VB_Name = "Typen"
Dim sql$

Public Type namen
 Pat_id As Long 'Pat_ID 3 '3000
 lfdnr As Long 'lfdnr 3 'laufende Patientennummer
 NVorsatz As String 'NVorsatz 129 '3100
 NachName As String 'Nachname 129 '3101
 VorName As String 'Vorname 129 '3102
 GebDat As Date 'GebDat 135 '3103
 Straße As String 'Straße 129 '3107
 KVKStatus As String 'KVKStatus 129 '3108
 Geschlecht As String 'Geschlecht 129 '3110
 Plz As String 'Plz 129 '3112
 Ort As String 'Ort 129 '3113
 Weggeldzone As String 'Weggeldzone 129 '3631
 AufnDat As Date 'AufnDat 135 '3610
 intZOGP As String 'intZOGP 129 '3635, interne Zuordnung Arzt bei GP
 Titel As String 'Titel 129 '3104
 Versichertennummer As String 'Versichertennummer 129 '3105
 PrivatTel As String 'PrivatTel 129 '3629
 KVNr As String 'KVNr 129 '3630
 PrivatTel_2 As String 'PrivatTel_2 129 '3629
 PrivatFax As String 'PrivatFax 129 '3629
 DienstTel As String 'DienstTel 129 '3629
 PrivatMobil As String 'PrivatMobil 129 '3629
 Email As String 'Email 129 'Email
 AnAllgda As Integer 'AnAllgda 11 'Anamnese allgemein da
 An1da As Integer 'An1da 11 'Anamnese S.1 da
 An2da As Integer 'An2da 11 'Anamnese S.2 da
 Checkda As Integer 'Checkda 11 'Checkliste da
 DMTypaD As String 'DMTypaD 129 'aus Diagnosen
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Cave As String 'Cave 129 '3654
 notiz As String 'Notiz 129 '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>
 zubenach As String 'zubenach 129 '3633
 Verwandt As String 'Verwandt 129 '
End Type

Public Type Faelle
 FID As Long 'FID 3 'Primärschlüssel
 Pat_id As Long 'Pat_ID 3 '3000
 quartal As String 'Quartal 129 '4101
 NachName As String 'Nachname 129 '3101
 VorName As String 'Vorname 129 '3102
 lfdnr As Long 'lfdnr 3 'laufende Fallnummer
 TMFNr As String 'TMFNr 129 '4144 Fallnummer in Turbomed
 VKNr As String 'VKNr 129 '4104
 BhFB As Date 'BhFB 135 '4150
 BhFE1 As Date 'BhFE1 135 '4151
 BhFE2 As Date 'BhFE2 135 '4152
 f4202 As String 'f4202 129 '4202
 ausgst As Date 'ausgst 135 '4102 ('ausgestellt am')
 KtrAbrB As String 'KtrAbrB 129 '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr As String 'AbrAr 129 '4107, Abrechnungsart (1 = Primärkassen)
 lVorl As Date 'lVorl 135 '4109, letzte Vorlage
 IK As String 'IK 129 '4111 (auch patientenbezogen)
 KVKs As String 'KVKs 129 '4112
 KVKserg As String 'KVKserg 129 '4113
 GebOr As String 'GebOr 129 '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As String 'AbrGb 129 '4122, Abrechnungsgebiet (07 = Diabetes)
 f4206 As String 'f4206 129 '4206, ?
 ÜwText As String 'ÜwText 129 '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 As String 'f4210 129 '4210, unbekannt
 statNuller As String 'statNuller 129 '4216, nu bei Musterfrau 16 Nuller
 ÜbwV As String 'ÜbwV 129 '4218, überwiesen von
 AndÜw As String 'AndÜw 129 '4219, anderer Überweiser
 Übw As String 'Übw 129 '4218 oder 4219, je nachdem, was befüllt
 ÜWZiel As String 'ÜWZiel 129 '4220
 ÜWNNr As String 'ÜWNNr 129 '4231(4): KV-Nummer des Überweisers
 ÜWNaN As String 'ÜWNaN 129 '4231(3): Nachname des Überweisers
 ÜWTit As String 'ÜWTit 129 '4231(3): Titel des Überweisers
 ÜWVor As String 'ÜWVor 129 '
 ÜWVsw As String 'ÜWVsw 129 '4231(2b): Vorsatzwort des Überweisers
 statklasse As String 'statklasse 129 '4236
 f4237 As String 'f4237 129 '4237, ? (nur bei Musterw)
 statBehTage As Long 'statBehTage 3 '4238
 SchGr As String 'SchGr 129 '4239, Scheingruppe
 Weiterbeh As String 'Weiterbeh 129 '4243, Weiterbehandelnder
 PGeb As String 'PGeb 129 '4401, Praxisgebühr
 PGebErg As String 'PGebErg 129 '4402, ?
 Mahnfrist As String 'Mahnfrist 129 '4403, Mahnfrist bis
 GOÄKatNr As String 'GOÄKatNr 129 '4580 (1): Katalog-Nummer
 GOÄKatName As String 'GOÄKatName 129 '4580 (2): Privat-Abrechnungskatalog
 AdNam As String 'AdNam 129 '
 AdStr As String 'AdStr 129 '4602(2) Straße bei Musterfrau
 AdPlz As String 'AdPlz 129 '4602(3) PLZ bei Musterfrau
 AdOrt As String 'AdOrt 129 '4602(4) Ort bei Musterfrau
 BhFE As Date 'BhFE 135 '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As String 's8000 129 '8000, ???
 s8100 As String 's8100 129 '8100, ???
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 Fanf As Date 'Fanf 135 'Fallanfang
 altQuart As String 'altQuart 129 '
 QAnf As Date 'QAnf 135 'Quartalsanfang
 QEnd As Date 'QEnd 135 'Quartalsende
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
End Type

Public Type au
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 Beginn As String 'Beginn 129 '6285 1. Hälfte
 Ende As String 'Ende 129 '6285 2. Hälfte
 ICDs As String 'ICDs 129 '6286
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type briefe
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 Pfad As String 'Pfad 129 '
 art As String 'Art 129 '
 Name As String 'Name 129 '
 Typ As String 'Typ 129 '
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As Long 'DokGroe 3 'Größe der Datei
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type Diagnosen
 ID1 As Long 'ID1 3 '
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_id 3 'Bezug auf Anamneseblattt
 GesName As String 'GesName 129 '
 DiagDatum As Date 'DiagDatum 135 '
 DiagSicherheit As String 'DiagSicherheit 129 '6003
 DiagText As String 'DiagText 129 '
 DiagSeite As String 'DiagSeite 129 '6004
 DiagAttr As String 'DiagAttr 129 '6006
 ICD As String 'ICD 129 '
 obDauer As Integer 'obDauer 11 'ob Dauerdiagnose
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type dokumente
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 DokPfad As String 'DokPfad 129 '
 DokArt As String 'DokArt 129 '
 DokName As String 'DokName 129 '
 Quelldatum As Date 'Quelldatum 135 'Datum, auf das sich das Dokument bezieht
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As Long 'DokGroe 3 'Dokument-Größe
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type eintraege
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '
 art As String 'Art 129 '6330
 Inhalt As String 'Inhalt 129 '8480
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhaltform_abk
 Form_AbkVW As Long 'Form_AbkVW 3 '
 Form_Abk As String 'Form_Abk 129 '
End Type

Public Type formulare
 FormID As Long 'FormID 3 '
 Form_Abk As String 'Form_Abk 129 '
 FormBez As String 'FormBez 129 '
 FormVorl As String 'FormVorl 129 '
 AktZeit As Date 'AktZeit 135 'Zeitpunkt der Aktualisierung
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhkopf
 FoID As Long 'FoID 3 '
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Form_ID As Long 'Form_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Satzart As String 'Satzart 129 '8000
 Satzlänge As String 'Satzlänge 129 '8100
End Type

Public Type forminhfeld
 FoID As Long 'FoID 3 '
 nr As Integer 'Nr 2 '
 FeldNr As Integer 'FeldNr 2 '
 FeldVW As Long 'FeldVW 3 '
 Feld As String
 FeldInhVW As Long 'FeldInhVW 3 '
 FeldInh As String
End Type

Public Type kheinweis
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '
 Ziel As String 'Ziel 129 '6291
 Diagnose As String 'Diagnose 129 '6230
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type lbanforderungen
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 AnfText As String 'AnfText 129 '6280
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Statusbyte
End Type

Public Type laborneu
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '
 FertigStGrad As String 'FertigStGrad 129 '8401
 Abkü As String 'Abkü 129 '
 LangtextVW As Long 'LangtextVW 3 '8411
 Langtext As String
 Wert As String 'Wert 129 '
 Einheit As String 'Einheit 129 '8421
 KommentarVW As Long 'KommentarVW 3 '8480
 Kommentar As String
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 RefNr As Long 'Refnr 3 'Bezug auf LaborXUS
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type leistungen
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '5000 + 6201
 Leistung As String 'Leistung 129 '5001
 f5002 As String 'f5002 129 '5002
 f5005 As String 'f5005 129 '5005
 f5006 As String 'f5006 129 '5006
 f5009 As String 'f5009 129 '5009
 Med As String 'Med 129 '5010
 f5015 As String 'f5015 129 '5015
 f5016 As String 'f5016 129 '5016
 f5021 As String 'f5021 129 '5021
 f5026 As String 'f5026 129 '5026
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type medplan
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 MPNr As Long 'MPNr 3 'Ordnungsziffer für Medikamentenplan
 Zeitpunkt As Date 'ZeitPunkt 135 'Zeitpunkt, der Speicherung im Turbomed
 datum As Date 'Datum 135 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament As String 'Medikament 129 '
 MedAnfang As String 'MedAnfang 129 '
 FeldNr As Integer 'FeldNr 2 '
 mo As String 'mo 129 '
 mi As String 'mi 129 '
 nm As String 'nm 129 '
 ab As String 'ab 129 '
 Zn As String 'zn 129 '
 bBed As Integer 'bBed 11 '
 bemerkung As String 'Bemerkung 129 '
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type rezepteintraege
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 Rezept As String 'Rezept 129 '6210, 3652(1), 6218(1)
 Rezeptklasse As String 'Rezeptklasse 129 '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Medikament As String 'Medikament 129 '3652(2), 6218(4)
 PZN As String 'PZN 129 '6210(2), 6218(3)
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Statusbyte
End Type

Public Type RR
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 RR As String 'RR 129 '6230
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type kvnrue
 lfdnr As Long 'lfdnr 3 '
 Pat_id As Long 'Pat_ID 3 '
 KVNr As String 'KVNr 129 '
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type unbekannte_kennungen
 Kennung As String 'Kennung 129 '
 absPos As Long 'absPos 3 '
 StByte As Long 'StByte 3 '
 Pat_id As Long 'Pat_id 3 'zugehöriger Patient für spätere Ermittlungen
End Type

Public Type dmpreihe
 Abk As String 'Abk 129 'Abkürzung der DMP-Art
 art As String 'Art 129 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As Date 'KarteiDatum 133 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As Date 'exportiert 135 'Datum des Exports
 DokuDatum As Date 'DokuDatum 135 'Datum der Dokumentation
 obvoll As Integer 'obvoll 11 'ob vollständig
 NachName As String 'NachName 129 '
 VorName As String 'VorName 129 '
 GebDat As Date 'GebDat 133 '
 Pat_id As Long 'Pat_id 3 '
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AktZeit As Date 'AktZeit 135 'Aktualisierungzeit
End Type

Public Type laborxsaetze
 SatzID As Long 'SatzID 3 'zum Bezug für LaborUS
 DatID As Long 'DatID 3 'Bezug zu LaborEingelesen
 Satzart As String 'Satzart 129 '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge 129 '8100 Satzlänge (Turbomed)
 SatzlängeSchluss As String 'SatzlängeSchluss 129 '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb As String 'VersionSatzb 129 '9212 Version der Satzbeschreibung (Turbomed)
 Arztnr As String 'Arztnr 129 '201 Arztnummer (Turbomed)
 Arztname As String 'Arztname 129 '203 Arztname (Turbomed)
 StraßePraxis As String 'StraßePraxis 129 '205 Straße der Praxis (Turbomed)
 PLZPraxis As String 'PLZPraxis 129 '215 PLZ der Praxis (Turbomed)
 OrtPraxis As String 'OrtPraxis 129 '216 Ort der Praxis (Turbomed)
 Labor As String 'Labor 129 '8320 Labor
 StraßeLabor As String 'StraßeLabor 129 '8321 Straße der Laboradresse (Turbomed)
 PLZLabor As String 'PLZLabor 129 '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor As String 'OrtLabor 129 '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr As String 'KBVPrüfnr 129 '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz As String 'Zeichensatz 129 '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr As String 'Kundenarztnr 129 '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum As String 'Erstellungsdatum 129 '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge As String 'Gesamtlänge 129 '9202 Gesamtlänge des Datenpaketes (Turbomed)
End Type

Public Type laborxeingel
 DatID As Long 'DatID 3 'Bezug auf LaborEingelesen
 Pfad As String 'Pfad 129 'Pfadname
 Name As String 'Name 129 'Name der eingelesenen Labordatei ohne Endung
 zp As Date 'Zp 135 'Einlesezeitpunkt
 fertig As Integer 'fertig 11 '
End Type

Public Type laborxus
 RefNr As Long 'RefNr 3 'Bezug auf LaborWert
 DatID As Long 'DatID 3 'Bezug auf LaborEingelesen
 SatzID As Long 'SatzID 3 'Bezug auf LaborXSätze
 Satzart As String 'Satzart 129 '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge 129 '8100 Satzlänge (Turbomed)
 Auftragsnummer As String 'Auftragsnummer 129 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As String 'Auftragsschlüssel 129 '
 Eingang As Date 'Eingang 135 '8301 Eingangsdatum in Datumsform
 Berichtsdatum As String 'Berichtsdatum 129 '
 Pat_id As Long 'Pat_id 3 '
 NachName As String 'Nachname 129 '3101
 VorName As String 'Vorname 129 '3102
 GebDat As String 'GebDat 129 '3103
 Titel As String 'Titel 129 '3104
 NVorsatz As String 'NVorsatz 129 '3100
 BefArt As String 'BefArt 129 '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Abrechnungstyp As String 'Abrechnungstyp 129 '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd As String 'GebüOrd 129 '8403 Gebührenordnung (Turbomed)
 Patienteninformation As String 'Patienteninformation 129 '8405 Patienteninformation (Turbomed)
 Geschlecht As String 'Geschlecht 129 '8407 Geschlecht (Turbomed)
 AuftrHinw As String 'AuftrHinw 129 '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp As String 'Pat_idUrsp 129 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG As String 'Pat_idErwVNG 129 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN As String 'Pat_idErwVN 129 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG As String 'Pat_idErwG 129 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB As String 'Pat_idErwGB 129 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL As String 'Pat_idErwGL 129 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu As String 'Pat_idLaborNeu 129 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu As Date 'ZeitpunktLaborneu 135 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP As Integer 'ZdüP 2 'Zahl der verglichenen Parameter
 ZdiP As Long 'ZdiP 3 'Zahl der infragekommenden Patienten
 LWerte As String 'LWerte 129 'Laborwerte, die zur Zuordnung geführt haben
 verglichen As Date 'verglichen 135 'Datum, zu dem Datensatz zuletzt verglichen wurde
 afn As Integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
End Type

Public Type laborxbakt
 RefNr As Long 'RefNr 3 '
 Verf As String 'Verf 129 '
 KuQu As String 'KuQu 129 '8428 Probenmaterial-Ident (Turbomed)
 Quelle As String 'Quelle 129 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As String 'QSpez 129 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Kommentar As String 'Kommentar 129 '8480 Ergebnistest (Turbomed)
 Erklärung As String 'Erklärung 129 '
 Keimzahl As String 'Keimzahl 129 '
End Type

Public Type laborxwert
 RefNr As Long 'RefNr 3 'Bezug auf LaborUS
 Abkü As String 'Abkü 129 '8410 Test-Ident  (Turbomed)
 Langname As String 'Langname 129 '8411 Testbezeichnung (Turbomed)
 Quelle As String 'Quelle 129 '
 QSpez As String 'QSpez 129 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Wert As String 'Wert 129 '8420 Ergebniswert (Turbomed)
 Einheit As String 'Einheit 129 '8421 Einheit (Turbomed)
 Grenzwerti As String 'Grenzwerti 129 '8422 Grenzwertindikator (Turbomed)
 Kommentar As String 'Kommentar 129 '8480 Ergebnistext (Turbomed)
 Teststatus As String 'Teststatus 129 '8418 Teststatus (Turbomed)
 Erklärung As String 'Erklärung 129 '8470 Testbezogene Hinweise (Turbomed)
 Normbereich As String 'Normbereich 129 '8460 Normalwert-Text (Turbomed)
 NormU As String 'NormU 129 '8461 Normuntergrenze
 NormO As String 'NormO 129 '8461 Normobergrenze
 AuftrHinw As String 'AuftrHinw 129 '8490 Auftragsbezogene Hinweise (Turbomed)
End Type

Public Type laborxleist
 RefNr As Long 'RefNr 3 'Bezug auf LaborUS
 Abkü As String 'Abkü 129 '8410 Test-Ident (Turbomed)
 Verf As String 'Verf 129 '
 EBM As String 'EBM 129 '5001 GNR (Turbomed)
 goä As String 'goä 129 '8406
 Anzahl As String 'Anzahl 129 '5005
End Type

Public Type Anamnesebogen
 Pat_id As Long 'Pat_id 3 '
 NachName As String 'Nachname 129 '-
 VorName As String 'Vorname 129 '
 NVorsatz As String 'NVorsatz 129 '
 Titel As String 'Titel 129 '
 Anrede As String 'Anrede 129 '
 GebDat As Date 'GebDat 135 ', geb.
 Tkz As Integer 'Tkz 11 'Tod-Kennzeichen
 Versicherungsart As String 'Versicherungsart 129 '
 Diabetestyp As String 'Diabetestyp 129 '^Diabetes Typ
 Diabetes_seit As String 'Diabetes_seit 129 '<seit
 Tabletten_seit As String 'Tabletten_seit 129 ', Tabletten seit
 Insulin_seit As String 'Insulin_seit 129 ', Insulin seit
 Grund_für_Vorstellung As String 'Grund_für_Vorstellung 129 '^:
 Familienanamnese As String 'Familienanamnese 129 '^:
 Größe As Double 'Größe 5 '^:
 Gewicht As Double 'Gewicht 5 ',:
 Tendenz As String 'Tendenz 129 '<, Tendenz
 DiabetesMedikament_1 As String 'DiabetesMedikament_1 129 '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge As String 'DiabetesMedikament_1_Menge 129 '<
 DiabetesMedikament_2 As String 'DiabetesMedikament_2 129 '<,
 DiabetesMedikament_2_Menge As String 'DiabetesMedikament_2_Menge 129 '<
 DiabetesMedikament_3 As String 'DiabetesMedikament_3 129 '<,
 DiabetesMedikament_3_Menge As String 'DiabetesMedikament_3_Menge 129 '<
 DiabetesMedikament_4 As String 'DiabetesMedikament_4 129 '<,
 DiabetesMedikament_4_Menge As String 'DiabetesMedikament_4_Menge 129 '<,
 Insulinpumpe As Integer 'Insulinpumpe 11 '^:
 Insulinpumpe_seit As String 'Insulinpumpe_seit 129 '<seit
 Insulinpumpe_Marke As String 'Insulinpumpe_Marke 129 '<, Marke:
 Broteinheiten_gesamt As String 'Broteinheiten_gesamt 129 '^Broteinheiten:gesamt
 Broteinheiten_früh As String 'Broteinheiten_früh 129 '<, früh
 Broteinheiten_ZM_früh As String 'Broteinheiten_ZM_früh 129 '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags As String 'Broteinheiten_mittags 129 '<, mittags
 Broteinheiten_nachmittags As String 'Broteinheiten_nachmittags 129 '<, nachmittags
 Broteinheiten_abends As String 'Broteinheiten_abends 129 '<, abends
 Broteinheiten_nachts As String 'Broteinheiten_nachts 129 '<, nachts
 Essenszeit_früh As String 'Essenszeit_früh 129 '^Essenszeiten:früh
 Essenszeit_vormittags As String 'Essenszeit_vormittags 129 '<, vormittags
 Essenszeit_mittags As String 'Essenszeit_mittags 129 '<, mittags
 Essenszeit_nachmittags As String 'Essenszeit_nachmittags 129 '<, nachmittags
 Essenszeit_abends As String 'Essenszeit_abends 129 '<, abends
 Essenszeit_spät As String 'Essenszeit_spät 129 '<, spät
 Spritz_Eß_Abstand_früh As String 'Spritz_Eß_Abstand_früh 129 '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags As String 'Spritz_Eß_Abstand_mittags 129 '<, mittags
 Spritz_Eß_Abstand_abends As String 'Spritz_Eß_Abstand_abends 129 '<, abends
 Spritzstelle_früh As String 'Spritzstelle_früh 129 '^Spritzstellen:früh
 Spritzstelle_mittags As String 'Spritzstelle_mittags 129 '<, mittags
 Spritzstelle_abends As String 'Spritzstelle_abends 129 '<, abends
 Spritzstelle_nachts As String 'Spritzstelle_nachts 129 '<, nachts
 Jahr_letzte_Diabetesschulung As String 'Jahr_letzte_Diabetesschulung 129 '^Letzte Diabetesschulung:
 Ort_Schulung As String 'Ort_Schulung 129 '<in
 letztes_HbA1c As String 'letztes_HbA1c 129 '^Letztes HbA1c:
 gemessen_am As Date 'gemessen_am 135 '<, gemessen
 vorherige_Werte As String 'vorherige_Werte 129 '<, vorher:
 BZMessungen_selbst As String 'BZMessungen_selbst 129 '^Blutzuckermessung:Selbstmessung?
 Gerät As String 'Gerät 129 '<:
 BZMessungen_pW As String 'BZMessungen_pW 129 '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE As String 'BZMessungen_pW_ndE 129 '<, davon nach dem Essen:
 BZMessungen_p_W_nachts As String 'BZMessungen_p_W_nachts 129 '<, nachts:
 Aufschreiben As String 'Aufschreiben 129 '<, Dokumentation:
 BZWerte_v_d_Essen As String 'BZWerte_v_d_Essen 129 '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen As String 'BZWerte_n_d_Essen 129 '<, nach dem Essen:
 UZ_Tageszeit As String 'UZ_Tageszeit 129 '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM As String 'Unterzucker_pM 129 '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig As String 'UZ_rechtzeitig 129 '<, rechtzeitig bemerkt:
 Fremde_Hilfe_pa As String 'Fremde_Hilfe_pa 129 '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa As String 'Bewußtlos_pa 129 '<, bewußtlos deshalb:
 Keto_pa As String 'Keto_pa 129 '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM As String 'BZgr300_pM 129 ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck As String 'Bluthochdruck 129 '^Bluthochdruck:
 BHD_seit As String 'BHD_seit 129 '<seit:
 BHD_beh_mit As String 'BHD_beh_mit 129 '<, behandelt mit:
 Blutdruckwerte As String 'Blutdruckwerte 129 '^Blutdruckwerte:
 BDselbst As String 'BDselbst 129 '^Blutdruckselbstmessung:
 Schwanger As String 'Schwanger 129 '^Aktuelle Schwangerschaft:
 Schwanger_seit As String 'Schwanger_seit 129 '<, seit:
 Augensp_zuletzt As String 'Augensp_zuletzt 129 '^Letzte Augenspiegelung:
 Augensp_Befund As String 'Augensp_Befund 129 '<, Befund:
 Netzhaut_gelasert As String 'Netzhaut_gelasert 129 ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar As String 'Sehminderung_unbehebbar 129 ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden As String 'Diabet_Nierenschaden 129 '^Diabetischer Nierenschaden:
 Albumin_zuletzt As String 'Albumin_zuletzt 129 ', letztes Albumin:
 erhöht As String 'erhöht 129 '<, Befund:
 Dialyse As Integer 'Dialyse 11 ',:
 Dialyse_seit As String 'Dialyse_seit 129 '<seit
 andere_Nierenerkrankung As String 'andere_Nierenerkrankung 129 ', andere Nierenerkrankung:
 Herzkrankheit As String 'Herzkrankheit 129 '^Herzkrankheit:
 Angina_pectoris As String 'Angina_pectoris 129 ',:
 Herzinfarkt As String 'Herzinfarkt 129 ',:
 Herzinfarkt_wann As String 'Herzinfarkt_wann 129 '<, wann:
 PTCA_oder_Stent As String 'PTCA_oder_Stent 129 ',:
 Bypass_kardial As Integer 'Bypass_kardial 11 ',:
 Bypass_wann As String 'Bypass_wann 129 '<, wann:
 Herzschwäche As String 'Herzschwäche 129 ',:
 Herzkrankheit_Beschreibung As String 'Herzkrankheit_Beschreibung 129 ', Beschreibung:
 Hirndurchblutungsstörung As String 'Hirndurchblutungsstörung 129 '
 Schlaganfall As String 'Schlaganfall 129 ',:
 Beindurchblutungsstörung As String 'Beindurchblutungsstörung 129 '^:
 Schaufensterkrankheit As String 'Schaufensterkrankheit 129 ',:
 Bypaß_peripher As Integer 'Bypaß_peripher 11 ',:
 Geschwür As String 'Geschwür 129 ',:
 Amputation As String 'Amputation 129 ',:
 pAVK_Beschreibung As String 'pAVK_Beschreibung 129 ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen As String 'Ameisenlaufen 129 '^:
 Ameisen_Ausmaß As String 'Ameisen_Ausmaß 129 '<, Ausmaß:
 Druckstellen As String 'Druckstellen 129 ',:
 Verformungen As String 'Verformungen 129 ',:
 Verformungen_Beschreibung As String 'Verformungen_Beschreibung 129 '<Beschreibung:
 Fußpflege As String 'Fußpflege 129 '^:
 Podologie As String 'Podologie 129 ',:
 Einlagen As String 'Einlagen 129 ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen As String 'Neue_Fußkomplikationen 129 '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen As String 'Entleerungsstörungen_Magen 129 '^:
 Entleerungsstörungen_Harnblase As String 'Entleerungsstörungen_Harnblase 129 ',:
 Schwindel_Aufstehen As String 'Schwindel_Aufstehen 129 ',:
 Folgeerkrankungen_Haut As String 'Folgeerkrankungen_Haut 129 '^:
 Bewegungseinschränkungen As String 'Bewegungseinschränkungen 129 ',:
 Sexualstörung As String 'Sexualstörung 129 '^:
 Sexualstörung_seit As String 'Sexualstörung_seit 129 '<seit
 Weitere_Anamnese As String 'Weitere_Anamnese 129 '^:
 Alkohol As String 'Alkohol 129 '^:
 tabak As String 'Tabak 129 ',:
 Weitere_Medikation As String 'Weitere_Medikation 129 '^:
 Liphypertrophien_Abdomen As String 'Liphypertrophien_Abdomen 129 '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine As String 'Liphypertrophien_Beine 129 '<, Beine:
 Liphypertrophien_Arme As String 'Liphypertrophien_Arme 129 '<, Arme:
 Beinbefund As String 'Beinbefund 129 '^:
 Hyperkeratosen As String 'Hyperkeratosen 129 ',:
 Ulcera As String 'Ulcera 129 ',:
 Kraft_Zehenheber As String 'Kraft_Zehenheber 129 '^Kraft:Zehenheber
 Kraft_Zehenbeuger As String 'Kraft_Zehenbeuger 129 '<, Zehenbeuger:
 Kraft_Knie As String 'Kraft_Knie 129 '<, Knie:
 ASR As String 'ASR 129 ',:
 PSR As String 'PSR 129 ',:
 Oberflächensensibilität As String 'Oberflächensensibilität 129 '^:
 Monofilamenttest As String 'Monofilamenttest 129 ',:
 Kalt_Warm As String 'Kalt_Warm 129 ', Kalt-Warm-Diskrimination:
 Vibration_IK As String 'Vibration_IK 129 ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe As String 'Vibration_Großzehe 129 '<, Großzehe:
 Puls_Leiste As String 'Puls_Leiste 129 '^Pulse:Leiste
 Puls_Kniekehle As String 'Puls_Kniekehle 129 '<,Kniekehle:
 Puls_Atp As String 'Puls_Atp 129 '<,Innenknöchel:
 Puls_Adp As String 'Puls_Adp 129 '<,Fußrücken:
 RR As String 'RR 129 '^Blutdruck:
 RRTurboMed As String 'RRTurboMed 129 '
 Herz As String 'Herz 129 '^:
 Lunge As String 'Lunge 129 ',:
 Bauch As String 'Bauch 129 ', Abdomen:
 WS As String 'WS 129 ', Wirbelsäule:
 NL As String 'NL 129 ', Nierenlager:
 SD As String 'SD 129 ', Schilddrüse:
 Carotiden As String 'Carotiden 129 ', Halsschlagadern:
 NNH As String 'NNH 129 ', Nasennebenhöhlen:
 Zähne As String 'Zähne 129 ',:
 Mundhöhle As String 'Mundhöhle 129 ',:
 LK As String 'LK 129 ', Lymphknoten:
 BeinödVen As String 'BeinödVen 129 ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst As String 'Neuro_sonst 129 '^Sonstige neurologische Befunde:
 Weitere_Befunde As String 'Weitere_Befunde 129 ', weitere Befunde:
 Schulung As String 'Schulung 129 'ob Schulungsbedarf
 DMP As String 'DMP 129 'ob Pat. bei HA im DMP
 DMSchulz As Integer 'DMSchulz 2 'Zahl der DMP-Schulungen hier
 DMSchL As Integer 'DMSchL 2 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz As Integer 'RRSchulz 2 'Zahl der Hypertonie-Schulungen hier
 DMPhier As Date 'DMPhier 135 'ob Pat hier im DMP
 HANr As String 'HANr 129 'mit "/"
 HANr2 As String 'HANr2 129 'mit "/"
 letzte_Änderung As Date 'letzte_Änderung 135 'Datum der letzten Änderung
 Diagnosen As String 'Diagnosen 129 '
 Vorgestellt As Date 'Vorgestellt 135 'Erstvorstellung
 Versicherung As String 'Versicherung 129 '
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 Ther1 As String 'Ther1 129 'Diät, OAD, CT, ICT, CSII
 TherAkt As String 'TherAkt 129 'Diät, OAD, CT, ICT, CSII
 Prim As Long 'Prim 3 'Primärschlüssel
 obAn1eing As Integer 'obAn1eing 11 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As Integer 'obAn2eing 11 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As Integer 'obAnAeing 11 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck As Integer 'obCheck 11 'ob Checkliste vorliegt
 obBZausgew As Integer 'obBZausgew 11 'ob Blutzuckergerät ausgewechselt
 obOSaufgek As Integer 'obOSaufgek 11 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As Integer 'obPodAufgek 11 'ob über Podologie aufgeklärt
 obMBlAusgeh As Integer 'obMBlAusgeh 11 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek As String 'obSchulaufgek 129 'ob über Schulung aufgeklärt
 obDMPaufgekl As String 'obDMPaufgekl 129 'ob über DMP aufgeklärt
 obMedNetz As Integer 'obMedNetz 11 'ob von Med. Netz geschickt
 Hausarzt As String 'Hausarzt 129 'Hausarzt laut Anamnesebogen
 ob As Integer 'ob 11 'für verschiedene Aktionen
 QS As String 'QS 129 'Quartal sortiert von vorgestellt
 QT As String 'QT 129 'Quartal sortiert von vorgestellt
End Type

Public rNa() As namen
Public rFa() As Faelle
Public rAu() As au
Public rBr() As briefe
Public rDi() As Diagnosen
Public rDo() As dokumente
Public rEi() As eintraege
Public rFi() As forminhaltform_abk
Public rFo() As formulare
Public rFr() As forminhkopf
Public rFm() As forminhfeld
Public rKH() As kheinweis
Public rLB() As lbanforderungen
Public rLa() As laborneu
Public rLe() As leistungen
Public rMe() As medplan
Public rRe() As rezepteintraege
Public rRr() As RR
Public rKV() As kvnrue
Public rUn() As unbekannte_kennungen
Public rDm() As dmpreihe
Public rLs() As laborxsaetze
Public rLg() As laborxeingel
Public rLu() As laborxus
Public rLo() As laborxbakt
Public rLw() As laborxwert
Public rLL() As laborxleist
Public rana() As Anamnesebogen

Public Function Tinit()
 Static wdh%
 ReDim rana(0)
 ReDim rNa(0)
 ReDim rFa(0)
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 If wdh = 0 Then ReDim rFi(0)
 If wdh = 0 Then ReDim rFo(0)
 ReDim rFr(0)
 ReDim rFm(0)
 ReDim rKH(0)
 ReDim rLB(0)
 ReDim rLa(0)
 ReDim rLe(0)
 ReDim rMe(0)
 ReDim rRe(0)
 ReDim rRr(0)
 ReDim rKV(0)
 ReDim rUn(0)
 ReDim rDm(0)
 wdh = -1
End Function ' Tinit

Public Function LabInit()
 Static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 wdh = -1
End Function ' Tinit

Public Function doLösch(frm As Lese, tbl$)
 Set rs = DBCn.Execute("select count(*) as ct from " + kla + tbl + klz)
 frm.Ausgabe = "Lösche: " & kla & tbl & klz & " (" & rs!ct & " Datensätze)" & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 sql = sqlDeletefrom & kla & tbl & klz
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doLösch

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doLösch(frm, "dmpreihe")
 Call doLösch(frm, "unbekannte kennungen")
 Call doLösch(frm, "kvnrue")
 Call doLösch(frm, "rr")
 Call doLösch(frm, "rezepteintraege")
 Call doLösch(frm, "medplan")
 Call doLösch(frm, "leistungen")
 Call doLösch(frm, "laborneu")
 Call doLösch(frm, "lbanforderungen")
 Call doLösch(frm, "kheinweis")
 Call doLösch(frm, "forminhfeld")
 Call doLösch(frm, "forminhkopf")
 Call doLösch(frm, "formulare")
 Call doLösch(frm, "forminhaltform_abk")
 Call doLösch(frm, "eintraege")
 Call doLösch(frm, "dokumente")
 Call doLösch(frm, "diagnosen")
 Call doLösch(frm, "briefe")
 Call doLösch(frm, "au")
 Call doLösch(frm, "faelle")
 Call doLösch(frm, "namen")
 Call ForeignYes0
 Call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' AllesLösch

Public Function LabLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doLösch(frm, "laborxleist")
 Call doLösch(frm, "laborxwert")
 Call doLösch(frm, "laborxbakt")
 Call doLösch(frm, "laborxus")
 Call doLösch(frm, "laborxeingel")
 Call doLösch(frm, "laborxsaetze")
 Call ForeignYes0
 Call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' LabLösch


Public Function namenSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "namen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "namen" & klz & " (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,AufnDat," & _
     "intZOGP,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil,Email," & _
     "AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte,Cave,Notiz," & _
     "zubenach,Verwandt) values"
 For i = 0 To UBound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rNa(i).Pat_id & "," & rNa(i).lfdnr & ",'" & rNa(i).NVorsatz & "','" & rNa(i).NachName & "','" & rNa(i).VorName & "'," & datform(rNa(i).GebDat) & ",'" & _
   rNa(i).Straße & "','" & rNa(i).KVKStatus & "','" & rNa(i).Geschlecht & "','" & rNa(i).Plz & "','" & rNa(i).Ort & "','" & _
   rNa(i).Weggeldzone & "'," & datform(rNa(i).AufnDat) & ",'" & rNa(i).intZOGP & "','" & rNa(i).Titel & "','" & rNa(i).Versichertennummer & "','" & _
   rNa(i).PrivatTel & "','" & rNa(i).KVNr & "','" & rNa(i).PrivatTel_2 & "','" & rNa(i).PrivatFax & "','" & rNa(i).DienstTel & "','" & rNa(i).PrivatMobil & "','" & _
   rNa(i).Email & "'," & CStr(CInt(rNa(i).AnAllgda)) & "," & CStr(CInt(rNa(i).An1da)) & "," & CStr(CInt(rNa(i).An2da)) & "," & CStr(CInt( _
   rNa(i).Checkda)) & ",'" & rNa(i).DMTypaD & "'," & datform(0) & "," & rNa(i).absPos & "," & rNa(i).StByte & ",'" & rNa(i).Cave & "','" & _
   rNa(i).notiz & "','" & rNa(i).zubenach & "','" & rNa(i).Verwandt & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "NVorsatz":  If SpMod(rNa(i).NVorsatz, "namen", rSc) Then Exit Do
   Case "Nachname":  If SpMod(rNa(i).NachName, "namen", rSc) Then Exit Do
   Case "Vorname":  If SpMod(rNa(i).VorName, "namen", rSc) Then Exit Do
   Case "Straße":  If SpMod(rNa(i).Straße, "namen", rSc) Then Exit Do
   Case "KVKStatus":  If SpMod(rNa(i).KVKStatus, "namen", rSc) Then Exit Do
   Case "Geschlecht":  If SpMod(rNa(i).Geschlecht, "namen", rSc) Then Exit Do
   Case "Plz":  If SpMod(rNa(i).Plz, "namen", rSc) Then Exit Do
   Case "Ort":  If SpMod(rNa(i).Ort, "namen", rSc) Then Exit Do
   Case "Weggeldzone":  If SpMod(rNa(i).Weggeldzone, "namen", rSc) Then Exit Do
   Case "intZOGP":  If SpMod(rNa(i).intZOGP, "namen", rSc) Then Exit Do
   Case "Titel":  If SpMod(rNa(i).Titel, "namen", rSc) Then Exit Do
   Case "Versichertennummer":  If SpMod(rNa(i).Versichertennummer, "namen", rSc) Then Exit Do
   Case "PrivatTel":  If SpMod(rNa(i).PrivatTel, "namen", rSc) Then Exit Do
   Case "KVNr":  If SpMod(rNa(i).KVNr, "namen", rSc) Then Exit Do
   Case "PrivatTel_2":  If SpMod(rNa(i).PrivatTel_2, "namen", rSc) Then Exit Do
   Case "PrivatFax":  If SpMod(rNa(i).PrivatFax, "namen", rSc) Then Exit Do
   Case "DienstTel":  If SpMod(rNa(i).DienstTel, "namen", rSc) Then Exit Do
   Case "PrivatMobil":  If SpMod(rNa(i).PrivatMobil, "namen", rSc) Then Exit Do
   Case "Email":  If SpMod(rNa(i).Email, "namen", rSc) Then Exit Do
   Case "DMTypaD":  If SpMod(rNa(i).DMTypaD, "namen", rSc) Then Exit Do
   Case "Cave":  If SpMod(rNa(i).Cave, "namen", rSc) Then Exit Do
   Case "Notiz":  If SpMod(rNa(i).notiz, "namen", rSc) Then Exit Do
   Case "zubenach":  If SpMod(rNa(i).zubenach, "namen", rSc) Then Exit Do
   Case "Verwandt":  If SpMod(rNa(i).Verwandt, "namen", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
ElseIf Err.Number = -2147217871 Or Err.Number = -2147217859 Then
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
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "faelle" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "faelle" & klz & " (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,f4206,ÜwText,f4210," & _
     "statNuller,ÜbwV,AndÜw,Übw,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw," & _
     "statklasse,f4237,statBehTage,SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName," & _
     "AdNam,AdStr,AdPlz,AdOrt,BhFE,s8000,s8100,AktZeit,Fanf,altQuart," & _
     "QAnf,QEnd,QS,QT,StByte,absPos) values"
 For i = 1 To UBound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFa(i).Pat_id & ",'" & rFa(i).quartal & "','" & rFa(i).NachName & "','" & rFa(i).VorName & "'," & rFa(i).lfdnr & ",'" & rFa(i).TMFNr & "','" & _
   rFa(i).VKNr & "'," & datform(rFa(i).BhFB) & "," & datform(rFa(i).BhFE1) & "," & datform(rFa(i).BhFE2) & ",'" & rFa(i).f4202 & "'," & datform( _
   rFa(i).ausgst) & ",'" & rFa(i).KtrAbrB & "','" & rFa(i).AbrAr & "'," & datform(rFa(i).lVorl) & ",'" & rFa(i).IK & "','" & rFa(i).KVKs & "','" & _
   rFa(i).KVKserg & "','" & rFa(i).GebOr & "','" & rFa(i).AbrGb & "','" & rFa(i).f4206 & "','" & rFa(i).ÜwText & "','" & rFa(i).f4210 & "','" & _
   rFa(i).statNuller & "','" & rFa(i).ÜbwV & "','" & rFa(i).AndÜw & "','" & rFa(i).Übw & "','" & rFa(i).ÜWZiel & "','" & rFa(i).ÜWNNr & "','" & _
   rFa(i).ÜWNaN & "','" & rFa(i).ÜWTit & "','" & rFa(i).ÜWVor & "','" & rFa(i).ÜWVsw & "','" & rFa(i).statklasse & "','" & rFa(i).f4237 & "'," & _
   rFa(i).statBehTage & ",'" & rFa(i).SchGr & "','" & rFa(i).Weiterbeh & "','" & rFa(i).PGeb & "','" & rFa(i).PGebErg & "','" & rFa(i).Mahnfrist & "','" & _
   rFa(i).GOÄKatNr & "','" & rFa(i).GOÄKatName & "','" & rFa(i).AdNam & "','" & rFa(i).AdStr & "','" & rFa(i).AdPlz & "','" & rFa(i).AdOrt & "'," & datform( _
   rFa(i).BhFE) & ",'" & rFa(i).s8000 & "','" & rFa(i).s8100 & "'," & datform(rFa(i).AktZeit) & "," & datform(rFa(i).Fanf) & ",'" & _
   rFa(i).altQuart & "'," & datform(rFa(i).QAnf) & "," & datform(rFa(i).QEnd) & ",'" & rFa(i).QS & "','" & rFa(i).QT & "'," & rFa(i).StByte & "," & _
   rFa(i).absPos & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_id & " and quartal = '" & rFa(i).quartal & "' and bhfb = " & datform(rFa(i).BhFB) & " and bhfe1 = " & datform(rFa(i).BhFE1) & " and ausgst = " & datform(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgabe = "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID & " -> " & rs!FID & vbCrLf & Lese.Ausgabe
    altAusgabe = Lese.Ausgabe
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
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Quartal":  If SpMod(rFa(i).quartal, "faelle", rSc) Then Exit Do
   Case "Nachname":  If SpMod(rFa(i).NachName, "faelle", rSc) Then Exit Do
   Case "Vorname":  If SpMod(rFa(i).VorName, "faelle", rSc) Then Exit Do
   Case "TMFNr":  If SpMod(rFa(i).TMFNr, "faelle", rSc) Then Exit Do
   Case "VKNr":  If SpMod(rFa(i).VKNr, "faelle", rSc) Then Exit Do
   Case "f4202":  If SpMod(rFa(i).f4202, "faelle", rSc) Then Exit Do
   Case "KtrAbrB":  If SpMod(rFa(i).KtrAbrB, "faelle", rSc) Then Exit Do
   Case "AbrAr":  If SpMod(rFa(i).AbrAr, "faelle", rSc) Then Exit Do
   Case "IK":  If SpMod(rFa(i).IK, "faelle", rSc) Then Exit Do
   Case "KVKs":  If SpMod(rFa(i).KVKs, "faelle", rSc) Then Exit Do
   Case "KVKserg":  If SpMod(rFa(i).KVKserg, "faelle", rSc) Then Exit Do
   Case "GebOr":  If SpMod(rFa(i).GebOr, "faelle", rSc) Then Exit Do
   Case "AbrGb":  If SpMod(rFa(i).AbrGb, "faelle", rSc) Then Exit Do
   Case "f4206":  If SpMod(rFa(i).f4206, "faelle", rSc) Then Exit Do
   Case "ÜwText":  If SpMod(rFa(i).ÜwText, "faelle", rSc) Then Exit Do
   Case "f4210":  If SpMod(rFa(i).f4210, "faelle", rSc) Then Exit Do
   Case "statNuller":  If SpMod(rFa(i).statNuller, "faelle", rSc) Then Exit Do
   Case "ÜbwV":  If SpMod(rFa(i).ÜbwV, "faelle", rSc) Then Exit Do
   Case "AndÜw":  If SpMod(rFa(i).AndÜw, "faelle", rSc) Then Exit Do
   Case "Übw":  If SpMod(rFa(i).Übw, "faelle", rSc) Then Exit Do
   Case "ÜWZiel":  If SpMod(rFa(i).ÜWZiel, "faelle", rSc) Then Exit Do
   Case "ÜWNNr":  If SpMod(rFa(i).ÜWNNr, "faelle", rSc) Then Exit Do
   Case "ÜWNaN":  If SpMod(rFa(i).ÜWNaN, "faelle", rSc) Then Exit Do
   Case "ÜWTit":  If SpMod(rFa(i).ÜWTit, "faelle", rSc) Then Exit Do
   Case "ÜWVor":  If SpMod(rFa(i).ÜWVor, "faelle", rSc) Then Exit Do
   Case "ÜWVsw":  If SpMod(rFa(i).ÜWVsw, "faelle", rSc) Then Exit Do
   Case "statklasse":  If SpMod(rFa(i).statklasse, "faelle", rSc) Then Exit Do
   Case "f4237":  If SpMod(rFa(i).f4237, "faelle", rSc) Then Exit Do
   Case "SchGr":  If SpMod(rFa(i).SchGr, "faelle", rSc) Then Exit Do
   Case "Weiterbeh":  If SpMod(rFa(i).Weiterbeh, "faelle", rSc) Then Exit Do
   Case "PGeb":  If SpMod(rFa(i).PGeb, "faelle", rSc) Then Exit Do
   Case "PGebErg":  If SpMod(rFa(i).PGebErg, "faelle", rSc) Then Exit Do
   Case "Mahnfrist":  If SpMod(rFa(i).Mahnfrist, "faelle", rSc) Then Exit Do
   Case "GOÄKatNr":  If SpMod(rFa(i).GOÄKatNr, "faelle", rSc) Then Exit Do
   Case "GOÄKatName":  If SpMod(rFa(i).GOÄKatName, "faelle", rSc) Then Exit Do
   Case "AdNam":  If SpMod(rFa(i).AdNam, "faelle", rSc) Then Exit Do
   Case "AdStr":  If SpMod(rFa(i).AdStr, "faelle", rSc) Then Exit Do
   Case "AdPlz":  If SpMod(rFa(i).AdPlz, "faelle", rSc) Then Exit Do
   Case "AdOrt":  If SpMod(rFa(i).AdOrt, "faelle", rSc) Then Exit Do
   Case "s8000":  If SpMod(rFa(i).s8000, "faelle", rSc) Then Exit Do
   Case "s8100":  If SpMod(rFa(i).s8100, "faelle", rSc) Then Exit Do
   Case "altQuart":  If SpMod(rFa(i).altQuart, "faelle", rSc) Then Exit Do
   Case "QS":  If SpMod(rFa(i).QS, "faelle", rSc) Then Exit Do
   Case "QT":  If SpMod(rFa(i).QT, "faelle", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "insert into " & kla & "kassenliste" & klz & "(" & kla & "GO" & klz & "," & kla & "VK" & klz & "," & kla & "IK" & klz & ") values (" & "'" & rFa(i).GOÄKatName & "', '" & rFa(i).VKNr & "', '" & rFa(i).VKNr & "')"
 Call DBCn.Execute(sqlquer, rAF)
 Resume
End If ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "au" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "au" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values"
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rAu(i).FID & "," & rAu(i).Pat_id & "," & datform(rAu(i).Zeitpunkt) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "','" & rAu(i).ICDs & "'," & _
   rAu(i).absPos & "," & datform(rAu(i).AktZeit) & "," & rAu(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Beginn":  If SpMod(rAu(i).Beginn, "au", rSc) Then Exit Do
   Case "Ende":  If SpMod(rAu(i).Ende, "au", rSc) Then Exit Do
   Case "ICDs":  If SpMod(rAu(i).ICDs, "au", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "briefe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "briefe" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte) values"
 For i = 1 To UBound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rBr(i).FID & "," & rBr(i).Pat_id & "," & datform(rBr(i).Zeitpunkt) & ",'" & rBr(i).Pfad & "','" & rBr(i).art & "','" & rBr(i).Name & "','" & _
   rBr(i).Typ & "'," & datform(rBr(i).AktZeit) & "," & rBr(i).DokGroe & ",'" & rBr(i).QS & "','" & rBr(i).QT & "'," & rBr(i).absPos & "," & _
   rBr(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Pfad":  If SpMod(rBr(i).Pfad, "briefe", rSc) Then Exit Do
   Case "Art":  If SpMod(rBr(i).art, "briefe", rSc) Then Exit Do
   Case "Name":  If SpMod(rBr(i).Name, "briefe", rSc) Then Exit Do
   Case "Typ":  If SpMod(rBr(i).Typ, "briefe", rSc) Then Exit Do
   Case "QS":  If SpMod(rBr(i).QS, "briefe", rSc) Then Exit Do
   Case "QT":  If SpMod(rBr(i).QT, "briefe", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "diagnosen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "diagnosen" & klz & " (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,absPos,AktZeit,StByte) values"
 For i = 1 To UBound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDi(i).FID & "," & rDi(i).Pat_id & ",'" & rDi(i).GesName & "'," & datform(rDi(i).DiagDatum) & ",'" & rDi(i).DiagSicherheit & "','" & _
   rDi(i).DiagText & "','" & rDi(i).DiagSeite & "','" & rDi(i).DiagAttr & "','" & rDi(i).ICD & "'," & CStr(CInt(rDi(i).obDauer)) & "," & rDi(i).absPos & "," & datform( _
   rDi(i).AktZeit) & "," & rDi(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "GesName":  If SpMod(rDi(i).GesName, "diagnosen", rSc) Then Exit Do
   Case "DiagSicherheit":  If SpMod(rDi(i).DiagSicherheit, "diagnosen", rSc) Then Exit Do
   Case "DiagText":  If SpMod(rDi(i).DiagText, "diagnosen", rSc) Then Exit Do
   Case "DiagSeite":  If SpMod(rDi(i).DiagSeite, "diagnosen", rSc) Then Exit Do
   Case "DiagAttr":  If SpMod(rDi(i).DiagAttr, "diagnosen", rSc) Then Exit Do
   Case "ICD":  If SpMod(rDi(i).ICD, "diagnosen", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "dokumente" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "dokumente" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte) values"
 For i = 1 To UBound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rDo(i).FID & "," & rDo(i).Pat_id & "," & datform(rDo(i).Zeitpunkt) & ",'" & rDo(i).DokPfad & "','" & rDo(i).DokArt & "','" & rDo(i).DokName & "'," & datform( _
   rDo(i).Quelldatum) & "," & rDo(i).absPos & "," & datform(rDo(i).AktZeit) & "," & rDo(i).DokGroe & ",'" & rDo(i).QS & "','" & _
   rDo(i).QT & "'," & rDo(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "DokPfad":  If SpMod(rDo(i).DokPfad, "dokumente", rSc) Then Exit Do
   Case "DokArt":  If SpMod(rDo(i).DokArt, "dokumente", rSc) Then Exit Do
   Case "DokName":  If SpMod(rDo(i).DokName, "dokumente", rSc) Then Exit Do
   Case "QS":  If SpMod(rDo(i).QS, "dokumente", rSc) Then Exit Do
   Case "QT":  If SpMod(rDo(i).QT, "dokumente", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "eintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "eintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte) values"
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rEi(i).FID & "," & rEi(i).Pat_id & "," & datform(rEi(i).Zeitpunkt) & ",'" & rEi(i).art & "','" & rEi(i).Inhalt & "'," & rEi(i).absPos & "," & datform( _
   rEi(i).AktZeit) & ",'" & rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Art":  If SpMod(rEi(i).art, "eintraege", rSc) Then Exit Do
   Case "Inhalt":  If SpMod(rEi(i).Inhalt, "eintraege", rSc) Then Exit Do
   Case "QS":  If SpMod(rEi(i).QS, "eintraege", rSc) Then Exit Do
   Case "QT":  If SpMod(rEi(i).QT, "eintraege", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "forminhaltform_abk" & klz & " (Form_Abk) values"
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
  sql = sql0 & "('" & rFi(i).Form_Abk & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFi1 = UBound(rFi)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Form_Abk":  If SpMod(rFi(i).Form_Abk, "forminhaltform_abk", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "formulare" & klz & " (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values"
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFo(i).FormID & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "','" & rFo(i).FormVorl & "'," & datform(rFo(i).AktZeit) & "," & _
   rFo(i).absPos & "," & rFo(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rFo1 = UBound(rFo)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Form_Abk":  If SpMod(rFo(i).Form_Abk, "formulare", rSc) Then Exit Do
   Case "FormBez":  If SpMod(rFo(i).FormBez, "formulare", rSc) Then Exit Do
   Case "FormVorl":  If SpMod(rFo(i).FormVorl, "formulare", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
'  sql = "delete FROM " & kla & "forminhfeld" & klz & " where foid in (select foid from " & kla & "forminhkopf" & klz & " where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "delete from " & kla & "forminhkopf" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "forminhkopf" & klz & " (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge) values"
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rFr(i).FoID & "," & rFr(i).FID & "," & rFr(i).Pat_id & "," & rFr(i).Form_ID & "," & datform(rFr(i).Zeitpunkt) & "," & rFr(i).absPos & "," & datform( _
   rFr(i).AktZeit) & "," & rFr(i).StByte & ",'" & rFr(i).Satzart & "','" & rFr(i).Satzlänge & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Satzart":  If SpMod(rFr(i).Satzart, "forminhkopf", rSc) Then Exit Do
   Case "Satzlänge":  If SpMod(rFr(i).Satzlänge, "forminhkopf", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "forminhfeld" & klz & " (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values"
 For i = 1 To UBound(rFm)
'  rFm(i).AktZeit = now()
  sql = sql0 & "(" & rFm(i).FoID & "," & rFm(i).nr & "," & rFm(i).FeldNr & "," & rFm(i).FeldVW & "," & rFm(i).FeldInhVW & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "kheinweis" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "kheinweis" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values"
 For i = 1 To UBound(rKH)
'  rKh(i).AktZeit = now()
  rKH(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKH(i).FID & "," & rKH(i).Pat_id & "," & datform(rKH(i).Zeitpunkt) & ",'" & rKH(i).Ziel & "','" & rKH(i).Diagnose & "'," & rKH(i).absPos & "," & datform( _
   rKH(i).AktZeit) & "," & rKH(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Ziel":  If SpMod(rKH(i).Ziel, "kheinweis", rSc) Then Exit Do
   Case "Diagnose":  If SpMod(rKH(i).Diagnose, "kheinweis", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "lbanforderungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "lbanforderungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values"
 For i = 1 To UBound(rLB)
'  rLb(i).AktZeit = now()
  rLB(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLB(i).FID & "," & rLB(i).Pat_id & "," & datform(rLB(i).Zeitpunkt) & ",'" & rLB(i).AnfText & "'," & rLB(i).absPos & "," & datform( _
   rLB(i).AktZeit) & "," & rLB(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "AnfText":  If SpMod(rLB(i).AnfText, "lbanforderungen", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "laborneu" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborneu" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values"
 For i = 1 To UBound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLa(i).FID & "," & rLa(i).Pat_id & "," & datform(rLa(i).Zeitpunkt) & ",'" & rLa(i).FertigStGrad & "','" & rLa(i).Abkü & "'," & rLa(i).LangtextVW & ",'" & _
   rLa(i).Wert & "','" & rLa(i).Einheit & "'," & rLa(i).KommentarVW & "," & rLa(i).absPos & "," & datform(rLa(i).AktZeit) & "," & _
   rLa(i).RefNr & "," & rLa(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "FertigStGrad":  If SpMod(rLa(i).FertigStGrad, "laborneu", rSc) Then Exit Do
   Case "Abkü":  If SpMod(rLa(i).Abkü, "laborneu", rSc) Then Exit Do
   Case "Wert":  If SpMod(rLa(i).Wert, "laborneu", rSc) Then Exit Do
   Case "Einheit":  If SpMod(rLa(i).Einheit, "laborneu", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "leistungen" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "leistungen" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "absPos,AktZeit,QS,QT,StByte) values"
 For i = 1 To UBound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rLe(i).FID & "," & rLe(i).Pat_id & "," & datform(rLe(i).Zeitpunkt) & ",'" & rLe(i).Leistung & "','" & rLe(i).f5002 & "','" & rLe(i).f5005 & "','" & _
   rLe(i).f5006 & "','" & rLe(i).f5009 & "','" & rLe(i).Med & "','" & rLe(i).f5015 & "','" & rLe(i).f5016 & "','" & rLe(i).f5021 & "','" & _
   rLe(i).f5026 & "'," & rLe(i).absPos & "," & datform(rLe(i).AktZeit) & ",'" & rLe(i).QS & "','" & rLe(i).QT & "'," & rLe(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Leistung":  If SpMod(rLe(i).Leistung, "leistungen", rSc) Then Exit Do
   Case "f5002":  If SpMod(rLe(i).f5002, "leistungen", rSc) Then Exit Do
   Case "f5005":  If SpMod(rLe(i).f5005, "leistungen", rSc) Then Exit Do
   Case "f5006":  If SpMod(rLe(i).f5006, "leistungen", rSc) Then Exit Do
   Case "f5009":  If SpMod(rLe(i).f5009, "leistungen", rSc) Then Exit Do
   Case "Med":  If SpMod(rLe(i).Med, "leistungen", rSc) Then Exit Do
   Case "f5015":  If SpMod(rLe(i).f5015, "leistungen", rSc) Then Exit Do
   Case "f5016":  If SpMod(rLe(i).f5016, "leistungen", rSc) Then Exit Do
   Case "f5021":  If SpMod(rLe(i).f5021, "leistungen", rSc) Then Exit Do
   Case "f5026":  If SpMod(rLe(i).f5026, "leistungen", rSc) Then Exit Do
   Case "QS":  If SpMod(rLe(i).QS, "leistungen", rSc) Then Exit Do
   Case "QT":  If SpMod(rLe(i).QT, "leistungen", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "medplan" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "medplan" & klz & " (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values"
 For i = 1 To UBound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rMe(i).FID & "," & rMe(i).Pat_id & "," & rMe(i).MPNr & "," & datform(rMe(i).Zeitpunkt) & "," & datform(rMe(i).datum) & ",'" & rMe(i).Medikament & "','" & _
   rMe(i).MedAnfang & "'," & rMe(i).FeldNr & ",'" & rMe(i).mo & "','" & rMe(i).mi & "','" & rMe(i).nm & "','" & rMe(i).ab & "','" & _
   rMe(i).Zn & "'," & CStr(CInt(rMe(i).bBed)) & ",'" & rMe(i).bemerkung & "'," & rMe(i).absPos & "," & datform(rMe(i).AktZeit) & "," & rMe(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Medikament":  If SpMod(rMe(i).Medikament, "medplan", rSc) Then Exit Do
   Case "MedAnfang":  If SpMod(rMe(i).MedAnfang, "medplan", rSc) Then Exit Do
   Case "mo":  If SpMod(rMe(i).mo, "medplan", rSc) Then Exit Do
   Case "mi":  If SpMod(rMe(i).mi, "medplan", rSc) Then Exit Do
   Case "nm":  If SpMod(rMe(i).nm, "medplan", rSc) Then Exit Do
   Case "ab":  If SpMod(rMe(i).ab, "medplan", rSc) Then Exit Do
   Case "zn":  If SpMod(rMe(i).Zn, "medplan", rSc) Then Exit Do
   Case "Bemerkung":  If SpMod(rMe(i).bemerkung, "medplan", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 syscmd 4, "Speichere rezepteintraege"
 If Not AllePat Then
  sql = "delete from " & kla & "rezepteintraege" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "rezepteintraege" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte) values"
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRe(i).FID & "," & rRe(i).Pat_id & "," & datform(rRe(i).Zeitpunkt) & ",'" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "','" & _
   rRe(i).Medikament & "','" & rRe(i).PZN & "'," & rRe(i).absPos & "," & datform(rRe(i).AktZeit) & ",'" & rRe(i).QS & "','" & rRe(i).QT & "'," & rRe(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 syscmd 5, ""
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Rezept":  If SpMod(rRe(i).Rezept, "rezepteintraege", rSc) Then Exit Do
   Case "Rezeptklasse":  If SpMod(rRe(i).Rezeptklasse, "rezepteintraege", rSc) Then Exit Do
   Case "Medikament":  If SpMod(rRe(i).Medikament, "rezepteintraege", rSc) Then Exit Do
   Case "PZN":  If SpMod(rRe(i).PZN, "rezepteintraege", rSc) Then Exit Do
   Case "QS":  If SpMod(rRe(i).QS, "rezepteintraege", rSc) Then Exit Do
   Case "QT":  If SpMod(rRe(i).QT, "rezepteintraege", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 syscmd 4, "Speichere rr"
 If Not AllePat Then
  sql = "delete from " & kla & "rr" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "rr" & klz & " (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte) values"
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rRr(i).FID & "," & rRr(i).Pat_id & "," & datform(rRr(i).Zeitpunkt) & ",'" & rRr(i).RR & "'," & rRr(i).absPos & "," & datform(rRr(i).AktZeit) & "," & _
   rRr(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "RR":  If SpMod(rRr(i).RR, "rr", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "kvnrue" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "kvnrue" & klz & " (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values"
 For i = 1 To UBound(rKV)
'  rKv(i).AktZeit = now()
  rKV(i).StByte = CStr(AktByte)
  sql = sql0 & "(" & rKV(i).Pat_id & ",'" & rKV(i).KVNr & "'," & rKV(i).absPos & "," & datform(rKV(i).AktZeit) & "," & rKV(i).StByte & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "KVNr":  If SpMod(rKV(i).KVNr, "kvnrue", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern()
 Dim i%, sql0$, rAF&
 On Error Resume Next
 If Not AllePat Then
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "unbekannte kennungen" & klz & " (Kennung,absPos,StByte," & _
     "Pat_id) values"
 For i = rUn1 + 1 To UBound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rUn(i).Kennung & "'," & rUn(i).absPos & "," & rUn(i).StByte & "," & rUn(i).Pat_id & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 rUn1 = UBound(rUn)
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Kennung":  If SpMod(rUn(i).Kennung, "unbekannte kennungen", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 If Not AllePat Then
  sql = "delete from " & kla & "dmpreihe" & klz & " where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
 sql0 = " insert " & sqlIgnore & "into " & kla & "dmpreihe" & klz & " (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit) values"
 For i = 1 To UBound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  sql = sql0 & "('" & rDm(i).Abk & "','" & rDm(i).art & "'," & datform(rDm(i).KarteiDatum) & "," & datform(rDm(i).exportiert) & "," & datform(rDm(i).DokuDatum) & "," & CStr(CInt( _
   rDm(i).obvoll)) & ",'" & rDm(i).NachName & "','" & rDm(i).VorName & "'," & datform(rDm(i).GebDat) & "," & rDm(i).Pat_id & "," & _
   rDm(i).StByte & "," & datform(rDm(i).AktZeit) & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 DBCn.CommitTrans
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Abk":  If SpMod(rDm(i).Abk, "dmpreihe", rSc) Then Exit Do
   Case "Art":  If SpMod(rDm(i).art, "dmpreihe", rSc) Then Exit Do
   Case "NachName":  If SpMod(rDm(i).NachName, "dmpreihe", rSc) Then Exit Do
   Case "VorName":  If SpMod(rDm(i).VorName, "dmpreihe", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 DBCn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function laborxsaetzeSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxsaetze" & klz & " (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,PLZPraxis,OrtPraxis,Labor,StraßeLabor,PLZLabor," & _
     "OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values"
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
  sql = sql0 & "(" & rLs(i).DatID & ",'" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).VersionSatzb & "','" & _
   rLs(i).Arztnr & "','" & rLs(i).Arztname & "','" & rLs(i).StraßePraxis & "','" & rLs(i).PLZPraxis & "','" & rLs(i).OrtPraxis & "','" & rLs(i).Labor & "','" & _
   rLs(i).StraßeLabor & "','" & rLs(i).PLZLabor & "','" & rLs(i).OrtLabor & "','" & rLs(i).KBVPrüfnr & "','" & rLs(i).Zeichensatz & "','" & _
   rLs(i).Kundenarztnr & "','" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Satzart":  If SpMod(rLs(i).Satzart, "laborxsaetze", rSc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLs(i).Satzlänge, "laborxsaetze", rSc) Then Exit Do
   Case "SatzlängeSchluss":  If SpMod(rLs(i).SatzlängeSchluss, "laborxsaetze", rSc) Then Exit Do
   Case "VersionSatzb":  If SpMod(rLs(i).VersionSatzb, "laborxsaetze", rSc) Then Exit Do
   Case "Arztnr":  If SpMod(rLs(i).Arztnr, "laborxsaetze", rSc) Then Exit Do
   Case "Arztname":  If SpMod(rLs(i).Arztname, "laborxsaetze", rSc) Then Exit Do
   Case "StraßePraxis":  If SpMod(rLs(i).StraßePraxis, "laborxsaetze", rSc) Then Exit Do
   Case "PLZPraxis":  If SpMod(rLs(i).PLZPraxis, "laborxsaetze", rSc) Then Exit Do
   Case "OrtPraxis":  If SpMod(rLs(i).OrtPraxis, "laborxsaetze", rSc) Then Exit Do
   Case "Labor":  If SpMod(rLs(i).Labor, "laborxsaetze", rSc) Then Exit Do
   Case "StraßeLabor":  If SpMod(rLs(i).StraßeLabor, "laborxsaetze", rSc) Then Exit Do
   Case "PLZLabor":  If SpMod(rLs(i).PLZLabor, "laborxsaetze", rSc) Then Exit Do
   Case "OrtLabor":  If SpMod(rLs(i).OrtLabor, "laborxsaetze", rSc) Then Exit Do
   Case "KBVPrüfnr":  If SpMod(rLs(i).KBVPrüfnr, "laborxsaetze", rSc) Then Exit Do
   Case "Zeichensatz":  If SpMod(rLs(i).Zeichensatz, "laborxsaetze", rSc) Then Exit Do
   Case "Kundenarztnr":  If SpMod(rLs(i).Kundenarztnr, "laborxsaetze", rSc) Then Exit Do
   Case "Erstellungsdatum":  If SpMod(rLs(i).Erstellungsdatum, "laborxsaetze", rSc) Then Exit Do
   Case "Gesamtlänge":  If SpMod(rLs(i).Gesamtlänge, "laborxsaetze", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxeingel" & klz & " (Pfad,Name,Zp," & _
     "fertig) values"
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
  sql = sql0 & "('" & rLg(i).Pfad & "','" & rLg(i).Name & "'," & datform(rLg(i).zp) & "," & CStr(CInt(rLg(i).fertig)) & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Pfad":  If SpMod(rLg(i).Pfad, "laborxeingel", rSc) Then Exit Do
   Case "Name":  If SpMod(rLg(i).Name, "laborxeingel", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxus" & klz & " (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN) values"
 For i = j To j
'  rLu(i).AktZeit = now()
  sql = sql0 & "(" & rLu(i).DatID & "," & rLu(i).SatzID & ",'" & rLu(i).Satzart & "','" & rLu(i).Satzlänge & "','" & rLu(i).Auftragsnummer & "','" & rLu(i).Auftragsschlüssel & "'," & datform( _
   rLu(i).Eingang) & ",'" & rLu(i).Berichtsdatum & "'," & rLu(i).Pat_id & ",'" & rLu(i).NachName & "','" & _
   rLu(i).VorName & "','" & rLu(i).GebDat & "','" & rLu(i).Titel & "','" & rLu(i).NVorsatz & "','" & rLu(i).BefArt & "','" & rLu(i).Abrechnungstyp & "','" & _
   rLu(i).GebüOrd & "','" & rLu(i).Patienteninformation & "','" & rLu(i).Geschlecht & "','" & rLu(i).AuftrHinw & "','" & rLu(i).Pat_idUrsp & "','" & _
   rLu(i).Pat_idErwVNG & "','" & rLu(i).Pat_idErwVN & "','" & rLu(i).Pat_idErwG & "','" & rLu(i).Pat_idErwGB & "','" & rLu(i).Pat_idErwGL & "','" & _
   rLu(i).Pat_idLaborNeu & "'," & datform(rLu(i).ZeitpunktLaborneu) & "," & rLu(i).ZdüP & "," & rLu(i).ZdiP & ",'" & rLu(i).LWerte & "'," & datform( _
   rLu(i).verglichen) & "," & rLu(i).afn & ")"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Satzart":  If SpMod(rLu(i).Satzart, "laborxus", rSc) Then Exit Do
   Case "Satzlänge":  If SpMod(rLu(i).Satzlänge, "laborxus", rSc) Then Exit Do
   Case "Auftragsnummer":  If SpMod(rLu(i).Auftragsnummer, "laborxus", rSc) Then Exit Do
   Case "Auftragsschlüssel":  If SpMod(rLu(i).Auftragsschlüssel, "laborxus", rSc) Then Exit Do
   Case "Berichtsdatum":  If SpMod(rLu(i).Berichtsdatum, "laborxus", rSc) Then Exit Do
   Case "Nachname":  If SpMod(rLu(i).NachName, "laborxus", rSc) Then Exit Do
   Case "Vorname":  If SpMod(rLu(i).VorName, "laborxus", rSc) Then Exit Do
   Case "GebDat":  If SpMod(rLu(i).GebDat, "laborxus", rSc) Then Exit Do
   Case "Titel":  If SpMod(rLu(i).Titel, "laborxus", rSc) Then Exit Do
   Case "NVorsatz":  If SpMod(rLu(i).NVorsatz, "laborxus", rSc) Then Exit Do
   Case "BefArt":  If SpMod(rLu(i).BefArt, "laborxus", rSc) Then Exit Do
   Case "Abrechnungstyp":  If SpMod(rLu(i).Abrechnungstyp, "laborxus", rSc) Then Exit Do
   Case "GebüOrd":  If SpMod(rLu(i).GebüOrd, "laborxus", rSc) Then Exit Do
   Case "Patienteninformation":  If SpMod(rLu(i).Patienteninformation, "laborxus", rSc) Then Exit Do
   Case "Geschlecht":  If SpMod(rLu(i).Geschlecht, "laborxus", rSc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLu(i).AuftrHinw, "laborxus", rSc) Then Exit Do
   Case "Pat_idUrsp":  If SpMod(rLu(i).Pat_idUrsp, "laborxus", rSc) Then Exit Do
   Case "Pat_idErwVNG":  If SpMod(rLu(i).Pat_idErwVNG, "laborxus", rSc) Then Exit Do
   Case "Pat_idErwVN":  If SpMod(rLu(i).Pat_idErwVN, "laborxus", rSc) Then Exit Do
   Case "Pat_idErwG":  If SpMod(rLu(i).Pat_idErwG, "laborxus", rSc) Then Exit Do
   Case "Pat_idErwGB":  If SpMod(rLu(i).Pat_idErwGB, "laborxus", rSc) Then Exit Do
   Case "Pat_idErwGL":  If SpMod(rLu(i).Pat_idErwGL, "laborxus", rSc) Then Exit Do
   Case "Pat_idLaborNeu":  If SpMod(rLu(i).Pat_idLaborNeu, "laborxus", rSc) Then Exit Do
   Case "LWerte":  If SpMod(rLu(i).LWerte, "laborxus", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxbakt" & klz & " (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values"
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
  sql = sql0 & "(" & rLo(i).RefNr & ",'" & rLo(i).Verf & "','" & rLo(i).KuQu & "','" & rLo(i).Quelle & "','" & rLo(i).QSpez & "'," & datform(rLo(i).AbnDat) & ",'" & _
   rLo(i).Kommentar & "','" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Verf":  If SpMod(rLo(i).Verf, "laborxbakt", rSc) Then Exit Do
   Case "KuQu":  If SpMod(rLo(i).KuQu, "laborxbakt", rSc) Then Exit Do
   Case "Quelle":  If SpMod(rLo(i).Quelle, "laborxbakt", rSc) Then Exit Do
   Case "QSpez":  If SpMod(rLo(i).QSpez, "laborxbakt", rSc) Then Exit Do
   Case "Kommentar":  If SpMod(rLo(i).Kommentar, "laborxbakt", rSc) Then Exit Do
   Case "Erklärung":  If SpMod(rLo(i).Erklärung, "laborxbakt", rSc) Then Exit Do
   Case "Keimzahl":  If SpMod(rLo(i).Keimzahl, "laborxbakt", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxwert" & klz & " (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw) values"
 For i = 1 To UBound(rLw)
'  rLw(i).AktZeit = now()
  sql = sql0 & "(" & rLw(i).RefNr & ",'" & rLw(i).Abkü & "','" & rLw(i).Langname & "','" & rLw(i).Quelle & "','" & rLw(i).QSpez & "'," & datform(rLw(i).AbnDat) & ",'" & _
   rLw(i).Wert & "','" & rLw(i).Einheit & "','" & rLw(i).Grenzwerti & "','" & rLw(i).Kommentar & "','" & rLw(i).Teststatus & "','" & _
   rLw(i).Erklärung & "','" & rLw(i).Normbereich & "','" & rLw(i).NormU & "','" & rLw(i).NormO & "','" & rLw(i).AuftrHinw & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Abkü":  If SpMod(rLw(i).Abkü, "laborxwert", rSc) Then Exit Do
   Case "Langname":  If SpMod(rLw(i).Langname, "laborxwert", rSc) Then Exit Do
   Case "Quelle":  If SpMod(rLw(i).Quelle, "laborxwert", rSc) Then Exit Do
   Case "QSpez":  If SpMod(rLw(i).QSpez, "laborxwert", rSc) Then Exit Do
   Case "Wert":  If SpMod(rLw(i).Wert, "laborxwert", rSc) Then Exit Do
   Case "Einheit":  If SpMod(rLw(i).Einheit, "laborxwert", rSc) Then Exit Do
   Case "Grenzwerti":  If SpMod(rLw(i).Grenzwerti, "laborxwert", rSc) Then Exit Do
   Case "Kommentar":  If SpMod(rLw(i).Kommentar, "laborxwert", rSc) Then Exit Do
   Case "Teststatus":  If SpMod(rLw(i).Teststatus, "laborxwert", rSc) Then Exit Do
   Case "Erklärung":  If SpMod(rLw(i).Erklärung, "laborxwert", rSc) Then Exit Do
   Case "Normbereich":  If SpMod(rLw(i).Normbereich, "laborxwert", rSc) Then Exit Do
   Case "NormU":  If SpMod(rLw(i).NormU, "laborxwert", rSc) Then Exit Do
   Case "NormO":  If SpMod(rLw(i).NormO, "laborxwert", rSc) Then Exit Do
   Case "AuftrHinw":  If SpMod(rLw(i).AuftrHinw, "laborxwert", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern()
 Dim i%, sql0$, rAF&
 On Error GoTo fehler
 sql0 = " insert " & sqlIgnore & "into " & kla & "laborxleist" & klz & " (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl) values"
 For i = 1 To UBound(rLL)
'  rLL(i).AktZeit = now()
  sql = sql0 & "(" & rLL(i).RefNr & ",'" & rLL(i).Abkü & "','" & rLL(i).Verf & "','" & rLL(i).EBM & "','" & rLL(i).goä & "','" & rLL(i).Anzahl & "')"
  Call DBCn.Execute(sql, rAF) ', , adAsyncExecute)
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rSc As ADODB.Recordset
 Set rSc = New ADODB.Recordset
 Set rSc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 Do While Not rSc.EOF
  Select Case rSc!column_name
   Case "Abkü":  If SpMod(rLL(i).Abkü, "laborxleist", rSc) Then Exit Do
   Case "Verf":  If SpMod(rLL(i).Verf, "laborxleist", rSc) Then Exit Do
   Case "EBM":  If SpMod(rLL(i).EBM, "laborxleist", rSc) Then Exit Do
   Case "goä":  If SpMod(rLL(i).goä, "laborxleist", rSc) Then Exit Do
   Case "Anzahl":  If SpMod(rLL(i).Anzahl, "laborxleist", rSc) Then Exit Do
  End Select
  rSc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function doSpeichern(frm As Lese)
 Dim rAF&
 On Error GoTo fehler
 Call namenSpeichern
 Call faelleSpeichern
   If Not lies.obMySQL Then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   End If
 Call auSpeichern
 Call briefeSpeichern
 Call diagnosenSpeichern
 Call dokumenteSpeichern
 Call eintraegeSpeichern
 Call forminhaltform_abkSpeichern
   If Not lies.obMySQL Then
    Call DBCn.CommitTrans
    Call DBCn.BeginTrans
   End If
 Call formulareSpeichern
 Call forminhkopfSpeichern
 Call forminhfeldSpeichern
 Call kheinweisSpeichern
 Call lbanforderungenSpeichern
 Call laborneuSpeichern
 Call leistungenSpeichern
 Call medplanSpeichern
 Call rezepteintraegeSpeichern
 Call rrSpeichern
 Call kvnrueSpeichern
 Call unbekannte_kennungenSpeichern
 Call dmpreiheSpeichern
 Call DBCn.Execute("update namen set aktZeit = " & datform(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_id, rAF)
 If rAF <> 1 Then
  frm.Ausgabe = "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_id & " " & rNa(0).NachName & " " & rNa(0).VorName & vbcrfl & altAusgabe
  altAusgabe = frm.Ausgabe
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

