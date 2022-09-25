Attribute VB_Name = "Typen"
Const BezFeh$ = "p:\ImportBeziehungsfehler.txt"
Dim sql$, T1!, T2!

Public Type namen
 Pat_id As Long 'Pat_ID 3 '3000
 lfdnr As Long 'lfdnr 3 'laufende Patientennummer
 NVorsatz As String 'NVorsatz 129 '3100
 Nachname As String 'Nachname 129 '3101
 Vorname As String 'Vorname 129 '3102
 GebDat As Date 'GebDat 135 '3103
 Straße As String 'Straße 129 '3107
 KVKStatus As String 'KVKStatus 129 '3108
 Geschlecht As String 'Geschlecht 129 '3110
 Plz As String 'Plz 129 '3112
 Ort As String 'Ort 129 '3113
 Weggeldzone As String 'Weggeldzone 129 '3631 (1) Weggeldzone mit Z
 WeggzZahl As Double 'WeggzZahl 131 '3631 (2) Weggeldzone, Zahl in Feld 2
 AufnDat As Date 'AufnDat 135 '3610
 LANR As String 'LANR 129 '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr As String 'BStNr 129 '3536 Betriebsstättennummer
 Titel As String 'Titel 129 '3104
 Versichertennummer As String 'Versichertennummer 129 '3105
 PrivatTel As String 'PrivatTel 129 '3629
 KVNr As String 'KVNr 129 '3630
 PrivatTel_2 As String 'PrivatTel_2 129 '3629
 PrivatFax As String 'PrivatFax 129 '3629
 DienstTel As String 'DienstTel 129 '3629OrdnerCtOrdnerCtOrdnerCt
 PrivatMobil As String 'PrivatMobil 129 '3629
 Email As String 'Email 129 'Email
 Arbeitgeber As String 'Arbeitgeber 129 '3625
 AnAllgda As Integer 'AnAllgda 11 'Anamnese allgemein da
 An1da As Integer 'An1da 11 'Anamnese S.1 da
 An2da As Integer 'An2da 11 'Anamnese S.2 da
 Checkda As Integer 'Checkda 11 'Checkliste da
 DMTypaD As String 'DMTypaD 129 'aus Diagnosen
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Cave As String 'Cave 129 '3654
 Notiz As String 'Notiz 129 '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>
 zubenach As String 'zubenach 129 '3633
 Verwandt As String 'Verwandt 129 '3632
 Sprache As String 'Sprache 129 '3628
 lAktTM As Date 'lAktTM 135 'letzte Aktualisierung in Turbomed
End Type

Public Type Faelle
 FID As Long 'FID 3 'Primärschlüssel
 Pat_id As Long 'Pat_ID 3 '3000
 Quartal As String 'Quartal 129 '4101
 Nachname As String 'Nachname 129 '3101
 Vorname As String 'Vorname 129 '3102
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
 IK As String 'IK 129 '4111 Krankenkassennummer (IK)
 KVKs As String 'KVKs 129 '4112 Versichertenstatus VK
 KVKserg As String 'KVKserg 129 '4113 Ost/West-Status VK
 GebOr As String 'GebOr 129 '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As String 'AbrGb 129 '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis As String 'PersKreis 129 '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz As String 'SKtZusatz 129 '4124 SKT-Zusatzangaben
 f4206 As String 'f4206 129 '4206, mutmasslicher Tag der Entbindung
 ÜwText As String 'ÜwText 129 '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 As Byte 'f4210 16 '4210, Ankreuzfeld LSR
 AkfHAH As Byte 'AkfHAH 16 '4211 Ankreuzfeld HAH
 AkfAB0 As Byte 'AkfAB0 16 '4212 Ankreuzfeld AB0.RH
 AkfAK As Byte 'AkfAK 16 '4213 Ankreuzfeld AK
 statNuller As String 'statNuller 129 '4216, nu bei Musterfrau 16 Nuller
 ÜbwV As String 'ÜbwV 129 '4218, überwiesen von Arztnummer
 AndÜw As String 'AndÜw 129 '4219, anderer Überweiser
 Übw As String 'Übw 129 '4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR As String 'ÜbwLANR 129 '4242 LANR des Überweisers
 ÜWZiel As String 'ÜWZiel 129 '4220 Überweisung an
 ÜWNNr As String 'ÜWNNr 129 '4231(4): KV-Nummer des Überweisers
 ÜWNaN As String 'ÜWNaN 129 '4231(3): Nachname des Überweisers
 ÜWTit As String 'ÜWTit 129 '4231(3): Titel des Überweisers
 ÜWVor As String 'ÜWVor 129 '4231(2): Vorname des Überweisers
 ÜWVsw As String 'ÜWVsw 129 '4231(2b): Vorsatzwort des Überweisers
 üwvid As Long 'üwvid 3 '4247 Bezug auf ueberwvon
 statKlasse As String 'statKlasse 129 '4236 Klasse bei Behandlung
 f4237 As String 'f4237 129 '4237 Krankenhausname
 statBehTage As Long 'statBehTage 3 '4238 Krankenhausaufenthalt
 SchGr As Double 'SchGr 131 '4239, Schein(unter)gruppe
 Weiterbeh As String 'Weiterbeh 129 '4243, Weiterbehandelnder
 PGeb As String 'PGeb 129 '4401, Praxisgebühr
 PGebErg As String 'PGebErg 129 '4402, ?
 Mahnfrist As String 'Mahnfrist 129 '4403, Mahnfrist bis
 GOÄKatNr As String 'GOÄKatNr 129 '4580 (1): Katalog-Nummer
 GOÄKatName As String 'GOÄKatName 129 '4580 (2): Privat-Abrechnungskatalog
 abrArzt As String 'abrArzt 129 '4585 abrechnender Arzt
 privVers As String 'privVers 129 '4586 private Versicherung
 AdNam As String 'AdNam 129 '4602(1) Name Rechnungsanschrift
 AdStr As String 'AdStr 129 '4602(2) Straße Rechnungsanschrift
 AdPlz As String 'AdPlz 129 '4602(3) PLZ Rechnungsanschrift
 AdOrt As String 'AdOrt 129 '4602(4) Ort Rechnungsanschrift
 BhFE As Date 'BhFE 135 '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As String 's8000 129 '8000, Satzidentifikation
 s8100 As String 's8100 129 '8100 Satzlänge
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 Fanf As Date 'Fanf 135 'Fallanfang
 altQuart As String 'altQuart 129 '
 QAnf As Date 'QAnf 135 'Quartalsanfang
 QEnd As Date 'QEnd 135 'Quartalsende
 QS As String 'QS 129 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 129 'Quartal des Behandlungsfallbeginns
 TherArt As Long 'TherArt 3 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)
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
 Art As String 'Art 129 '
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
 obDauer As Byte 'obDauer 16 'ob Dauerdiagnose
 Ausnahme As String 'Ausnahme 129 '3677 Ausnahme / Begründung für abweichendes Geschlecht
 intBemerk As String 'intBemerk 129 '6009 interne Bemerkung
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
 Art As String 'Art 129 '6330
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
 Nr As Integer 'Nr 2 '
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
 Abkü As String 'Abkü 129 '8410
 LangtextVW As Long 'LangtextVW 3 '8411
 Langtext As String
 Wert As String 'Wert 129 '8420
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
 Faktor As String 'Faktor 129 '5062 Multiplikator für GOÄ-Rechnung
 f5098 As String 'f5098 129 '5098 0000000000
 LANR As String 'LANR 129 '5099 LANR
 letzVorg As Date 'letzVorg 135 '5101 letzter Vorgang
 Ausn As String 'Ausn 129 '3677 Ausnahme/Begründung für abweichendes Geschlecht
 Beme As String 'Beme 129 '         Bemerkung
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
 Datum As Date 'Datum 135 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament As String 'Medikament 129 '
 MedAnfang As String 'MedAnfang 129 '
 FeldNr As Integer 'FeldNr 2 '
 mo As String 'mo 129 '
 mi As String 'mi 129 '
 nm As String 'nm 129 '
 ab As String 'ab 129 '
 Zn As String 'zn 129 '
 bBed As Integer 'bBed 11 '
 Bemerkung As String 'Bemerkung 129 '
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
 Inhalt As String 'Inhalt 129 'Inhalt Zeile zum Wiederauffinden
End Type

Public Type dmpreihe
 Abk As String 'Abk 129 'Abkürzung der DMP-Art
 Art As String 'Art 129 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As Date 'KarteiDatum 133 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As Date 'exportiert 135 'Datum des Exports
 DokuDatum As Date 'DokuDatum 135 'Datum der Dokumentation
 obvoll As Integer 'obvoll 11 'ob vollständig
 Nachname As String 'NachName 129 '
 Vorname As String 'VorName 129 '
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
 Arzt As String 'Arzt 129 ' 211 Ausführender Arzt
 LANR As String 'LANR 129 ' 212 LANR
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
 Zp As Date 'Zp 135 'Einlesezeitpunkt
 fertig As Integer 'fertig 11 'ob Einlesen fertig
End Type

Public Type laborxus
 RefNr As Long 'RefNr 3 'Bezug auf LaborWert
 DatID As Long 'DatID 3 'Bezug auf LaborEingelesen
 SatzID As Long 'SatzID 3 'Bezug auf LaborXSätze
 Satzart As String 'Satzart 129 '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge 129 '8100 Satzlänge (Turbomed)
 Auftragsnummer As String 'Auftragsnummer 129 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As String 'Auftragsschlüssel 129 '8311 Anforderungsnr d Labors (Turbomed)
 Eingang As Date 'Eingang 135 '8301 Eingangsdatum in Datumsform
 Berichtsdatum As String 'Berichtsdatum 129 '8302 Berichtsdatum
 Pat_id As Long 'Pat_id 3 '
 Nachname As String 'Nachname 129 '3101
 Vorname As String 'Vorname 129 '3102
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
 AfN As Integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
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
 Quelle As String 'Quelle 129 '8430 Probenmaterial-Bezeichnung (Turbomed)
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
 Verf As String 'Verf 129 '8434
 EBM As String 'EBM 129 '5001 GNR (Turbomed)
 goä As String 'goä 129 '8406
 Anzahl As String 'Anzahl 129 '5005
End Type

Public Type Anamnesebogen
 Prim As Long 'Prim 3 'Primärschlüssel
 Pat_id As Long 'Pat_id 3 '
 Nachname As String 'Nachname 129 '-
 Vorname As String 'Vorname 129 '
 NVorsatz As String 'NVorsatz 129 '
 Titel As String 'Titel 129 '
 Anrede As String 'Anrede 129 '
 GebDat As Date 'GebDat 135 ', geb.
 Tkz As Byte 'Tkz 16 'Tod-Kennzeichen
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
 Insulinpumpe As Byte 'Insulinpumpe 16 '^:
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
 Dialyse As Byte 'Dialyse 16 ',:
 Dialyse_seit As String 'Dialyse_seit 129 '<seit
 andere_Nierenerkrankung As String 'andere_Nierenerkrankung 129 ', andere Nierenerkrankung:
 Herzkrankheit As String 'Herzkrankheit 129 '^Herzkrankheit:
 Angina_pectoris As String 'Angina_pectoris 129 ',:
 Herzinfarkt As String 'Herzinfarkt 129 ',:
 Herzinfarkt_wann As String 'Herzinfarkt_wann 129 '<, wann:
 PTCA_oder_Stent As String 'PTCA_oder_Stent 129 ',:
 Bypass_kardial As Byte 'Bypass_kardial 16 ',:
 Bypass_wann As String 'Bypass_wann 129 '<, wann:
 Herzschwäche As String 'Herzschwäche 129 ',:
 Herzkrankheit_Beschreibung As String 'Herzkrankheit_Beschreibung 129 ', Beschreibung:
 Hirndurchblutungsstörung As String 'Hirndurchblutungsstörung 129 '^:
 Schlaganfall As String 'Schlaganfall 129 ',:
 Beindurchblutungsstörung As String 'Beindurchblutungsstörung 129 '^:
 Schaufensterkrankheit As String 'Schaufensterkrankheit 129 ',:
 Bypaß_peripher As Byte 'Bypaß_peripher 16 ',:
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
 Tabak As String 'Tabak 129 ',:
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
 obAn1eing As Byte 'obAn1eing 16 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As Byte 'obAn2eing 16 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As Byte 'obAnAeing 16 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck As Byte 'obCheck 16 'ob Checkliste vorliegt
 obBZausgew As Byte 'obBZausgew 16 'ob Blutzuckergerät ausgewechselt
 obOSaufgek As Byte 'obOSaufgek 16 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As Byte 'obPodAufgek 16 'ob über Podologie aufgeklärt
 obMBlAusgeh As Byte 'obMBlAusgeh 16 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek As String 'obSchulaufgek 129 'ob über Podologie aufgeklärt
 obDMPaufgekl As String 'obDMPaufgekl 129 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz As Byte 'obMedNetz 16 'ob von Med. Netz geschickt
 Hausarzt As String 'Hausarzt 129 'Hausarzt laut Anamnesebogen
 ob As Byte 'ob 16 'für verschiedene Aktionen
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
Public rKh() As kheinweis
Public rLb() As lbanforderungen
Public rLa() As laborneu
Public rLe() As leistungen
Public rMe() As medplan
Public rRe() As rezepteintraege
Public rRr() As RR
Public rKv() As kvnrue
Public rUn() As unbekannte_kennungen
Public rDm() As dmpreihe
Public rLs() As laborxsaetze
Public rLg() As laborxeingel
Public rLu() As laborxus
Public rLo() As laborxbakt
Public rLw() As laborxwert
Public rLL() As laborxleist
Public rAna() As Anamnesebogen

Public Function Tinit()
 Static wdh%
 ReDim rAna(0)
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
 Static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 wdh = -1
End Function ' Tinit

Public Function doEntleer(frm As Lese, Tbl$)
 Set rs = DBCn.Execute("select count(*) as ct from `" & Tbl & "`")
 frm.Ausgeb "Entleere: `" & DefDB(DBCn) & "`.`" & Tbl & "` (" & rs!ct & " Datensätze)", True
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call DBCn.Execute(sql, rAF) ' ,,adasyncexecute
 DoEvents
End Function ' doEntleer

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 If LVobMySQL <> 0 Then
'  DBCn.Execute ("reset master")
  DBCn.Execute ("reset query cache")
'  DBCn.Execute ("reset slave")
 End If
 Call doEntleer(frm, "dmpreihe")
 Call doEntleer(frm, "unbekannte kennungen")
 Call doEntleer(frm, "kvnrue")
 Call doEntleer(frm, "rr")
 Call doEntleer(frm, "rezepteintraege")
 Call doEntleer(frm, "medplan")
 Call doEntleer(frm, "leistungen")
 Call doEntleer(frm, "laborneu")
 Call doEntleer(frm, "lbanforderungen")
 Call doEntleer(frm, "kheinweis")
 Call doEntleer(frm, "forminhfeld")
 Call doEntleer(frm, "forminhkopf")
 Call doEntleer(frm, "formulare")
 Call doEntleer(frm, "forminhaltform_abk")
 Call doEntleer(frm, "eintraege")
 Call doEntleer(frm, "dokumente")
 Call doEntleer(frm, "diagnosen")
 Call doEntleer(frm, "briefe")
 Call doEntleer(frm, "au")
 Call doEntleer(frm, "faelle")
 Call doEntleer(frm, "namen")
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
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
 Call doEntleer(frm, "laborxleist")
 Call doEntleer(frm, "laborxwert")
 Call doEntleer(frm, "laborxbakt")
 Call doEntleer(frm, "laborxus")
 Call doEntleer(frm, "laborxeingel")
 Call doEntleer(frm, "laborxsaetze")
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' LabLösch


Public Function namenSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere namen"
 If Not AllePat Then
  sql = "delete from `namen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,WeggzZahl," & _
     "AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel," & _
     "PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos," & _
     "StByte,Cave,Notiz,zubenach,Verwandt,Sprache,lAktTM) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,WeggzZahl," & _
     "AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel," & _
     "PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos," & _
     "StByte,Cave,Notiz,zubenach,Verwandt,Sprache,lAktTM)              values"))
 For i = 0 To UBound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rNa(i).Pat_ID & "," & rNa(i).lfdnr & ",'" & rNa(i).NVorsatz & "','" & rNa(i).Nachname & "','" & rNa(i).Vorname & "'," & datform(rNa(i).GebDat) & ",'" & _
   rNa(i).Straße & "','" & rNa(i).KVKStatus & "','" & rNa(i).Geschlecht & "','" & rNa(i).Plz & "','" & rNa(i).Ort & "','" & _
   rNa(i).Weggeldzone & "'," & rNa(i).WeggzZahl & "," & datform(rNa(i).AufnDat) & ",'" & rNa(i).LANR & "','" & rNa(i).BStNr & "','" & rNa(i).Titel & "','" & _
   rNa(i).Versichertennummer & "','" & rNa(i).PrivatTel & "','" & rNa(i).KVNr & "','" & rNa(i).PrivatTel_2 & "','" & rNa(i).PrivatFax & "','" & _
   rNa(i).DienstTel & "','" & rNa(i).PrivatMobil & "','" & rNa(i).Email & "','" & rNa(i).Arbeitgeber & "'," & cstr(cint(rNa(i).AnAllgda)) & "," & cstr(cint( _
   rNa(i).An1da)) & "," & cstr(cint(rNa(i).An2da)) & "," & cstr(cint(rNa(i).Checkda)) & ",'" & rNa(i).DMTypaD & "'," & datform( _
    0 ) & "," & rNa(i).absPos & "," & rNa(i).StByte & ",'" & rNa(i).Cave & "','" & rNa(i).Notiz & "','" & rNa(i).zubenach & "','" & rNa(i).Verwandt & "','" & _
   rNa(i).Sprache & "'," & datform(rNa(i).lAktTM) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rNa(i).Pat_id, ",", rNa(i).lfdnr, ",'", rNa(i).NVorsatz, "','", rNa(i).Nachname, "','", rNa(i).Vorname, "',", DatForm( _
   rNa(i).GebDat), ",'", rNa(i).Straße, "','", rNa(i).KVKStatus, "','", rNa(i).Geschlecht, "','", rNa(i).Plz, "','", rNa(i).Ort, "','", _
   rNa(i).Weggeldzone, "',", rNa(i).WeggzZahl, ",", DatForm(rNa(i).AufnDat), ",'", rNa(i).LANR, "','", rNa(i).BStNr, "','", rNa(i).Titel, "','", _
   rNa(i).Versichertennummer, "','", rNa(i).PrivatTel, "','", rNa(i).KVNr, "','", rNa(i).PrivatTel_2, "','", rNa(i).PrivatFax, "','", _
   rNa(i).DienstTel, "','", rNa(i).PrivatMobil, "','", rNa(i).Email, "','", rNa(i).Arbeitgeber, "',", CStr(CInt(rNa(i).AnAllgda)), ",", CStr(CInt( _
   rNa(i).An1da)), ",", CStr(CInt(rNa(i).An2da)), ",", CStr(CInt(rNa(i).Checkda)), ",'", rNa(i).DMTypaD, "',", DatForm( _
    0), ",", rNa(i).absPos, ",", rNa(i).StByte, ",'", rNa(i).Cave, "','", rNa(i).Notiz, "','", rNa(i).zubenach, "','", rNa(i).Verwandt, "','", _
   rNa(i).Sprache, "',", DatForm(rNa(i).lAktTM), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rNa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(26)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rNa), i)
  If Len(rNa(k).NVorsatz) > maxi(0) Then maxi(0) = Len(rNa(k).NVorsatz)
  If Len(rNa(k).Nachname) > maxi(1) Then maxi(1) = Len(rNa(k).Nachname)
  If Len(rNa(k).Vorname) > maxi(2) Then maxi(2) = Len(rNa(k).Vorname)
  If Len(rNa(k).Straße) > maxi(3) Then maxi(3) = Len(rNa(k).Straße)
  If Len(rNa(k).KVKStatus) > maxi(4) Then maxi(4) = Len(rNa(k).KVKStatus)
  If Len(rNa(k).Geschlecht) > maxi(5) Then maxi(5) = Len(rNa(k).Geschlecht)
  If Len(rNa(k).Plz) > maxi(6) Then maxi(6) = Len(rNa(k).Plz)
  If Len(rNa(k).Ort) > maxi(7) Then maxi(7) = Len(rNa(k).Ort)
  If Len(rNa(k).Weggeldzone) > maxi(8) Then maxi(8) = Len(rNa(k).Weggeldzone)
  If Len(rNa(k).LANR) > maxi(9) Then maxi(9) = Len(rNa(k).LANR)
  If Len(rNa(k).BStNr) > maxi(10) Then maxi(10) = Len(rNa(k).BStNr)
  If Len(rNa(k).Titel) > maxi(11) Then maxi(11) = Len(rNa(k).Titel)
  If Len(rNa(k).Versichertennummer) > maxi(12) Then maxi(12) = Len(rNa(k).Versichertennummer)
  If Len(rNa(k).PrivatTel) > maxi(13) Then maxi(13) = Len(rNa(k).PrivatTel)
  If Len(rNa(k).KVNr) > maxi(14) Then maxi(14) = Len(rNa(k).KVNr)
  If Len(rNa(k).PrivatTel_2) > maxi(15) Then maxi(15) = Len(rNa(k).PrivatTel_2)
  If Len(rNa(k).PrivatFax) > maxi(16) Then maxi(16) = Len(rNa(k).PrivatFax)
  If Len(rNa(k).DienstTel) > maxi(17) Then maxi(17) = Len(rNa(k).DienstTel)
  If Len(rNa(k).PrivatMobil) > maxi(18) Then maxi(18) = Len(rNa(k).PrivatMobil)
  If Len(rNa(k).Email) > maxi(19) Then maxi(19) = Len(rNa(k).Email)
  If Len(rNa(k).Arbeitgeber) > maxi(20) Then maxi(20) = Len(rNa(k).Arbeitgeber)
  If Len(rNa(k).DMTypaD) > maxi(21) Then maxi(21) = Len(rNa(k).DMTypaD)
  If Len(rNa(k).Cave) > maxi(22) Then maxi(22) = Len(rNa(k).Cave)
  If Len(rNa(k).Notiz) > maxi(23) Then maxi(23) = Len(rNa(k).Notiz)
  If Len(rNa(k).zubenach) > maxi(24) Then maxi(24) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(25) Then maxi(25) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(26) Then maxi(26) = Len(rNa(k).Sprache)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rNa), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & Left$(rNa(k).NVorsatz, maxL) & "'", True: rNa(k).NVorsatz = Left$(rNa(k).NVorsatz, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & Left$(rNa(k).Nachname, maxL) & "'", True: rNa(k).Nachname = Left$(rNa(k).Nachname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & Left$(rNa(k).Vorname, maxL) & "'", True: rNa(k).Vorname = Left$(rNa(k).Vorname, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & Left$(rNa(k).Straße, maxL) & "'", True: rNa(k).Straße = Left$(rNa(k).Straße, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & Left$(rNa(k).KVKStatus, maxL) & "'", True: rNa(k).KVKStatus = Left$(rNa(k).KVKStatus, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).Geschlecht & "' -> '" & Left$(rNa(k).Geschlecht, maxL) & "'", True: rNa(k).Geschlecht = Left$(rNa(k).Geschlecht, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).Plz & "' -> '" & Left$(rNa(k).Plz, maxL) & "'", True: rNa(k).Plz = Left$(rNa(k).Plz, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).Ort & "' -> '" & Left$(rNa(k).Ort, maxL) & "'", True: rNa(k).Ort = Left$(rNa(k).Ort, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & Left$(rNa(k).Weggeldzone, maxL) & "'", True: rNa(k).Weggeldzone = Left$(rNa(k).Weggeldzone, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).LANR & "' -> '" & Left$(rNa(k).LANR, maxL) & "'", True: rNa(k).LANR = Left$(rNa(k).LANR, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rNa.BStNr: '" & rNa(k).BStNr & "' -> '" & Left$(rNa(k).BStNr, maxL) & "'", True: rNa(k).BStNr = Left$(rNa(k).BStNr, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rNa.Titel: '" & rNa(k).Titel & "' -> '" & Left$(rNa(k).Titel, maxL) & "'", True: rNa(k).Titel = Left$(rNa(k).Titel, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rNa.Versichertennummer: '" & rNa(k).Versichertennummer & "' -> '" & Left$(rNa(k).Versichertennummer, maxL) & "'", True: rNa(k).Versichertennummer = Left$(rNa(k).Versichertennummer, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel: '" & rNa(k).PrivatTel & "' -> '" & Left$(rNa(k).PrivatTel, maxL) & "'", True: rNa(k).PrivatTel = Left$(rNa(k).PrivatTel, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr: '" & rNa(k).KVNr & "' -> '" & Left$(rNa(k).KVNr, maxL) & "'", True: rNa(k).KVNr = Left$(rNa(k).KVNr, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel_2: '" & rNa(k).PrivatTel_2 & "' -> '" & Left$(rNa(k).PrivatTel_2, maxL) & "'", True: rNa(k).PrivatTel_2 = Left$(rNa(k).PrivatTel_2, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatFax: '" & rNa(k).PrivatFax & "' -> '" & Left$(rNa(k).PrivatFax, maxL) & "'", True: rNa(k).PrivatFax = Left$(rNa(k).PrivatFax, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rNa.DienstTel: '" & rNa(k).DienstTel & "' -> '" & Left$(rNa(k).DienstTel, maxL) & "'", True: rNa(k).DienstTel = Left$(rNa(k).DienstTel, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatMobil: '" & rNa(k).PrivatMobil & "' -> '" & Left$(rNa(k).PrivatMobil, maxL) & "'", True: rNa(k).PrivatMobil = Left$(rNa(k).PrivatMobil, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rNa.Email: '" & rNa(k).Email & "' -> '" & Left$(rNa(k).Email, maxL) & "'", True: rNa(k).Email = Left$(rNa(k).Email, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rNa.Arbeitgeber: '" & rNa(k).Arbeitgeber & "' -> '" & Left$(rNa(k).Arbeitgeber, maxL) & "'", True: rNa(k).Arbeitgeber = Left$(rNa(k).Arbeitgeber, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rNa.DMTypaD: '" & rNa(k).DMTypaD & "' -> '" & Left$(rNa(k).DMTypaD, maxL) & "'", True: rNa(k).DMTypaD = Left$(rNa(k).DMTypaD, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rNa.Cave: '" & rNa(k).Cave & "' -> '" & Left$(rNa(k).Cave, maxL) & "'", True: rNa(k).Cave = Left$(rNa(k).Cave, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rNa.Notiz: '" & rNa(k).Notiz & "' -> '" & Left$(rNa(k).Notiz, maxL) & "'", True: rNa(k).Notiz = Left$(rNa(k).Notiz, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL) & "'", True: rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL) & "'", True: rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL) & "'", True: rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
ElseIf Err.Number = -2147217871 Or Err.Number = -2147217859 Or Err.Number = -2147467259 Then
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere faelle"
 If Not AllePat Then
  sql = "delete from `faelle` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,PersKreis,SKtZusatz,f4206," & _
     "ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,AndÜw,Übw,ÜbwLANR," & _
     "ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,statKlasse,f4237,statBehTage," & _
     "SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam," & _
     "AdStr,AdPlz,AdOrt,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf," & _
     "QEnd,QS,QT,TherArt,StByte,absPos) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,PersKreis,SKtZusatz,f4206," & _
     "ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,AndÜw,Übw,ÜbwLANR," & _
     "ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,statKlasse,f4237,statBehTage," & _
     "SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam," & _
     "AdStr,AdPlz,AdOrt,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf," & _
     "QEnd,QS,QT,TherArt,StByte,absPos)    values"))
 For i = 1 To UBound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFa(i).Pat_ID & ",'" & rFa(i).Quartal & "','" & rFa(i).Nachname & "','" & rFa(i).Vorname & "'," & rFa(i).lfdnr & ",'" & rFa(i).TMFNr & "','" & _
   rFa(i).VKNr & "'," & datform(rFa(i).BhFB) & "," & datform(rFa(i).BhFE1) & "," & datform(rFa(i).BhFE2) & ",'" & rFa(i).f4202 & "'," & datform( _
   rFa(i).ausgst) & ",'" & rFa(i).KtrAbrB & "','" & rFa(i).AbrAr & "'," & datform(rFa(i).lVorl) & ",'" & rFa(i).IK & "','" & rFa(i).KVKs & "','" & _
   rFa(i).KVKserg & "','" & rFa(i).GebOr & "','" & rFa(i).AbrGb & "','" & rFa(i).PersKreis & "','" & rFa(i).SKtZusatz & "','" & rFa(i).f4206 & "','" & _
   rFa(i).ÜwText & "'," & rFa(i).f4210 & "," & rFa(i).AkfHAH & "," & rFa(i).AkfAB0 & "," & rFa(i).AkfAK & ",'" & rFa(i).statNuller & "','" & _
   rFa(i).ÜbwV & "','" & rFa(i).AndÜw & "','" & rFa(i).Übw & "','" & rFa(i).ÜbwLANR & "','" & rFa(i).ÜWZiel & "','" & rFa(i).ÜWNNr & "','" & _
   rFa(i).ÜWNaN & "','" & rFa(i).ÜWTit & "','" & rFa(i).ÜWVor & "','" & rFa(i).ÜWVsw & "'," & rFa(i).üwvid & ",'" & rFa(i).statKlasse & "','" & _
   rFa(i).f4237 & "'," & rFa(i).statBehTage & "," & rFa(i).SchGr & ",'" & rFa(i).Weiterbeh & "','" & rFa(i).PGeb & "','" & rFa(i).PGebErg & "','" & _
   rFa(i).Mahnfrist & "','" & rFa(i).GOÄKatNr & "','" & rFa(i).GOÄKatName & "','" & rFa(i).abrArzt & "','" & rFa(i).privVers & "','" & rFa(i).AdNam & "','" & _
   rFa(i).AdStr & "','" & rFa(i).AdPlz & "','" & rFa(i).AdOrt & "'," & datform(rFa(i).BhFE) & ",'" & rFa(i).s8000 & "','" & rFa(i).s8100 & "'," & datform( _
   rFa(i).AktZeit) & "," & datform(rFa(i).Fanf) & ",'" & rFa(i).altQuart & "'," & datform(rFa(i).QAnf) & "," & datform( _
   rFa(i).QEnd) & ",'" & rFa(i).QS & "','" & rFa(i).QT & "'," & rFa(i).TherArt & "," & rFa(i).StByte & "," & rFa(i).absPos & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFa(i).Pat_id, ",'", rFa(i).Quartal, "','", rFa(i).Nachname, "','", rFa(i).Vorname, "',", rFa(i).lfdnr, ",'", rFa(i).TMFNr, "','", _
   rFa(i).VKNr, "',", DatForm(rFa(i).BhFB), ",", DatForm(rFa(i).BhFE1), ",", DatForm(rFa(i).BhFE2), ",'", rFa(i).f4202, "',", DatForm( _
   rFa(i).ausgst), ",'", rFa(i).KtrAbrB, "','", rFa(i).AbrAr, "',", DatForm(rFa(i).lVorl), ",'", rFa(i).IK, "','", rFa(i).KVKs, "','", _
   rFa(i).KVKserg, "','", rFa(i).GebOr, "','", rFa(i).AbrGb, "','", rFa(i).PersKreis, "','", rFa(i).SKtZusatz, "','", rFa(i).f4206, "','", _
   rFa(i).ÜwText, "',", rFa(i).f4210, ",", rFa(i).AkfHAH, ",", rFa(i).AkfAB0, ",", rFa(i).AkfAK, ",'", rFa(i).statNuller, "','", _
   rFa(i).ÜbwV, "','", rFa(i).AndÜw, "','", rFa(i).Übw, "','", rFa(i).ÜbwLANR, "','", rFa(i).ÜWZiel, "','", rFa(i).ÜWNNr, "','", _
   rFa(i).ÜWNaN, "','", rFa(i).ÜWTit, "','", rFa(i).ÜWVor, "','", rFa(i).ÜWVsw, "',", rFa(i).üwvid, ",'", rFa(i).statKlasse, "','", _
   rFa(i).f4237, "',", rFa(i).statBehTage, ",", rFa(i).SchGr, ",'", rFa(i).Weiterbeh, "','", rFa(i).PGeb, "','", rFa(i).PGebErg, "','", _
   rFa(i).Mahnfrist, "','", rFa(i).GOÄKatNr, "','", rFa(i).GOÄKatName, "','", rFa(i).abrArzt, "','", rFa(i).privVers, "','", _
   rFa(i).AdNam, "','", rFa(i).AdStr, "','", rFa(i).AdPlz, "','", rFa(i).AdOrt, "',", DatForm(rFa(i).BhFE), ",'", rFa(i).s8000, "','", _
   rFa(i).s8100, "',", DatForm(rFa(i).AktZeit), ",", DatForm(rFa(i).Fanf), ",'", rFa(i).altQuart, "',", DatForm(rFa(i).QAnf), ",", DatForm( _
   rFa(i).QEnd), ",'", rFa(i).QS, "','", rFa(i).QT, "',", rFa(i).TherArt, ",", rFa(i).StByte, ",", rFa(i).absPos, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_id & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & DatForm(rFa(i).BhFB) & " and bhfe1 = " & DatForm(rFa(i).BhFE1) & " and ausgst = " & DatForm(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID & " -> " & rs!FID, True
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
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(46)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFa), i)
  If Len(rFa(k).Quartal) > maxi(0) Then maxi(0) = Len(rFa(k).Quartal)
  If Len(rFa(k).Nachname) > maxi(1) Then maxi(1) = Len(rFa(k).Nachname)
  If Len(rFa(k).Vorname) > maxi(2) Then maxi(2) = Len(rFa(k).Vorname)
  If Len(rFa(k).TMFNr) > maxi(3) Then maxi(3) = Len(rFa(k).TMFNr)
  If Len(rFa(k).VKNr) > maxi(4) Then maxi(4) = Len(rFa(k).VKNr)
  If Len(rFa(k).f4202) > maxi(5) Then maxi(5) = Len(rFa(k).f4202)
  If Len(rFa(k).KtrAbrB) > maxi(6) Then maxi(6) = Len(rFa(k).KtrAbrB)
  If Len(rFa(k).AbrAr) > maxi(7) Then maxi(7) = Len(rFa(k).AbrAr)
  If Len(rFa(k).IK) > maxi(8) Then maxi(8) = Len(rFa(k).IK)
  If Len(rFa(k).KVKs) > maxi(9) Then maxi(9) = Len(rFa(k).KVKs)
  If Len(rFa(k).KVKserg) > maxi(10) Then maxi(10) = Len(rFa(k).KVKserg)
  If Len(rFa(k).GebOr) > maxi(11) Then maxi(11) = Len(rFa(k).GebOr)
  If Len(rFa(k).AbrGb) > maxi(12) Then maxi(12) = Len(rFa(k).AbrGb)
  If Len(rFa(k).PersKreis) > maxi(13) Then maxi(13) = Len(rFa(k).PersKreis)
  If Len(rFa(k).SKtZusatz) > maxi(14) Then maxi(14) = Len(rFa(k).SKtZusatz)
  If Len(rFa(k).f4206) > maxi(15) Then maxi(15) = Len(rFa(k).f4206)
  If Len(rFa(k).ÜwText) > maxi(16) Then maxi(16) = Len(rFa(k).ÜwText)
  If Len(rFa(k).statNuller) > maxi(17) Then maxi(17) = Len(rFa(k).statNuller)
  If Len(rFa(k).ÜbwV) > maxi(18) Then maxi(18) = Len(rFa(k).ÜbwV)
  If Len(rFa(k).AndÜw) > maxi(19) Then maxi(19) = Len(rFa(k).AndÜw)
  If Len(rFa(k).Übw) > maxi(20) Then maxi(20) = Len(rFa(k).Übw)
  If Len(rFa(k).ÜbwLANR) > maxi(21) Then maxi(21) = Len(rFa(k).ÜbwLANR)
  If Len(rFa(k).ÜWZiel) > maxi(22) Then maxi(22) = Len(rFa(k).ÜWZiel)
  If Len(rFa(k).ÜWNNr) > maxi(23) Then maxi(23) = Len(rFa(k).ÜWNNr)
  If Len(rFa(k).ÜWNaN) > maxi(24) Then maxi(24) = Len(rFa(k).ÜWNaN)
  If Len(rFa(k).ÜWTit) > maxi(25) Then maxi(25) = Len(rFa(k).ÜWTit)
  If Len(rFa(k).ÜWVor) > maxi(26) Then maxi(26) = Len(rFa(k).ÜWVor)
  If Len(rFa(k).ÜWVsw) > maxi(27) Then maxi(27) = Len(rFa(k).ÜWVsw)
  If Len(rFa(k).statKlasse) > maxi(28) Then maxi(28) = Len(rFa(k).statKlasse)
  If Len(rFa(k).f4237) > maxi(29) Then maxi(29) = Len(rFa(k).f4237)
  If Len(rFa(k).Weiterbeh) > maxi(30) Then maxi(30) = Len(rFa(k).Weiterbeh)
  If Len(rFa(k).PGeb) > maxi(31) Then maxi(31) = Len(rFa(k).PGeb)
  If Len(rFa(k).PGebErg) > maxi(32) Then maxi(32) = Len(rFa(k).PGebErg)
  If Len(rFa(k).Mahnfrist) > maxi(33) Then maxi(33) = Len(rFa(k).Mahnfrist)
  If Len(rFa(k).GOÄKatNr) > maxi(34) Then maxi(34) = Len(rFa(k).GOÄKatNr)
  If Len(rFa(k).GOÄKatName) > maxi(35) Then maxi(35) = Len(rFa(k).GOÄKatName)
  If Len(rFa(k).abrArzt) > maxi(36) Then maxi(36) = Len(rFa(k).abrArzt)
  If Len(rFa(k).privVers) > maxi(37) Then maxi(37) = Len(rFa(k).privVers)
  If Len(rFa(k).AdNam) > maxi(38) Then maxi(38) = Len(rFa(k).AdNam)
  If Len(rFa(k).AdStr) > maxi(39) Then maxi(39) = Len(rFa(k).AdStr)
  If Len(rFa(k).AdPlz) > maxi(40) Then maxi(40) = Len(rFa(k).AdPlz)
  If Len(rFa(k).AdOrt) > maxi(41) Then maxi(41) = Len(rFa(k).AdOrt)
  If Len(rFa(k).s8000) > maxi(42) Then maxi(42) = Len(rFa(k).s8000)
  If Len(rFa(k).s8100) > maxi(43) Then maxi(43) = Len(rFa(k).s8100)
  If Len(rFa(k).altQuart) > maxi(44) Then maxi(44) = Len(rFa(k).altQuart)
  If Len(rFa(k).QS) > maxi(45) Then maxi(45) = Len(rFa(k).QS)
  If Len(rFa(k).QT) > maxi(46) Then maxi(46) = Len(rFa(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFa), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & Left$(rFa(k).Quartal, maxL) & "'", True: rFa(k).Quartal = Left$(rFa(k).Quartal, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & Left$(rFa(k).Nachname, maxL) & "'", True: rFa(k).Nachname = Left$(rFa(k).Nachname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & Left$(rFa(k).Vorname, maxL) & "'", True: rFa(k).Vorname = Left$(rFa(k).Vorname, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & Left$(rFa(k).TMFNr, maxL) & "'", True: rFa(k).TMFNr = Left$(rFa(k).TMFNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & Left$(rFa(k).VKNr, maxL) & "'", True: rFa(k).VKNr = Left$(rFa(k).VKNr, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4202: '" & rFa(k).f4202 & "' -> '" & Left$(rFa(k).f4202, maxL) & "'", True: rFa(k).f4202 = Left$(rFa(k).f4202, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & Left$(rFa(k).KtrAbrB, maxL) & "'", True: rFa(k).KtrAbrB = Left$(rFa(k).KtrAbrB, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & Left$(rFa(k).AbrAr, maxL) & "'", True: rFa(k).AbrAr = Left$(rFa(k).AbrAr, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & Left$(rFa(k).IK, maxL) & "'", True: rFa(k).IK = Left$(rFa(k).IK, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & Left$(rFa(k).KVKs, maxL) & "'", True: rFa(k).KVKs = Left$(rFa(k).KVKs, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & Left$(rFa(k).KVKserg, maxL) & "'", True: rFa(k).KVKserg = Left$(rFa(k).KVKserg, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & Left$(rFa(k).GebOr, maxL) & "'", True: rFa(k).GebOr = Left$(rFa(k).GebOr, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & Left$(rFa(k).AbrGb, maxL) & "'", True: rFa(k).AbrGb = Left$(rFa(k).AbrGb, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & Left$(rFa(k).PersKreis, maxL) & "'", True: rFa(k).PersKreis = Left$(rFa(k).PersKreis, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & Left$(rFa(k).SKtZusatz, maxL) & "'", True: rFa(k).SKtZusatz = Left$(rFa(k).SKtZusatz, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4206: '" & rFa(k).f4206 & "' -> '" & Left$(rFa(k).f4206, maxL) & "'", True: rFa(k).f4206 = Left$(rFa(k).f4206, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & Left$(rFa(k).ÜwText, maxL) & "'", True: rFa(k).ÜwText = Left$(rFa(k).ÜwText, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & Left$(rFa(k).statNuller, maxL) & "'", True: rFa(k).statNuller = Left$(rFa(k).statNuller, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & Left$(rFa(k).ÜbwV, maxL) & "'", True: rFa(k).ÜbwV = Left$(rFa(k).ÜbwV, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & Left$(rFa(k).AndÜw, maxL) & "'", True: rFa(k).AndÜw = Left$(rFa(k).AndÜw, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übw: '" & rFa(k).Übw & "' -> '" & Left$(rFa(k).Übw, maxL) & "'", True: rFa(k).Übw = Left$(rFa(k).Übw, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & Left$(rFa(k).ÜbwLANR, maxL) & "'", True: rFa(k).ÜbwLANR = Left$(rFa(k).ÜbwLANR, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & Left$(rFa(k).ÜWZiel, maxL) & "'", True: rFa(k).ÜWZiel = Left$(rFa(k).ÜWZiel, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & Left$(rFa(k).ÜWNNr, maxL) & "'", True: rFa(k).ÜWNNr = Left$(rFa(k).ÜWNNr, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & Left$(rFa(k).ÜWNaN, maxL) & "'", True: rFa(k).ÜWNaN = Left$(rFa(k).ÜWNaN, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & Left$(rFa(k).ÜWTit, maxL) & "'", True: rFa(k).ÜWTit = Left$(rFa(k).ÜWTit, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & Left$(rFa(k).ÜWVor, maxL) & "'", True: rFa(k).ÜWVor = Left$(rFa(k).ÜWVor, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & Left$(rFa(k).ÜWVsw, maxL) & "'", True: rFa(k).ÜWVsw = Left$(rFa(k).ÜWVsw, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & Left$(rFa(k).statKlasse, maxL) & "'", True: rFa(k).statKlasse = Left$(rFa(k).statKlasse, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4237: '" & rFa(k).f4237 & "' -> '" & Left$(rFa(k).f4237, maxL) & "'", True: rFa(k).f4237 = Left$(rFa(k).f4237, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & Left$(rFa(k).Weiterbeh, maxL) & "'", True: rFa(k).Weiterbeh = Left$(rFa(k).Weiterbeh, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & Left$(rFa(k).PGeb, maxL) & "'", True: rFa(k).PGeb = Left$(rFa(k).PGeb, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & Left$(rFa(k).PGebErg, maxL) & "'", True: rFa(k).PGebErg = Left$(rFa(k).PGebErg, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & Left$(rFa(k).Mahnfrist, maxL) & "'", True: rFa(k).Mahnfrist = Left$(rFa(k).Mahnfrist, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & Left$(rFa(k).GOÄKatNr, maxL) & "'", True: rFa(k).GOÄKatNr = Left$(rFa(k).GOÄKatNr, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & Left$(rFa(k).GOÄKatName, maxL) & "'", True: rFa(k).GOÄKatName = Left$(rFa(k).GOÄKatName, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & Left$(rFa(k).abrArzt, maxL) & "'", True: rFa(k).abrArzt = Left$(rFa(k).abrArzt, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & Left$(rFa(k).privVers, maxL) & "'", True: rFa(k).privVers = Left$(rFa(k).privVers, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & Left$(rFa(k).AdNam, maxL) & "'", True: rFa(k).AdNam = Left$(rFa(k).AdNam, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & Left$(rFa(k).AdStr, maxL) & "'", True: rFa(k).AdStr = Left$(rFa(k).AdStr, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & Left$(rFa(k).AdPlz, maxL) & "'", True: rFa(k).AdPlz = Left$(rFa(k).AdPlz, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & Left$(rFa(k).AdOrt, maxL) & "'", True: rFa(k).AdOrt = Left$(rFa(k).AdOrt, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & Left$(rFa(k).s8000, maxL) & "'", True: rFa(k).s8000 = Left$(rFa(k).s8000, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & Left$(rFa(k).s8100, maxL) & "'", True: rFa(k).s8100 = Left$(rFa(k).s8100, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & Left$(rFa(k).altQuart, maxL) & "'", True: rFa(k).altQuart = Left$(rFa(k).altQuart, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & Left$(rFa(k).QS, maxL) & "'", True: rFa(k).QS = Left$(rFa(k).QS, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & Left$(rFa(k).QT, maxL) & "'", True: rFa(k).QT = Left$(rFa(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "insert into `kassenliste`(`GO`,`VK`,`IK`) values (" & "'" & rFa(i).GOÄKatName & "', '" & rFa(i).VKNr & "', '" & rFa(i).IK & "')"
 Call DBCn.Execute(sqlquer, rAF)
 Resume
End If ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere au"
 If Not AllePat Then
  sql = "delete from `au` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            values"))
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rAu(i).FID & "," & rAu(i).Pat_ID & "," & datform(rAu(i).ZeitPunkt) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "','" & rAu(i).ICDs & "'," & _
   rAu(i).absPos & "," & datform(rAu(i).AktZeit) & "," & rAu(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rAu(i).FID, ",", rAu(i).Pat_id, ",", DatForm(rAu(i).Zeitpunkt), ",'", rAu(i).Beginn, "','", rAu(i).Ende, "','", rAu(i).ICDs, "',", _
   rAu(i).absPos, ",", DatForm(rAu(i).AktZeit), ",", rAu(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rAu) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rAu), i)
  If Len(rAu(k).Beginn) > maxi(0) Then maxi(0) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(1) Then maxi(1) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(2) Then maxi(2) = Len(rAu(k).ICDs)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rAu), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & Left$(rAu(k).Beginn, maxL) & "'", True: rAu(k).Beginn = Left$(rAu(k).Beginn, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & Left$(rAu(k).Ende, maxL) & "'", True: rAu(k).Ende = Left$(rAu(k).Ende, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & Left$(rAu(k).ICDs, maxL) & "'", True: rAu(k).ICDs = Left$(rAu(k).ICDs, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere briefe"
 If Not AllePat Then
  sql = "delete from `briefe` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte)           values"))
 For i = 1 To UBound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rBr(i).FID & "," & rBr(i).Pat_ID & "," & datform(rBr(i).ZeitPunkt) & ",'" & rBr(i).Pfad & "','" & rBr(i).Art & "','" & rBr(i).Name & "','" & _
   rBr(i).Typ & "'," & datform(rBr(i).AktZeit) & "," & rBr(i).DokGroe & ",'" & rBr(i).QS & "','" & rBr(i).QT & "'," & rBr(i).absPos & "," & _
   rBr(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rBr(i).FID, ",", rBr(i).Pat_id, ",", DatForm(rBr(i).Zeitpunkt), ",'", rBr(i).Pfad, "','", rBr(i).Art, "','", rBr(i).Name, "','", _
   rBr(i).Typ, "',", DatForm(rBr(i).AktZeit), ",", rBr(i).DokGroe, ",'", rBr(i).QS, "','", rBr(i).QT, "',", rBr(i).absPos, ",", _
   rBr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rBr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147467259 Then ' forein key fehler
 GoTo entscheid
ElseIf Err.Number = 0 Then ' server has gone away
 Dim cnstrzwi$
 cnstrzwi = DBCn
 Set DBCn = Nothing
 Err.Clear
 DBCn.Open cnstrzwi
 Resume
ElseIf Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rBr), i)
  If Len(rBr(k).Pfad) > maxi(0) Then maxi(0) = Len(rBr(k).Pfad)
  If Len(rBr(k).Art) > maxi(1) Then maxi(1) = Len(rBr(k).Art)
  If Len(rBr(k).Name) > maxi(2) Then maxi(2) = Len(rBr(k).Name)
  If Len(rBr(k).Typ) > maxi(3) Then maxi(3) = Len(rBr(k).Typ)
  If Len(rBr(k).QS) > maxi(4) Then maxi(4) = Len(rBr(k).QS)
  If Len(rBr(k).QT) > maxi(5) Then maxi(5) = Len(rBr(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rBr), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL) & "'", True: rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).Art & "' -> '" & Left$(rBr(k).Art, maxL) & "'", True: rBr(k).Art = Left$(rBr(k).Art, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).Name & "' -> '" & Left$(rBr(k).Name, maxL) & "'", True: rBr(k).Name = Left$(rBr(k).Name, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & Left$(rBr(k).Typ, maxL) & "'", True: rBr(k).Typ = Left$(rBr(k).Typ, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & Left$(rBr(k).QS, maxL) & "'", True: rBr(k).QS = Left$(rBr(k).QS, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & Left$(rBr(k).QT, maxL) & "'", True: rBr(k).QT = Left$(rBr(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
entscheid:
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere diagnosen"
 If Not AllePat Then
  sql = "delete from `diagnosen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,Ausnahme,intBemerk,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,Ausnahme,intBemerk,absPos," & _
     "AktZeit,StByte)        values"))
 For i = 1 To UBound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rDi(i).FID & "," & rDi(i).Pat_id & ",'" & rDi(i).GesName & "'," & datform(rDi(i).DiagDatum) & ",'" & rDi(i).DiagSicherheit & "','" & _
   rDi(i).DiagText & "','" & rDi(i).DiagSeite & "','" & rDi(i).DiagAttr & "','" & rDi(i).ICD & "'," & rDi(i).obDauer & ",'" & rDi(i).Ausnahme & "','" & _
   rDi(i).intBemerk & "'," & rDi(i).absPos & "," & datform(rDi(i).AktZeit) & "," & rDi(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDi(i).FID, ",", rDi(i).Pat_id, ",'", rDi(i).GesName, "',", DatForm(rDi(i).DiagDatum), ",'", rDi(i).DiagSicherheit, "','", _
   rDi(i).DiagText, "','", rDi(i).DiagSeite, "','", rDi(i).DiagAttr, "','", rDi(i).ICD, "',", rDi(i).obDauer, ",'", rDi(i).Ausnahme, "','", _
   rDi(i).intBemerk, "',", rDi(i).absPos, ",", DatForm(rDi(i).AktZeit), ",", rDi(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(7)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDi), i)
  If Len(rDi(k).GesName) > maxi(0) Then maxi(0) = Len(rDi(k).GesName)
  If Len(rDi(k).DiagSicherheit) > maxi(1) Then maxi(1) = Len(rDi(k).DiagSicherheit)
  If Len(rDi(k).DiagText) > maxi(2) Then maxi(2) = Len(rDi(k).DiagText)
  If Len(rDi(k).DiagSeite) > maxi(3) Then maxi(3) = Len(rDi(k).DiagSeite)
  If Len(rDi(k).DiagAttr) > maxi(4) Then maxi(4) = Len(rDi(k).DiagAttr)
  If Len(rDi(k).ICD) > maxi(5) Then maxi(5) = Len(rDi(k).ICD)
  If Len(rDi(k).Ausnahme) > maxi(6) Then maxi(6) = Len(rDi(k).Ausnahme)
  If Len(rDi(k).intBemerk) > maxi(7) Then maxi(7) = Len(rDi(k).intBemerk)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDi), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.GesName: '" & rDi(k).GesName & "' -> '" & Left$(rDi(k).GesName, maxL) & "'", True: rDi(k).GesName = Left$(rDi(k).GesName, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & Left$(rDi(k).DiagSicherheit, maxL) & "'", True: rDi(k).DiagSicherheit = Left$(rDi(k).DiagSicherheit, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & Left$(rDi(k).DiagText, maxL) & "'", True: rDi(k).DiagText = Left$(rDi(k).DiagText, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & Left$(rDi(k).DiagSeite, maxL) & "'", True: rDi(k).DiagSeite = Left$(rDi(k).DiagSeite, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & Left$(rDi(k).DiagAttr, maxL) & "'", True: rDi(k).DiagAttr = Left$(rDi(k).DiagAttr, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & Left$(rDi(k).ICD, maxL) & "'", True: rDi(k).ICD = Left$(rDi(k).ICD, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.Ausnahme: '" & rDi(k).Ausnahme & "' -> '" & Left$(rDi(k).Ausnahme, maxL) & "'", True: rDi(k).Ausnahme = Left$(rDi(k).Ausnahme, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & Left$(rDi(k).intBemerk, maxL) & "'", True: rDi(k).intBemerk = Left$(rDi(k).intBemerk, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere dokumente"
 If Not AllePat Then
  sql = "delete from `dokumente` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte)         values"))
 For i = 1 To UBound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rDo(i).FID & "," & rDo(i).Pat_ID & "," & datform(rDo(i).ZeitPunkt) & ",'" & rDo(i).DokPfad & "','" & rDo(i).DokArt & "','" & rDo(i).DokName & "'," & datform( _
   rDo(i).Quelldatum) & "," & rDo(i).absPos & "," & datform(rDo(i).AktZeit) & "," & rDo(i).DokGroe & ",'" & rDo(i).QS & "','" & _
   rDo(i).QT & "'," & rDo(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDo(i).FID, ",", rDo(i).Pat_id, ",", DatForm(rDo(i).Zeitpunkt), ",'", rDo(i).DokPfad, "','", rDo(i).DokArt, "','", _
   rDo(i).DokName, "',", DatForm(rDo(i).Quelldatum), ",", rDo(i).absPos, ",", DatForm(rDo(i).AktZeit), ",", rDo(i).DokGroe, ",'", rDo(i).QS, "','", _
   rDo(i).QT, "',", rDo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDo), i)
  If Len(rDo(k).DokPfad) > maxi(0) Then maxi(0) = Len(rDo(k).DokPfad)
  If Len(rDo(k).DokArt) > maxi(1) Then maxi(1) = Len(rDo(k).DokArt)
  If Len(rDo(k).DokName) > maxi(2) Then maxi(2) = Len(rDo(k).DokName)
  If Len(rDo(k).QS) > maxi(3) Then maxi(3) = Len(rDo(k).QS)
  If Len(rDo(k).QT) > maxi(4) Then maxi(4) = Len(rDo(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDo), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & Left$(rDo(k).DokPfad, maxL) & "'", True: rDo(k).DokPfad = Left$(rDo(k).DokPfad, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & Left$(rDo(k).DokArt, maxL) & "'", True: rDo(k).DokArt = Left$(rDo(k).DokArt, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & Left$(rDo(k).DokName, maxL) & "'", True: rDo(k).DokName = Left$(rDo(k).DokName, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & Left$(rDo(k).QS, maxL) & "'", True: rDo(k).QS = Left$(rDo(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & Left$(rDo(k).QT, maxL) & "'", True: rDo(k).QT = Left$(rDo(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere eintraege"
 If Not AllePat Then
  sql = "delete from `eintraege` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte)            values"))
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rEi(i).FID & "," & rEi(i).Pat_ID & "," & datform(rEi(i).ZeitPunkt) & ",'" & rEi(i).Art & "','" & rEi(i).Inhalt & "'," & rEi(i).absPos & "," & datform( _
   rEi(i).AktZeit) & ",'" & rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rEi(i).FID, ",", rEi(i).Pat_id, ",", DatForm(rEi(i).Zeitpunkt), ",'", rEi(i).Art, "','", rEi(i).Inhalt, "',", rEi(i).absPos, ",", DatForm( _
   rEi(i).AktZeit), ",'", rEi(i).QS, "','", rEi(i).QT, "',", rEi(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rEi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rEi), i)
  If Len(rEi(k).Art) > maxi(0) Then maxi(0) = Len(rEi(k).Art)
  If Len(rEi(k).Inhalt) > maxi(1) Then maxi(1) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(2) Then maxi(2) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(3) Then maxi(3) = Len(rEi(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rEi), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).Art & "' -> '" & Left$(rEi(k).Art, maxL) & "'", True: rEi(k).Art = Left$(rEi(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & Left$(rEi(k).Inhalt, maxL) & "'", True: rEi(k).Inhalt = Left$(rEi(k).Inhalt, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & Left$(rEi(k).QS, maxL) & "'", True: rEi(k).QS = Left$(rEi(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & Left$(rEi(k).QT, maxL) & "'", True: rEi(k).QT = Left$(rEi(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere forminhaltform_abk"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhaltform_abk` (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhaltform_abk` (Form_Abk)             values"))
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
'   sql = sql0 & "('" & rFi(i).Form_Abk & "')"
  If Lese.dlg.SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rFi(i).Form_Abk, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rFi1 = UBound(rFi)
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rFi1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFi), i)
  If Len(rFi(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFi(k).Form_Abk)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rFi1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFi), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFi.Form_Abk: '" & rFi(k).Form_Abk & "' -> '" & Left$(rFi(k).Form_Abk, maxL) & "'", True: rFi(k).Form_Abk = Left$(rFi(k).Form_Abk, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere formulare"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      values"))
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFo(i).FormID & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "','" & rFo(i).FormVorl & "'," & datform(rFo(i).AktZeit) & "," & _
   rFo(i).absPos & "," & rFo(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFo(i).FormID, ",'", rFo(i).Form_Abk, "','", rFo(i).FormBez, "','", rFo(i).FormVorl, "',", DatForm(rFo(i).AktZeit), ",", _
   rFo(i).absPos, ",", rFo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rFo1 = UBound(rFo)
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rFo1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFo), i)
  If Len(rFo(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFo(k).Form_Abk)
  If Len(rFo(k).FormBez) > maxi(1) Then maxi(1) = Len(rFo(k).FormBez)
  If Len(rFo(k).FormVorl) > maxi(2) Then maxi(2) = Len(rFo(k).FormVorl)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rFo1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFo), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & Left$(rFo(k).Form_Abk, maxL) & "'", True: rFo(k).Form_Abk = Left$(rFo(k).Form_Abk, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & Left$(rFo(k).FormBez, maxL) & "'", True: rFo(k).FormBez = Left$(rFo(k).FormBez, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & Left$(rFo(k).FormVorl, maxL) & "'", True: rFo(k).FormVorl = Left$(rFo(k).FormVorl, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere forminhkopf"
 If Not AllePat Then
'  sql = "delete FROM `forminhfeld` where foid in (select foid from `forminhkopf` where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "delete from `forminhkopf` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge)       values"))
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFr(i).FoID & "," & rFr(i).FID & "," & rFr(i).Pat_ID & "," & rFr(i).Form_ID & "," & datform(rFr(i).ZeitPunkt) & "," & rFr(i).AbsPos & "," & datform( _
   rFr(i).AktZeit) & "," & rFr(i).StByte & ",'" & rFr(i).Satzart & "','" & rFr(i).Satzlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFr(i).FoID, ",", rFr(i).FID, ",", rFr(i).Pat_id, ",", rFr(i).Form_ID, ",", DatForm(rFr(i).Zeitpunkt), ",", rFr(i).absPos, ",", DatForm( _
   rFr(i).AktZeit), ",", rFr(i).StByte, ",'", rFr(i).Satzart, "','", rFr(i).Satzlänge, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFr), i)
  If Len(rFr(k).Satzart) > maxi(0) Then maxi(0) = Len(rFr(k).Satzart)
  If Len(rFr(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rFr(k).Satzlänge)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFr), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & Left$(rFr(k).Satzart, maxL) & "'", True: rFr(k).Satzart = Left$(rFr(k).Satzart, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & Left$(rFr(k).Satzlänge, maxL) & "'", True: rFr(k).Satzlänge = Left$(rFr(k).Satzlänge, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere forminhfeld"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      values"))
 For i = 1 To UBound(rFm)
'  rFm(i).AktZeit = now()
'   sql = sql0 & "(" & rFm(i).FoID & "," & rFm(i).Nr & "," & rFm(i).FeldNr & "," & rFm(i).FeldVW & "," & rFm(i).FeldInhVW & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFm(i).FoID, ",", rFm(i).Nr, ",", rFm(i).FeldNr, ",", rFm(i).FeldVW, ",", rFm(i).FeldInhVW, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFm) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(-1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFm), i)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhfeld", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFm), i)
      Select Case m
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere kheinweis"
 If Not AllePat Then
  sql = "delete from `kheinweis` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               values"))
 For i = 1 To UBound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKh(i).FID & "," & rKh(i).Pat_ID & "," & datform(rKh(i).ZeitPunkt) & ",'" & rKh(i).Ziel & "','" & rKh(i).Diagnose & "'," & rKh(i).absPos & "," & datform( _
   rKh(i).AktZeit) & "," & rKh(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKh(i).FID, ",", rKh(i).Pat_id, ",", DatForm(rKh(i).Zeitpunkt), ",'", rKh(i).Ziel, "','", rKh(i).Diagnose, "',", rKh(i).absPos, ",", DatForm( _
   rKh(i).AktZeit), ",", rKh(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKh) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKh), i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKh), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & Left$(rKh(k).Ziel, maxL) & "'", True: rKh(k).Ziel = Left$(rKh(k).Ziel, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & Left$(rKh(k).Diagnose, maxL) & "'", True: rKh(k).Diagnose = Left$(rKh(k).Diagnose, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere lbanforderungen"
 If Not AllePat Then
  sql = "delete from `lbanforderungen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       values"))
 For i = 1 To UBound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLb(i).FID & "," & rLb(i).Pat_ID & "," & datform(rLb(i).ZeitPunkt) & ",'" & rLb(i).AnfText & "'," & rLb(i).absPos & "," & datform( _
   rLb(i).AktZeit) & "," & rLb(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLb(i).FID, ",", rLb(i).Pat_id, ",", DatForm(rLb(i).Zeitpunkt), ",'", rLb(i).AnfText, "',", rLb(i).absPos, ",", DatForm( _
   rLb(i).AktZeit), ",", rLb(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLb) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLb), i)
  If Len(rLb(k).AnfText) > maxi(0) Then maxi(0) = Len(rLb(k).AnfText)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLb), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & Left$(rLb(k).AnfText, maxL) & "'", True: rLb(k).AnfText = Left$(rLb(k).AnfText, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborneu"
 If Not AllePat Then
  sql = "delete from `laborneu` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte)           values"))
 For i = 1 To UBound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLa(i).FID & "," & rLa(i).Pat_ID & "," & datform(rLa(i).ZeitPunkt) & ",'" & rLa(i).FertigStGrad & "','" & rLa(i).Abkü & "'," & rLa(i).LangtextVW & ",'" & _
   rLa(i).Wert & "','" & rLa(i).Einheit & "'," & rLa(i).KommentarVW & "," & rLa(i).AbsPos & "," & datform(rLa(i).AktZeit) & "," & _
   rLa(i).Refnr & "," & rLa(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLa(i).FID, ",", rLa(i).Pat_id, ",", DatForm(rLa(i).Zeitpunkt), ",'", rLa(i).FertigStGrad, "','", rLa(i).Abkü, "',", _
   rLa(i).LangtextVW, ",'", rLa(i).Wert, "','", rLa(i).Einheit, "',", rLa(i).KommentarVW, ",", rLa(i).absPos, ",", DatForm(rLa(i).AktZeit), ",", _
   rLa(i).RefNr, ",", rLa(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLa), i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLa), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & Left$(rLa(k).FertigStGrad, maxL) & "'", True: rLa(k).FertigStGrad = Left$(rLa(k).FertigStGrad, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & Left$(rLa(k).Abkü, maxL) & "'", True: rLa(k).Abkü = Left$(rLa(k).Abkü, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & Left$(rLa(k).Wert, maxL) & "'", True: rLa(k).Wert = Left$(rLa(k).Wert, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & Left$(rLa(k).Einheit, maxL) & "'", True: rLa(k).Einheit = Left$(rLa(k).Einheit, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere leistungen"
 If Not AllePat Then
  sql = "delete from `leistungen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte)  values"))
 For i = 1 To UBound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLe(i).FID & "," & rLe(i).Pat_ID & "," & datform(rLe(i).ZeitPunkt) & ",'" & rLe(i).Leistung & "','" & rLe(i).f5002 & "','" & rLe(i).f5005 & "','" & _
   rLe(i).f5006 & "','" & rLe(i).f5009 & "','" & rLe(i).Med & "','" & rLe(i).f5015 & "','" & rLe(i).f5016 & "','" & rLe(i).f5021 & "','" & _
   rLe(i).f5026 & "','" & rLe(i).Faktor & "','" & rLe(i).f5098 & "','" & rLe(i).LANR & "'," & datform(rLe(i).letzVorg) & ",'" & rLe(i).Ausn & "','" & _
   rLe(i).Beme & "'," & rLe(i).absPos & "," & datform(rLe(i).AktZeit) & ",'" & rLe(i).QS & "','" & rLe(i).QT & "'," & rLe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLe(i).FID, ",", rLe(i).Pat_id, ",", DatForm(rLe(i).Zeitpunkt), ",'", rLe(i).Leistung, "','", rLe(i).f5002, "','", _
   rLe(i).f5005, "','", rLe(i).f5006, "','", rLe(i).f5009, "','", rLe(i).Med, "','", rLe(i).f5015, "','", rLe(i).f5016, "','", rLe(i).f5021, "','", _
   rLe(i).f5026, "','", rLe(i).Faktor, "','", rLe(i).f5098, "','", rLe(i).LANR, "',", DatForm(rLe(i).letzVorg), ",'", rLe(i).Ausn, "','", _
   rLe(i).Beme, "',", rLe(i).absPos, ",", DatForm(rLe(i).AktZeit), ",'", rLe(i).QS, "','", rLe(i).QT, "',", rLe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(16)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLe), i)
  If Len(rLe(k).Leistung) > maxi(0) Then maxi(0) = Len(rLe(k).Leistung)
  If Len(rLe(k).f5002) > maxi(1) Then maxi(1) = Len(rLe(k).f5002)
  If Len(rLe(k).f5005) > maxi(2) Then maxi(2) = Len(rLe(k).f5005)
  If Len(rLe(k).f5006) > maxi(3) Then maxi(3) = Len(rLe(k).f5006)
  If Len(rLe(k).f5009) > maxi(4) Then maxi(4) = Len(rLe(k).f5009)
  If Len(rLe(k).Med) > maxi(5) Then maxi(5) = Len(rLe(k).Med)
  If Len(rLe(k).f5015) > maxi(6) Then maxi(6) = Len(rLe(k).f5015)
  If Len(rLe(k).f5016) > maxi(7) Then maxi(7) = Len(rLe(k).f5016)
  If Len(rLe(k).f5021) > maxi(8) Then maxi(8) = Len(rLe(k).f5021)
  If Len(rLe(k).f5026) > maxi(9) Then maxi(9) = Len(rLe(k).f5026)
  If Len(rLe(k).Faktor) > maxi(10) Then maxi(10) = Len(rLe(k).Faktor)
  If Len(rLe(k).f5098) > maxi(11) Then maxi(11) = Len(rLe(k).f5098)
  If Len(rLe(k).LANR) > maxi(12) Then maxi(12) = Len(rLe(k).LANR)
  If Len(rLe(k).Ausn) > maxi(13) Then maxi(13) = Len(rLe(k).Ausn)
  If Len(rLe(k).Beme) > maxi(14) Then maxi(14) = Len(rLe(k).Beme)
  If Len(rLe(k).QS) > maxi(15) Then maxi(15) = Len(rLe(k).QS)
  If Len(rLe(k).QT) > maxi(16) Then maxi(16) = Len(rLe(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLe), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLe.Leistung: '" & rLe(k).Leistung & "' -> '" & Left$(rLe(k).Leistung, maxL) & "'", True: rLe(k).Leistung = Left$(rLe(k).Leistung, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5002: '" & rLe(k).f5002 & "' -> '" & Left$(rLe(k).f5002, maxL) & "'", True: rLe(k).f5002 = Left$(rLe(k).f5002, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5005: '" & rLe(k).f5005 & "' -> '" & Left$(rLe(k).f5005, maxL) & "'", True: rLe(k).f5005 = Left$(rLe(k).f5005, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5006: '" & rLe(k).f5006 & "' -> '" & Left$(rLe(k).f5006, maxL) & "'", True: rLe(k).f5006 = Left$(rLe(k).f5006, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5009: '" & rLe(k).f5009 & "' -> '" & Left$(rLe(k).f5009, maxL) & "'", True: rLe(k).f5009 = Left$(rLe(k).f5009, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLe.Med: '" & rLe(k).Med & "' -> '" & Left$(rLe(k).Med, maxL) & "'", True: rLe(k).Med = Left$(rLe(k).Med, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5015: '" & rLe(k).f5015 & "' -> '" & Left$(rLe(k).f5015, maxL) & "'", True: rLe(k).f5015 = Left$(rLe(k).f5015, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5016: '" & rLe(k).f5016 & "' -> '" & Left$(rLe(k).f5016, maxL) & "'", True: rLe(k).f5016 = Left$(rLe(k).f5016, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5021: '" & rLe(k).f5021 & "' -> '" & Left$(rLe(k).f5021, maxL) & "'", True: rLe(k).f5021 = Left$(rLe(k).f5021, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5026: '" & rLe(k).f5026 & "' -> '" & Left$(rLe(k).f5026, maxL) & "'", True: rLe(k).f5026 = Left$(rLe(k).f5026, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLe.Faktor: '" & rLe(k).Faktor & "' -> '" & Left$(rLe(k).Faktor, maxL) & "'", True: rLe(k).Faktor = Left$(rLe(k).Faktor, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5098: '" & rLe(k).f5098 & "' -> '" & Left$(rLe(k).f5098, maxL) & "'", True: rLe(k).f5098 = Left$(rLe(k).f5098, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).LANR & "' -> '" & Left$(rLe(k).LANR, maxL) & "'", True: rLe(k).LANR = Left$(rLe(k).LANR, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL) & "'", True: rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).Beme & "' -> '" & Left$(rLe(k).Beme, maxL) & "'", True: rLe(k).Beme = Left$(rLe(k).Beme, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL) & "'", True: rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL) & "'", True: rLe(k).QT = Left$(rLe(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim obfor%
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere medplan"
 If Not AllePat Then
  sql = "delete from `medplan` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte)              values"))
 For i = 1 To UBound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rMe(i).FID & "," & rMe(i).Pat_ID & "," & rMe(i).MPNr & "," & datform(rMe(i).ZeitPunkt) & "," & datform(rMe(i).Datum) & ",'" & rMe(i).Medikament & "','" & _
   rMe(i).MedAnfang & "'," & rMe(i).FeldNr & ",'" & rMe(i).mo & "','" & rMe(i).mi & "','" & rMe(i).nm & "','" & rMe(i).ab & "','" & _
   rMe(i).zn & "'," & cstr(cint(rMe(i).bBed)) & ",'" & rMe(i).Bemerkung & "'," & rMe(i).AbsPos & "," & datform(rMe(i).AktZeit) & "," & rMe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rMe(i).FID, ",", rMe(i).Pat_id, ",", rMe(i).MPNr, ",", DatForm(rMe(i).Zeitpunkt), ",", DatForm(rMe(i).Datum), ",'", _
   rMe(i).Medikament, "','", rMe(i).MedAnfang, "',", rMe(i).FeldNr, ",'", rMe(i).mo, "','", rMe(i).mi, "','", rMe(i).nm, "','", rMe(i).ab, "','", _
   rMe(i).Zn, "',", CStr(CInt(rMe(i).bBed)), ",'", rMe(i).Bemerkung, "',", rMe(i).absPos, ",", DatForm(rMe(i).AktZeit), ",", _
   rMe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rMe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147467259 Then  ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add or update a child row: a foreign key constraint fails
 ErrDescription = Err.Description
 Call ForeignNo0
 Call ForeignNo1
 Open BezFeh For Append As #299
 Print #299, vbCrLf & vbCrLf & Now() & ": " & csql.Value
 Print #299, vbCrLf & ErrDescription
 Close #299
 obfor = True
 Resume
ElseIf Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(7)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rMe), i)
  If Len(rMe(k).Medikament) > maxi(0) Then maxi(0) = Len(rMe(k).Medikament)
  If Len(rMe(k).MedAnfang) > maxi(1) Then maxi(1) = Len(rMe(k).MedAnfang)
  If Len(rMe(k).mo) > maxi(2) Then maxi(2) = Len(rMe(k).mo)
  If Len(rMe(k).mi) > maxi(3) Then maxi(3) = Len(rMe(k).mi)
  If Len(rMe(k).nm) > maxi(4) Then maxi(4) = Len(rMe(k).nm)
  If Len(rMe(k).ab) > maxi(5) Then maxi(5) = Len(rMe(k).ab)
  If Len(rMe(k).Zn) > maxi(6) Then maxi(6) = Len(rMe(k).Zn)
  If Len(rMe(k).Bemerkung) > maxi(7) Then maxi(7) = Len(rMe(k).Bemerkung)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rMe), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & Left$(rMe(k).Medikament, maxL) & "'", True: rMe(k).Medikament = Left$(rMe(k).Medikament, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & Left$(rMe(k).MedAnfang, maxL) & "'", True: rMe(k).MedAnfang = Left$(rMe(k).MedAnfang, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & Left$(rMe(k).mo, maxL) & "'", True: rMe(k).mo = Left$(rMe(k).mo, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & Left$(rMe(k).mi, maxL) & "'", True: rMe(k).mi = Left$(rMe(k).mi, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & Left$(rMe(k).nm, maxL) & "'", True: rMe(k).nm = Left$(rMe(k).nm, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & Left$(rMe(k).ab, maxL) & "'", True: rMe(k).ab = Left$(rMe(k).ab, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).Zn & "' -> '" & Left$(rMe(k).Zn, maxL) & "'", True: rMe(k).Zn = Left$(rMe(k).Zn, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & Left$(rMe(k).Bemerkung, maxL) & "'", True: rMe(k).Bemerkung = Left$(rMe(k).Bemerkung, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere rezepteintraege"
 If Not AllePat Then
  sql = "delete from `rezepteintraege` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte)  values"))
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRe(i).FID & "," & rRe(i).Pat_ID & "," & datform(rRe(i).ZeitPunkt) & ",'" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "','" & _
   rRe(i).Medikament & "','" & rRe(i).PZN & "'," & rRe(i).absPos & "," & datform(rRe(i).AktZeit) & ",'" & rRe(i).QS & "','" & rRe(i).QT & "'," & rRe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRe(i).FID, ",", rRe(i).Pat_id, ",", DatForm(rRe(i).Zeitpunkt), ",'", rRe(i).Rezept, "','", rRe(i).Rezeptklasse, "','", _
   rRe(i).Medikament, "','", rRe(i).PZN, "',", rRe(i).absPos, ",", DatForm(rRe(i).AktZeit), ",'", rRe(i).QS, "','", rRe(i).QT, "',", _
   rRe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRe), i)
  If Len(rRe(k).Rezept) > maxi(0) Then maxi(0) = Len(rRe(k).Rezept)
  If Len(rRe(k).Rezeptklasse) > maxi(1) Then maxi(1) = Len(rRe(k).Rezeptklasse)
  If Len(rRe(k).Medikament) > maxi(2) Then maxi(2) = Len(rRe(k).Medikament)
  If Len(rRe(k).PZN) > maxi(3) Then maxi(3) = Len(rRe(k).PZN)
  If Len(rRe(k).QS) > maxi(4) Then maxi(4) = Len(rRe(k).QS)
  If Len(rRe(k).QT) > maxi(5) Then maxi(5) = Len(rRe(k).QT)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRe), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & Left$(rRe(k).Rezept, maxL) & "'", True: rRe(k).Rezept = Left$(rRe(k).Rezept, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & Left$(rRe(k).Rezeptklasse, maxL) & "'", True: rRe(k).Rezeptklasse = Left$(rRe(k).Rezeptklasse, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & Left$(rRe(k).Medikament, maxL) & "'", True: rRe(k).Medikament = Left$(rRe(k).Medikament, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & Left$(rRe(k).PZN, maxL) & "'", True: rRe(k).PZN = Left$(rRe(k).PZN, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & Left$(rRe(k).QS, maxL) & "'", True: rRe(k).QS = Left$(rRe(k).QS, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & Left$(rRe(k).QT, maxL) & "'", True: rRe(k).QT = Left$(rRe(k).QT, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere rr"
 If Not AllePat Then
  sql = "delete from `rr` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte)            values"))
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRr(i).FID & "," & rRr(i).Pat_ID & "," & datform(rRr(i).ZeitPunkt) & ",'" & rRr(i).RR & "'," & rRr(i).absPos & "," & datform(rRr(i).AktZeit) & "," & _
   rRr(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRr(i).FID, ",", rRr(i).Pat_id, ",", DatForm(rRr(i).Zeitpunkt), ",'", rRr(i).RR, "',", rRr(i).absPos, ",", DatForm( _
   rRr(i).AktZeit), ",", rRr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
End If
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRr), i)
  If Len(rRr(k).RR) > maxi(0) Then maxi(0) = Len(rRr(k).RR)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRr), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & Left$(rRr(k).RR, maxL) & "'", True: rRr(k).RR = Left$(rRr(k).RR, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere kvnrue"
 If Not AllePat Then
  sql = "delete from `kvnrue` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        values"))
 For i = 1 To UBound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKv(i).Pat_ID & ",'" & rKv(i).KVNr & "'," & rKv(i).absPos & "," & datform(rKv(i).AktZeit) & "," & rKv(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKv(i).Pat_id, ",'", rKv(i).KVNr, "',", rKv(i).absPos, ",", DatForm(rKv(i).AktZeit), ",", rKv(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKv) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKv), i)
  If Len(rKv(k).KVNr) > maxi(0) Then maxi(0) = Len(rKv(k).KVNr)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKv), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & Left$(rKv(k).KVNr, maxL) & "'", True: rKv(k).KVNr = Left$(rKv(k).KVNr, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 syscmd 4, PID & ": Speichere unbekannte_kennungen"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt)         values"))
 For i = rUn1 + 1 To UBound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
'   sql = sql0 & "('" & rUn(i).Kennung & "'," & rUn(i).absPos & "," & rUn(i).StByte & "," & rUn(i).Pat_id & ",'" & rUn(i).Inhalt & "')"
  If Lese.dlg.SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rUn(i).Kennung, "',", rUn(i).absPos, ",", rUn(i).StByte, ",", rUn(i).Pat_id, ",'", rUn(i).Inhalt, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rUn) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rUn) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 rUn1 = UBound(rUn)
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rUn1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rUn), i)
  If Len(rUn(k).Kennung) > maxi(0) Then maxi(0) = Len(rUn(k).Kennung)
  If Len(rUn(k).Inhalt) > maxi(1) Then maxi(1) = Len(rUn(k).Inhalt)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rUn1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rUn), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & Left$(rUn(k).Kennung, maxL) & "'", True: rUn(k).Kennung = Left$(rUn(k).Kennung, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & Left$(rUn(k).Inhalt, maxL) & "'", True: rUn(k).Inhalt = Left$(rUn(k).Inhalt, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere dmpreihe"
 If Not AllePat Then
  sql = "delete from `dmpreihe` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit)     values"))
 For i = 1 To UBound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
'   sql = sql0 & "('" & rDm(i).Abk & "','" & rDm(i).Art & "'," & datform(rDm(i).KarteiDatum) & "," & datform(rDm(i).exportiert) & "," & datform(rDm(i).DokuDatum) & "," & cstr(cint( _
   rDm(i).obvoll)) & ",'" & rDm(i).NachName & "','" & rDm(i).VorName & "'," & datform(rDm(i).GebDat) & "," & rDm(i).Pat_id & "," & _
   rDm(i).StByte & "," & datform(rDm(i).AktZeit) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rDm(i).Abk, "','", rDm(i).Art, "',", DatForm(rDm(i).KarteiDatum), ",", DatForm(rDm(i).exportiert), ",", DatForm(rDm(i).DokuDatum), ",", CStr(CInt( _
   rDm(i).obvoll)), ",'", rDm(i).Nachname, "','", rDm(i).Vorname, "',", DatForm(rDm(i).GebDat), ",", rDm(i).Pat_id, ",", _
   rDm(i).StByte, ",", DatForm(rDm(i).AktZeit), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDm) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDm), i)
  If Len(rDm(k).Abk) > maxi(0) Then maxi(0) = Len(rDm(k).Abk)
  If Len(rDm(k).Art) > maxi(1) Then maxi(1) = Len(rDm(k).Art)
  If Len(rDm(k).Nachname) > maxi(2) Then maxi(2) = Len(rDm(k).Nachname)
  If Len(rDm(k).Vorname) > maxi(3) Then maxi(3) = Len(rDm(k).Vorname)
 Next k
 DBCn.CommitTrans
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDm), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL) & "'", True: rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).Art & "' -> '" & Left$(rDm(k).Art, maxL) & "'", True: rDm(k).Art = Left$(rDm(k).Art, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).Nachname & "' -> '" & Left$(rDm(k).Nachname, maxL) & "'", True: rDm(k).Nachname = Left$(rDm(k).Nachname, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).Vorname & "' -> '" & Left$(rDm(k).Vorname, maxL) & "'", True: rDm(k).Vorname = Left$(rDm(k).Vorname, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function laborxsaetzeSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             values"))
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
'   sql = sql0 & "(" & rLs(i).DatID & ",'" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).VersionSatzb & "','" & _
   rLs(i).Arztnr & "','" & rLs(i).Arztname & "','" & rLs(i).StraßePraxis & "','" & rLs(i).Arzt & "','" & rLs(i).LANR & "','" & rLs(i).PLZPraxis & "','" & _
   rLs(i).OrtPraxis & "','" & rLs(i).Labor & "','" & rLs(i).StraßeLabor & "','" & rLs(i).PLZLabor & "','" & rLs(i).OrtLabor & "','" & _
   rLs(i).KBVPrüfnr & "','" & rLs(i).Zeichensatz & "','" & rLs(i).Kundenarztnr & "','" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLs(i).DatID, ",'", rLs(i).Satzart, "','", rLs(i).Satzlänge, "','", rLs(i).SatzlängeSchluss, "','", rLs(i).VersionSatzb, "','", _
   rLs(i).Arztnr, "','", rLs(i).Arztname, "','", rLs(i).StraßePraxis, "','", rLs(i).Arzt, "','", rLs(i).LANR, "','", rLs(i).PLZPraxis, "','", _
   rLs(i).OrtPraxis, "','", rLs(i).Labor, "','", rLs(i).StraßeLabor, "','", rLs(i).PLZLabor, "','", rLs(i).OrtLabor, "','", _
   rLs(i).KBVPrüfnr, "','", rLs(i).Zeichensatz, "','", rLs(i).Kundenarztnr, "','", rLs(i).Erstellungsdatum, "','", rLs(i).Gesamtlänge, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLs) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(19)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLs), i)
  If Len(rLs(k).Satzart) > maxi(0) Then maxi(0) = Len(rLs(k).Satzart)
  If Len(rLs(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rLs(k).Satzlänge)
  If Len(rLs(k).SatzlängeSchluss) > maxi(2) Then maxi(2) = Len(rLs(k).SatzlängeSchluss)
  If Len(rLs(k).VersionSatzb) > maxi(3) Then maxi(3) = Len(rLs(k).VersionSatzb)
  If Len(rLs(k).Arztnr) > maxi(4) Then maxi(4) = Len(rLs(k).Arztnr)
  If Len(rLs(k).Arztname) > maxi(5) Then maxi(5) = Len(rLs(k).Arztname)
  If Len(rLs(k).StraßePraxis) > maxi(6) Then maxi(6) = Len(rLs(k).StraßePraxis)
  If Len(rLs(k).Arzt) > maxi(7) Then maxi(7) = Len(rLs(k).Arzt)
  If Len(rLs(k).LANR) > maxi(8) Then maxi(8) = Len(rLs(k).LANR)
  If Len(rLs(k).PLZPraxis) > maxi(9) Then maxi(9) = Len(rLs(k).PLZPraxis)
  If Len(rLs(k).OrtPraxis) > maxi(10) Then maxi(10) = Len(rLs(k).OrtPraxis)
  If Len(rLs(k).Labor) > maxi(11) Then maxi(11) = Len(rLs(k).Labor)
  If Len(rLs(k).StraßeLabor) > maxi(12) Then maxi(12) = Len(rLs(k).StraßeLabor)
  If Len(rLs(k).PLZLabor) > maxi(13) Then maxi(13) = Len(rLs(k).PLZLabor)
  If Len(rLs(k).OrtLabor) > maxi(14) Then maxi(14) = Len(rLs(k).OrtLabor)
  If Len(rLs(k).KBVPrüfnr) > maxi(15) Then maxi(15) = Len(rLs(k).KBVPrüfnr)
  If Len(rLs(k).Zeichensatz) > maxi(16) Then maxi(16) = Len(rLs(k).Zeichensatz)
  If Len(rLs(k).Kundenarztnr) > maxi(17) Then maxi(17) = Len(rLs(k).Kundenarztnr)
  If Len(rLs(k).Erstellungsdatum) > maxi(18) Then maxi(18) = Len(rLs(k).Erstellungsdatum)
  If Len(rLs(k).Gesamtlänge) > maxi(19) Then maxi(19) = Len(rLs(k).Gesamtlänge)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLs), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzart: '" & rLs(k).Satzart & "' -> '" & Left$(rLs(k).Satzart, maxL) & "'", True: rLs(k).Satzart = Left$(rLs(k).Satzart, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzlänge: '" & rLs(k).Satzlänge & "' -> '" & Left$(rLs(k).Satzlänge, maxL) & "'", True: rLs(k).Satzlänge = Left$(rLs(k).Satzlänge, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLs.SatzlängeSchluss: '" & rLs(k).SatzlängeSchluss & "' -> '" & Left$(rLs(k).SatzlängeSchluss, maxL) & "'", True: rLs(k).SatzlängeSchluss = Left$(rLs(k).SatzlängeSchluss, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLs.VersionSatzb: '" & rLs(k).VersionSatzb & "' -> '" & Left$(rLs(k).VersionSatzb, maxL) & "'", True: rLs(k).VersionSatzb = Left$(rLs(k).VersionSatzb, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztnr: '" & rLs(k).Arztnr & "' -> '" & Left$(rLs(k).Arztnr, maxL) & "'", True: rLs(k).Arztnr = Left$(rLs(k).Arztnr, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztname: '" & rLs(k).Arztname & "' -> '" & Left$(rLs(k).Arztname, maxL) & "'", True: rLs(k).Arztname = Left$(rLs(k).Arztname, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßePraxis: '" & rLs(k).StraßePraxis & "' -> '" & Left$(rLs(k).StraßePraxis, maxL) & "'", True: rLs(k).StraßePraxis = Left$(rLs(k).StraßePraxis, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arzt: '" & rLs(k).Arzt & "' -> '" & Left$(rLs(k).Arzt, maxL) & "'", True: rLs(k).Arzt = Left$(rLs(k).Arzt, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLs.LANR: '" & rLs(k).LANR & "' -> '" & Left$(rLs(k).LANR, maxL) & "'", True: rLs(k).LANR = Left$(rLs(k).LANR, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZPraxis: '" & rLs(k).PLZPraxis & "' -> '" & Left$(rLs(k).PLZPraxis, maxL) & "'", True: rLs(k).PLZPraxis = Left$(rLs(k).PLZPraxis, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtPraxis: '" & rLs(k).OrtPraxis & "' -> '" & Left$(rLs(k).OrtPraxis, maxL) & "'", True: rLs(k).OrtPraxis = Left$(rLs(k).OrtPraxis, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLs.Labor: '" & rLs(k).Labor & "' -> '" & Left$(rLs(k).Labor, maxL) & "'", True: rLs(k).Labor = Left$(rLs(k).Labor, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßeLabor: '" & rLs(k).StraßeLabor & "' -> '" & Left$(rLs(k).StraßeLabor, maxL) & "'", True: rLs(k).StraßeLabor = Left$(rLs(k).StraßeLabor, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZLabor: '" & rLs(k).PLZLabor & "' -> '" & Left$(rLs(k).PLZLabor, maxL) & "'", True: rLs(k).PLZLabor = Left$(rLs(k).PLZLabor, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtLabor: '" & rLs(k).OrtLabor & "' -> '" & Left$(rLs(k).OrtLabor, maxL) & "'", True: rLs(k).OrtLabor = Left$(rLs(k).OrtLabor, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLs.KBVPrüfnr: '" & rLs(k).KBVPrüfnr & "' -> '" & Left$(rLs(k).KBVPrüfnr, maxL) & "'", True: rLs(k).KBVPrüfnr = Left$(rLs(k).KBVPrüfnr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLs.Zeichensatz: '" & rLs(k).Zeichensatz & "' -> '" & Left$(rLs(k).Zeichensatz, maxL) & "'", True: rLs(k).Zeichensatz = Left$(rLs(k).Zeichensatz, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLs.Kundenarztnr: '" & rLs(k).Kundenarztnr & "' -> '" & Left$(rLs(k).Kundenarztnr, maxL) & "'", True: rLs(k).Kundenarztnr = Left$(rLs(k).Kundenarztnr, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLs.Erstellungsdatum: '" & rLs(k).Erstellungsdatum & "' -> '" & Left$(rLs(k).Erstellungsdatum, maxL) & "'", True: rLs(k).Erstellungsdatum = Left$(rLs(k).Erstellungsdatum, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLs.Gesamtlänge: '" & rLs(k).Gesamtlänge & "' -> '" & Left$(rLs(k).Gesamtlänge, maxL) & "'", True: rLs(k).Gesamtlänge = Left$(rLs(k).Gesamtlänge, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  values"))
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
'   sql = sql0 & "('" & rLg(i).Pfad & "','" & rLg(i).Name & "'," & datform(rLg(i).Zp) & "," & cstr(cint(rLg(i).fertig)) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rLg(i).Pfad, "','", rLg(i).Name, "',", DatForm(rLg(i).Zp), ",", CStr(CInt(rLg(i).fertig)), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLg) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLg), i)
  If Len(rLg(k).Pfad) > maxi(0) Then maxi(0) = Len(rLg(k).Pfad)
  If Len(rLg(k).Name) > maxi(1) Then maxi(1) = Len(rLg(k).Name)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLg), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL) & "'", True: rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).Name & "' -> '" & Left$(rLg(k).Name, maxL) & "'", True: rLg(k).Name = Left$(rLg(k).Name, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxus"
' sql0 = " insert " & sqlignore &  "into `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN)       values"))
 For i = j To j
'  rLu(i).AktZeit = now()
'   sql = sql0 & "(" & rLu(i).DatID & "," & rLu(i).SatzID & ",'" & rLu(i).Satzart & "','" & rLu(i).Satzlänge & "','" & rLu(i).Auftragsnummer & "','" & rLu(i).Auftragsschlüssel & "'," & datform( _
   rLu(i).Eingang) & ",'" & rLu(i).Berichtsdatum & "'," & rLu(i).Pat_id & ",'" & rLu(i).Nachname & "','" & _
   rLu(i).Vorname & "','" & rLu(i).GebDat & "','" & rLu(i).Titel & "','" & rLu(i).NVorsatz & "','" & rLu(i).BefArt & "','" & rLu(i).Abrechnungstyp & "','" & _
   rLu(i).GebüOrd & "','" & rLu(i).Patienteninformation & "','" & rLu(i).Geschlecht & "','" & rLu(i).AuftrHinw & "','" & rLu(i).Pat_idUrsp & "','" & _
   rLu(i).Pat_idErwVNG & "','" & rLu(i).Pat_idErwVN & "','" & rLu(i).Pat_idErwG & "','" & rLu(i).Pat_idErwGB & "','" & rLu(i).Pat_idErwGL & "','" & _
   rLu(i).Pat_idLaborNeu & "'," & datform(rLu(i).ZeitpunktLaborneu) & "," & rLu(i).ZdüP & "," & rLu(i).ZdiP & ",'" & rLu(i).LWerte & "'," & datform( _
   rLu(i).verglichen) & "," & rLu(i).AfN & ")"
  If Lese.dlg.SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLu(i).DatID, ",", rLu(i).SatzID, ",'", rLu(i).Satzart, "','", rLu(i).Satzlänge, "','", rLu(i).Auftragsnummer, "','", _
   rLu(i).Auftragsschlüssel, "',", DatForm(rLu(i).Eingang), ",'", rLu(i).Berichtsdatum, "',", rLu(i).Pat_id, ",'", rLu(i).Nachname, "','", _
   rLu(i).Vorname, "','", rLu(i).GebDat, "','", rLu(i).Titel, "','", rLu(i).NVorsatz, "','", rLu(i).BefArt, "','", rLu(i).Abrechnungstyp, "','", _
   rLu(i).GebüOrd, "','", rLu(i).Patienteninformation, "','", rLu(i).Geschlecht, "','", rLu(i).AuftrHinw, "','", rLu(i).Pat_idUrsp, "','", _
   rLu(i).Pat_idErwVNG, "','", rLu(i).Pat_idErwVN, "','", rLu(i).Pat_idErwG, "','", rLu(i).Pat_idErwGB, "','", rLu(i).Pat_idErwGL, "','", _
   rLu(i).Pat_idLaborNeu, "',", DatForm(rLu(i).ZeitpunktLaborneu), ",", rLu(i).ZdüP, ",", rLu(i).ZdiP, ",'", rLu(i).LWerte, "',", DatForm( _
   rLu(i).verglichen), ",", rLu(i).AfN, ")")
  If Lese.dlg.SammelInsert <> 0 And i < j Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = j Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(23)
 For k = IIf(Lese.dlg.SammelInsert <> 0, j, i) To IIf(Lese.dlg.SammelInsert <> 0, j, i)
  If Len(rLu(k).Satzart) > maxi(0) Then maxi(0) = Len(rLu(k).Satzart)
  If Len(rLu(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rLu(k).Satzlänge)
  If Len(rLu(k).Auftragsnummer) > maxi(2) Then maxi(2) = Len(rLu(k).Auftragsnummer)
  If Len(rLu(k).Auftragsschlüssel) > maxi(3) Then maxi(3) = Len(rLu(k).Auftragsschlüssel)
  If Len(rLu(k).Berichtsdatum) > maxi(4) Then maxi(4) = Len(rLu(k).Berichtsdatum)
  If Len(rLu(k).Nachname) > maxi(5) Then maxi(5) = Len(rLu(k).Nachname)
  If Len(rLu(k).Vorname) > maxi(6) Then maxi(6) = Len(rLu(k).Vorname)
  If Len(rLu(k).GebDat) > maxi(7) Then maxi(7) = Len(rLu(k).GebDat)
  If Len(rLu(k).Titel) > maxi(8) Then maxi(8) = Len(rLu(k).Titel)
  If Len(rLu(k).NVorsatz) > maxi(9) Then maxi(9) = Len(rLu(k).NVorsatz)
  If Len(rLu(k).BefArt) > maxi(10) Then maxi(10) = Len(rLu(k).BefArt)
  If Len(rLu(k).Abrechnungstyp) > maxi(11) Then maxi(11) = Len(rLu(k).Abrechnungstyp)
  If Len(rLu(k).GebüOrd) > maxi(12) Then maxi(12) = Len(rLu(k).GebüOrd)
  If Len(rLu(k).Patienteninformation) > maxi(13) Then maxi(13) = Len(rLu(k).Patienteninformation)
  If Len(rLu(k).Geschlecht) > maxi(14) Then maxi(14) = Len(rLu(k).Geschlecht)
  If Len(rLu(k).AuftrHinw) > maxi(15) Then maxi(15) = Len(rLu(k).AuftrHinw)
  If Len(rLu(k).Pat_idUrsp) > maxi(16) Then maxi(16) = Len(rLu(k).Pat_idUrsp)
  If Len(rLu(k).Pat_idErwVNG) > maxi(17) Then maxi(17) = Len(rLu(k).Pat_idErwVNG)
  If Len(rLu(k).Pat_idErwVN) > maxi(18) Then maxi(18) = Len(rLu(k).Pat_idErwVN)
  If Len(rLu(k).Pat_idErwG) > maxi(19) Then maxi(19) = Len(rLu(k).Pat_idErwG)
  If Len(rLu(k).Pat_idErwGB) > maxi(20) Then maxi(20) = Len(rLu(k).Pat_idErwGB)
  If Len(rLu(k).Pat_idErwGL) > maxi(21) Then maxi(21) = Len(rLu(k).Pat_idErwGL)
  If Len(rLu(k).Pat_idLaborNeu) > maxi(22) Then maxi(22) = Len(rLu(k).Pat_idLaborNeu)
  If Len(rLu(k).LWerte) > maxi(23) Then maxi(23) = Len(rLu(k).LWerte)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, j, i) To IIf(Lese.dlg.SammelInsert <> 0, j, i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzart: '" & rLu(k).Satzart & "' -> '" & Left$(rLu(k).Satzart, maxL) & "'", True: rLu(k).Satzart = Left$(rLu(k).Satzart, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzlänge: '" & rLu(k).Satzlänge & "' -> '" & Left$(rLu(k).Satzlänge, maxL) & "'", True: rLu(k).Satzlänge = Left$(rLu(k).Satzlänge, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsnummer: '" & rLu(k).Auftragsnummer & "' -> '" & Left$(rLu(k).Auftragsnummer, maxL) & "'", True: rLu(k).Auftragsnummer = Left$(rLu(k).Auftragsnummer, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsschlüssel: '" & rLu(k).Auftragsschlüssel & "' -> '" & Left$(rLu(k).Auftragsschlüssel, maxL) & "'", True: rLu(k).Auftragsschlüssel = Left$(rLu(k).Auftragsschlüssel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLu.Berichtsdatum: '" & rLu(k).Berichtsdatum & "' -> '" & Left$(rLu(k).Berichtsdatum, maxL) & "'", True: rLu(k).Berichtsdatum = Left$(rLu(k).Berichtsdatum, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLu.Nachname: '" & rLu(k).Nachname & "' -> '" & Left$(rLu(k).Nachname, maxL) & "'", True: rLu(k).Nachname = Left$(rLu(k).Nachname, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLu.Vorname: '" & rLu(k).Vorname & "' -> '" & Left$(rLu(k).Vorname, maxL) & "'", True: rLu(k).Vorname = Left$(rLu(k).Vorname, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebDat: '" & rLu(k).GebDat & "' -> '" & Left$(rLu(k).GebDat, maxL) & "'", True: rLu(k).GebDat = Left$(rLu(k).GebDat, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLu.Titel: '" & rLu(k).Titel & "' -> '" & Left$(rLu(k).Titel, maxL) & "'", True: rLu(k).Titel = Left$(rLu(k).Titel, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLu.NVorsatz: '" & rLu(k).NVorsatz & "' -> '" & Left$(rLu(k).NVorsatz, maxL) & "'", True: rLu(k).NVorsatz = Left$(rLu(k).NVorsatz, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLu.BefArt: '" & rLu(k).BefArt & "' -> '" & Left$(rLu(k).BefArt, maxL) & "'", True: rLu(k).BefArt = Left$(rLu(k).BefArt, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLu.Abrechnungstyp: '" & rLu(k).Abrechnungstyp & "' -> '" & Left$(rLu(k).Abrechnungstyp, maxL) & "'", True: rLu(k).Abrechnungstyp = Left$(rLu(k).Abrechnungstyp, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebüOrd: '" & rLu(k).GebüOrd & "' -> '" & Left$(rLu(k).GebüOrd, maxL) & "'", True: rLu(k).GebüOrd = Left$(rLu(k).GebüOrd, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLu.Patienteninformation: '" & rLu(k).Patienteninformation & "' -> '" & Left$(rLu(k).Patienteninformation, maxL) & "'", True: rLu(k).Patienteninformation = Left$(rLu(k).Patienteninformation, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLu.Geschlecht: '" & rLu(k).Geschlecht & "' -> '" & Left$(rLu(k).Geschlecht, maxL) & "'", True: rLu(k).Geschlecht = Left$(rLu(k).Geschlecht, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLu.AuftrHinw: '" & rLu(k).AuftrHinw & "' -> '" & Left$(rLu(k).AuftrHinw, maxL) & "'", True: rLu(k).AuftrHinw = Left$(rLu(k).AuftrHinw, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idUrsp: '" & rLu(k).Pat_idUrsp & "' -> '" & Left$(rLu(k).Pat_idUrsp, maxL) & "'", True: rLu(k).Pat_idUrsp = Left$(rLu(k).Pat_idUrsp, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVNG: '" & rLu(k).Pat_idErwVNG & "' -> '" & Left$(rLu(k).Pat_idErwVNG, maxL) & "'", True: rLu(k).Pat_idErwVNG = Left$(rLu(k).Pat_idErwVNG, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVN: '" & rLu(k).Pat_idErwVN & "' -> '" & Left$(rLu(k).Pat_idErwVN, maxL) & "'", True: rLu(k).Pat_idErwVN = Left$(rLu(k).Pat_idErwVN, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwG: '" & rLu(k).Pat_idErwG & "' -> '" & Left$(rLu(k).Pat_idErwG, maxL) & "'", True: rLu(k).Pat_idErwG = Left$(rLu(k).Pat_idErwG, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGB: '" & rLu(k).Pat_idErwGB & "' -> '" & Left$(rLu(k).Pat_idErwGB, maxL) & "'", True: rLu(k).Pat_idErwGB = Left$(rLu(k).Pat_idErwGB, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGL: '" & rLu(k).Pat_idErwGL & "' -> '" & Left$(rLu(k).Pat_idErwGL, maxL) & "'", True: rLu(k).Pat_idErwGL = Left$(rLu(k).Pat_idErwGL, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idLaborNeu: '" & rLu(k).Pat_idLaborNeu & "' -> '" & Left$(rLu(k).Pat_idLaborNeu, maxL) & "'", True: rLu(k).Pat_idLaborNeu = Left$(rLu(k).Pat_idLaborNeu, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLu.LWerte: '" & rLu(k).LWerte & "' -> '" & Left$(rLu(k).LWerte, maxL) & "'", True: rLu(k).LWerte = Left$(rLu(k).LWerte, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  values"))
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
'   sql = sql0 & "(" & rLo(i).RefNr & ",'" & rLo(i).Verf & "','" & rLo(i).KuQu & "','" & rLo(i).Quelle & "','" & rLo(i).QSpez & "'," & datform(rLo(i).AbnDat) & ",'" & _
   rLo(i).Kommentar & "','" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLo(i).RefNr, ",'", rLo(i).Verf, "','", rLo(i).KuQu, "','", rLo(i).Quelle, "','", rLo(i).QSpez, "',", DatForm(rLo(i).AbnDat), ",'", _
   rLo(i).Kommentar, "','", rLo(i).Erklärung, "','", rLo(i).Keimzahl, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(6)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLo), i)
  If Len(rLo(k).Verf) > maxi(0) Then maxi(0) = Len(rLo(k).Verf)
  If Len(rLo(k).KuQu) > maxi(1) Then maxi(1) = Len(rLo(k).KuQu)
  If Len(rLo(k).Quelle) > maxi(2) Then maxi(2) = Len(rLo(k).Quelle)
  If Len(rLo(k).QSpez) > maxi(3) Then maxi(3) = Len(rLo(k).QSpez)
  If Len(rLo(k).Kommentar) > maxi(4) Then maxi(4) = Len(rLo(k).Kommentar)
  If Len(rLo(k).Erklärung) > maxi(5) Then maxi(5) = Len(rLo(k).Erklärung)
  If Len(rLo(k).Keimzahl) > maxi(6) Then maxi(6) = Len(rLo(k).Keimzahl)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLo), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & Left$(rLo(k).Verf, maxL) & "'", True: rLo(k).Verf = Left$(rLo(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & Left$(rLo(k).KuQu, maxL) & "'", True: rLo(k).KuQu = Left$(rLo(k).KuQu, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & Left$(rLo(k).Quelle, maxL) & "'", True: rLo(k).Quelle = Left$(rLo(k).Quelle, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & Left$(rLo(k).QSpez, maxL) & "'", True: rLo(k).QSpez = Left$(rLo(k).QSpez, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & Left$(rLo(k).Kommentar, maxL) & "'", True: rLo(k).Kommentar = Left$(rLo(k).Kommentar, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & Left$(rLo(k).Erklärung, maxL) & "'", True: rLo(k).Erklärung = Left$(rLo(k).Erklärung, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & Left$(rLo(k).Keimzahl, maxL) & "'", True: rLo(k).Keimzahl = Left$(rLo(k).Keimzahl, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw)               values"))
 For i = 1 To UBound(rLw)
'  rLw(i).AktZeit = now()
'   sql = sql0 & "(" & rLw(i).RefNr & ",'" & rLw(i).Abkü & "','" & rLw(i).Langname & "','" & rLw(i).Quelle & "','" & rLw(i).QSpez & "'," & datform(rLw(i).AbnDat) & ",'" & _
   rLw(i).Wert & "','" & rLw(i).Einheit & "','" & rLw(i).Grenzwerti & "','" & rLw(i).Kommentar & "','" & rLw(i).Teststatus & "','" & _
   rLw(i).Erklärung & "','" & rLw(i).Normbereich & "','" & rLw(i).NormU & "','" & rLw(i).NormO & "','" & rLw(i).AuftrHinw & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLw(i).RefNr, ",'", rLw(i).Abkü, "','", rLw(i).Langname, "','", rLw(i).Quelle, "','", rLw(i).QSpez, "',", DatForm(rLw(i).AbnDat), ",'", _
   rLw(i).Wert, "','", rLw(i).Einheit, "','", rLw(i).Grenzwerti, "','", rLw(i).Kommentar, "','", rLw(i).Teststatus, "','", _
   rLw(i).Erklärung, "','", rLw(i).Normbereich, "','", rLw(i).NormU, "','", rLw(i).NormO, "','", rLw(i).AuftrHinw, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLw) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(13)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLw), i)
  If Len(rLw(k).Abkü) > maxi(0) Then maxi(0) = Len(rLw(k).Abkü)
  If Len(rLw(k).Langname) > maxi(1) Then maxi(1) = Len(rLw(k).Langname)
  If Len(rLw(k).Quelle) > maxi(2) Then maxi(2) = Len(rLw(k).Quelle)
  If Len(rLw(k).QSpez) > maxi(3) Then maxi(3) = Len(rLw(k).QSpez)
  If Len(rLw(k).Wert) > maxi(4) Then maxi(4) = Len(rLw(k).Wert)
  If Len(rLw(k).Einheit) > maxi(5) Then maxi(5) = Len(rLw(k).Einheit)
  If Len(rLw(k).Grenzwerti) > maxi(6) Then maxi(6) = Len(rLw(k).Grenzwerti)
  If Len(rLw(k).Kommentar) > maxi(7) Then maxi(7) = Len(rLw(k).Kommentar)
  If Len(rLw(k).Teststatus) > maxi(8) Then maxi(8) = Len(rLw(k).Teststatus)
  If Len(rLw(k).Erklärung) > maxi(9) Then maxi(9) = Len(rLw(k).Erklärung)
  If Len(rLw(k).Normbereich) > maxi(10) Then maxi(10) = Len(rLw(k).Normbereich)
  If Len(rLw(k).NormU) > maxi(11) Then maxi(11) = Len(rLw(k).NormU)
  If Len(rLw(k).NormO) > maxi(12) Then maxi(12) = Len(rLw(k).NormO)
  If Len(rLw(k).AuftrHinw) > maxi(13) Then maxi(13) = Len(rLw(k).AuftrHinw)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLw), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLw.Abkü: '" & rLw(k).Abkü & "' -> '" & Left$(rLw(k).Abkü, maxL) & "'", True: rLw(k).Abkü = Left$(rLw(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLw.Langname: '" & rLw(k).Langname & "' -> '" & Left$(rLw(k).Langname, maxL) & "'", True: rLw(k).Langname = Left$(rLw(k).Langname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLw.Quelle: '" & rLw(k).Quelle & "' -> '" & Left$(rLw(k).Quelle, maxL) & "'", True: rLw(k).Quelle = Left$(rLw(k).Quelle, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLw.QSpez: '" & rLw(k).QSpez & "' -> '" & Left$(rLw(k).QSpez, maxL) & "'", True: rLw(k).QSpez = Left$(rLw(k).QSpez, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLw.Wert: '" & rLw(k).Wert & "' -> '" & Left$(rLw(k).Wert, maxL) & "'", True: rLw(k).Wert = Left$(rLw(k).Wert, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLw.Einheit: '" & rLw(k).Einheit & "' -> '" & Left$(rLw(k).Einheit, maxL) & "'", True: rLw(k).Einheit = Left$(rLw(k).Einheit, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLw.Grenzwerti: '" & rLw(k).Grenzwerti & "' -> '" & Left$(rLw(k).Grenzwerti, maxL) & "'", True: rLw(k).Grenzwerti = Left$(rLw(k).Grenzwerti, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLw.Kommentar: '" & rLw(k).Kommentar & "' -> '" & Left$(rLw(k).Kommentar, maxL) & "'", True: rLw(k).Kommentar = Left$(rLw(k).Kommentar, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLw.Teststatus: '" & rLw(k).Teststatus & "' -> '" & Left$(rLw(k).Teststatus, maxL) & "'", True: rLw(k).Teststatus = Left$(rLw(k).Teststatus, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLw.Erklärung: '" & rLw(k).Erklärung & "' -> '" & Left$(rLw(k).Erklärung, maxL) & "'", True: rLw(k).Erklärung = Left$(rLw(k).Erklärung, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLw.Normbereich: '" & rLw(k).Normbereich & "' -> '" & Left$(rLw(k).Normbereich, maxL) & "'", True: rLw(k).Normbereich = Left$(rLw(k).Normbereich, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLw.NormU: '" & rLw(k).NormU & "' -> '" & Left$(rLw(k).NormU, maxL) & "'", True: rLw(k).NormU = Left$(rLw(k).NormU, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLw.NormO: '" & rLw(k).NormO & "' -> '" & Left$(rLw(k).NormO, maxL) & "'", True: rLw(k).NormO = Left$(rLw(k).NormO, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLw.AuftrHinw: '" & rLw(k).AuftrHinw & "' -> '" & Left$(rLw(k).AuftrHinw, maxL) & "'", True: rLw(k).AuftrHinw = Left$(rLw(k).AuftrHinw, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern()
 Dim i%, rAF&, PID$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 PID = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, PID & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl)        values"))
 For i = 1 To UBound(rLL)
'  rLL(i).AktZeit = now()
'   sql = sql0 & "(" & rLL(i).RefNr & ",'" & rLL(i).Abkü & "','" & rLL(i).Verf & "','" & rLL(i).EBM & "','" & rLL(i).goä & "','" & rLL(i).Anzahl & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLL(i).RefNr, ",'", rLL(i).Abkü, "','", rLL(i).Verf, "','", rLL(i).EBM, "','", rLL(i).goä, "','", rLL(i).Anzahl, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLL) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLL) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = DBCn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLL), i)
  If Len(rLL(k).Abkü) > maxi(0) Then maxi(0) = Len(rLL(k).Abkü)
  If Len(rLL(k).Verf) > maxi(1) Then maxi(1) = Len(rLL(k).Verf)
  If Len(rLL(k).EBM) > maxi(2) Then maxi(2) = Len(rLL(k).EBM)
  If Len(rLL(k).goä) > maxi(3) Then maxi(3) = Len(rLL(k).goä)
  If Len(rLL(k).Anzahl) > maxi(4) Then maxi(4) = Len(rLL(k).Anzahl)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLL), i)
      Select Case m
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & Left$(rLL(k).Abkü, maxL) & "'", True: rLL(k).Abkü = Left$(rLL(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & Left$(rLL(k).Verf, maxL) & "'", True: rLL(k).Verf = Left$(rLL(k).Verf, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & Left$(rLL(k).EBM, maxL) & "'", True: rLL(k).EBM = Left$(rLL(k).EBM, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & Left$(rLL(k).goä, maxL) & "'", True: rLL(k).goä = Left$(rLL(k).goä, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & Left$(rLL(k).Anzahl, maxL) & "'", True: rLL(k).Anzahl = Left$(rLL(k).Anzahl, maxL)
      End Select
     Next
    ElseIf maxL < 0 Then
     GoTo nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
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
' Call ForeignNo0 ' 10.5.09 auf quelle1 nötig
' Call ForeignNo1
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
' Call ForeignYes0
' Call ForeignYes1
 Call kvnrueSpeichern
 Call unbekannte_kennungenSpeichern
 Call dmpreiheSpeichern
 Call DBCn.Execute("update namen set aktZeit = " & DatForm(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_id, rAF)
 If rAF <> 1 Then
  frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_id & " " & rNa(0).Nachname & " " & rNa(0).Vorname, True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Resume
End If
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

