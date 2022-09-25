Attribute VB_Name = "Typen"
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
 DienstTel As String 'DienstTel 129 '3629
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
 ZeitPunkt As Date 'ZeitPunkt 135 '6200 + 6201
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
 ZeitPunkt As Date 'ZeitPunkt 135 '
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
 ZeitPunkt As Date 'ZeitPunkt 135 '
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
 ZeitPunkt As Date 'ZeitPunkt 135 '
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
 ZeitPunkt As Date 'ZeitPunkt 135 '
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
 ZeitPunkt As Date 'ZeitPunkt 135 '
 Ziel As String 'Ziel 129 '6291
 Diagnose As String 'Diagnose 129 '6230
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type lbanforderungen
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 ZeitPunkt As Date 'ZeitPunkt 135 '6200 + 6201
 AnfText As String 'AnfText 129 '6280
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Statusbyte
End Type

Public Type laborneu
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 ZeitPunkt As Date 'ZeitPunkt 135 '
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
 ZeitPunkt As Date 'ZeitPunkt 135 '5000 + 6201
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
 ZeitPunkt As Date 'ZeitPunkt 135 'Zeitpunkt, der Speicherung im Turbomed
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
 ZeitPunkt As Date 'ZeitPunkt 135 '6200 + 6201
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
 ZeitPunkt As Date 'ZeitPunkt 135 '6200 + 6201
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

Public Function doLösch(frm As Lese, Tbl$)
 Set rs = dbcn.Execute("select count(*) as ct from `" & Tbl & "`")
 frm.Ausgabe = "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)" & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call dbcn.Execute(sql) ' ,,adasyncexecute
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' LabLösch


Public Function namenSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere namen"
 If Not AllePat Then
  sql = "delete from `namen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,WeggzZahl," & _
     "AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel," & _
     "PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos," & _
     "StByte,Cave,Notiz,zubenach,Verwandt,Sprache) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Weggeldzone,WeggzZahl," & _
     "AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax,DienstTel," & _
     "PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos," & _
     "StByte,Cave,Notiz,zubenach,Verwandt,Sprache)       values"))
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
   rNa(i).Sprache & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rNa(i).Pat_id, ",", rNa(i).lfdnr, ",'", rNa(i).NVorsatz, "','", rNa(i).Nachname, "','", rNa(i).Vorname, "',", datform( _
   rNa(i).GebDat), ",'", rNa(i).Straße, "','", rNa(i).KVKStatus, "','", rNa(i).Geschlecht, "','", rNa(i).Plz, "','", rNa(i).Ort, "','", _
   rNa(i).Weggeldzone, "',", rNa(i).WeggzZahl, ",", datform(rNa(i).AufnDat), ",'", rNa(i).LANR, "','", rNa(i).BStNr, "','", rNa(i).Titel, "','", _
   rNa(i).Versichertennummer, "','", rNa(i).PrivatTel, "','", rNa(i).KVNr, "','", rNa(i).PrivatTel_2, "','", rNa(i).PrivatFax, "','", _
   rNa(i).DienstTel, "','", rNa(i).PrivatMobil, "','", rNa(i).Email, "','", rNa(i).Arbeitgeber, "',", CStr(CInt(rNa(i).AnAllgda)), ",", CStr(CInt( _
   rNa(i).An1da)), ",", CStr(CInt(rNa(i).An2da)), ",", CStr(CInt(rNa(i).Checkda)), ",'", rNa(i).DMTypaD, "',", datform( _
    0), ",", rNa(i).absPos, ",", rNa(i).StByte, ",'", rNa(i).Cave, "','", rNa(i).Notiz, "','", rNa(i).zubenach, "','", rNa(i).Verwandt, "','", _
   rNa(i).Sprache, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rNa) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
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
 dbcn.BeginTrans
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere faelle"
 If Not AllePat Then
  sql = "delete from `faelle` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,GebOr,AbrGb,PersKreis,SKtZusatz,f4206," & _
     "ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,AndÜw,Übw,ÜbwLANR," & _
     "ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,statKlasse,f4237,statBehTage," & _
     "SchGr,Weiterbeh,PGeb,PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam," & _
     "AdStr,AdPlz,AdOrt,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf," & _
     "QEnd,QS,QT,TherArt,StByte,absPos) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Pat_ID,Quartal,Nachname," & _
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
   rFa(i).VKNr, "',", datform(rFa(i).BhFB), ",", datform(rFa(i).BhFE1), ",", datform(rFa(i).BhFE2), ",'", rFa(i).f4202, "',", datform( _
   rFa(i).ausgst), ",'", rFa(i).KtrAbrB, "','", rFa(i).AbrAr, "',", datform(rFa(i).lVorl), ",'", rFa(i).IK, "','", rFa(i).KVKs, "','", _
   rFa(i).KVKserg, "','", rFa(i).GebOr, "','", rFa(i).AbrGb, "','", rFa(i).PersKreis, "','", rFa(i).SKtZusatz, "','", rFa(i).f4206, "','", _
   rFa(i).ÜwText, "',", rFa(i).f4210, ",", rFa(i).AkfHAH, ",", rFa(i).AkfAB0, ",", rFa(i).AkfAK, ",'", rFa(i).statNuller, "','", _
   rFa(i).ÜbwV, "','", rFa(i).AndÜw, "','", rFa(i).Übw, "','", rFa(i).ÜbwLANR, "','", rFa(i).ÜWZiel, "','", rFa(i).ÜWNNr, "','", _
   rFa(i).ÜWNaN, "','", rFa(i).ÜWTit, "','", rFa(i).ÜWVor, "','", rFa(i).ÜWVsw, "',", rFa(i).üwvid, ",'", rFa(i).statKlasse, "','", _
   rFa(i).f4237, "',", rFa(i).statBehTage, ",", rFa(i).SchGr, ",'", rFa(i).Weiterbeh, "','", rFa(i).PGeb, "','", rFa(i).PGebErg, "','", _
   rFa(i).Mahnfrist, "','", rFa(i).GOÄKatNr, "','", rFa(i).GOÄKatName, "','", rFa(i).abrArzt, "','", rFa(i).privVers, "','", _
   rFa(i).AdNam, "','", rFa(i).AdStr, "','", rFa(i).AdPlz, "','", rFa(i).AdOrt, "',", datform(rFa(i).BhFE), ",'", rFa(i).s8000, "','", _
   rFa(i).s8100, "',", datform(rFa(i).AktZeit), ",", datform(rFa(i).Fanf), ",'", rFa(i).altQuart, "',", datform(rFa(i).QAnf), ",", datform( _
   rFa(i).QEnd), ",'", rFa(i).QS, "','", rFa(i).QT, "',", rFa(i).TherArt, ",", rFa(i).StByte, ",", rFa(i).absPos, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFa) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
   If Not rs.BOF() Then
    If rs!code = 1265 Then
     Err.Raise -2147217833
    End If
   End If
  End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = dbcn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_id & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & datform(rFa(i).BhFB) & " and bhfe1 = " & datform(rFa(i).BhFE1) & " and ausgst = " & datform(rFa(i).ausgst))
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
  End If
 Next i
 syscmd 5
 Exit Function
fehler:
If Err.Number = -2147217833 Then
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
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
   Case "PersKreis":  If SpMod(rFa(i).PersKreis, "faelle", rsc) Then Exit Do
   Case "SKtZusatz":  If SpMod(rFa(i).SKtZusatz, "faelle", rsc) Then Exit Do
   Case "f4206":  If SpMod(rFa(i).f4206, "faelle", rsc) Then Exit Do
   Case "ÜwText":  If SpMod(rFa(i).ÜwText, "faelle", rsc) Then Exit Do
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
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "insert into `kassenliste`(`GO`,`VK`,`IK`) values (" & "'" & rFa(i).GOÄKatName & "', '" & rFa(i).VKNr & "', '" & rFa(i).IK & "')"
 Call dbcn.Execute(sqlquer, rAf)
 Resume
End If ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere au"
 If Not AllePat Then
  sql = "delete from `au` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            values"))
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rAu(i).FID & "," & rAu(i).Pat_ID & "," & datform(rAu(i).ZeitPunkt) & ",'" & rAu(i).Beginn & "','" & rAu(i).Ende & "','" & rAu(i).ICDs & "'," & _
   rAu(i).absPos & "," & datform(rAu(i).AktZeit) & "," & rAu(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rAu(i).FID, ",", rAu(i).Pat_id, ",", datform(rAu(i).ZeitPunkt), ",'", rAu(i).Beginn, "','", rAu(i).Ende, "','", rAu(i).ICDs, "',", _
   rAu(i).absPos, ",", datform(rAu(i).AktZeit), ",", rAu(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rAu) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Beginn":  If SpMod(rAu(i).Beginn, "au", rsc) Then Exit Do
   Case "Ende":  If SpMod(rAu(i).Ende, "au", rsc) Then Exit Do
   Case "ICDs":  If SpMod(rAu(i).ICDs, "au", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere briefe"
 If Not AllePat Then
  sql = "delete from `briefe` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,QS,QT,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
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
  csql.AppVar Array("(", rBr(i).FID, ",", rBr(i).Pat_id, ",", datform(rBr(i).ZeitPunkt), ",'", rBr(i).Pfad, "','", rBr(i).Art, "','", rBr(i).Name, "','", _
   rBr(i).Typ, "',", datform(rBr(i).AktZeit), ",", rBr(i).DokGroe, ",'", rBr(i).QS, "','", rBr(i).QT, "',", rBr(i).absPos, ",", _
   rBr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rBr) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
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
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere diagnosen"
 If Not AllePat Then
  sql = "delete from `diagnosen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,Ausnahme,intBemerk,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_id,GesName," & _
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
  csql.AppVar Array("(", rDi(i).FID, ",", rDi(i).Pat_id, ",'", rDi(i).GesName, "',", datform(rDi(i).DiagDatum), ",'", rDi(i).DiagSicherheit, "','", _
   rDi(i).DiagText, "','", rDi(i).DiagSeite, "','", rDi(i).DiagAttr, "','", rDi(i).ICD, "',", rDi(i).obDauer, ",'", rDi(i).Ausnahme, "','", _
   rDi(i).intBemerk, "',", rDi(i).absPos, ",", datform(rDi(i).AktZeit), ",", rDi(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDi) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "GesName":  If SpMod(rDi(i).GesName, "diagnosen", rsc) Then Exit Do
   Case "DiagSicherheit":  If SpMod(rDi(i).DiagSicherheit, "diagnosen", rsc) Then Exit Do
   Case "DiagText":  If SpMod(rDi(i).DiagText, "diagnosen", rsc) Then Exit Do
   Case "DiagSeite":  If SpMod(rDi(i).DiagSeite, "diagnosen", rsc) Then Exit Do
   Case "DiagAttr":  If SpMod(rDi(i).DiagAttr, "diagnosen", rsc) Then Exit Do
   Case "ICD":  If SpMod(rDi(i).ICD, "diagnosen", rsc) Then Exit Do
   Case "Ausnahme":  If SpMod(rDi(i).Ausnahme, "diagnosen", rsc) Then Exit Do
   Case "intBemerk":  If SpMod(rDi(i).intBemerk, "diagnosen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere dokumente"
 If Not AllePat Then
  sql = "delete from `dokumente` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
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
  csql.AppVar Array("(", rDo(i).FID, ",", rDo(i).Pat_id, ",", datform(rDo(i).ZeitPunkt), ",'", rDo(i).DokPfad, "','", rDo(i).DokArt, "','", _
   rDo(i).DokName, "',", datform(rDo(i).Quelldatum), ",", rDo(i).absPos, ",", datform(rDo(i).AktZeit), ",", rDo(i).DokGroe, ",'", rDo(i).QS, "','", _
   rDo(i).QT, "',", rDo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDo) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
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
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere eintraege"
 If Not AllePat Then
  sql = "delete from `eintraege` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte)            values"))
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rEi(i).FID & "," & rEi(i).Pat_ID & "," & datform(rEi(i).ZeitPunkt) & ",'" & rEi(i).Art & "','" & rEi(i).Inhalt & "'," & rEi(i).absPos & "," & datform( _
   rEi(i).AktZeit) & ",'" & rEi(i).QS & "','" & rEi(i).QT & "'," & rEi(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rEi(i).FID, ",", rEi(i).Pat_id, ",", datform(rEi(i).ZeitPunkt), ",'", rEi(i).Art, "','", rEi(i).Inhalt, "',", rEi(i).absPos, ",", datform( _
   rEi(i).AktZeit), ",'", rEi(i).QS, "','", rEi(i).QT, "',", rEi(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rEi) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Art":  If SpMod(rEi(i).Art, "eintraege", rsc) Then Exit Do
   Case "Inhalt":  If SpMod(rEi(i).Inhalt, "eintraege", rsc) Then Exit Do
   Case "QS":  If SpMod(rEi(i).QS, "eintraege", rsc) Then Exit Do
   Case "QT":  If SpMod(rEi(i).QT, "eintraege", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere forminhaltform_abk"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhaltform_abk` (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Form_Abk)         values"))
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
'   sql = sql0 & "('" & rFi(i).Form_Abk & "')"
  If Lese.dlg.SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rFi(i).Form_Abk, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFi) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFi(i).Form_Abk, "forminhaltform_abk", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere formulare"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      values"))
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFo(i).FormID & ",'" & rFo(i).Form_Abk & "','" & rFo(i).FormBez & "','" & rFo(i).FormVorl & "'," & datform(rFo(i).AktZeit) & "," & _
   rFo(i).absPos & "," & rFo(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFo(i).FormID, ",'", rFo(i).Form_Abk, "','", rFo(i).FormBez, "','", rFo(i).FormVorl, "',", datform(rFo(i).AktZeit), ",", _
   rFo(i).absPos, ",", rFo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFo) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Form_Abk":  If SpMod(rFo(i).Form_Abk, "formulare", rsc) Then Exit Do
   Case "FormBez":  If SpMod(rFo(i).FormBez, "formulare", rsc) Then Exit Do
   Case "FormVorl":  If SpMod(rFo(i).FormVorl, "formulare", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere forminhkopf"
 If Not AllePat Then
'  sql = "delete FROM `forminhfeld` where foid in (select foid from `forminhkopf` where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "delete from `forminhkopf` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge)       values"))
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rFr(i).FoID & "," & rFr(i).FID & "," & rFr(i).Pat_ID & "," & rFr(i).Form_ID & "," & datform(rFr(i).ZeitPunkt) & "," & rFr(i).AbsPos & "," & datform( _
   rFr(i).AktZeit) & "," & rFr(i).StByte & ",'" & rFr(i).Satzart & "','" & rFr(i).Satzlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFr(i).FoID, ",", rFr(i).FID, ",", rFr(i).Pat_id, ",", rFr(i).Form_ID, ",", datform(rFr(i).ZeitPunkt), ",", rFr(i).absPos, ",", datform( _
   rFr(i).AktZeit), ",", rFr(i).StByte, ",'", rFr(i).Satzart, "','", rFr(i).Satzlänge, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFr) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Satzart":  If SpMod(rFr(i).Satzart, "forminhkopf", rsc) Then Exit Do
   Case "Satzlänge":  If SpMod(rFr(i).Satzlänge, "forminhkopf", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere forminhfeld"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FoID,Nr,FeldNr," & _
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
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere kheinweis"
 If Not AllePat Then
  sql = "delete from `kheinweis` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               values"))
 For i = 1 To UBound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKh(i).FID & "," & rKh(i).Pat_ID & "," & datform(rKh(i).ZeitPunkt) & ",'" & rKh(i).Ziel & "','" & rKh(i).Diagnose & "'," & rKh(i).absPos & "," & datform( _
   rKh(i).AktZeit) & "," & rKh(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKh(i).FID, ",", rKh(i).Pat_id, ",", datform(rKh(i).ZeitPunkt), ",'", rKh(i).Ziel, "','", rKh(i).Diagnose, "',", rKh(i).absPos, ",", datform( _
   rKh(i).AktZeit), ",", rKh(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKh) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Ziel":  If SpMod(rKh(i).Ziel, "kheinweis", rsc) Then Exit Do
   Case "Diagnose":  If SpMod(rKh(i).Diagnose, "kheinweis", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere lbanforderungen"
 If Not AllePat Then
  sql = "delete from `lbanforderungen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       values"))
 For i = 1 To UBound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rLb(i).FID & "," & rLb(i).Pat_ID & "," & datform(rLb(i).ZeitPunkt) & ",'" & rLb(i).AnfText & "'," & rLb(i).absPos & "," & datform( _
   rLb(i).AktZeit) & "," & rLb(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLb(i).FID, ",", rLb(i).Pat_id, ",", datform(rLb(i).ZeitPunkt), ",'", rLb(i).AnfText, "',", rLb(i).absPos, ",", datform( _
   rLb(i).AktZeit), ",", rLb(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLb) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "AnfText":  If SpMod(rLb(i).AnfText, "lbanforderungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborneu"
 If Not AllePat Then
  sql = "delete from `laborneu` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
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
  csql.AppVar Array("(", rLa(i).FID, ",", rLa(i).Pat_id, ",", datform(rLa(i).ZeitPunkt), ",'", rLa(i).FertigStGrad, "','", rLa(i).Abkü, "',", _
   rLa(i).LangtextVW, ",'", rLa(i).Wert, "','", rLa(i).Einheit, "',", rLa(i).KommentarVW, ",", rLa(i).absPos, ",", datform(rLa(i).AktZeit), ",", _
   rLa(i).RefNr, ",", rLa(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLa) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "FertigStGrad":  If SpMod(rLa(i).FertigStGrad, "laborneu", rsc) Then Exit Do
   Case "Abkü":  If SpMod(rLa(i).Abkü, "laborneu", rsc) Then Exit Do
   Case "Wert":  If SpMod(rLa(i).Wert, "laborneu", rsc) Then Exit Do
   Case "Einheit":  If SpMod(rLa(i).Einheit, "laborneu", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere leistungen"
 If Not AllePat Then
  sql = "delete from `leistungen` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
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
  csql.AppVar Array("(", rLe(i).FID, ",", rLe(i).Pat_id, ",", datform(rLe(i).ZeitPunkt), ",'", rLe(i).Leistung, "','", rLe(i).f5002, "','", _
   rLe(i).f5005, "','", rLe(i).f5006, "','", rLe(i).f5009, "','", rLe(i).Med, "','", rLe(i).f5015, "','", rLe(i).f5016, "','", rLe(i).f5021, "','", _
   rLe(i).f5026, "','", rLe(i).Faktor, "','", rLe(i).f5098, "','", rLe(i).LANR, "',", datform(rLe(i).letzVorg), ",'", rLe(i).Ausn, "','", _
   rLe(i).Beme, "',", rLe(i).absPos, ",", datform(rLe(i).AktZeit), ",'", rLe(i).QS, "','", rLe(i).QT, "',", rLe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLe) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
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
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere medplan"
 If Not AllePat Then
  sql = "delete from `medplan` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,MPNr," & _
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
  csql.AppVar Array("(", rMe(i).FID, ",", rMe(i).Pat_id, ",", rMe(i).MPNr, ",", datform(rMe(i).ZeitPunkt), ",", datform(rMe(i).Datum), ",'", _
   rMe(i).Medikament, "','", rMe(i).MedAnfang, "',", rMe(i).FeldNr, ",'", rMe(i).mo, "','", rMe(i).mi, "','", rMe(i).nm, "','", rMe(i).ab, "','", _
   rMe(i).Zn, "',", CStr(CInt(rMe(i).bBed)), ",'", rMe(i).Bemerkung, "',", rMe(i).absPos, ",", datform(rMe(i).AktZeit), ",", _
   rMe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rMe) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Medikament":  If SpMod(rMe(i).Medikament, "medplan", rsc) Then Exit Do
   Case "MedAnfang":  If SpMod(rMe(i).MedAnfang, "medplan", rsc) Then Exit Do
   Case "mo":  If SpMod(rMe(i).mo, "medplan", rsc) Then Exit Do
   Case "mi":  If SpMod(rMe(i).mi, "medplan", rsc) Then Exit Do
   Case "nm":  If SpMod(rMe(i).nm, "medplan", rsc) Then Exit Do
   Case "ab":  If SpMod(rMe(i).ab, "medplan", rsc) Then Exit Do
   Case "zn":  If SpMod(rMe(i).Zn, "medplan", rsc) Then Exit Do
   Case "Bemerkung":  If SpMod(rMe(i).Bemerkung, "medplan", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere rezepteintraege"
 If Not AllePat Then
  sql = "delete from `rezepteintraege` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte)  values"))
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRe(i).FID & "," & rRe(i).Pat_ID & "," & datform(rRe(i).ZeitPunkt) & ",'" & rRe(i).Rezept & "','" & rRe(i).Rezeptklasse & "','" & _
   rRe(i).Medikament & "','" & rRe(i).PZN & "'," & rRe(i).absPos & "," & datform(rRe(i).AktZeit) & ",'" & rRe(i).QS & "','" & rRe(i).QT & "'," & rRe(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRe(i).FID, ",", rRe(i).Pat_id, ",", datform(rRe(i).ZeitPunkt), ",'", rRe(i).Rezept, "','", rRe(i).Rezeptklasse, "','", _
   rRe(i).Medikament, "','", rRe(i).PZN, "',", rRe(i).absPos, ",", datform(rRe(i).AktZeit), ",'", rRe(i).QS, "','", rRe(i).QT, "',", _
   rRe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRe) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
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
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere rr"
 If Not AllePat Then
  sql = "delete from `rr` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (FID,Pat_ID,ZeitPunkt," & _
     "RR,absPos,AktZeit,StByte)            values"))
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rRr(i).FID & "," & rRr(i).Pat_ID & "," & datform(rRr(i).ZeitPunkt) & ",'" & rRr(i).RR & "'," & rRr(i).absPos & "," & datform(rRr(i).AktZeit) & "," & _
   rRr(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRr(i).FID, ",", rRr(i).Pat_id, ",", datform(rRr(i).ZeitPunkt), ",'", rRr(i).RR, "',", rRr(i).absPos, ",", datform( _
   rRr(i).AktZeit), ",", rRr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRr) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "RR":  If SpMod(rRr(i).RR, "rr", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere kvnrue"
 If Not AllePat Then
  sql = "delete from `kvnrue` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        values"))
 For i = 1 To UBound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
'   sql = sql0 & "(" & rKv(i).Pat_ID & ",'" & rKv(i).KVNr & "'," & rKv(i).absPos & "," & datform(rKv(i).AktZeit) & "," & rKv(i).StByte & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKv(i).Pat_id, ",'", rKv(i).KVNr, "',", rKv(i).absPos, ",", datform(rKv(i).AktZeit), ",", rKv(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKv) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "KVNr":  If SpMod(rKv(i).KVNr, "kvnrue", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error Resume Next
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere unbekannte_kennungen"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Kennung,absPos,StByte," & _
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
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Kennung":  If SpMod(rUn(i).Kennung, "unbekannte kennungen", rsc) Then Exit Do
   Case "Inhalt":  If SpMod(rUn(i).Inhalt, "unbekannte kennungen", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere dmpreihe"
 If Not AllePat Then
  sql = "delete from `dmpreihe` where Pat_ID = " + CStr(rNa(0).Pat_id)
  Call dbcn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,NachName,VorName,GebDat,Pat_id,StByte,AktZeit) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Abk,Art,KarteiDatum," & _
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
  csql.AppVar Array("('", rDm(i).Abk, "','", rDm(i).Art, "',", datform(rDm(i).KarteiDatum), ",", datform(rDm(i).exportiert), ",", datform(rDm(i).DokuDatum), ",", CStr(CInt( _
   rDm(i).obvoll)), ",'", rDm(i).Nachname, "','", rDm(i).Vorname, "',", datform(rDm(i).GebDat), ",", rDm(i).Pat_id, ",", _
   rDm(i).StByte, ",", datform(rDm(i).AktZeit), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDm) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 dbcn.CommitTrans
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Abk":  If SpMod(rDm(i).Abk, "dmpreihe", rsc) Then Exit Do
   Case "Art":  If SpMod(rDm(i).Art, "dmpreihe", rsc) Then Exit Do
   Case "NachName":  If SpMod(rDm(i).Nachname, "dmpreihe", rsc) Then Exit Do
   Case "VorName":  If SpMod(rDm(i).Vorname, "dmpreihe", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 dbcn.BeginTrans
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function laborxsaetzeSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,PLZPraxis,OrtPraxis,Labor,StraßeLabor,PLZLabor," & _
     "OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,PLZPraxis,OrtPraxis,Labor,StraßeLabor,PLZLabor," & _
     "OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)      values"))
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
'   sql = sql0 & "(" & rLs(i).DatID & ",'" & rLs(i).Satzart & "','" & rLs(i).Satzlänge & "','" & rLs(i).SatzlängeSchluss & "','" & rLs(i).VersionSatzb & "','" & _
   rLs(i).Arztnr & "','" & rLs(i).Arztname & "','" & rLs(i).StraßePraxis & "','" & rLs(i).PLZPraxis & "','" & rLs(i).OrtPraxis & "','" & rLs(i).Labor & "','" & _
   rLs(i).StraßeLabor & "','" & rLs(i).PLZLabor & "','" & rLs(i).OrtLabor & "','" & rLs(i).KBVPrüfnr & "','" & rLs(i).Zeichensatz & "','" & _
   rLs(i).Kundenarztnr & "','" & rLs(i).Erstellungsdatum & "','" & rLs(i).Gesamtlänge & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLs(i).DatID, ",'", rLs(i).Satzart, "','", rLs(i).Satzlänge, "','", rLs(i).SatzlängeSchluss, "','", rLs(i).VersionSatzb, "','", _
   rLs(i).Arztnr, "','", rLs(i).Arztname, "','", rLs(i).StraßePraxis, "','", rLs(i).PLZPraxis, "','", rLs(i).OrtPraxis, "','", _
   rLs(i).Labor, "','", rLs(i).StraßeLabor, "','", rLs(i).PLZLabor, "','", rLs(i).OrtLabor, "','", rLs(i).KBVPrüfnr, "','", rLs(i).Zeichensatz, "','", _
   rLs(i).Kundenarztnr, "','", rLs(i).Erstellungsdatum, "','", rLs(i).Gesamtlänge, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLs) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
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
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (Pfad,Name,Zp," & _
     "fertig)  values"))
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
'   sql = sql0 & "('" & rLg(i).Pfad & "','" & rLg(i).Name & "'," & datform(rLg(i).Zp) & "," & cstr(cint(rLg(i).fertig)) & ")"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rLg(i).Pfad, "','", rLg(i).Name, "',", datform(rLg(i).Zp), ",", CStr(CInt(rLg(i).fertig)), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLg) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 Do While Not rsc.EOF
  Select Case rsc!column_name
   Case "Pfad":  If SpMod(rLg(i).Pfad, "laborxeingel", rsc) Then Exit Do
   Case "Name":  If SpMod(rLg(i).Name, "laborxeingel", rsc) Then Exit Do
  End Select
  rsc.Move 1
 Loop
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxus"
' sql0 = " insert " & sqlignore &  "into `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG,Pat_idErwVN," & _
     "Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen,AfN) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (DatID,SatzID,Satzart," & _
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
   rLu(i).Auftragsschlüssel, "',", datform(rLu(i).Eingang), ",'", rLu(i).Berichtsdatum, "',", rLu(i).Pat_id, ",'", rLu(i).Nachname, "','", _
   rLu(i).Vorname, "','", rLu(i).GebDat, "','", rLu(i).Titel, "','", rLu(i).NVorsatz, "','", rLu(i).BefArt, "','", rLu(i).Abrechnungstyp, "','", _
   rLu(i).GebüOrd, "','", rLu(i).Patienteninformation, "','", rLu(i).Geschlecht, "','", rLu(i).AuftrHinw, "','", rLu(i).Pat_idUrsp, "','", _
   rLu(i).Pat_idErwVNG, "','", rLu(i).Pat_idErwVN, "','", rLu(i).Pat_idErwG, "','", rLu(i).Pat_idErwGB, "','", rLu(i).Pat_idErwGL, "','", _
   rLu(i).Pat_idLaborNeu, "',", datform(rLu(i).ZeitpunktLaborneu), ",", rLu(i).ZdüP, ",", rLu(i).ZdiP, ",'", rLu(i).LWerte, "',", datform( _
   rLu(i).verglichen), ",", rLu(i).AfN, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLu) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
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
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  values"))
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
'   sql = sql0 & "(" & rLo(i).RefNr & ",'" & rLo(i).Verf & "','" & rLo(i).KuQu & "','" & rLo(i).Quelle & "','" & rLo(i).QSpez & "'," & datform(rLo(i).AbnDat) & ",'" & _
   rLo(i).Kommentar & "','" & rLo(i).Erklärung & "','" & rLo(i).Keimzahl & "')"
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLo(i).RefNr, ",'", rLo(i).Verf, "','", rLo(i).KuQu, "','", rLo(i).Quelle, "','", rLo(i).QSpez, "',", datform(rLo(i).AbnDat), ",'", _
   rLo(i).Kommentar, "','", rLo(i).Erklärung, "','", rLo(i).Keimzahl, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLo) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
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
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,Normbereich," & _
     "NormU,NormO,AuftrHinw) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (RefNr,Abkü,Langname," & _
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
  csql.AppVar Array("(", rLw(i).RefNr, ",'", rLw(i).Abkü, "','", rLw(i).Langname, "','", rLw(i).Quelle, "','", rLw(i).QSpez, "',", datform(rLw(i).AbnDat), ",'", _
   rLw(i).Wert, "','", rLw(i).Einheit, "','", rLw(i).Grenzwerti, "','", rLw(i).Kommentar, "','", rLw(i).Teststatus, "','", _
   rLw(i).Erklärung, "','", rLw(i).Normbereich, "','", rLw(i).NormU, "','", rLw(i).NormO, "','", rLw(i).AuftrHinw, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLw) Then
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
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
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern()
 Dim i%, rAf& ',sql0$
 Dim csql0 As New CString, csql As New CString
 T1 = Timer
 On Error GoTo fehler
 syscmd 4, CStr(rNa(0).Pat_id) & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `", LCase(Tbn(i)), "`  (RefNr,Abkü,Verf," & _
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
  Call dbcn.Execute(csql.Value, rAf) ', , adAsyncExecute)
  If Lese.dlg.SammelInsert = 0 Then csql.m_Len = 0
  If lies.obMySQL And obMitAlterTab Then
   Set rs = dbcn.Execute("show warnings")
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
 Dim rsc As ADODB.Recordset
 Set rsc = New ADODB.Recordset
 Set rsc = dbcn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
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
 If lies.obMySQL Then Resume Next Else Resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function doSpeichern(frm As Lese)
 Dim rAf&
 On Error GoTo fehler
 Call namenSpeichern
 Call faelleSpeichern
   If Not lies.obMySQL Then
    Call dbcn.CommitTrans
    Call dbcn.BeginTrans
   End If
 Call auSpeichern
 Call briefeSpeichern
 Call diagnosenSpeichern
 Call dokumenteSpeichern
 Call eintraegeSpeichern
 Call forminhaltform_abkSpeichern
   If Not lies.obMySQL Then
    Call dbcn.CommitTrans
    Call dbcn.BeginTrans
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
 Call dbcn.Execute("update namen set aktZeit = " & datform(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_id, rAf)
 If rAf <> 1 Then
  frm.Ausgabe = "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_id & " " & rNa(0).Nachname & " " & rNa(0).Vorname & vbcrfl & altAusgabe
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

