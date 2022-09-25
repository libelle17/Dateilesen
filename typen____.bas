Attribute VB_Name = "Typen"
Option Explicit
Public obFor%
Dim sql$, T1!, T2!, rs As ADODB.Recordset, maxL%

Public Type namen
 Pat_id As Long 'Pat_ID 3 '3000
 lfdnr As Long 'lfdnr 3 'laufende Patientennummer
 NVorsatz As String 'NVorsatz 130 '3100
 Nachname As String 'Nachname 130 '3101
 Vorname As String 'Vorname 130 '3102
 GebDat As Date 'GebDat 135 '3103
 Straße As String 'Straße 130 '3107
 KVKStatus As String 'KVKStatus 130 '3108
 Geschlecht As String 'Geschlecht 130 '3110
 Plz As String 'Plz 130 '3112
 Ort As String 'Ort 130 '3113
 Postfach As String 'Postfach 130 '3216
 Weggeldzone As String 'Weggeldzone 130 '3631 (1) Weggeldzone mit Z
 WeggzZahl As Double 'WeggzZahl 131 '3631 (2) Weggeldzone, Zahl in Feld 2
 AufnDat As Date 'AufnDat 135 '3610
 Lanr As String 'LANR 130 '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr As String 'BStNr 130 '3536 Betriebsstättennummer
 Titel As String 'Titel 130 '3104
 Versichertennummer As String 'Versichertennummer 130 '3105
 PrivatTel As String 'PrivatTel 130 '3629
 KVNr As String 'KVNr 130 '3630
 PrivatTel_2 As String 'PrivatTel_2 130 '3629
 PrivatFax As String 'PrivatFax 130 '3629
 DienstTel As String 'DienstTel 130 '3629
 PrivatMobil As String 'PrivatMobil 130 '3629
 Email As String 'Email 130 'Email
 Arbeitgeber As String 'Arbeitgeber 130 '3625
 AnAllgda As Integer 'AnAllgda 11 'Anamnese allgemein da
 An1da As Integer 'An1da 11 'Anamnese S.1 da
 An2da As Integer 'An2da 11 'Anamnese S.2 da
 Checkda As Integer 'Checkda 11 'Checkliste da
 DMTypaD As String 'DMTypaD 130 'aus Diagnosen
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Cave As String 'Cave 130 '3654
 Notiz As String 'Notiz 130 '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>
 f3800 As String 'f3800 130 '3800
 dmpklass As Long 'dmpklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpbeg As Date 'dmpbeg 133 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass As Long 'dmpkhkklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg As Date 'dmpkhkbeg 133 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass As Long 'dmpcopdklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg As Date 'dmpcopdbeg 133 'Datum der aktuellen DMP-Klassifikation
 getHA0 As Long 'getHA0 3 'KVNr aus getHausarzt -> Üw(12,0)
 fnHA0 As String 'fnHA0 130 'Funktion aus getHausarzt -> Üw(10,0)
 getHA1 As Long 'getHA1 3 'KVNr aus getHausarzt -> Üw(12,1)
 fnHA1 As String 'fnHA1 130 'Funktion aus getHausarzt -> Üw(10,1)
 getHA2 As Long 'getHA2 3 'KVNr aus getHausarzt -> Üw(12,2)
 fnHA2 As String 'fnHA2 130 'Funktion aus getHausarzt -> Üw(10,2)
 zubenach As String 'zubenach 130 '3633
 Verwandt As String 'Verwandt 130 '3632
 Sprache As String 'Sprache 130 '3628
 lAktTM As Date 'lAktTM 135 'letzte Aktualisierung in Turbomed
End Type

Public Type Faelle
 FID As Long 'FID 3 '
 Pat_id As Long 'Pat_ID 3 '3000
 Quartal As String 'Quartal 130 '4101
 Nachname As String 'Nachname 130 '3101
 Vorname As String 'Vorname 130 '3102
 lfdnr As Long 'lfdnr 3 'laufende Fallnummer
 TMFNr As String 'TMFNr 130 '4144 Fallnummer in Turbomed
 VKNr As String 'VKNr 130 '4104
 BhFB As Date 'BhFB 135 '4150
 BhFE1 As Date 'BhFE1 135 '4151
 BhFE2 As Date 'BhFE2 135 '4152
 f4202 As String 'f4202 130 '4202
 ausgst As Date 'ausgst 135 '4102 ('ausgestellt am')
 KtrAbrB As String 'KtrAbrB 130 '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr As String 'AbrAr 130 '4107, Abrechnungsart (1 = Primärkassen)
 lVorl As Date 'lVorl 135 '4109, letzte Vorlage
 IK As String 'IK 130 '4111 Krankenkassennummer (IK)
 KVKs As String 'KVKs 130 '4112 Versichertenstatus VK
 KVKserg As String 'KVKserg 130 '4113 Ost/West-Status VK
 Kasse As String 'Kasse 130 '6299 Kasse (aus Formularen)
 GebOr As String 'GebOr 130 '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As String 'AbrGb 130 '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis As String 'PersKreis 130 '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz As String 'SKtZusatz 130 '4124 SKT-Zusatzangaben
 f4206 As String 'f4206 130 '4206, mutmasslicher Tag der Entbindung
 ÜwText As String 'ÜwText 130 '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 As Byte 'f4210 16 '4210, Ankreuzfeld LSR
 AkfHAH As Byte 'AkfHAH 16 '4211 Ankreuzfeld HAH
 AkfAB0 As Byte 'AkfAB0 16 '4212 Ankreuzfeld AB0.RH
 AkfAK As Byte 'AkfAK 16 '4213 Ankreuzfeld AK
 statNuller As String 'statNuller 130 '4216, nu bei Musterfrau 16 Nuller
 ÜbwV As String 'ÜbwV 130 '4218, überwiesen von Arztnummer
 ÜbWVLANR As String 'ÜbWVLANR 130 '4218(1) überwiesen von LANR
 ÜbWVBSNR As String 'ÜbWVBSNR 130 '4218(2) überwiesen von BSNR
 ÜbWVKVNR As String 'ÜbWVKVNR 130 '4218(3) überwiesen von KVNR
 AndÜw As String 'AndÜw 130 '4219, anderer Überweiser
 Übwr As String 'Übwr 130 'resultierender Überweiser (BSNR): 4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR As String 'ÜbwLANR 130 '4242 LANR des Überweisers
 ÜWZiel As String 'ÜWZiel 130 '4220 Überweisung an
 ÜWNNr As String 'ÜWNNr 130 '4231(4): KV-Nummer des Überweisers
 ÜWNaN As String 'ÜWNaN 130 '4231(3): Nachname des Überweisers
 ÜWTit As String 'ÜWTit 130 '4231(3): Titel des Überweisers
 ÜWVor As String 'ÜWVor 130 '4231(2): Vorname des Überweisers
 ÜWVsw As String 'ÜWVsw 130 '4231(2b): Vorsatzwort des Überweisers
 üwvid As Long 'üwvid 3 '4247 Bezug auf ueberwvon
 Auftrag As String 'Auftrag 130 '4205 Auftrag bei Überweisung
 Verdacht As String 'Verdacht 130 '4207 Verdacht bei Überweisung

 Befund As String 'Befund 130 '4208 Befund bei Überweisung
 statKlasse As String 'statKlasse 130 '4236 Klasse bei Behandlung
 f4237 As String 'f4237 130 '4237 Krankenhausname
 statBehTage As Long 'statBehTage 3 '4238 Krankenhausaufenthalt
 SchGr As Double 'SchGr 131 '4239, Schein(unter)gruppe
 Weiterbeh As String 'Weiterbeh 130 '4243, Weiterbehandelnder
 PGeb As String 'PGeb 130 '4401, Praxisgebühr
 PGebErg As String 'PGebErg 130 '4402, ?
 Mahnfrist As String 'Mahnfrist 130 '4403, Mahnfrist bis
 GOÄKatNr As String 'GOÄKatNr 130 '4580 (1): Katalog-Nummer
 GOÄKatName As String 'GOÄKatName 130 '4580 (2): Privat-Abrechnungskatalog
 abrArzt As String 'abrArzt 130 '4585 abrechnender Arzt
 privVers As String 'privVers 130 '4586 private Versicherung
 AdNam As String 'AdNam 130 '4602(1) Name Rechnungsanschrift
 AdStr As String 'AdStr 130 '4602(2) Straße Rechnungsanschrift
 AdPlz As String 'AdPlz 130 '4602(3) PLZ Rechnungsanschrift
 AdOrt As String 'AdOrt 130 '4602(4) Ort Rechnungsanschrift
 BhFE As Date 'BhFE 135 '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As String 's8000 130 '8000, Satzidentifikation
 s8100 As String 's8100 130 '8100 Satzlänge
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 Fanf As Date 'Fanf 135 'Fallanfang
 altQuart As String 'altQuart 130 '
 QAnf As Date 'QAnf 135 'Quartalsanfang
 QEnd As Date 'QEnd 135 'Quartalsende
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 therart As Long 'TherArt 3 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 lanrid As Long 'LANRid 3 'Bezug auf lanrpraxis.id
 f4108 As String 'f4108 130 '4108
 BGFallNr As String 'BGFallNr 130 '3603 BG-Fall-Nummer
 lGewicht As Double 'lGewicht 131 'letztes Gewicht in kg
End Type

Public Type au
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 Beginn As String 'Beginn 130 '6285 1. Hälfte
 Ende As String 'Ende 130 '6285 2. Hälfte
 ICDs As String 'ICDs 130 '6286
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type briefe
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 Pfad As String 'Pfad 130 '
 Art As String 'Art 130 '
 Name As String 'Name 130 '
 Typ As String 'Typ 130 '
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As Long 'DokGroe 3 'Größe der Datei
 DokAenD As Date 'DokAenD 135 'Dokument-letzte Änderung
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type Diagnosen
 ID1 As Long 'ID1 3 '
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_id 3 'Bezug auf Anamneseblattt
 GesName As String 'GesName 130 '
 DiagDatum As Date 'DiagDatum 135 '
 DiagSicherheit As String 'DiagSicherheit 130 '6003
 DiagText As String 'DiagText 130 '
 DiagSeite As String 'DiagSeite 130 '6004
 DiagAttr As String 'DiagAttr 130 '6006 Diagnosenattribut (optionale Erläuterung)
 ICD As String 'ICD 130 '
 obDauer As Byte 'obDauer 16 'ob Dauerdiagnose
 intBemerk As String 'intBemerk 130 '6009 interne Bemerkung
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AusnBegr As String 'AusnBegr 130 '6008 Ausnahmebegründung
 f6010 As Byte 'f6010 16 '6010 8.12.10: bisher nur "TM#Falsch"
 f6011 As String 'f6011 130 '6011 8.12.10: bisher nur "TM#?"
End Type

Public Type dokumente
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 DokPfad As String 'DokPfad 130 '
 DokArt As String 'DokArt 130 '
 DokName As String 'DokName 130 '
 Quelldatum As Date 'Quelldatum 135 'Datum, auf das sich das Dokument bezieht
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As Long 'DokGroe 3 'Dokument-Größe
 DokAenD As Date 'DokAenD 135 'Dokument-letzte Änderung
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type eintraege
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '
 Art As String 'Art 130 '6330
 Inhalt As String 'Inhalt 130 '8480
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 ID As Long 'id 3 '
 inhNum As Double 'inhNum 5 'Inhalt numerisch
End Type

Public Type forminhaltform_abk
 Form_AbkVW As Long 'Form_AbkVW 3 '
 Form_Abk As String 'Form_Abk 130 '
End Type

Public Type formulare
 FormID As Long 'FormID 3 '
 Form_Abk As String 'Form_Abk 130 '
 FormBez As String 'FormBez 130 '
 FormVorl As String 'FormVorl 130 '
 AktZeit As Date 'AktZeit 135 'Zeitpunkt der Aktualisierung
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhkopf
 Foid As Long 'FoID 3 '
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '
 Form_ID As Long 'Form_ID 3 '
 Zeitpunkt As Date 'ZeitPunkt 135 '
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Satzart As String 'Satzart 130 '8000
 Satzlänge As String 'Satzlänge 130 '8100
 lanrid As Long 'LANRid 3 'Bezug auf lanrpraxis.id
End Type

Public Type forminhfeld
 Foid As Long 'FoID 3 '
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
 Ziel As String 'Ziel 130 '6291
 Diagnose As String 'Diagnose 130 '6230
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type lbanforderungen
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 AnfText As String 'AnfText 130 '6280
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Statusbyte
End Type

Public Type laborneu
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '
 FertigStGrad As String 'FertigStGrad 130 '8401
 Abkü As String 'Abkü 130 '8410
 LangtextVW As Long 'LangtextVW 3 '8411
 Langtext As String
 Wert As String 'Wert 130 '8420
 Einheit As String 'Einheit 130 '8421
 KommentarVW As Long 'KommentarVW 3 '8480
 Kommentar As String
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 RefNr As Long 'Refnr 3 'Bezug auf LaborXUS
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type leistungen
 ID As Long 'id 3 'eindeutige ID, hinzugefügt 26.3.11
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '5000 + 6201
 Leistung As String 'Leistung 130 '5001 Leistungsziffer
 f5002 As String 'f5002 130 '5002 Art der Untersuchung
 f5005 As String 'f5005 130 '5005 Anzahl
 f5006 As String 'f5006 130 '5006 um Uhrzeit
 f5009 As String 'f5009 130 '5009 freier Begründungstext
 Med As String 'Med 130 '5010 Medikament
 f5015 As String 'f5015 130 '5015 Organ
 f5016 As String 'f5016 130 '5016 Name des Arztes (Briefempfänger)
 f5021 As String 'f5021 130 '5021 Datum letzte Krebsvorsorge
 f5026 As String 'f5026 130 '5026 Entlassungsdatum
 Faktor As String 'Faktor 130 '5062 Multiplikator für GOÄ-Rechnung
 f5098 As String 'f5098 130 '5098 0000000000
 Lanr As String 'LANR 130 '5099 LANR
 letzVorg As Date 'letzVorg 135 '5101 letzter Vorgang
 Ausn As String 'Ausn 130 '3677 Ausnahme/Begründung für abweichendes Geschlecht
 Beme As String 'Beme 130 '         Bemerkung
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 lanrid As Long 'LANRid 3 'Bezug auf lanrpraxis.id
 Sachkbez As String 'Sachkbez 130 '5011 Sachkostenbezeichnung
 Sachkct As Long 'Sachkct 3 '5012 Sach-/Materialkosten in ct
 Zone As String 'Zone 130 '5018 Zone bei Besuchen
End Type

Public Type medplan
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 mpnr As Long 'MPNr 3 'Ordnungsziffer für Medikamentenplan
 Zeitpunkt As Date 'ZeitPunkt 135 'Zeitpunkt, der Speicherung im Turbomed
 Datum As Date 'Datum 135 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament As String 'Medikament 130 '
 MedAnfang As String 'MedAnfang 130 '
 FeldNr As Integer 'FeldNr 2 '
 mo As String 'mo 130 '
 mi As String 'mi 130 '
 nm As String 'nm 130 '
 ab As String 'ab 130 '
 Zn As String 'zn 130 '
 bBed As Integer 'bBed 11 '
 Bemerkung As String 'Bemerkung 130 '
 absPos As Long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type rezepteintraege
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 Rezept As String 'Rezept 130 '6210, 3652(1), 6218(1)
 Rezeptklasse As String 'Rezeptklasse 130 '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Medikament As String 'Medikament 130 '3652(2), 6218(4)
 PZN As String 'PZN 130 '6210(2), 6218(3)
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 QS As String 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte 3 'Statusbyte
 lanrid As Long 'LANRid 3 'Bezug auf lanrpraxis.id
 ID As Long 'id 3 '
End Type

Public Type RR
 FID As Long 'FID 3 'Fall-Bezug
 Pat_id As Long 'Pat_ID 3 '3000
 Zeitpunkt As Date 'ZeitPunkt 135 '6200 + 6201
 RR As String 'RR 130 '6230
 Puls As Long 'Puls 3 'Puls
 Quelle As String 'Quelle 130 'Informationsquelle
 Bemerkung As String 'Bemerkung 130 'Bemerkung
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type kvnrue
 lfdnr As Long 'lfdnr 3 '
 Pat_id As Long 'Pat_ID 3 '
 KVNr As String 'KVNr 130 '
 absPos As Long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit 135 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type unbekannte_kennungen
 Kennung As String 'Kennung 130 '
 absPos As Long 'absPos 3 '
 StByte As Long 'StByte 3 '
 Pat_id As Long 'Pat_id 3 'zugehöriger Patient für spätere Ermittlungen
 Inhalt As String 'Inhalt 130 'Inhalt Zeile zum Wiederauffinden
End Type

Public Type dmpreihe
 Abk As String 'Abk 129 'Abkürzung der DMP-Art
 Art As String 'Art 129 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As Date 'KarteiDatum 133 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As Date 'exportiert 135 'Datum des Exports
 DokuDatum As Date 'DokuDatum 135 'Datum der Dokumentation
 obvoll As Integer 'obvoll 11 'ob vollständig
 Ok As Integer 'ok 11 'ob "ok"
 ausgedruckt As Integer 'ausgedruckt 11 'ob "ausgedruckt"
 Nachname As String 'NachName 129 '
 Vorname As String 'VorName 129 '
 GebDat As Date 'GebDat 133 '
 Pat_id As Long 'Pat_id 3 '
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AktZeit As Date 'AktZeit 135 'Aktualisierungzeit
 lanrid As Long 'lanrid 3 'Bezug auf lanrpraxis.id
End Type

Public Type therarten
 ID As Long 'id 3 '
 Pat_id As Long 'pat_id 3 '
 Zp As Date 'zp 133 '
 mpnr As Long 'mpnr 3 '
 therart As String 'therart 130 '
 absPos As Long 'absPos 3 '
 AktZeit As Date 'AktZeit 135 '
 StByte As Long 'StByte 3 '
End Type

Public Type desktop
 ID As Long 'id 3 'Primärschlüssel
 ids As String 'IDS 130 'id=
 Pat_id As Long 'Pat_ID 3 '3000
 erstZP As Date 'erstZP 135 'erstellungsZeitpunkt
 exoL As String 'exoL 130 'executeonLoad
 hideT As Byte 'hideT 16 'hideTitel
 iconPath As String 'iconPath 130 'iconPath
 noteBkColor As Long 'noteBkColor 3 'noteBkColor
 noteFgColor As Long 'noteFgColor 3 'noteFgColor
 positionBottom As Long 'positionBottom 3 'positionBottom
 positionLeft As Long 'positionLeft 3 'positionLeft
 positionRight As Long 'positionRight 3 'positionRight
 positionTop As Long 'positionTop 3 'positionTop
 showAsNote As Byte 'showAsNote 16 'showAsNote
 syncInfoList As String 'syncInfoList 130 'syncInfoList
 Titel As String 'titel 130 'titel
 toolTipText As String 'toolTipText 130 'toolTipText
 verankert As Byte 'verankert 16 'verankert
 absPos As Long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 StByte As Long 'StByte 3 'Ordnungsnummer der Datenübertragung
End Type

Public Type laborxsaetze
 SatzID As Long 'SatzID 3 'zum Bezug für LaborUS
 DatID As Long 'DatID 3 'Bezug zu LaborEingelesen
 Satzart As String 'Satzart 130 '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 SatzlängeSchluss As String 'SatzlängeSchluss 130 '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb As String 'VersionSatzb 130 '9212 Version der Satzbeschreibung (Turbomed)
 ArztNr As String 'Arztnr 130 '201 Arztnummer (Turbomed)
 Arztname As String 'Arztname 130 '203 Arztname (Turbomed)
 StraßePraxis As String 'StraßePraxis 130 '205 Straße der Praxis (Turbomed)
 Arzt As String 'Arzt 130 ' 211 Ausführender Arzt
 Lanr As String 'LANR 130 ' 212 LANR
 PLZPraxis As String 'PLZPraxis 130 '215 PLZ der Praxis (Turbomed)
 OrtPraxis As String 'OrtPraxis 130 '216 Ort der Praxis (Turbomed)
 Labor As String 'Labor 130 '8320 Labor
 StraßeLabor As String 'StraßeLabor 130 '8321 Straße der Laboradresse (Turbomed)
 PLZLabor As String 'PLZLabor 130 '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor As String 'OrtLabor 130 '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr As String 'KBVPrüfnr 130 '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz As String 'Zeichensatz 130 '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr As String 'Kundenarztnr 130 '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum As String 'Erstellungsdatum 130 '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge As String 'Gesamtlänge 130 '9202 Gesamtlänge des Datenpaketes (Turbomed)
End Type

Public Type laborxeingel
 DatID As Long 'DatID 3 'Bezug auf LaborEingelesen
 Pfad As String 'Pfad 130 'Pfadname
 Name As String 'Name 130 'Name der eingelesenen Labordatei ohne Endung
 Zp As Date 'Zp 135 'Einlesezeitpunkt
 Fertig As Integer 'fertig 11 'ob Einlesen fertig
End Type

Public Type laborxus
 RefNr As Long 'RefNr 3 'Bezug auf LaborWert
 DatID As Long 'DatID 3 'Bezug auf LaborEingelesen
 SatzID As Long 'SatzID 3 'Bezug auf LaborXSätze
 Satzart As String 'Satzart 130 '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 Auftragsnummer As String 'Auftragsnummer 130 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As String 'Auftragsschlüssel 130 '8311 Anforderungsnr d Labors (Turbomed)
 Eingang As Date 'Eingang 135 '8301 Eingangsdatum in Datumsform
 Berichtsdatum As String 'Berichtsdatum 130 '8302 Berichtsdatum
 Pat_id As Long 'Pat_id 3 '
 Nachname As String 'Nachname 130 '3101
 Vorname As String 'Vorname 130 '3102
 GebDat As String 'GebDat 130 '3103
 Titel As String 'Titel 130 '3104
 NVorsatz As String 'NVorsatz 130 '3100
 BefArt As String 'BefArt 130 '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Abrechnungstyp As String 'Abrechnungstyp 130 '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd As String 'GebüOrd 130 '8403 Gebührenordnung (Turbomed)
 Auftraggeber As String 'Auftraggeber 130 '8615 Auftraggeber (LANR)
 Patienteninformation As String 'Patienteninformation 130 '8405 Patienteninformation (Turbomed)
 Geschlecht As String 'Geschlecht 130 '8407 Geschlecht (Turbomed)
 AuftrHinw As String 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp As String 'Pat_idUrsp 130 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG As String 'Pat_idErwVNG 130 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN As String 'Pat_idErwVN 130 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG As String 'Pat_idErwG 130 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB As String 'Pat_idErwGB 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL As String 'Pat_idErwGL 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu As String 'Pat_idLaborNeu 130 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu As Date 'ZeitpunktLaborneu 135 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP As Integer 'ZdüP 2 'Zahl der verglichenen Parameter
 ZdiP As Long 'ZdiP 3 'Zahl der infragekommenden Patienten
 LWerte As String 'LWerte 130 'Laborwerte, die zur Zuordnung geführt haben
 verglichen As Date 'verglichen 135 'Datum, zu dem Datensatz zuletzt verglichen wurde
 AfN As Integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
End Type

Public Type laborxbakt
 RefNr As Long 'RefNr 3 '
 Verf As String 'Verf 130 '
 KuQu As String 'KuQu 130 '8428 Probenmaterial-Ident (Turbomed)
 Quelle As String 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As String 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Kommentar As String 'Kommentar 130 '8480 Ergebnistest (Turbomed)
 Erklärung As String 'Erklärung 130 '
 Keimzahl As String 'Keimzahl 130 '
End Type

Public Type laborxwert
 RefNr As Long 'RefNr 3 'Bezug auf LaborUS
 Abkü As String 'Abkü 130 '8410 Test-Ident  (Turbomed)
 Langname As String 'Langname 130 '8411 Testbezeichnung (Turbomed)
 Quelle As String 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As String 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Wert As String 'Wert 130 '8420 Ergebniswert (Turbomed)
 Einheit As String 'Einheit 130 '8421 Einheit (Turbomed)
 Grenzwerti As String 'Grenzwerti 130 '8422 Grenzwertindikator (Turbomed)
 Kommentar As String 'Kommentar 130 '8480 Ergebnistext (Turbomed)
 Teststatus As String 'Teststatus 130 '8418 Teststatus (Turbomed)
 Erklärung As String 'Erklärung 130 '8470 Testbezogene Hinweise (Turbomed)
 AuftrHinw As String 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 nbid As Long 'nbid 3 'Bezug zu laborxplab.id
End Type

Public Type laborxleist
 RefNr As Long 'RefNr 3 'Bezug auf LaborUS
 Abkü As String 'Abkü 130 '8410 Test-Ident (Turbomed)
 Verf As String 'Verf 130 '8434
 EBM As String 'EBM 130 '5001 GNR (Turbomed)
 goä As String 'goä 130 '8406
 Anzahl As String 'Anzahl 130 '5005
 abrd As String 'abrd 130 '8614 Abrechnung durch: 1 = Labor, 2 = Einweiser
End Type

Public Type Anamnesebogen
 Prim As Long 'Prim 3 'Primärschlüssel
 Pat_id As Long 'Pat_id 3 '
 Nachname As String 'Nachname 130 '-
 Vorname As String 'Vorname 130 '
 NVorsatz As String 'NVorsatz 130 '
 Titel As String 'Titel 130 '
 Anrede As String 'Anrede 130 '
 GebDat As Date 'GebDat 135 ', geb.
 Tkz As Byte 'Tkz 16 'Tod-Kennzeichen
 Versicherungsart As String 'Versicherungsart 130 '
 Diabetestyp As String 'Diabetestyp 130 '^Diabetes Typ
 Diabetes_seit As String 'Diabetes_seit 130 '<seit
 Tabletten_seit As String 'Tabletten_seit 130 ', Tabletten seit
 Insulin_seit As String 'Insulin_seit 130 ', Insulin seit
 Grund_für_Vorstellung As String 'Grund_für_Vorstellung 130 '^:
 Familienanamnese As String 'Familienanamnese 130 '^:
 Größe As Double 'Größe 5 '^:
 Gewicht As Double 'Gewicht 5 ',:
 bmi As Double 'bmi 131 '
 Tendenz As String 'Tendenz 130 '<, Tendenz
 DiabetesMedikament_1 As String 'DiabetesMedikament_1 130 '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge As String 'DiabetesMedikament_1_Menge 130 '<
 DiabetesMedikament_2 As String 'DiabetesMedikament_2 130 '<,
 DiabetesMedikament_2_Menge As String 'DiabetesMedikament_2_Menge 130 '<
 DiabetesMedikament_3 As String 'DiabetesMedikament_3 130 '<,
 DiabetesMedikament_3_Menge As String 'DiabetesMedikament_3_Menge 130 '<
 DiabetesMedikament_4 As String 'DiabetesMedikament_4 130 '<,
 DiabetesMedikament_4_Menge As String 'DiabetesMedikament_4_Menge 130 '<,
 Insulinpumpe As Byte 'Insulinpumpe 16 '^:
 Insulinpumpe_seit As String 'Insulinpumpe_seit 130 '<seit
 Insulinpumpe_Marke As String 'Insulinpumpe_Marke 130 '<, Marke:
 Broteinheiten_gesamt As String 'Broteinheiten_gesamt 130 '^Broteinheiten:gesamt
 Broteinheiten_früh As String 'Broteinheiten_früh 130 '<, früh
 Broteinheiten_ZM_früh As String 'Broteinheiten_ZM_früh 130 '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags As String 'Broteinheiten_mittags 130 '<, mittags
 Broteinheiten_nachmittags As String 'Broteinheiten_nachmittags 130 '<, nachmittags
 Broteinheiten_abends As String 'Broteinheiten_abends 130 '<, abends
 Broteinheiten_nachts As String 'Broteinheiten_nachts 130 '<, nachts
 Essenszeit_früh As String 'Essenszeit_früh 130 '^Essenszeiten:früh
 Essenszeit_vormittags As String 'Essenszeit_vormittags 130 '<, vormittags
 Essenszeit_mittags As String 'Essenszeit_mittags 130 '<, mittags
 Essenszeit_nachmittags As String 'Essenszeit_nachmittags 130 '<, nachmittags
 Essenszeit_abends As String 'Essenszeit_abends 130 '<, abends
 Essenszeit_spät As String 'Essenszeit_spät 130 '<, spät
 Spritz_Eß_Abstand_früh As String 'Spritz_Eß_Abstand_früh 130 '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags As String 'Spritz_Eß_Abstand_mittags 130 '<, mittags
 Spritz_Eß_Abstand_abends As String 'Spritz_Eß_Abstand_abends 130 '<, abends
 Spritzstelle_früh As String 'Spritzstelle_früh 130 '^Spritzstellen:früh
 Spritzstelle_mittags As String 'Spritzstelle_mittags 130 '<, mittags
 Spritzstelle_abends As String 'Spritzstelle_abends 130 '<, abends
 Spritzstelle_nachts As String 'Spritzstelle_nachts 130 '<, nachts
 Jahr_letzte_Diabetesschulung As String 'Jahr_letzte_Diabetesschulung 130 '^Letzte Diabetesschulung:
 Ort_Schulung As String 'Ort_Schulung 130 '<in
 letztes_HbA1c As String 'letztes_HbA1c 130 '^Letztes HbA1c:
 gemessen_am As Date 'gemessen_am 135 '<, gemessen
 vorherige_Werte As String 'vorherige_Werte 130 '<, vorher:
 BZMessungen_selbst As String 'BZMessungen_selbst 130 '^Blutzuckermessung:Selbstmessung?
 Gerät As String 'Gerät 130 '<:
 BZMessungen_pW As String 'BZMessungen_pW 130 '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE As String 'BZMessungen_pW_ndE 130 '<, davon nach dem Essen:
 BZMessungen_p_W_nachts As String 'BZMessungen_p_W_nachts 130 '<, nachts:
 Aufschreiben As String 'Aufschreiben 130 '<, Dokumentation:
 BZWerte_v_d_Essen As String 'BZWerte_v_d_Essen 130 '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen As String 'BZWerte_n_d_Essen 130 '<, nach dem Essen:
 UZ_Tageszeit As String 'UZ_Tageszeit 130 '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM As String 'Unterzucker_pM 130 '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig As String 'UZ_rechtzeitig 130 '<, rechtzeitig bemerkt:
 Fremde_Hilfe_pa As String 'Fremde_Hilfe_pa 130 '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa As String 'Bewußtlos_pa 130 '<, bewußtlos deshalb:
 Keto_pa As String 'Keto_pa 130 '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM As String 'BZgr300_pM 130 ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck As String 'Bluthochdruck 130 '^Bluthochdruck:
 BHD_seit As String 'BHD_seit 130 '<seit:
 BHD_beh_mit As String 'BHD_beh_mit 130 '<, behandelt mit:
 Blutdruckwerte As String 'Blutdruckwerte 130 '^Blutdruckwerte:
 BDselbst As String 'BDselbst 130 '^Blutdruckselbstmessung:
 Schwanger As String 'Schwanger 130 '^Aktuelle Schwangerschaft:
 Schwanger_seit As String 'Schwanger_seit 130 '<, seit:
 Augensp_zuletzt As String 'Augensp_zuletzt 130 '^Letzte Augenspiegelung:
 Augensp_Befund As String 'Augensp_Befund 130 '<, Befund:
 Netzhaut_gelasert As String 'Netzhaut_gelasert 130 ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar As String 'Sehminderung_unbehebbar 130 ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden As String 'Diabet_Nierenschaden 130 '^Diabetischer Nierenschaden:
 Albumin_zuletzt As String 'Albumin_zuletzt 130 ', letztes Albumin:
 erhöht As String 'erhöht 130 '<, Befund:
 Dialyse As Byte 'Dialyse 16 ',:
 Dialyse_seit As String 'Dialyse_seit 130 '<seit
 andere_Nierenerkrankung As String 'andere_Nierenerkrankung 130 ', andere Nierenerkrankung:
 Herzkrankheit As String 'Herzkrankheit 130 '^Herzkrankheit:
 Angina_pectoris As String 'Angina_pectoris 130 ',:
 Herzinfarkt As String 'Herzinfarkt 130 ',:
 Herzinfarkt_wann As String 'Herzinfarkt_wann 130 '<, wann:
 PTCA_oder_Stent As String 'PTCA_oder_Stent 130 ',:
 Bypass_kardial As Byte 'Bypass_kardial 16 ',:
 Bypass_wann As String 'Bypass_wann 130 '<, wann:
 Herzschwäche As String 'Herzschwäche 130 ',:
 Herzkrankheit_Beschreibung As String 'Herzkrankheit_Beschreibung 130 ', Beschreibung:
 Hirndurchblutungsstörung As String 'Hirndurchblutungsstörung 130 '^:
 Schlaganfall As String 'Schlaganfall 130 ',:
 Beindurchblutungsstörung As String 'Beindurchblutungsstörung 130 '^:
 Schaufensterkrankheit As String 'Schaufensterkrankheit 130 ',:
 Bypaß_peripher As Byte 'Bypaß_peripher 16 ',:
 Geschwür As String 'Geschwür 130 ',:
 Amputation As String 'Amputation 130 ',:
 pAVK_Beschreibung As String 'pAVK_Beschreibung 130 ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen As String 'Ameisenlaufen 130 '^:
 Ameisen_Ausmaß As String 'Ameisen_Ausmaß 130 '<, Ausmaß:
 Druckstellen As String 'Druckstellen 130 ',:
 Verformungen As String 'Verformungen 130 ',:
 Verformungen_Beschreibung As String 'Verformungen_Beschreibung 130 '<Beschreibung:
 Fußpflege As String 'Fußpflege 130 '^:
 Podologie As String 'Podologie 130 ',:
 Einlagen As String 'Einlagen 130 ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen As String 'Neue_Fußkomplikationen 130 '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen As String 'Entleerungsstörungen_Magen 130 '^:
 Entleerungsstörungen_Harnblase As String 'Entleerungsstörungen_Harnblase 130 ',:
 Schwindel_Aufstehen As String 'Schwindel_Aufstehen 130 ',:
 Folgeerkrankungen_Haut As String 'Folgeerkrankungen_Haut 130 '^:
 Bewegungseinschränkungen As String 'Bewegungseinschränkungen 130 ',:
 Sexualstörung As String 'Sexualstörung 130 '^:
 Sexualstörung_seit As String 'Sexualstörung_seit 130 '<seit
 Weitere_Anamnese As String 'Weitere_Anamnese 130 '^:
 Alkohol As String 'Alkohol 130 '^:
 Tabak As String 'Tabak 130 ',:
 tabakex As String 'tabakex 130 '
 tabakbis As String 'tabakbis 130 '
 tabakakt As String 'tabakakt 130 '
 tabakmenge As String 'tabakmenge 130 '
 Weitere_Medikation As String 'Weitere_Medikation 130 '^:
 Liphypertrophien_Abdomen As String 'Liphypertrophien_Abdomen 130 '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine As String 'Liphypertrophien_Beine 130 '<, Beine:
 Liphypertrophien_Arme As String 'Liphypertrophien_Arme 130 '<, Arme:
 Beinbefund As String 'Beinbefund 130 '^:
 Hyperkeratosen As String 'Hyperkeratosen 130 ',:
 Ulcera As String 'Ulcera 130 ',:
 Kraft_Zehenheber As String 'Kraft_Zehenheber 130 '^Kraft:Zehenheber
 Kraft_Zehenbeuger As String 'Kraft_Zehenbeuger 130 '<, Zehenbeuger:
 Kraft_Knie As String 'Kraft_Knie 130 '<, Knie:
 ASR As String 'ASR 130 ',:
 PSR As String 'PSR 130 ',:
 Oberflächensensibilität As String 'Oberflächensensibilität 130 '^:
 Monofilamenttest As String 'Monofilamenttest 130 ',:
 Kalt_Warm As String 'Kalt_Warm 130 ', Kalt-Warm-Diskrimination:
 Vibration_IK As String 'Vibration_IK 130 ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe As String 'Vibration_Großzehe 130 '<, Großzehe:
 Puls_Leiste As String 'Puls_Leiste 130 '^Pulse:Leiste
 Puls_Kniekehle As String 'Puls_Kniekehle 130 '<,Kniekehle:
 Puls_Atp As String 'Puls_Atp 130 '<,Innenknöchel:
 Puls_Adp As String 'Puls_Adp 130 '<,Fußrücken:
 RR As String 'RR 130 '^Blutdruck:
 RRTurboMed As String 'RRTurboMed 130 '
 Herz As String 'Herz 130 '^:
 Lunge As String 'Lunge 130 ',:
 Bauch As String 'Bauch 130 ', Abdomen:
 WS As String 'WS 130 ', Wirbelsäule:
 NL As String 'NL 130 ', Nierenlager:
 SD As String 'SD 130 ', Schilddrüse:
 Carotiden As String 'Carotiden 130 ', Halsschlagadern:
 NNH As String 'NNH 130 ', Nasennebenhöhlen:
 Zähne As String 'Zähne 130 ',:
 Mundhöhle As String 'Mundhöhle 130 ',:
 LK As String 'LK 130 ', Lymphknoten:
 BeinödVen As String 'BeinödVen 130 ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst As String 'Neuro_sonst 130 '^Sonstige neurologische Befunde:
 Weitere_Befunde As String 'Weitere_Befunde 130 ', weitere Befunde:
 Schulung As String 'Schulung 130 'ob Schulungsbedarf
 DMP As String 'DMP 130 'ob Pat. bei HA im DMP
 DMSchulz As Integer 'DMSchulz 2 'Zahl der DMP-Schulungen hier
 DMSchL As Integer 'DMSchL 2 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz As Integer 'RRSchulz 2 'Zahl der Hypertonie-Schulungen hier
 DMPhier As Date 'DMPhier 135 'ob Pat hier im DMP
 HANr As String 'HANr 130 'mit "/"
 HANr2 As String 'HANr2 130 'mit "/"
 letzte_Änderung As Date 'letzte_Änderung 135 'Datum der letzten Änderung
 Diagnosen As String 'Diagnosen 130 '
 Vorgestellt As Date 'Vorgestellt 135 'Erstvorstellung
 Versicherung As String 'Versicherung 130 '
 AktZeit As Date 'AktZeit 135 'Aktualisierungszeit
 Ther1 As String 'Ther1 130 'Diät, OAD, CT, Komb, ICT, CSII
 TherAkt As String 'TherAkt 130 'Diät, OAD, CT, Komb, ICT, CSII
 obAn1eing As Byte 'obAn1eing 16 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As Byte 'obAn2eing 16 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As Byte 'obAnAeing 16 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck As Byte 'obCheck 16 'ob Checkliste vorliegt
 obBZausgew As Byte 'obBZausgew 16 'ob Blutzuckergerät ausgewechselt
 obOSaufgek As Byte 'obOSaufgek 16 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As Byte 'obPodAufgek 16 'ob über Podologie aufgeklärt
 obMBlAusgeh As Byte 'obMBlAusgeh 16 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek As String 'obSchulaufgek 130 'ob über Podologie aufgeklärt
 obDMPaufgekl As String 'obDMPaufgekl 130 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz As Byte 'obMedNetz 16 'ob von Med. Netz geschickt
 Hausarzt As String 'Hausarzt 130 'Hausarzt laut Anamnesebogen
 ob As Byte 'ob 16 'für verschiedene Aktionen
 QS As String 'QS 130 'Quartal sortiert von vorgestellt
 QT As String 'QT 130 'Quartal sortiert von vorgestellt
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
Public rTh() As therarten
Public rDe() As desktop
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
 ReDim rTh(0)
 ReDim rDe(0)
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
 Dim rs As New ADODB.Recordset
 Set rs = DBCn.Execute("select count(0) As ct from `" & Tbl & "`")
 frm.Ausgeb "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)", True
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doEntleer

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doEntleer(frm, "desktop")
 Call doEntleer(frm, "therarten")
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
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
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
 Call doEntleer(frm, "laborxplab")
 Call doEntleer(frm, "laborxpneu")
 Call doEntleer(frm, "laborxpnb")
 Call ForeignYes0
 Call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' LabLösch

Function doBezFeh(csqlVal$, ErrDes$)
 Call ForeignNo0
 Call ForeignNo1
 obFor = True
 If lies.dlg.BeziehungsfehlerSpeichern <> 0 Then
  Open BezFeh For Append As #299
  Print #299, vbCrLf & vbCrLf & Now() & ": " & csqlVal
  Print #299, vbCrLf & ErrDes
  Close #299
 End If
End Function 'doBezFeh

Public Function namenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere namen"
 If Not AllePat Then
   sql = "delete from `namen` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Postfach,Weggeldzone," & _
     "WeggzZahl,AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax," & _
     "DienstTel,PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit," & _
     "absPos,StByte,Cave,Notiz,f3800,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass," & _
     "dmpcopdbeg,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache," & _
     "lAktTM) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Postfach,Weggeldzone," & _
     "WeggzZahl,AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax," & _
     "DienstTel,PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit," & _
     "absPos,StByte,Cave,Notiz,f3800,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass," & _
     "dmpcopdbeg,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache," & _
     "lAktTM)  values"))
 For i = 0 To UBound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rNa(i).Pat_id, ",", rNa(i).lfdnr, ",'", rNa(i).NVorsatz, "','", rNa(i).Nachname, "','", rNa(i).Vorname, "',", DatFormK( _
   rNa(i).GebDat), ",'", rNa(i).Straße, "','", rNa(i).KVKStatus, "','", rNa(i).Geschlecht, "','", rNa(i).Plz, "','", rNa(i).Ort, "','", _
   rNa(i).Postfach, "','", rNa(i).Weggeldzone, "',", rNa(i).WeggzZahl, ",", DatFormK(rNa(i).AufnDat), ",'", rNa(i).Lanr, "','", rNa(i).BStNr, "','", _
   rNa(i).Titel, "','", rNa(i).Versichertennummer, "','", rNa(i).PrivatTel, "','", rNa(i).KVNr, "','", rNa(i).PrivatTel_2, "','", _
   rNa(i).PrivatFax, "','", rNa(i).DienstTel, "','", rNa(i).PrivatMobil, "','", rNa(i).Email, "','", rNa(i).Arbeitgeber, "',", CStr(-( _
   rNa(i).AnAllgda <> 0)), ",", CStr(-(rNa(i).An1da <> 0)), ",", CStr(-(rNa(i).An2da <> 0)), ",", CStr(-(rNa(i).Checkda <> 0)), ",'", _
   rNa(i).DMTypaD, "',", DatFormK(0), ",", rNa(i).absPos, ",", rNa(i).StByte, ",'", rNa(i).Cave, "','", rNa(i).Notiz, "','", rNa(i).f3800, "',", _
   rNa(i).dmpklass, ",", DatFormK(rNa(i).dmpbeg), ",", rNa(i).dmpkhkklass, ",", DatFormK(rNa(i).dmpkhkbeg), ",", rNa(i).dmpcopdklass, ",", DatFormK( _
   rNa(i).dmpcopdbeg), ",", rNa(i).getHA0, ",'", rNa(i).fnHA0, "',", rNa(i).getHA1, ",'", rNa(i).fnHA1, "',", _
   rNa(i).getHA2, ",'", rNa(i).fnHA2, "','", rNa(i).zubenach, "','", rNa(i).Verwandt, "','", rNa(i).Sprache, "',", DatFormK(rNa(i).lAktTM), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rNa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(31)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rNa), i)
  If Len(rNa(k).NVorsatz) > maxi(0) Then maxi(0) = Len(rNa(k).NVorsatz)
  If Len(rNa(k).Nachname) > maxi(1) Then maxi(1) = Len(rNa(k).Nachname)
  If Len(rNa(k).Vorname) > maxi(2) Then maxi(2) = Len(rNa(k).Vorname)
  If Len(rNa(k).Straße) > maxi(3) Then maxi(3) = Len(rNa(k).Straße)
  If Len(rNa(k).KVKStatus) > maxi(4) Then maxi(4) = Len(rNa(k).KVKStatus)
  If Len(rNa(k).Geschlecht) > maxi(5) Then maxi(5) = Len(rNa(k).Geschlecht)
  If Len(rNa(k).Plz) > maxi(6) Then maxi(6) = Len(rNa(k).Plz)
  If Len(rNa(k).Ort) > maxi(7) Then maxi(7) = Len(rNa(k).Ort)
  If Len(rNa(k).Postfach) > maxi(8) Then maxi(8) = Len(rNa(k).Postfach)
  If Len(rNa(k).Weggeldzone) > maxi(9) Then maxi(9) = Len(rNa(k).Weggeldzone)
  If Len(rNa(k).Lanr) > maxi(10) Then maxi(10) = Len(rNa(k).Lanr)
  If Len(rNa(k).BStNr) > maxi(11) Then maxi(11) = Len(rNa(k).BStNr)
  If Len(rNa(k).Titel) > maxi(12) Then maxi(12) = Len(rNa(k).Titel)
  If Len(rNa(k).Versichertennummer) > maxi(13) Then maxi(13) = Len(rNa(k).Versichertennummer)
  If Len(rNa(k).PrivatTel) > maxi(14) Then maxi(14) = Len(rNa(k).PrivatTel)
  If Len(rNa(k).KVNr) > maxi(15) Then maxi(15) = Len(rNa(k).KVNr)
  If Len(rNa(k).PrivatTel_2) > maxi(16) Then maxi(16) = Len(rNa(k).PrivatTel_2)
  If Len(rNa(k).PrivatFax) > maxi(17) Then maxi(17) = Len(rNa(k).PrivatFax)
  If Len(rNa(k).DienstTel) > maxi(18) Then maxi(18) = Len(rNa(k).DienstTel)
  If Len(rNa(k).PrivatMobil) > maxi(19) Then maxi(19) = Len(rNa(k).PrivatMobil)
  If Len(rNa(k).Email) > maxi(20) Then maxi(20) = Len(rNa(k).Email)
  If Len(rNa(k).Arbeitgeber) > maxi(21) Then maxi(21) = Len(rNa(k).Arbeitgeber)
  If Len(rNa(k).DMTypaD) > maxi(22) Then maxi(22) = Len(rNa(k).DMTypaD)
  If Len(rNa(k).Cave) > maxi(23) Then maxi(23) = Len(rNa(k).Cave)
  If Len(rNa(k).Notiz) > maxi(24) Then maxi(24) = Len(rNa(k).Notiz)
  If Len(rNa(k).f3800) > maxi(25) Then maxi(25) = Len(rNa(k).f3800)
  If Len(rNa(k).fnHA0) > maxi(26) Then maxi(26) = Len(rNa(k).fnHA0)
  If Len(rNa(k).fnHA1) > maxi(27) Then maxi(27) = Len(rNa(k).fnHA1)
  If Len(rNa(k).fnHA2) > maxi(28) Then maxi(28) = Len(rNa(k).fnHA2)
  If Len(rNa(k).zubenach) > maxi(29) Then maxi(29) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(30) Then maxi(30) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(31) Then maxi(31) = Len(rNa(k).Sprache)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rNa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & Left$(rNa(k).NVorsatz, maxL) & "'", True: rNa(k).NVorsatz = Left$(rNa(k).NVorsatz, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & Left$(rNa(k).Nachname, maxL) & "'", True: rNa(k).Nachname = Left$(rNa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & Left$(rNa(k).Vorname, maxL) & "'", True: rNa(k).Vorname = Left$(rNa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & Left$(rNa(k).Straße, maxL) & "'", True: rNa(k).Straße = Left$(rNa(k).Straße, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & Left$(rNa(k).KVKStatus, maxL) & "'", True: rNa(k).KVKStatus = Left$(rNa(k).KVKStatus, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).Geschlecht & "' -> '" & Left$(rNa(k).Geschlecht, maxL) & "'", True: rNa(k).Geschlecht = Left$(rNa(k).Geschlecht, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).Plz & "' -> '" & Left$(rNa(k).Plz, maxL) & "'", True: rNa(k).Plz = Left$(rNa(k).Plz, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).Ort & "' -> '" & Left$(rNa(k).Ort, maxL) & "'", True: rNa(k).Ort = Left$(rNa(k).Ort, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach: '" & rNa(k).Postfach & "' -> '" & Left$(rNa(k).Postfach, maxL) & "'", True: rNa(k).Postfach = Left$(rNa(k).Postfach, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & Left$(rNa(k).Weggeldzone, maxL) & "'", True: rNa(k).Weggeldzone = Left$(rNa(k).Weggeldzone, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).Lanr & "' -> '" & Left$(rNa(k).Lanr, maxL) & "'", True: rNa(k).Lanr = Left$(rNa(k).Lanr, maxL)
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
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3800: '" & rNa(k).f3800 & "' -> '" & Left$(rNa(k).f3800, maxL) & "'", True: rNa(k).f3800 = Left$(rNa(k).f3800, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & Left$(rNa(k).fnHA0, maxL) & "'", True: rNa(k).fnHA0 = Left$(rNa(k).fnHA0, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & Left$(rNa(k).fnHA1, maxL) & "'", True: rNa(k).fnHA1 = Left$(rNa(k).fnHA1, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & Left$(rNa(k).fnHA2, maxL) & "'", True: rNa(k).fnHA2 = Left$(rNa(k).fnHA2, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL) & "'", True: rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL) & "'", True: rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL) & "'", True: rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere faelle"
 If Not AllePat Then
  sql = "select pat_id from `faelle` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `faelle` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,Kasse,GebOr,AbrGb,PersKreis,SKtZusatz," & _
     "f4206,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR," & _
     "ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw," & _
     "üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,PGeb," & _
     "PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt," & _
     "BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT," & _
     "TherArt,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB," & _
     "AbrAr,lVorl,IK,KVKs,KVKserg,Kasse,GebOr,AbrGb,PersKreis,SKtZusatz," & _
     "f4206,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR," & _
     "ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw," & _
     "üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,PGeb," & _
     "PGebErg,Mahnfrist,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt," & _
     "BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT," & _
     "TherArt,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht)            values"))
 For i = 1 To UBound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFa(i).Pat_id, ",'", rFa(i).Quartal, "','", rFa(i).Nachname, "','", rFa(i).Vorname, "',", rFa(i).lfdnr, ",'", rFa(i).TMFNr, "','", _
   rFa(i).VKNr, "',", DatFormK(rFa(i).BhFB), ",", DatFormK(rFa(i).BhFE1), ",", DatFormK(rFa(i).BhFE2), ",'", rFa(i).f4202, "',", DatFormK( _
   rFa(i).ausgst), ",'", rFa(i).KtrAbrB, "','", rFa(i).AbrAr, "',", DatFormK(rFa(i).lVorl), ",'", rFa(i).IK, "','", _
   rFa(i).KVKs, "','", rFa(i).KVKserg, "','", rFa(i).Kasse, "','", rFa(i).GebOr, "','", rFa(i).AbrGb, "','", rFa(i).PersKreis, "','", _
   rFa(i).SKtZusatz, "','", rFa(i).f4206, "','", rFa(i).ÜwText, "',", rFa(i).f4210, ",", rFa(i).AkfHAH, ",", rFa(i).AkfAB0, ",", rFa(i).AkfAK, ",'", _
   rFa(i).statNuller, "','", rFa(i).ÜbwV, "','", rFa(i).ÜbWVLANR, "','", rFa(i).ÜbWVBSNR, "','", rFa(i).ÜbWVKVNR, "','", _
   rFa(i).AndÜw, "','", rFa(i).Übwr, "','", rFa(i).ÜbwLANR, "','", rFa(i).ÜWZiel, "','", rFa(i).ÜWNNr, "','", rFa(i).ÜWNaN, "','", rFa(i).ÜWTit, "','", _
   rFa(i).ÜWVor, "','", rFa(i).ÜWVsw, "',", rFa(i).üwvid, ",'", rFa(i).Auftrag, "','", rFa(i).Verdacht, "','", rFa(i).Befund, "','", _
   rFa(i).statKlasse, "','", rFa(i).f4237, "',", rFa(i).statBehTage, ",", rFa(i).SchGr, ",'", rFa(i).Weiterbeh, "','", _
   rFa(i).PGeb, "','", rFa(i).PGebErg, "','", rFa(i).Mahnfrist, "','", rFa(i).GOÄKatNr, "','", rFa(i).GOÄKatName, "','", rFa(i).abrArzt, "','", _
   rFa(i).privVers, "','", rFa(i).AdNam, "','", rFa(i).AdStr, "','", rFa(i).AdPlz, "','", rFa(i).AdOrt, "',", DatFormK(rFa(i).BhFE), ",'", _
   rFa(i).s8000, "','", rFa(i).s8100, "',", DatFormK(rFa(i).AktZeit), ",", DatFormK(rFa(i).Fanf), ",'", rFa(i).altQuart, "',", DatFormK( _
   rFa(i).QAnf), ",", DatFormK(rFa(i).QEnd), ",'", rFa(i).QS, "','", rFa(i).QT, "',", rFa(i).therart, ",", rFa(i).StByte, ",", _
   rFa(i).absPos, ",", rFa(i).lanrid, ",'", rFa(i).f4108, "','", rFa(i).BGFallNr, "',", rFa(i).lGewicht, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_id & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & DatFormK(rFa(i).BhFB) & " and bhfe1 = " & DatFormK(rFa(i).BhFE1) & " and ausgst = " & DatFormK(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID & " -> " & rs!FID, True
    Dim jjj&
    For jjj = 1 To UBound(rAu)
     If rAu(jjj).FID = rFa(i).FID Then
      rAu(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rBr)
     If rBr(jjj).FID = rFa(i).FID Then
      rBr(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rDi)
     If rDi(jjj).FID = rFa(i).FID Then
      rDi(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rDo)
     If rDo(jjj).FID = rFa(i).FID Then
      rDo(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rEi)
     If rEi(jjj).FID = rFa(i).FID Then
      rEi(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rFr)
     If rFr(jjj).FID = rFa(i).FID Then
      rFr(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rKh)
     If rKh(jjj).FID = rFa(i).FID Then
      rKh(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLb)
     If rLb(jjj).FID = rFa(i).FID Then
      rLb(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLa)
     If rLa(jjj).FID = rFa(i).FID Then
      rLa(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rLe)
     If rLe(jjj).FID = rFa(i).FID Then
      rLe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rMe)
     If rMe(jjj).FID = rFa(i).FID Then
      rMe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rRe)
     If rRe(jjj).FID = rFa(i).FID Then
      rRe(jjj).FID = rs!FID
     End If
    Next jjj
    For jjj = 1 To UBound(rRr)
     If rRr(jjj).FID = rFa(i).FID Then
      rRr(jjj).FID = rs!FID
     End If
    Next jjj
   End If
  End If
  End If
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(55)
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
  If Len(rFa(k).Kasse) > maxi(11) Then maxi(11) = Len(rFa(k).Kasse)
  If Len(rFa(k).GebOr) > maxi(12) Then maxi(12) = Len(rFa(k).GebOr)
  If Len(rFa(k).AbrGb) > maxi(13) Then maxi(13) = Len(rFa(k).AbrGb)
  If Len(rFa(k).PersKreis) > maxi(14) Then maxi(14) = Len(rFa(k).PersKreis)
  If Len(rFa(k).SKtZusatz) > maxi(15) Then maxi(15) = Len(rFa(k).SKtZusatz)
  If Len(rFa(k).f4206) > maxi(16) Then maxi(16) = Len(rFa(k).f4206)
  If Len(rFa(k).ÜwText) > maxi(17) Then maxi(17) = Len(rFa(k).ÜwText)
  If Len(rFa(k).statNuller) > maxi(18) Then maxi(18) = Len(rFa(k).statNuller)
  If Len(rFa(k).ÜbwV) > maxi(19) Then maxi(19) = Len(rFa(k).ÜbwV)
  If Len(rFa(k).ÜbWVLANR) > maxi(20) Then maxi(20) = Len(rFa(k).ÜbWVLANR)
  If Len(rFa(k).ÜbWVBSNR) > maxi(21) Then maxi(21) = Len(rFa(k).ÜbWVBSNR)
  If Len(rFa(k).ÜbWVKVNR) > maxi(22) Then maxi(22) = Len(rFa(k).ÜbWVKVNR)
  If Len(rFa(k).AndÜw) > maxi(23) Then maxi(23) = Len(rFa(k).AndÜw)
  If Len(rFa(k).Übwr) > maxi(24) Then maxi(24) = Len(rFa(k).Übwr)
  If Len(rFa(k).ÜbwLANR) > maxi(25) Then maxi(25) = Len(rFa(k).ÜbwLANR)
  If Len(rFa(k).ÜWZiel) > maxi(26) Then maxi(26) = Len(rFa(k).ÜWZiel)
  If Len(rFa(k).ÜWNNr) > maxi(27) Then maxi(27) = Len(rFa(k).ÜWNNr)
  If Len(rFa(k).ÜWNaN) > maxi(28) Then maxi(28) = Len(rFa(k).ÜWNaN)
  If Len(rFa(k).ÜWTit) > maxi(29) Then maxi(29) = Len(rFa(k).ÜWTit)
  If Len(rFa(k).ÜWVor) > maxi(30) Then maxi(30) = Len(rFa(k).ÜWVor)
  If Len(rFa(k).ÜWVsw) > maxi(31) Then maxi(31) = Len(rFa(k).ÜWVsw)
  If Len(rFa(k).Auftrag) > maxi(32) Then maxi(32) = Len(rFa(k).Auftrag)
  If Len(rFa(k).Verdacht) > maxi(33) Then maxi(33) = Len(rFa(k).Verdacht)
  If Len(rFa(k).Befund) > maxi(34) Then maxi(34) = Len(rFa(k).Befund)
  If Len(rFa(k).statKlasse) > maxi(35) Then maxi(35) = Len(rFa(k).statKlasse)
  If Len(rFa(k).f4237) > maxi(36) Then maxi(36) = Len(rFa(k).f4237)
  If Len(rFa(k).Weiterbeh) > maxi(37) Then maxi(37) = Len(rFa(k).Weiterbeh)
  If Len(rFa(k).PGeb) > maxi(38) Then maxi(38) = Len(rFa(k).PGeb)
  If Len(rFa(k).PGebErg) > maxi(39) Then maxi(39) = Len(rFa(k).PGebErg)
  If Len(rFa(k).Mahnfrist) > maxi(40) Then maxi(40) = Len(rFa(k).Mahnfrist)
  If Len(rFa(k).GOÄKatNr) > maxi(41) Then maxi(41) = Len(rFa(k).GOÄKatNr)
  If Len(rFa(k).GOÄKatName) > maxi(42) Then maxi(42) = Len(rFa(k).GOÄKatName)
  If Len(rFa(k).abrArzt) > maxi(43) Then maxi(43) = Len(rFa(k).abrArzt)
  If Len(rFa(k).privVers) > maxi(44) Then maxi(44) = Len(rFa(k).privVers)
  If Len(rFa(k).AdNam) > maxi(45) Then maxi(45) = Len(rFa(k).AdNam)
  If Len(rFa(k).AdStr) > maxi(46) Then maxi(46) = Len(rFa(k).AdStr)
  If Len(rFa(k).AdPlz) > maxi(47) Then maxi(47) = Len(rFa(k).AdPlz)
  If Len(rFa(k).AdOrt) > maxi(48) Then maxi(48) = Len(rFa(k).AdOrt)
  If Len(rFa(k).s8000) > maxi(49) Then maxi(49) = Len(rFa(k).s8000)
  If Len(rFa(k).s8100) > maxi(50) Then maxi(50) = Len(rFa(k).s8100)
  If Len(rFa(k).altQuart) > maxi(51) Then maxi(51) = Len(rFa(k).altQuart)
  If Len(rFa(k).QS) > maxi(52) Then maxi(52) = Len(rFa(k).QS)
  If Len(rFa(k).QT) > maxi(53) Then maxi(53) = Len(rFa(k).QT)
  If Len(rFa(k).f4108) > maxi(54) Then maxi(54) = Len(rFa(k).f4108)
  If Len(rFa(k).BGFallNr) > maxi(55) Then maxi(55) = Len(rFa(k).BGFallNr)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & Left$(rFa(k).Quartal, maxL) & "'", True: rFa(k).Quartal = Left$(rFa(k).Quartal, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & Left$(rFa(k).Nachname, maxL) & "'", True: rFa(k).Nachname = Left$(rFa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & Left$(rFa(k).Vorname, maxL) & "'", True: rFa(k).Vorname = Left$(rFa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & Left$(rFa(k).TMFNr, maxL) & "'", True: rFa(k).TMFNr = Left$(rFa(k).TMFNr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & Left$(rFa(k).VKNr, maxL) & "'", True: rFa(k).VKNr = Left$(rFa(k).VKNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4202: '" & rFa(k).f4202 & "' -> '" & Left$(rFa(k).f4202, maxL) & "'", True: rFa(k).f4202 = Left$(rFa(k).f4202, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & Left$(rFa(k).KtrAbrB, maxL) & "'", True: rFa(k).KtrAbrB = Left$(rFa(k).KtrAbrB, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & Left$(rFa(k).AbrAr, maxL) & "'", True: rFa(k).AbrAr = Left$(rFa(k).AbrAr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & Left$(rFa(k).IK, maxL) & "'", True: rFa(k).IK = Left$(rFa(k).IK, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & Left$(rFa(k).KVKs, maxL) & "'", True: rFa(k).KVKs = Left$(rFa(k).KVKs, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & Left$(rFa(k).KVKserg, maxL) & "'", True: rFa(k).KVKserg = Left$(rFa(k).KVKserg, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.Kasse: '" & rFa(k).Kasse & "' -> '" & Left$(rFa(k).Kasse, maxL) & "'", True: rFa(k).Kasse = Left$(rFa(k).Kasse, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & Left$(rFa(k).GebOr, maxL) & "'", True: rFa(k).GebOr = Left$(rFa(k).GebOr, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & Left$(rFa(k).AbrGb, maxL) & "'", True: rFa(k).AbrGb = Left$(rFa(k).AbrGb, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & Left$(rFa(k).PersKreis, maxL) & "'", True: rFa(k).PersKreis = Left$(rFa(k).PersKreis, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & Left$(rFa(k).SKtZusatz, maxL) & "'", True: rFa(k).SKtZusatz = Left$(rFa(k).SKtZusatz, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4206: '" & rFa(k).f4206 & "' -> '" & Left$(rFa(k).f4206, maxL) & "'", True: rFa(k).f4206 = Left$(rFa(k).f4206, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & Left$(rFa(k).ÜwText, maxL) & "'", True: rFa(k).ÜwText = Left$(rFa(k).ÜwText, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & Left$(rFa(k).statNuller, maxL) & "'", True: rFa(k).statNuller = Left$(rFa(k).statNuller, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & Left$(rFa(k).ÜbwV, maxL) & "'", True: rFa(k).ÜbwV = Left$(rFa(k).ÜbwV, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVLANR: '" & rFa(k).ÜbWVLANR & "' -> '" & Left$(rFa(k).ÜbWVLANR, maxL) & "'", True: rFa(k).ÜbWVLANR = Left$(rFa(k).ÜbWVLANR, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVBSNR: '" & rFa(k).ÜbWVBSNR & "' -> '" & Left$(rFa(k).ÜbWVBSNR, maxL) & "'", True: rFa(k).ÜbWVBSNR = Left$(rFa(k).ÜbWVBSNR, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVKVNR: '" & rFa(k).ÜbWVKVNR & "' -> '" & Left$(rFa(k).ÜbWVKVNR, maxL) & "'", True: rFa(k).ÜbWVKVNR = Left$(rFa(k).ÜbWVKVNR, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & Left$(rFa(k).AndÜw, maxL) & "'", True: rFa(k).AndÜw = Left$(rFa(k).AndÜw, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übwr: '" & rFa(k).Übwr & "' -> '" & Left$(rFa(k).Übwr, maxL) & "'", True: rFa(k).Übwr = Left$(rFa(k).Übwr, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & Left$(rFa(k).ÜbwLANR, maxL) & "'", True: rFa(k).ÜbwLANR = Left$(rFa(k).ÜbwLANR, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & Left$(rFa(k).ÜWZiel, maxL) & "'", True: rFa(k).ÜWZiel = Left$(rFa(k).ÜWZiel, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & Left$(rFa(k).ÜWNNr, maxL) & "'", True: rFa(k).ÜWNNr = Left$(rFa(k).ÜWNNr, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & Left$(rFa(k).ÜWNaN, maxL) & "'", True: rFa(k).ÜWNaN = Left$(rFa(k).ÜWNaN, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & Left$(rFa(k).ÜWTit, maxL) & "'", True: rFa(k).ÜWTit = Left$(rFa(k).ÜWTit, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & Left$(rFa(k).ÜWVor, maxL) & "'", True: rFa(k).ÜWVor = Left$(rFa(k).ÜWVor, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & Left$(rFa(k).ÜWVsw, maxL) & "'", True: rFa(k).ÜWVsw = Left$(rFa(k).ÜWVsw, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.Auftrag: '" & rFa(k).Auftrag & "' -> '" & Left$(rFa(k).Auftrag, maxL) & "'", True: rFa(k).Auftrag = Left$(rFa(k).Auftrag, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.Verdacht: '" & rFa(k).Verdacht & "' -> '" & Left$(rFa(k).Verdacht, maxL) & "'", True: rFa(k).Verdacht = Left$(rFa(k).Verdacht, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.Befund: '" & rFa(k).Befund & "' -> '" & Left$(rFa(k).Befund, maxL) & "'", True: rFa(k).Befund = Left$(rFa(k).Befund, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & Left$(rFa(k).statKlasse, maxL) & "'", True: rFa(k).statKlasse = Left$(rFa(k).statKlasse, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4237: '" & rFa(k).f4237 & "' -> '" & Left$(rFa(k).f4237, maxL) & "'", True: rFa(k).f4237 = Left$(rFa(k).f4237, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & Left$(rFa(k).Weiterbeh, maxL) & "'", True: rFa(k).Weiterbeh = Left$(rFa(k).Weiterbeh, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & Left$(rFa(k).PGeb, maxL) & "'", True: rFa(k).PGeb = Left$(rFa(k).PGeb, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & Left$(rFa(k).PGebErg, maxL) & "'", True: rFa(k).PGebErg = Left$(rFa(k).PGebErg, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & Left$(rFa(k).Mahnfrist, maxL) & "'", True: rFa(k).Mahnfrist = Left$(rFa(k).Mahnfrist, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & Left$(rFa(k).GOÄKatNr, maxL) & "'", True: rFa(k).GOÄKatNr = Left$(rFa(k).GOÄKatNr, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & Left$(rFa(k).GOÄKatName, maxL) & "'", True: rFa(k).GOÄKatName = Left$(rFa(k).GOÄKatName, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & Left$(rFa(k).abrArzt, maxL) & "'", True: rFa(k).abrArzt = Left$(rFa(k).abrArzt, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & Left$(rFa(k).privVers, maxL) & "'", True: rFa(k).privVers = Left$(rFa(k).privVers, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & Left$(rFa(k).AdNam, maxL) & "'", True: rFa(k).AdNam = Left$(rFa(k).AdNam, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & Left$(rFa(k).AdStr, maxL) & "'", True: rFa(k).AdStr = Left$(rFa(k).AdStr, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & Left$(rFa(k).AdPlz, maxL) & "'", True: rFa(k).AdPlz = Left$(rFa(k).AdPlz, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & Left$(rFa(k).AdOrt, maxL) & "'", True: rFa(k).AdOrt = Left$(rFa(k).AdOrt, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & Left$(rFa(k).s8000, maxL) & "'", True: rFa(k).s8000 = Left$(rFa(k).s8000, maxL)
       Case 50: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & Left$(rFa(k).s8100, maxL) & "'", True: rFa(k).s8100 = Left$(rFa(k).s8100, maxL)
       Case 51: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & Left$(rFa(k).altQuart, maxL) & "'", True: rFa(k).altQuart = Left$(rFa(k).altQuart, maxL)
       Case 52: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & Left$(rFa(k).QS, maxL) & "'", True: rFa(k).QS = Left$(rFa(k).QS, maxL)
       Case 53: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & Left$(rFa(k).QT, maxL) & "'", True: rFa(k).QT = Left$(rFa(k).QT, maxL)
       Case 54: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4108: '" & rFa(k).f4108 & "' -> '" & Left$(rFa(k).f4108, maxL) & "'", True: rFa(k).f4108 = Left$(rFa(k).f4108, maxL)
       Case 55: Lese.Ausgeb "   Verkürze Inhalt von rFa.BGFallNr: '" & rFa(k).BGFallNr & "' -> '" & Left$(rFa(k).BGFallNr, maxL) & "'", True: rFa(k).BGFallNr = Left$(rFa(k).BGFallNr, maxL)
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
 InsKorr DBCn, DBCnS, sqlquer, rAF
 Resume
End If ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere au"
 If Not AllePat Then
  sql = "select pat_id from `au` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `au` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            values"))
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rAu(i).FID, ",", rAu(i).Pat_id, ",", DatFormK(rAu(i).Zeitpunkt), ",'", rAu(i).Beginn, "','", rAu(i).Ende, "','", rAu(i).ICDs, "',", _
   rAu(i).absPos, ",", DatFormK(rAu(i).AktZeit), ",", rAu(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rAu) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rAu), i)
  If Len(rAu(k).Beginn) > maxi(0) Then maxi(0) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(1) Then maxi(1) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(2) Then maxi(2) = Len(rAu(k).ICDs)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rAu), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & Left$(rAu(k).Beginn, maxL) & "'", True: rAu(k).Beginn = Left$(rAu(k).Beginn, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & Left$(rAu(k).Ende, maxL) & "'", True: rAu(k).Ende = Left$(rAu(k).Ende, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & Left$(rAu(k).ICDs, maxL) & "'", True: rAu(k).ICDs = Left$(rAu(k).ICDs, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere briefe"
 If Not AllePat Then
  sql = "select pat_id from `briefe` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `briefe` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,DokAenD,QS,QT,absPos," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,DokAenD,QS,QT,absPos," & _
     "StByte)  values"))
 For i = 1 To UBound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rBr(i).FID, ",", rBr(i).Pat_id, ",", DatFormK(rBr(i).Zeitpunkt), ",'", rBr(i).Pfad, "','", rBr(i).Art, "','", rBr(i).Name, "','", _
   rBr(i).Typ, "',", DatFormK(rBr(i).AktZeit), ",", rBr(i).DokGroe, ",", DatFormK(rBr(i).DokAenD), ",'", rBr(i).QS, "','", _
   rBr(i).QT, "',", rBr(i).absPos, ",", rBr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rBr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
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
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rBr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL) & "'", True: rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).Art & "' -> '" & Left$(rBr(k).Art, maxL) & "'", True: rBr(k).Art = Left$(rBr(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).Name & "' -> '" & Left$(rBr(k).Name, maxL) & "'", True: rBr(k).Name = Left$(rBr(k).Name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & Left$(rBr(k).Typ, maxL) & "'", True: rBr(k).Typ = Left$(rBr(k).Typ, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & Left$(rBr(k).QS, maxL) & "'", True: rBr(k).QS = Left$(rBr(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & Left$(rBr(k).QT, maxL) & "'", True: rBr(k).QT = Left$(rBr(k).QT, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere diagnosen"
 If Not AllePat Then
  sql = "select pat_id from `diagnosen` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `diagnosen` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit," & _
     "StByte,AusnBegr,f6010,f6011) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit," & _
     "StByte,AusnBegr,f6010,f6011)         values"))
 For i = 1 To UBound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDi(i).FID, ",", rDi(i).Pat_id, ",'", rDi(i).GesName, "',", DatFormK(rDi(i).DiagDatum), ",'", rDi(i).DiagSicherheit, "','", _
   rDi(i).DiagText, "','", rDi(i).DiagSeite, "','", rDi(i).DiagAttr, "','", rDi(i).ICD, "',", rDi(i).obDauer, ",'", rDi(i).intBemerk, "',", _
   rDi(i).absPos, ",", DatFormK(rDi(i).AktZeit), ",", rDi(i).StByte, ",'", rDi(i).AusnBegr, "',", rDi(i).f6010, ",'", rDi(i).f6011, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(8)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDi), i)
  If Len(rDi(k).GesName) > maxi(0) Then maxi(0) = Len(rDi(k).GesName)
  If Len(rDi(k).DiagSicherheit) > maxi(1) Then maxi(1) = Len(rDi(k).DiagSicherheit)
  If Len(rDi(k).DiagText) > maxi(2) Then maxi(2) = Len(rDi(k).DiagText)
  If Len(rDi(k).DiagSeite) > maxi(3) Then maxi(3) = Len(rDi(k).DiagSeite)
  If Len(rDi(k).DiagAttr) > maxi(4) Then maxi(4) = Len(rDi(k).DiagAttr)
  If Len(rDi(k).ICD) > maxi(5) Then maxi(5) = Len(rDi(k).ICD)
  If Len(rDi(k).intBemerk) > maxi(6) Then maxi(6) = Len(rDi(k).intBemerk)
  If Len(rDi(k).AusnBegr) > maxi(7) Then maxi(7) = Len(rDi(k).AusnBegr)
  If Len(rDi(k).f6011) > maxi(8) Then maxi(8) = Len(rDi(k).f6011)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDi.GesName: '" & rDi(k).GesName & "' -> '" & Left$(rDi(k).GesName, maxL) & "'", True: rDi(k).GesName = Left$(rDi(k).GesName, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & Left$(rDi(k).DiagSicherheit, maxL) & "'", True: rDi(k).DiagSicherheit = Left$(rDi(k).DiagSicherheit, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & Left$(rDi(k).DiagText, maxL) & "'", True: rDi(k).DiagText = Left$(rDi(k).DiagText, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & Left$(rDi(k).DiagSeite, maxL) & "'", True: rDi(k).DiagSeite = Left$(rDi(k).DiagSeite, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & Left$(rDi(k).DiagAttr, maxL) & "'", True: rDi(k).DiagAttr = Left$(rDi(k).DiagAttr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & Left$(rDi(k).ICD, maxL) & "'", True: rDi(k).ICD = Left$(rDi(k).ICD, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & Left$(rDi(k).intBemerk, maxL) & "'", True: rDi(k).intBemerk = Left$(rDi(k).intBemerk, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.AusnBegr: '" & rDi(k).AusnBegr & "' -> '" & Left$(rDi(k).AusnBegr, maxL) & "'", True: rDi(k).AusnBegr = Left$(rDi(k).AusnBegr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rDi.f6011: '" & rDi(k).f6011 & "' -> '" & Left$(rDi(k).f6011, maxL) & "'", True: rDi(k).f6011 = Left$(rDi(k).f6011, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere dokumente"
 If Not AllePat Then
  sql = "select pat_id from `dokumente` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `dokumente` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte)  values"))
 For i = 1 To UBound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDo(i).FID, ",", rDo(i).Pat_id, ",", DatFormK(rDo(i).Zeitpunkt), ",'", rDo(i).DokPfad, "','", rDo(i).DokArt, "','", _
   rDo(i).DokName, "',", DatFormK(rDo(i).Quelldatum), ",", rDo(i).absPos, ",", DatFormK(rDo(i).AktZeit), ",", rDo(i).DokGroe, ",", DatFormK( _
   rDo(i).DokAenD), ",'", rDo(i).QS, "','", rDo(i).QT, "',", rDo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDo), i)
  If Len(rDo(k).DokPfad) > maxi(0) Then maxi(0) = Len(rDo(k).DokPfad)
  If Len(rDo(k).DokArt) > maxi(1) Then maxi(1) = Len(rDo(k).DokArt)
  If Len(rDo(k).DokName) > maxi(2) Then maxi(2) = Len(rDo(k).DokName)
  If Len(rDo(k).QS) > maxi(3) Then maxi(3) = Len(rDo(k).QS)
  If Len(rDo(k).QT) > maxi(4) Then maxi(4) = Len(rDo(k).QT)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & Left$(rDo(k).DokPfad, maxL) & "'", True: rDo(k).DokPfad = Left$(rDo(k).DokPfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & Left$(rDo(k).DokArt, maxL) & "'", True: rDo(k).DokArt = Left$(rDo(k).DokArt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & Left$(rDo(k).DokName, maxL) & "'", True: rDo(k).DokName = Left$(rDo(k).DokName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & Left$(rDo(k).QS, maxL) & "'", True: rDo(k).QS = Left$(rDo(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & Left$(rDo(k).QT, maxL) & "'", True: rDo(k).QT = Left$(rDo(k).QT, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere eintraege"
 If Not AllePat Then
  sql = "select pat_id from `eintraege` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `eintraege` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum)     values"))
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rEi(i).FID, ",", rEi(i).Pat_id, ",", DatFormK(rEi(i).Zeitpunkt), ",'", rEi(i).Art, "','", rEi(i).Inhalt, "',", rEi(i).absPos, ",", DatFormK( _
   rEi(i).AktZeit), ",'", rEi(i).QS, "','", rEi(i).QT, "',", rEi(i).StByte, ",", rEi(i).inhNum, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rEi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rEi), i)
  If Len(rEi(k).Art) > maxi(0) Then maxi(0) = Len(rEi(k).Art)
  If Len(rEi(k).Inhalt) > maxi(1) Then maxi(1) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(2) Then maxi(2) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(3) Then maxi(3) = Len(rEi(k).QT)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rEi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).Art & "' -> '" & Left$(rEi(k).Art, maxL) & "'", True: rEi(k).Art = Left$(rEi(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & Left$(rEi(k).Inhalt, maxL) & "'", True: rEi(k).Inhalt = Left$(rEi(k).Inhalt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & Left$(rEi(k).QS, maxL) & "'", True: rEi(k).QS = Left$(rEi(k).QS, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & Left$(rEi(k).QT, maxL) & "'", True: rEi(k).QT = Left$(rEi(k).QT, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere forminhaltform_abk"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhaltform_abk` (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhaltform_abk` (Form_Abk)             values"))
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rFi(i).Form_Abk, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFi) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 rFi1 = UBound(rFi)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rFi1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFi), i)
  If Len(rFi(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFi(k).Form_Abk)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rFi1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFi.Form_Abk: '" & rFi(k).Form_Abk & "' -> '" & Left$(rFi(k).Form_Abk, maxL) & "'", True: rFi(k).Form_Abk = Left$(rFi(k).Form_Abk, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere formulare"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      values"))
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFo(i).FormID, ",'", rFo(i).Form_Abk, "','", rFo(i).FormBez, "','", rFo(i).FormVorl, "',", DatFormK(rFo(i).AktZeit), ",", _
   rFo(i).absPos, ",", rFo(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 rFo1 = UBound(rFo)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rFo1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFo), i)
  If Len(rFo(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFo(k).Form_Abk)
  If Len(rFo(k).FormBez) > maxi(1) Then maxi(1) = Len(rFo(k).FormBez)
  If Len(rFo(k).FormVorl) > maxi(2) Then maxi(2) = Len(rFo(k).FormVorl)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rFo1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & Left$(rFo(k).Form_Abk, maxL) & "'", True: rFo(k).Form_Abk = Left$(rFo(k).Form_Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & Left$(rFo(k).FormBez, maxL) & "'", True: rFo(k).FormBez = Left$(rFo(k).FormBez, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & Left$(rFo(k).FormVorl, maxL) & "'", True: rFo(k).FormVorl = Left$(rFo(k).FormVorl, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere forminhkopf"
 If Not AllePat Then
'  sql = "delete FROM `forminhfeld` where foid in (select foid from `forminhkopf` where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "select pat_id from `forminhkopf` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `forminhkopf` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid)              values"))
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFr(i).Foid, ",", rFr(i).FID, ",", rFr(i).Pat_id, ",", rFr(i).Form_ID, ",", DatFormK(rFr(i).Zeitpunkt), ",", rFr(i).absPos, ",", DatFormK( _
   rFr(i).AktZeit), ",", rFr(i).StByte, ",'", rFr(i).Satzart, "','", rFr(i).Satzlänge, "',", rFr(i).lanrid, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFr), i)
  If Len(rFr(k).Satzart) > maxi(0) Then maxi(0) = Len(rFr(k).Satzart)
  If Len(rFr(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rFr(k).Satzlänge)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & Left$(rFr(k).Satzart, maxL) & "'", True: rFr(k).Satzart = Left$(rFr(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & Left$(rFr(k).Satzlänge, maxL) & "'", True: rFr(k).Satzlänge = Left$(rFr(k).Satzlänge, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere forminhfeld"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      values"))
 For i = 1 To UBound(rFm)
'  rFm(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFm(i).Foid, ",", rFm(i).Nr, ",", rFm(i).FeldNr, ",", rFm(i).FeldVW, ",", rFm(i).FeldInhVW, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rFm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rFm) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(-1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rFm), i)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere kheinweis"
 If Not AllePat Then
  sql = "select pat_id from `kheinweis` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `kheinweis` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               values"))
 For i = 1 To UBound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKh(i).FID, ",", rKh(i).Pat_id, ",", DatFormK(rKh(i).Zeitpunkt), ",'", rKh(i).Ziel, "','", rKh(i).Diagnose, "',", rKh(i).absPos, ",", DatFormK( _
   rKh(i).AktZeit), ",", rKh(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKh) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKh), i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKh), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & Left$(rKh(k).Ziel, maxL) & "'", True: rKh(k).Ziel = Left$(rKh(k).Ziel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & Left$(rKh(k).Diagnose, maxL) & "'", True: rKh(k).Diagnose = Left$(rKh(k).Diagnose, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere lbanforderungen"
 If Not AllePat Then
  sql = "select pat_id from `lbanforderungen` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `lbanforderungen` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       values"))
 For i = 1 To UBound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLb(i).FID, ",", rLb(i).Pat_id, ",", DatFormK(rLb(i).Zeitpunkt), ",'", rLb(i).AnfText, "',", rLb(i).absPos, ",", DatFormK( _
   rLb(i).AktZeit), ",", rLb(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLb) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLb), i)
  If Len(rLb(k).AnfText) > maxi(0) Then maxi(0) = Len(rLb(k).AnfText)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLb), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & Left$(rLb(k).AnfText, maxL) & "'", True: rLb(k).AnfText = Left$(rLb(k).AnfText, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborneu"
 If Not AllePat Then
  sql = "select pat_id from `laborneu` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `laborneu` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte)           values"))
 For i = 1 To UBound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLa(i).FID, ",", rLa(i).Pat_id, ",", DatFormK(rLa(i).Zeitpunkt), ",'", rLa(i).FertigStGrad, "','", rLa(i).Abkü, "',", _
   rLa(i).LangtextVW, ",'", rLa(i).Wert, "','", rLa(i).Einheit, "',", rLa(i).KommentarVW, ",", rLa(i).absPos, ",", DatFormK(rLa(i).AktZeit), ",", _
   rLa(i).RefNr, ",", rLa(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLa) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLa), i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & Left$(rLa(k).FertigStGrad, maxL) & "'", True: rLa(k).FertigStGrad = Left$(rLa(k).FertigStGrad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & Left$(rLa(k).Abkü, maxL) & "'", True: rLa(k).Abkü = Left$(rLa(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & Left$(rLa(k).Wert, maxL) & "'", True: rLa(k).Wert = Left$(rLa(k).Wert, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & Left$(rLa(k).Einheit, maxL) & "'", True: rLa(k).Einheit = Left$(rLa(k).Einheit, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere leistungen"
 If Not AllePat Then
  sql = "select pat_id from `leistungen` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `leistungen` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte,LANRid,Sachkbez,Sachkct,Zone) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte,LANRid,Sachkbez,Sachkct,Zone)               values"))
 For i = 1 To UBound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLe(i).FID, ",", rLe(i).Pat_id, ",", DatFormK(rLe(i).Zeitpunkt), ",'", rLe(i).Leistung, "','", rLe(i).f5002, "','", _
   rLe(i).f5005, "','", rLe(i).f5006, "','", rLe(i).f5009, "','", rLe(i).Med, "','", rLe(i).f5015, "','", rLe(i).f5016, "','", rLe(i).f5021, "','", _
   rLe(i).f5026, "','", rLe(i).Faktor, "','", rLe(i).f5098, "','", rLe(i).Lanr, "',", DatFormK(rLe(i).letzVorg), ",'", _
   rLe(i).Ausn, "','", rLe(i).Beme, "',", rLe(i).absPos, ",", DatFormK(rLe(i).AktZeit), ",'", rLe(i).QS, "','", rLe(i).QT, "',", rLe(i).StByte, ",", _
   rLe(i).lanrid, ",'", rLe(i).Sachkbez, "',", rLe(i).Sachkct, ",'", rLe(i).Zone, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(18)
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
  If Len(rLe(k).Lanr) > maxi(12) Then maxi(12) = Len(rLe(k).Lanr)
  If Len(rLe(k).Ausn) > maxi(13) Then maxi(13) = Len(rLe(k).Ausn)
  If Len(rLe(k).Beme) > maxi(14) Then maxi(14) = Len(rLe(k).Beme)
  If Len(rLe(k).QS) > maxi(15) Then maxi(15) = Len(rLe(k).QS)
  If Len(rLe(k).QT) > maxi(16) Then maxi(16) = Len(rLe(k).QT)
  If Len(rLe(k).Sachkbez) > maxi(17) Then maxi(17) = Len(rLe(k).Sachkbez)
  If Len(rLe(k).Zone) > maxi(18) Then maxi(18) = Len(rLe(k).Zone)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLe.Leistung: '" & rLe(k).Leistung & "' -> '" & Left$(rLe(k).Leistung, maxL) & "'", True: rLe(k).Leistung = Left$(rLe(k).Leistung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5002: '" & rLe(k).f5002 & "' -> '" & Left$(rLe(k).f5002, maxL) & "'", True: rLe(k).f5002 = Left$(rLe(k).f5002, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5005: '" & rLe(k).f5005 & "' -> '" & Left$(rLe(k).f5005, maxL) & "'", True: rLe(k).f5005 = Left$(rLe(k).f5005, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5006: '" & rLe(k).f5006 & "' -> '" & Left$(rLe(k).f5006, maxL) & "'", True: rLe(k).f5006 = Left$(rLe(k).f5006, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5009: '" & rLe(k).f5009 & "' -> '" & Left$(rLe(k).f5009, maxL) & "'", True: rLe(k).f5009 = Left$(rLe(k).f5009, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLe.Med: '" & rLe(k).Med & "' -> '" & Left$(rLe(k).Med, maxL) & "'", True: rLe(k).Med = Left$(rLe(k).Med, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5015: '" & rLe(k).f5015 & "' -> '" & Left$(rLe(k).f5015, maxL) & "'", True: rLe(k).f5015 = Left$(rLe(k).f5015, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5016: '" & rLe(k).f5016 & "' -> '" & Left$(rLe(k).f5016, maxL) & "'", True: rLe(k).f5016 = Left$(rLe(k).f5016, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5021: '" & rLe(k).f5021 & "' -> '" & Left$(rLe(k).f5021, maxL) & "'", True: rLe(k).f5021 = Left$(rLe(k).f5021, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5026: '" & rLe(k).f5026 & "' -> '" & Left$(rLe(k).f5026, maxL) & "'", True: rLe(k).f5026 = Left$(rLe(k).f5026, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLe.Faktor: '" & rLe(k).Faktor & "' -> '" & Left$(rLe(k).Faktor, maxL) & "'", True: rLe(k).Faktor = Left$(rLe(k).Faktor, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5098: '" & rLe(k).f5098 & "' -> '" & Left$(rLe(k).f5098, maxL) & "'", True: rLe(k).f5098 = Left$(rLe(k).f5098, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).Lanr & "' -> '" & Left$(rLe(k).Lanr, maxL) & "'", True: rLe(k).Lanr = Left$(rLe(k).Lanr, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL) & "'", True: rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).Beme & "' -> '" & Left$(rLe(k).Beme, maxL) & "'", True: rLe(k).Beme = Left$(rLe(k).Beme, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL) & "'", True: rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL) & "'", True: rLe(k).QT = Left$(rLe(k).QT, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & Left$(rLe(k).Sachkbez, maxL) & "'", True: rLe(k).Sachkbez = Left$(rLe(k).Sachkbez, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & Left$(rLe(k).Zone, maxL) & "'", True: rLe(k).Zone = Left$(rLe(k).Zone, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere medplan"
 If Not AllePat Then
  sql = "select pat_id from `medplan` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `medplan` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
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
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rMe(i).FID, ",", rMe(i).Pat_id, ",", rMe(i).mpnr, ",", DatFormK(rMe(i).Zeitpunkt), ",", DatFormK(rMe(i).Datum), ",'", _
   rMe(i).Medikament, "','", rMe(i).MedAnfang, "',", rMe(i).FeldNr, ",'", rMe(i).mo, "','", rMe(i).mi, "','", rMe(i).nm, "','", rMe(i).ab, "','", _
   rMe(i).Zn, "',", CStr(-(rMe(i).bBed <> 0)), ",'", rMe(i).Bemerkung, "',", rMe(i).absPos, ",", DatFormK(rMe(i).AktZeit), ",", _
   rMe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rMe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
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
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rMe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & Left$(rMe(k).Medikament, maxL) & "'", True: rMe(k).Medikament = Left$(rMe(k).Medikament, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & Left$(rMe(k).MedAnfang, maxL) & "'", True: rMe(k).MedAnfang = Left$(rMe(k).MedAnfang, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & Left$(rMe(k).mo, maxL) & "'", True: rMe(k).mo = Left$(rMe(k).mo, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & Left$(rMe(k).mi, maxL) & "'", True: rMe(k).mi = Left$(rMe(k).mi, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & Left$(rMe(k).nm, maxL) & "'", True: rMe(k).nm = Left$(rMe(k).nm, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & Left$(rMe(k).ab, maxL) & "'", True: rMe(k).ab = Left$(rMe(k).ab, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).Zn & "' -> '" & Left$(rMe(k).Zn, maxL) & "'", True: rMe(k).Zn = Left$(rMe(k).Zn, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & Left$(rMe(k).Bemerkung, maxL) & "'", True: rMe(k).Bemerkung = Left$(rMe(k).Bemerkung, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere rezepteintraege"
 If Not AllePat Then
  sql = "select pat_id from `rezepteintraege` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `rezepteintraege` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte,LANRid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte,LANRid)         values"))
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRe(i).FID, ",", rRe(i).Pat_id, ",", DatFormK(rRe(i).Zeitpunkt), ",'", rRe(i).Rezept, "','", rRe(i).Rezeptklasse, "','", _
   rRe(i).Medikament, "','", rRe(i).PZN, "',", rRe(i).absPos, ",", DatFormK(rRe(i).AktZeit), ",'", rRe(i).QS, "','", rRe(i).QT, "',", _
   rRe(i).StByte, ",", rRe(i).lanrid, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
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
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & Left$(rRe(k).Rezept, maxL) & "'", True: rRe(k).Rezept = Left$(rRe(k).Rezept, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & Left$(rRe(k).Rezeptklasse, maxL) & "'", True: rRe(k).Rezeptklasse = Left$(rRe(k).Rezeptklasse, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & Left$(rRe(k).Medikament, maxL) & "'", True: rRe(k).Medikament = Left$(rRe(k).Medikament, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & Left$(rRe(k).PZN, maxL) & "'", True: rRe(k).PZN = Left$(rRe(k).PZN, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & Left$(rRe(k).QS, maxL) & "'", True: rRe(k).QS = Left$(rRe(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & Left$(rRe(k).QT, maxL) & "'", True: rRe(k).QT = Left$(rRe(k).QT, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere rr"
 If Not AllePat Then
  sql = "select pat_id from `rr` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `rr` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,Quelle,Bemerkung,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,Quelle,Bemerkung,absPos,AktZeit,StByte)    values"))
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRr(i).FID, ",", rRr(i).Pat_id, ",", DatFormK(rRr(i).Zeitpunkt), ",'", rRr(i).RR, "',", rRr(i).Puls, ",'", rRr(i).Quelle, "','", _
   rRr(i).Bemerkung, "',", rRr(i).absPos, ",", DatFormK(rRr(i).AktZeit), ",", rRr(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rRr) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRr), i)
  If Len(rRr(k).RR) > maxi(0) Then maxi(0) = Len(rRr(k).RR)
  If Len(rRr(k).Quelle) > maxi(1) Then maxi(1) = Len(rRr(k).Quelle)
  If Len(rRr(k).Bemerkung) > maxi(2) Then maxi(2) = Len(rRr(k).Bemerkung)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rRr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & Left$(rRr(k).RR, maxL) & "'", True: rRr(k).RR = Left$(rRr(k).RR, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.Quelle: '" & rRr(k).Quelle & "' -> '" & Left$(rRr(k).Quelle, maxL) & "'", True: rRr(k).Quelle = Left$(rRr(k).Quelle, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRr.Bemerkung: '" & rRr(k).Bemerkung & "' -> '" & Left$(rRr(k).Bemerkung, maxL) & "'", True: rRr(k).Bemerkung = Left$(rRr(k).Bemerkung, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere kvnrue"
 If Not AllePat Then
  sql = "select pat_id from `kvnrue` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `kvnrue` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        values"))
 For i = 1 To UBound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKv(i).Pat_id, ",'", rKv(i).KVNr, "',", rKv(i).absPos, ",", DatFormK(rKv(i).AktZeit), ",", rKv(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rKv) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKv), i)
  If Len(rKv(k).KVNr) > maxi(0) Then maxi(0) = Len(rKv(k).KVNr)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rKv), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & Left$(rKv(k).KVNr, maxL) & "'", True: rKv(k).KVNr = Left$(rKv(k).KVNr, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 syscmd 4, Pid & ": Speichere unbekannte_kennungen"
 If Not AllePat Then
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt)         values"))
 For i = rUn1 + 1 To UBound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rUn(i).Kennung, "',", rUn(i).absPos, ",", rUn(i).StByte, ",", rUn(i).Pat_id, ",'", rUn(i).Inhalt, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rUn) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rUn) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 rUn1 = UBound(rUn)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, rUn1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rUn), i)
  If Len(rUn(k).Kennung) > maxi(0) Then maxi(0) = Len(rUn(k).Kennung)
  If Len(rUn(k).Inhalt) > maxi(1) Then maxi(1) = Len(rUn(k).Inhalt)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, rUn1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rUn), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & Left$(rUn(k).Kennung, maxL) & "'", True: rUn(k).Kennung = Left$(rUn(k).Kennung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & Left$(rUn(k).Inhalt, maxL) & "'", True: rUn(k).Inhalt = Left$(rUn(k).Inhalt, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere dmpreihe"
 If Not AllePat Then
  sql = "select pat_id from `dmpreihe` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `dmpreihe` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid)        values"))
 For i = 1 To UBound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rDm(i).Abk, "','", rDm(i).Art, "',", DatFormK(rDm(i).KarteiDatum), ",", DatFormK(rDm(i).exportiert), ",", DatFormK(rDm(i).DokuDatum), ",", CStr(-( _
   rDm(i).obvoll <> 0)), ",", CStr(-(rDm(i).Ok <> 0)), ",", CStr(-(rDm(i).ausgedruckt <> 0)), ",'", rDm(i).Nachname, "','", _
   rDm(i).Vorname, "',", DatFormK(rDm(i).GebDat), ",", rDm(i).Pat_id, ",", rDm(i).StByte, ",", DatFormK(rDm(i).AktZeit), ",", _
   rDm(i).lanrid, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDm) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDm), i)
  If Len(rDm(k).Abk) > maxi(0) Then maxi(0) = Len(rDm(k).Abk)
  If Len(rDm(k).Art) > maxi(1) Then maxi(1) = Len(rDm(k).Art)
  If Len(rDm(k).Nachname) > maxi(2) Then maxi(2) = Len(rDm(k).Nachname)
  If Len(rDm(k).Vorname) > maxi(3) Then maxi(3) = Len(rDm(k).Vorname)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDm), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL) & "'", True: rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).Art & "' -> '" & Left$(rDm(k).Art, maxL) & "'", True: rDm(k).Art = Left$(rDm(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).Nachname & "' -> '" & Left$(rDm(k).Nachname, maxL) & "'", True: rDm(k).Nachname = Left$(rDm(k).Nachname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).Vorname & "' -> '" & Left$(rDm(k).Vorname, maxL) & "'", True: rDm(k).Vorname = Left$(rDm(k).Vorname, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function therartenSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere therarten"
 If Not AllePat Then
  sql = "select pat_id from `therarten` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `therarten` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `therarten` (pat_id,zp,mpnr," & _
     "therart,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `therarten` (pat_id,zp,mpnr," & _
     "therart,absPos,AktZeit,StByte)       values"))
 For i = 1 To UBound(rTh)
'  rTh(i).AktZeit = now()
  rTh(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rTh(i).Pat_id, ",", DatFormK(rTh(i).Zp), ",", rTh(i).mpnr, ",'", rTh(i).therart, "',", rTh(i).absPos, ",", DatFormK( _
   rTh(i).AktZeit), ",", rTh(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rTh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rTh) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rTh), i)
  If Len(rTh(k).therart) > maxi(0) Then maxi(0) = Len(rTh(k).therart)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "therarten", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "therarten", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rTh), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rTh.therart: '" & rTh(k).therart & "' -> '" & Left$(rTh(k).therart, maxL) & "'", True: rTh(k).therart = Left$(rTh(k).therart, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in therartenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' therartenSpeichern

Public Function desktopSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere desktop"
 If Not AllePat Then
  sql = "select pat_id from `desktop` where Pat_ID = " & CStr(rNa(0).Pat_id)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   Set rs = Nothing
   sql = "delete from `desktop` where Pat_ID = " & CStr(rNa(0).Pat_id)
   Call DBCn.Execute(sql)
  End If
 End If ' not allePat
' sql0 = " insert " & sqlignore &  "into `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte)  values"))
 For i = 1 To UBound(rDe)
'  rDe(i).AktZeit = now()
  rDe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rDe(i).ids, "',", rDe(i).Pat_id, ",", DatFormK(rDe(i).erstZP), ",'", rDe(i).exoL, "',", rDe(i).hideT, ",'", rDe(i).iconPath, "',", _
   rDe(i).noteBkColor, ",", rDe(i).noteFgColor, ",", rDe(i).positionBottom, ",", rDe(i).positionLeft, ",", rDe(i).positionRight, ",", _
   rDe(i).positionTop, ",", rDe(i).showAsNote, ",'", rDe(i).syncInfoList, "','", rDe(i).Titel, "','", rDe(i).toolTipText, "',", _
   rDe(i).verankert, ",", rDe(i).absPos, ",", DatFormK(rDe(i).AktZeit), ",", rDe(i).StByte, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rDe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rDe) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDe), i)
  If Len(rDe(k).ids) > maxi(0) Then maxi(0) = Len(rDe(k).ids)
  If Len(rDe(k).exoL) > maxi(1) Then maxi(1) = Len(rDe(k).exoL)
  If Len(rDe(k).iconPath) > maxi(2) Then maxi(2) = Len(rDe(k).iconPath)
  If Len(rDe(k).syncInfoList) > maxi(3) Then maxi(3) = Len(rDe(k).syncInfoList)
  If Len(rDe(k).Titel) > maxi(4) Then maxi(4) = Len(rDe(k).Titel)
  If Len(rDe(k).toolTipText) > maxi(5) Then maxi(5) = Len(rDe(k).toolTipText)
 Next k
 If obTrans <> 0 Then DBCn.CommitTrans: obTrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "desktop", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "desktop", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rDe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDe.IDS: '" & rDe(k).ids & "' -> '" & Left$(rDe(k).ids, maxL) & "'", True: rDe(k).ids = Left$(rDe(k).ids, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDe.exoL: '" & rDe(k).exoL & "' -> '" & Left$(rDe(k).exoL, maxL) & "'", True: rDe(k).exoL = Left$(rDe(k).exoL, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDe.iconPath: '" & rDe(k).iconPath & "' -> '" & Left$(rDe(k).iconPath, maxL) & "'", True: rDe(k).iconPath = Left$(rDe(k).iconPath, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDe.syncInfoList: '" & rDe(k).syncInfoList & "' -> '" & Left$(rDe(k).syncInfoList, maxL) & "'", True: rDe(k).syncInfoList = Left$(rDe(k).syncInfoList, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDe.titel: '" & rDe(k).Titel & "' -> '" & Left$(rDe(k).Titel, maxL) & "'", True: rDe(k).Titel = Left$(rDe(k).Titel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDe.toolTipText: '" & rDe(k).toolTipText & "' -> '" & Left$(rDe(k).toolTipText, maxL) & "'", True: rDe(k).toolTipText = Left$(rDe(k).toolTipText, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in desktopSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' desktopSpeichern

Public Function laborxsaetzeSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             values"))
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLs(i).DatID, ",'", rLs(i).Satzart, "','", rLs(i).Satzlänge, "','", rLs(i).SatzlängeSchluss, "','", rLs(i).VersionSatzb, "','", _
   rLs(i).ArztNr, "','", rLs(i).Arztname, "','", rLs(i).StraßePraxis, "','", rLs(i).Arzt, "','", rLs(i).Lanr, "','", rLs(i).PLZPraxis, "','", _
   rLs(i).OrtPraxis, "','", rLs(i).Labor, "','", rLs(i).StraßeLabor, "','", rLs(i).PLZLabor, "','", rLs(i).OrtLabor, "','", _
   rLs(i).KBVPrüfnr, "','", rLs(i).Zeichensatz, "','", rLs(i).Kundenarztnr, "','", rLs(i).Erstellungsdatum, "','", rLs(i).Gesamtlänge, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLs) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(19)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLs), i)
  If Len(rLs(k).Satzart) > maxi(0) Then maxi(0) = Len(rLs(k).Satzart)
  If Len(rLs(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rLs(k).Satzlänge)
  If Len(rLs(k).SatzlängeSchluss) > maxi(2) Then maxi(2) = Len(rLs(k).SatzlängeSchluss)
  If Len(rLs(k).VersionSatzb) > maxi(3) Then maxi(3) = Len(rLs(k).VersionSatzb)
  If Len(rLs(k).ArztNr) > maxi(4) Then maxi(4) = Len(rLs(k).ArztNr)
  If Len(rLs(k).Arztname) > maxi(5) Then maxi(5) = Len(rLs(k).Arztname)
  If Len(rLs(k).StraßePraxis) > maxi(6) Then maxi(6) = Len(rLs(k).StraßePraxis)
  If Len(rLs(k).Arzt) > maxi(7) Then maxi(7) = Len(rLs(k).Arzt)
  If Len(rLs(k).Lanr) > maxi(8) Then maxi(8) = Len(rLs(k).Lanr)
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 0, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzart: '" & rLs(k).Satzart & "' -> '" & Left$(rLs(k).Satzart, maxL) & "'", True: rLs(k).Satzart = Left$(rLs(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzlänge: '" & rLs(k).Satzlänge & "' -> '" & Left$(rLs(k).Satzlänge, maxL) & "'", True: rLs(k).Satzlänge = Left$(rLs(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLs.SatzlängeSchluss: '" & rLs(k).SatzlängeSchluss & "' -> '" & Left$(rLs(k).SatzlängeSchluss, maxL) & "'", True: rLs(k).SatzlängeSchluss = Left$(rLs(k).SatzlängeSchluss, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLs.VersionSatzb: '" & rLs(k).VersionSatzb & "' -> '" & Left$(rLs(k).VersionSatzb, maxL) & "'", True: rLs(k).VersionSatzb = Left$(rLs(k).VersionSatzb, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztnr: '" & rLs(k).ArztNr & "' -> '" & Left$(rLs(k).ArztNr, maxL) & "'", True: rLs(k).ArztNr = Left$(rLs(k).ArztNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztname: '" & rLs(k).Arztname & "' -> '" & Left$(rLs(k).Arztname, maxL) & "'", True: rLs(k).Arztname = Left$(rLs(k).Arztname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßePraxis: '" & rLs(k).StraßePraxis & "' -> '" & Left$(rLs(k).StraßePraxis, maxL) & "'", True: rLs(k).StraßePraxis = Left$(rLs(k).StraßePraxis, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arzt: '" & rLs(k).Arzt & "' -> '" & Left$(rLs(k).Arzt, maxL) & "'", True: rLs(k).Arzt = Left$(rLs(k).Arzt, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLs.LANR: '" & rLs(k).Lanr & "' -> '" & Left$(rLs(k).Lanr, maxL) & "'", True: rLs(k).Lanr = Left$(rLs(k).Lanr, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZPraxis: '" & rLs(k).PLZPraxis & "' -> '" & Left$(rLs(k).PLZPraxis, maxL) & "'", True: rLs(k).PLZPraxis = Left$(rLs(k).PLZPraxis, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtPraxis: '" & rLs(k).OrtPraxis & "' -> '" & Left$(rLs(k).OrtPraxis, maxL) & "'", True: rLs(k).OrtPraxis = Left$(rLs(k).OrtPraxis, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLs.Labor: '" & rLs(k).Labor & "' -> '" & Left$(rLs(k).Labor, maxL) & "'", True: rLs(k).Labor = Left$(rLs(k).Labor, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßeLabor: '" & rLs(k).StraßeLabor & "' -> '" & Left$(rLs(k).StraßeLabor, maxL) & "'", True: rLs(k).StraßeLabor = Left$(rLs(k).StraßeLabor, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZLabor: '" & rLs(k).PLZLabor & "' -> '" & Left$(rLs(k).PLZLabor, maxL) & "'", True: rLs(k).PLZLabor = Left$(rLs(k).PLZLabor, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtLabor: '" & rLs(k).OrtLabor & "' -> '" & Left$(rLs(k).OrtLabor, maxL) & "'", True: rLs(k).OrtLabor = Left$(rLs(k).OrtLabor, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLs.KBVPrüfnr: '" & rLs(k).KBVPrüfnr & "' -> '" & Left$(rLs(k).KBVPrüfnr, maxL) & "'", True: rLs(k).KBVPrüfnr = Left$(rLs(k).KBVPrüfnr, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLs.Zeichensatz: '" & rLs(k).Zeichensatz & "' -> '" & Left$(rLs(k).Zeichensatz, maxL) & "'", True: rLs(k).Zeichensatz = Left$(rLs(k).Zeichensatz, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLs.Kundenarztnr: '" & rLs(k).Kundenarztnr & "' -> '" & Left$(rLs(k).Kundenarztnr, maxL) & "'", True: rLs(k).Kundenarztnr = Left$(rLs(k).Kundenarztnr, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLs.Erstellungsdatum: '" & rLs(k).Erstellungsdatum & "' -> '" & Left$(rLs(k).Erstellungsdatum, maxL) & "'", True: rLs(k).Erstellungsdatum = Left$(rLs(k).Erstellungsdatum, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLs.Gesamtlänge: '" & rLs(k).Gesamtlänge & "' -> '" & Left$(rLs(k).Gesamtlänge, maxL) & "'", True: rLs(k).Gesamtlänge = Left$(rLs(k).Gesamtlänge, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  values"))
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rLg(i).Pfad, "','", rLg(i).Name, "',", DatFormK(rLg(i).Zp), ",", CStr(-(rLg(i).Fertig <> 0)), ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLg) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLg), i)
  If Len(rLg(k).Pfad) > maxi(0) Then maxi(0) = Len(rLg(k).Pfad)
  If Len(rLg(k).Name) > maxi(1) Then maxi(1) = Len(rLg(k).Name)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLg), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL) & "'", True: rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).Name & "' -> '" & Left$(rLg(k).Name, maxL) & "'", True: rLg(k).Name = Left$(rLg(k).Name, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(j&)
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxus"
' sql0 = " insert " & sqlignore &  "into `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN)     values"))
 For i = j To j
'  rLu(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLu(i).DatID, ",", rLu(i).SatzID, ",'", rLu(i).Satzart, "','", rLu(i).Satzlänge, "','", rLu(i).Auftragsnummer, "','", _
   rLu(i).Auftragsschlüssel, "',", DatFormK(rLu(i).Eingang), ",'", rLu(i).Berichtsdatum, "',", rLu(i).Pat_id, ",'", rLu(i).Nachname, "','", _
   rLu(i).Vorname, "','", rLu(i).GebDat, "','", rLu(i).Titel, "','", rLu(i).NVorsatz, "','", rLu(i).BefArt, "','", rLu(i).Abrechnungstyp, "','", _
   rLu(i).GebüOrd, "','", rLu(i).Auftraggeber, "','", rLu(i).Patienteninformation, "','", rLu(i).Geschlecht, "','", rLu(i).AuftrHinw, "','", _
   rLu(i).Pat_idUrsp, "','", rLu(i).Pat_idErwVNG, "','", rLu(i).Pat_idErwVN, "','", rLu(i).Pat_idErwG, "','", rLu(i).Pat_idErwGB, "','", _
   rLu(i).Pat_idErwGL, "','", rLu(i).Pat_idLaborNeu, "',", DatFormK(rLu(i).ZeitpunktLaborneu), ",", rLu(i).ZdüP, ",", _
   rLu(i).ZdiP, ",'", rLu(i).LWerte, "',", DatFormK(rLu(i).verglichen), ",", rLu(i).AfN, ")")
  If Lese.dlg.SammelInsert <> 0 And i < j Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = j Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(24)
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
  If Len(rLu(k).Auftraggeber) > maxi(13) Then maxi(13) = Len(rLu(k).Auftraggeber)
  If Len(rLu(k).Patienteninformation) > maxi(14) Then maxi(14) = Len(rLu(k).Patienteninformation)
  If Len(rLu(k).Geschlecht) > maxi(15) Then maxi(15) = Len(rLu(k).Geschlecht)
  If Len(rLu(k).AuftrHinw) > maxi(16) Then maxi(16) = Len(rLu(k).AuftrHinw)
  If Len(rLu(k).Pat_idUrsp) > maxi(17) Then maxi(17) = Len(rLu(k).Pat_idUrsp)
  If Len(rLu(k).Pat_idErwVNG) > maxi(18) Then maxi(18) = Len(rLu(k).Pat_idErwVNG)
  If Len(rLu(k).Pat_idErwVN) > maxi(19) Then maxi(19) = Len(rLu(k).Pat_idErwVN)
  If Len(rLu(k).Pat_idErwG) > maxi(20) Then maxi(20) = Len(rLu(k).Pat_idErwG)
  If Len(rLu(k).Pat_idErwGB) > maxi(21) Then maxi(21) = Len(rLu(k).Pat_idErwGB)
  If Len(rLu(k).Pat_idErwGL) > maxi(22) Then maxi(22) = Len(rLu(k).Pat_idErwGL)
  If Len(rLu(k).Pat_idLaborNeu) > maxi(23) Then maxi(23) = Len(rLu(k).Pat_idLaborNeu)
  If Len(rLu(k).LWerte) > maxi(24) Then maxi(24) = Len(rLu(k).LWerte)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, j, i) To IIf(Lese.dlg.SammelInsert <> 0, j, i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzart: '" & rLu(k).Satzart & "' -> '" & Left$(rLu(k).Satzart, maxL) & "'", True: rLu(k).Satzart = Left$(rLu(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzlänge: '" & rLu(k).Satzlänge & "' -> '" & Left$(rLu(k).Satzlänge, maxL) & "'", True: rLu(k).Satzlänge = Left$(rLu(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsnummer: '" & rLu(k).Auftragsnummer & "' -> '" & Left$(rLu(k).Auftragsnummer, maxL) & "'", True: rLu(k).Auftragsnummer = Left$(rLu(k).Auftragsnummer, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsschlüssel: '" & rLu(k).Auftragsschlüssel & "' -> '" & Left$(rLu(k).Auftragsschlüssel, maxL) & "'", True: rLu(k).Auftragsschlüssel = Left$(rLu(k).Auftragsschlüssel, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLu.Berichtsdatum: '" & rLu(k).Berichtsdatum & "' -> '" & Left$(rLu(k).Berichtsdatum, maxL) & "'", True: rLu(k).Berichtsdatum = Left$(rLu(k).Berichtsdatum, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLu.Nachname: '" & rLu(k).Nachname & "' -> '" & Left$(rLu(k).Nachname, maxL) & "'", True: rLu(k).Nachname = Left$(rLu(k).Nachname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLu.Vorname: '" & rLu(k).Vorname & "' -> '" & Left$(rLu(k).Vorname, maxL) & "'", True: rLu(k).Vorname = Left$(rLu(k).Vorname, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebDat: '" & rLu(k).GebDat & "' -> '" & Left$(rLu(k).GebDat, maxL) & "'", True: rLu(k).GebDat = Left$(rLu(k).GebDat, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLu.Titel: '" & rLu(k).Titel & "' -> '" & Left$(rLu(k).Titel, maxL) & "'", True: rLu(k).Titel = Left$(rLu(k).Titel, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLu.NVorsatz: '" & rLu(k).NVorsatz & "' -> '" & Left$(rLu(k).NVorsatz, maxL) & "'", True: rLu(k).NVorsatz = Left$(rLu(k).NVorsatz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLu.BefArt: '" & rLu(k).BefArt & "' -> '" & Left$(rLu(k).BefArt, maxL) & "'", True: rLu(k).BefArt = Left$(rLu(k).BefArt, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLu.Abrechnungstyp: '" & rLu(k).Abrechnungstyp & "' -> '" & Left$(rLu(k).Abrechnungstyp, maxL) & "'", True: rLu(k).Abrechnungstyp = Left$(rLu(k).Abrechnungstyp, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebüOrd: '" & rLu(k).GebüOrd & "' -> '" & Left$(rLu(k).GebüOrd, maxL) & "'", True: rLu(k).GebüOrd = Left$(rLu(k).GebüOrd, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftraggeber: '" & rLu(k).Auftraggeber & "' -> '" & Left$(rLu(k).Auftraggeber, maxL) & "'", True: rLu(k).Auftraggeber = Left$(rLu(k).Auftraggeber, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  values"))
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLo(i).RefNr, ",'", rLo(i).Verf, "','", rLo(i).KuQu, "','", rLo(i).Quelle, "','", rLo(i).QSpez, "',", DatFormK(rLo(i).AbnDat), ",'", _
   rLo(i).Kommentar, "','", rLo(i).Erklärung, "','", rLo(i).Keimzahl, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLo) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & Left$(rLo(k).Verf, maxL) & "'", True: rLo(k).Verf = Left$(rLo(k).Verf, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & Left$(rLo(k).KuQu, maxL) & "'", True: rLo(k).KuQu = Left$(rLo(k).KuQu, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & Left$(rLo(k).Quelle, maxL) & "'", True: rLo(k).Quelle = Left$(rLo(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & Left$(rLo(k).QSpez, maxL) & "'", True: rLo(k).QSpez = Left$(rLo(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & Left$(rLo(k).Kommentar, maxL) & "'", True: rLo(k).Kommentar = Left$(rLo(k).Kommentar, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & Left$(rLo(k).Erklärung, maxL) & "'", True: rLo(k).Erklärung = Left$(rLo(k).Erklärung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & Left$(rLo(k).Keimzahl, maxL) & "'", True: rLo(k).Keimzahl = Left$(rLo(k).Keimzahl, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    values"))
 For i = 1 To UBound(rLw)
'  rLw(i).AktZeit = now()
    Dim j&
    For j = 1 To i - 1
    If rLw(i).RefNr <> rLw(j).RefNr Then GoTo nextj
    If rLw(i).Abkü <> rLw(j).Abkü Then GoTo nextj
    If rLw(i).Langname <> rLw(j).Langname Then GoTo nextj
    If rLw(i).Quelle <> rLw(j).Quelle Then GoTo nextj
    If rLw(i).QSpez <> rLw(j).QSpez Then GoTo nextj
    If rLw(i).AbnDat <> rLw(j).AbnDat Then GoTo nextj
    If rLw(i).Wert <> rLw(j).Wert Then GoTo nextj
    If rLw(i).Einheit <> rLw(j).Einheit Then GoTo nextj
    If rLw(i).Grenzwerti <> rLw(j).Grenzwerti Then GoTo nextj
    If rLw(i).Kommentar <> rLw(j).Kommentar Then GoTo nextj
    If rLw(i).Teststatus <> rLw(j).Teststatus Then GoTo nextj
    If rLw(i).Erklärung <> rLw(j).Erklärung Then GoTo nextj
    If rLw(i).AuftrHinw <> rLw(j).AuftrHinw Then GoTo nextj
    If rLw(i).nbid <> rLw(j).nbid Then GoTo nextj
     GoTo nexti
nextj:
     Dim rsdop As New ADODB.Recordset
     Set rsdop = Nothing
     rsdop.Open "select 0 from `laborxwert` where `RefNr` = " & rLw(i).RefNr & " and `Abkü` = '" & rLw(i).Abkü & "' and `Langname` = '" & rLw(i).Langname & "' and `Quelle` = '" & _
   rLw(i).Quelle & "' and `QSpez` = '" & rLw(i).QSpez & "' and `AbnDat` = " & DatFormK(rLw(i).AbnDat) & " and `Wert` = '" & _
   rLw(i).Wert & "' and `Einheit` = '" & rLw(i).Einheit & "' and `Grenzwerti` = '" & rLw(i).Grenzwerti & "' and `Kommentar` = '" & rLw(i).Kommentar & "' and `Teststatus` = '" & _
   rLw(i).Teststatus & "' and `Erklärung` = '" & rLw(i).Erklärung & "' and `AuftrHinw` = '" & rLw(i).AuftrHinw & "' and `nbid` = " & _
   rLw(i).nbid & "", DBCn, adOpenStatic, adLockReadOnly
     If Not rsdop.EOF Then GoTo nexti
    Next j
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLw(i).RefNr, ",'", rLw(i).Abkü, "','", rLw(i).Langname, "','", rLw(i).Quelle, "','", rLw(i).QSpez, "',", DatFormK( _
   rLw(i).AbnDat), ",'", rLw(i).Wert, "','", rLw(i).Einheit, "','", rLw(i).Grenzwerti, "','", rLw(i).Kommentar, "','", rLw(i).Teststatus, "','", _
   rLw(i).Erklärung, "','", rLw(i).AuftrHinw, "',", rLw(i).nbid, ")")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLw) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
nexti:
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(10)
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
  If Len(rLw(k).AuftrHinw) > maxi(10) Then maxi(10) = Len(rLw(k).AuftrHinw)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLw), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLw.Abkü: '" & rLw(k).Abkü & "' -> '" & Left$(rLw(k).Abkü, maxL) & "'", True: rLw(k).Abkü = Left$(rLw(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLw.Langname: '" & rLw(k).Langname & "' -> '" & Left$(rLw(k).Langname, maxL) & "'", True: rLw(k).Langname = Left$(rLw(k).Langname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLw.Quelle: '" & rLw(k).Quelle & "' -> '" & Left$(rLw(k).Quelle, maxL) & "'", True: rLw(k).Quelle = Left$(rLw(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLw.QSpez: '" & rLw(k).QSpez & "' -> '" & Left$(rLw(k).QSpez, maxL) & "'", True: rLw(k).QSpez = Left$(rLw(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLw.Wert: '" & rLw(k).Wert & "' -> '" & Left$(rLw(k).Wert, maxL) & "'", True: rLw(k).Wert = Left$(rLw(k).Wert, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLw.Einheit: '" & rLw(k).Einheit & "' -> '" & Left$(rLw(k).Einheit, maxL) & "'", True: rLw(k).Einheit = Left$(rLw(k).Einheit, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLw.Grenzwerti: '" & rLw(k).Grenzwerti & "' -> '" & Left$(rLw(k).Grenzwerti, maxL) & "'", True: rLw(k).Grenzwerti = Left$(rLw(k).Grenzwerti, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLw.Kommentar: '" & rLw(k).Kommentar & "' -> '" & Left$(rLw(k).Kommentar, maxL) & "'", True: rLw(k).Kommentar = Left$(rLw(k).Kommentar, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLw.Teststatus: '" & rLw(k).Teststatus & "' -> '" & Left$(rLw(k).Teststatus, maxL) & "'", True: rLw(k).Teststatus = Left$(rLw(k).Teststatus, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLw.Erklärung: '" & rLw(k).Erklärung & "' -> '" & Left$(rLw(k).Erklärung, maxL) & "'", True: rLw(k).Erklärung = Left$(rLw(k).Erklärung, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLw.AuftrHinw: '" & rLw(k).AuftrHinw & "' -> '" & Left$(rLw(k).AuftrHinw, maxL) & "'", True: rLw(k).AuftrHinw = Left$(rLw(k).AuftrHinw, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern()
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, Pid & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   values"))
 For i = 1 To UBound(rLL)
'  rLL(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLL(i).RefNr, ",'", rLL(i).Abkü, "','", rLL(i).Verf, "','", rLL(i).EBM, "','", rLL(i).goä, "','", rLL(i).Anzahl, "','", _
   rLL(i).abrd, "')")
  If Lese.dlg.SammelInsert <> 0 And i < UBound(rLL) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = UBound(rLL) Then
  Call DBCn.Execute(csql.Value, rAF) ', , adAsyncExecute)
  If obFor Then
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
 DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
If Err.Number = -2147467259 And InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLL), i)
  If Len(rLL(k).Abkü) > maxi(0) Then maxi(0) = Len(rLL(k).Abkü)
  If Len(rLL(k).Verf) > maxi(1) Then maxi(1) = Len(rLL(k).Verf)
  If Len(rLL(k).EBM) > maxi(2) Then maxi(2) = Len(rLL(k).EBM)
  If Len(rLL(k).goä) > maxi(3) Then maxi(3) = Len(rLL(k).goä)
  If Len(rLL(k).Anzahl) > maxi(4) Then maxi(4) = Len(rLL(k).Anzahl)
  If Len(rLL(k).abrd) > maxi(5) Then maxi(5) = Len(rLL(k).abrd)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0, 1, i) To IIf(Lese.dlg.SammelInsert <> 0, UBound(rLL), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & Left$(rLL(k).Abkü, maxL) & "'", True: rLL(k).Abkü = Left$(rLL(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & Left$(rLL(k).Verf, maxL) & "'", True: rLL(k).Verf = Left$(rLL(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & Left$(rLL(k).EBM, maxL) & "'", True: rLL(k).EBM = Left$(rLL(k).EBM, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & Left$(rLL(k).goä, maxL) & "'", True: rLL(k).goä = Left$(rLL(k).goä, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & Left$(rLL(k).Anzahl, maxL) & "'", True: rLL(k).Anzahl = Left$(rLL(k).Anzahl, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.abrd: '" & rLL(k).abrd & "' -> '" & Left$(rLL(k).abrd, maxL) & "'", True: rLL(k).abrd = Left$(rLL(k).abrd, maxL)
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
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
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
    If obTrans <> 0 Then Call DBCn.CommitTrans: obTrans = 0
    Call DBCn.BeginTrans
    obTrans = 1
   End If
 Call auSpeichern
 Call briefeSpeichern
 Call diagnosenSpeichern
 Call dokumenteSpeichern
 Call eintraegeSpeichern
 Call forminhaltform_abkSpeichern
   If Not lies.obMySQL Then
    If obTrans <> 0 Then Call DBCn.CommitTrans: obTrans = 0
    Call DBCn.BeginTrans
    obTrans = 1
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
 Call therartenSpeichern
 Call desktopSpeichern
 Call DBCn.Execute("update `namen` set aktZeit = " & DatFormK(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_id, rAF)
 If rAF <> 1 Then
  frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_id & " " & rNa(0).Nachname & " " & rNa(0).Vorname, True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
 ErrDescription = Err.Description
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
  Resume
 End If
 Select Case MsgBox("FNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

