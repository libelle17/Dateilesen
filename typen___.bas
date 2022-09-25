Option Explicit
Public obFor%
Dim sql$, T1!, T2!, rs As ADODB.Recordset, maxL%

Public type namen
 Pat_ID As long 'Pat_ID 3 '3000
 lfdnr As long 'lfdnr 3 'laufende Patientennummer
 NVorsatz As string 'NVorsatz 130 '3100
 Nachname As string 'Nachname 130 '3101
 Vorname As string 'Vorname 130 '3102
 GebDat As date 'GebDat 135 '3103
 Straße As string 'Straße 130 '3107
 KVKStatus As string 'KVKStatus 130 '3108
 Geschlecht As string 'Geschlecht 130 '3110
 Plz As string 'Plz 130 '3112
 Ort As string 'Ort 130 '3113
 Postfach As string 'Postfach 130 '3216
 Weggeldzone As string 'Weggeldzone 130 '3631 (1) Weggeldzone mit Z
 WeggzZahl As double 'WeggzZahl 131 '3631 (2) Weggeldzone, Zahl in Feld 2
 AufnDat As date 'AufnDat 135 '3610
 LANR As string 'LANR 130 '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr As string 'BStNr 130 '3536 Betriebsstättennummer
 Titel As string 'Titel 130 '3104
 Versichertennummer As string 'Versichertennummer 130 '3105
 PrivatTel As string 'PrivatTel 130 '3629
 KVNr As string 'KVNr 130 '3630
 PrivatTel_2 As string 'PrivatTel_2 130 '3629
 PrivatFax As string 'PrivatFax 130 '3629
 DienstTel As string 'DienstTel 130 '3629
 PrivatMobil As string 'PrivatMobil 130 '3629
 Email As string 'Email 130 'Email
 Arbeitgeber As string 'Arbeitgeber 130 '3625
 AnAllgda As integer 'AnAllgda 11 'Anamnese allgemein da
 An1da As integer 'An1da 11 'Anamnese S.1 da
 An2da As integer 'An2da 11 'Anamnese S.2 da
 Checkda As integer 'Checkda 11 'Checkliste da
 DMTypaD As string 'DMTypaD 130 'aus Diagnosen
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Cave As string 'Cave 130 '3654
 Notiz As string 'Notiz 130 '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>
 dmpklass As long 'dmpklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpbeg As date 'dmpbeg 133 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass As long 'dmpkhkklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg As date 'dmpkhkbeg 133 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass As long 'dmpcopdklass 3 '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg As date 'dmpcopdbeg 133 'Datum der aktuellen DMP-Klassifikation
 getHA0 As long 'getHA0 3 'KVNr aus getHausarzt -> Üw(12,0)
 fnHA0 As string 'fnHA0 130 'Funktion aus getHausarzt -> Üw(10,0)
 getHA1 As long 'getHA1 3 'KVNr aus getHausarzt -> Üw(12,1)
 fnHA1 As string 'fnHA1 130 'Funktion aus getHausarzt -> Üw(10,1)
 getHA2 As long 'getHA2 3 'KVNr aus getHausarzt -> Üw(12,2)
 fnHA2 As string 'fnHA2 130 'Funktion aus getHausarzt -> Üw(10,2)
 zubenach As string 'zubenach 130 '3633
 Verwandt As string 'Verwandt 130 '3632
 Sprache As string 'Sprache 130 '3628
 lAktTM As date 'lAktTM 135 'letzte Aktualisierung in Turbomed
end type

Public type faelle
 FID As long 'FID 3 '
 Pat_ID As long 'Pat_ID 3 '3000
 Quartal As string 'Quartal 130 '4101
 Nachname As string 'Nachname 130 '3101
 Vorname As string 'Vorname 130 '3102
 lfdnr As long 'lfdnr 3 'laufende Fallnummer
 TMFNr As string 'TMFNr 130 '4144 Fallnummer in Turbomed
 VKNr As string 'VKNr 130 '4104
 BhFB As date 'BhFB 135 '4150
 BhFE1 As date 'BhFE1 135 '4151
 BhFE2 As date 'BhFE2 135 '4152
 f4202 As string 'f4202 130 '4202
 ausgst As date 'ausgst 135 '4102 ('ausgestellt am')
 KtrAbrB As string 'KtrAbrB 130 '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr As string 'AbrAr 130 '4107, Abrechnungsart (1 = Primärkassen)
 lVorl As date 'lVorl 135 '4109, letzte Vorlage
 IK As string 'IK 130 '4111 Krankenkassennummer (IK)
 KVKs As string 'KVKs 130 '4112 Versichertenstatus VK
 KVKserg As string 'KVKserg 130 '4113 Ost/West-Status VK
 Kasse As string 'Kasse 130 '6299 Kasse (aus Formularen)
 GebOr As string 'GebOr 130 '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As string 'AbrGb 130 '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis As string 'PersKreis 130 '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz As string 'SKtZusatz 130 '4124 SKT-Zusatzangaben
 f4206 As string 'f4206 130 '4206, mutmasslicher Tag der Entbindung
 ÜwText As string 'ÜwText 130 '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 As byte 'f4210 16 '4210, Ankreuzfeld LSR
 AkfHAH As byte 'AkfHAH 16 '4211 Ankreuzfeld HAH
 AkfAB0 As byte 'AkfAB0 16 '4212 Ankreuzfeld AB0.RH
 AkfAK As byte 'AkfAK 16 '4213 Ankreuzfeld AK
 statNuller As string 'statNuller 130 '4216, nu bei Musterfrau 16 Nuller
 ÜbwV As string 'ÜbwV 130 '4218, überwiesen von Arztnummer
 ÜbWVLANR As string 'ÜbWVLANR 130 '4218(1) überwiesen von LANR
 ÜbWVBSNR As string 'ÜbWVBSNR 130 '4218(2) überwiesen von BSNR
 ÜbWVKVNR As string 'ÜbWVKVNR 130 '4218(3) überwiesen von KVNR
 AndÜw As string 'AndÜw 130 '4219, anderer Überweiser
 Übwr As string 'Übwr 130 'resultierender Überweiser (BSNR): 4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR As string 'ÜbwLANR 130 '4242 LANR des Überweisers
 ÜWZiel As string 'ÜWZiel 130 '4220 Überweisung an
 ÜWNNr As string 'ÜWNNr 130 '4231(4): KV-Nummer des Überweisers
 ÜWNaN As string 'ÜWNaN 130 '4231(3): Nachname des Überweisers
 ÜWTit As string 'ÜWTit 130 '4231(3): Titel des Überweisers
 ÜWVor As string 'ÜWVor 130 '4231(2): Vorname des Überweisers
 ÜWVsw As string 'ÜWVsw 130 '4231(2b): Vorsatzwort des Überweisers
 üwvid As long 'üwvid 3 '4247 Bezug auf ueberwvon
 Auftrag As string 'Auftrag 130 '4205 Auftrag bei Überweisung
 Verdacht As string 'Verdacht 130 '4207 Verdacht bei Überweisung

 Befund As string 'Befund 130 '4208 Befund bei Überweisung
 statKlasse As string 'statKlasse 130 '4236 Klasse bei Behandlung
 f4237 As string 'f4237 130 '4237 Krankenhausname
 statBehTage As long 'statBehTage 3 '4238 Krankenhausaufenthalt
 SchGr As double 'SchGr 131 '4239, Schein(unter)gruppe
 Weiterbeh As string 'Weiterbeh 130 '4243, Weiterbehandelnder
 PGeb As string 'PGeb 130 '4401, Praxisgebühr
 PGebErg As string 'PGebErg 130 '4402, ?
 Mahnfrist As string 'Mahnfrist 130 '4403, Mahnfrist bis
 GOÄKatNr As string 'GOÄKatNr 130 '4580 (1): Katalog-Nummer
 GOÄKatName As string 'GOÄKatName 130 '4580 (2): Privat-Abrechnungskatalog
 abrArzt As string 'abrArzt 130 '4585 abrechnender Arzt
 privVers As string 'privVers 130 '4586 private Versicherung
 AdNam As string 'AdNam 130 '4602(1) Name Rechnungsanschrift
 AdStr As string 'AdStr 130 '4602(2) Straße Rechnungsanschrift
 AdPlz As string 'AdPlz 130 '4602(3) PLZ Rechnungsanschrift
 AdOrt As string 'AdOrt 130 '4602(4) Ort Rechnungsanschrift
 BhFE As date 'BhFE 135 '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As string 's8000 130 '8000, Satzidentifikation
 s8100 As string 's8100 130 '8100 Satzlänge
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 Fanf As date 'Fanf 135 'Fallanfang
 altQuart As string 'altQuart 130 '
 QAnf As date 'QAnf 135 'Quartalsanfang
 QEnd As date 'QEnd 135 'Quartalsende
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 TherArt As long 'TherArt 3 'Therapieart: (0 = offen,  1= diät,  2= oad, 3= komb,  4= ct, 5= ict, 6 = csii)
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 LANRid As long 'LANRid 3 'Bezug auf lanrpraxis.id
 f4108 As string 'f4108 130 '4108
 BGFallNr As string 'BGFallNr 130 '3603 BG-Fall-Nummer	
 lGewicht As double 'lGewicht 131 'letztes Gewicht in kg
end type

Public type au
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '6200 + 6201
 Beginn As string 'Beginn 130 '6285 1. Hälfte
 Ende As string 'Ende 130 '6285 2. Hälfte
 ICDs As string 'ICDs 130 '6286
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type briefe
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '
 ZeitPunkt As date 'ZeitPunkt 135 '
 Pfad As string 'Pfad 130 '
 Art As string 'Art 130 '
 Name As string 'Name 130 '
 Typ As string 'Typ 130 '
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As long 'DokGroe 3 'Größe der Datei
 DokAenD As date 'DokAenD 135 'Dokument-letzte Änderung
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type diagnosen
 ID1 As long 'ID1 3 '
 FID As long 'FID 3 'Fall-Bezug
 Pat_id As long 'Pat_id 3 'Bezug auf Anamneseblattt
 GesName As string 'GesName 130 '
 DiagDatum As date 'DiagDatum 135 '
 DiagSicherheit As string 'DiagSicherheit 130 '6003
 DiagText As string 'DiagText 130 '
 DiagSeite As string 'DiagSeite 130 '6004
 DiagAttr As string 'DiagAttr 130 '6006 Diagnosenattribut (optionale Erläuterung)
 ICD As string 'ICD 130 '
 obDauer As byte 'obDauer 16 'ob Dauerdiagnose
 intBemerk As string 'intBemerk 130 '6009 interne Bemerkung
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AusnBegr As string 'AusnBegr 130 '6008 Ausnahmebegründung
 f6010 As byte 'f6010 16 '6010 8.12.10: bisher nur "TM#Falsch"
 f6011 As string 'f6011 130 '6011 8.12.10: bisher nur "TM#?"
end type

Public type dokumente
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '
 ZeitPunkt As date 'ZeitPunkt 135 '
 DokPfad As string 'DokPfad 130 '
 DokArt As string 'DokArt 130 '
 DokName As string 'DokName 130 '
 Quelldatum As date 'Quelldatum 135 'Datum, auf das sich das Dokument bezieht
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 DokGroe As long 'DokGroe 3 'Dokument-Größe
 DokAenD As date 'DokAenD 135 'Dokument-letzte Änderung
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type eintraege
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '
 Art As string 'Art 130 '6330
 Inhalt As string 'Inhalt 130 '8480
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 id As long 'id 3 '
 inhNum As double 'inhNum 5 'Inhalt numerisch
end type

Public type forminhaltform_abk
 Form_AbkVW As long 'Form_AbkVW 3 '
 Form_Abk As string 'Form_Abk 130 '
end type

Public type formulare
 FormID As long 'FormID 3 '
 Form_Abk As string 'Form_Abk 130 '
 FormBez As string 'FormBez 130 '
 FormVorl As string 'FormVorl 130 '
 AktZeit As date 'AktZeit 135 'Zeitpunkt der Aktualisierung
 absPos As long 'absPos 3 'Zeile in BDT-Datei
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type forminhkopf
 FoID As long 'FoID 3 '
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '
 Form_ID As long 'Form_ID 3 '
 ZeitPunkt As date 'ZeitPunkt 135 '
 AbsPos As long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 Satzart As string 'Satzart 130 '8000
 Satzlänge As string 'Satzlänge 130 '8100
 LANRid As long 'LANRid 3 'Bezug auf lanrpraxis.id
end type

Public type forminhfeld
 FoID As long 'FoID 3 '
 Nr As integer 'Nr 2 '
 FeldNr As integer 'FeldNr 2 '
 FeldVW As long 'FeldVW 3 '
 Feld As string
 FeldInhVW As long 'FeldInhVW 3 '
 FeldInh As string
end type

Public type kheinweis
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '
 Ziel As string 'Ziel 130 '6291
 Diagnose As string 'Diagnose 130 '6230
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type lbanforderungen
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '6200 + 6201
 AnfText As string 'AnfText 130 '6280
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Statusbyte
end type

Public type laborneu
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '
 FertigStGrad As string 'FertigStGrad 130 '8401
 Abkü As string 'Abkü 130 '8410
 LangtextVW As long 'LangtextVW 3 '8411
 Langtext As string
 Wert As string 'Wert 130 '8420
 Einheit As string 'Einheit 130 '8421
 KommentarVW As long 'KommentarVW 3 '8480
 Kommentar As string
 AbsPos As long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 Refnr As long 'Refnr 3 'Bezug auf LaborXUS
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type leistungen
 id As long 'id 3 'eindeutige ID, hinzugefügt 26.3.11
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '5000 + 6201
 Leistung As string 'Leistung 130 '5001 Leistungsziffer
 f5002 As string 'f5002 130 '5002 Art der Untersuchung
 f5005 As string 'f5005 130 '5005 Anzahl
 f5006 As string 'f5006 130 '5006 um Uhrzeit
 f5009 As string 'f5009 130 '5009 freier Begründungstext
 Med As string 'Med 130 '5010 Medikament
 f5015 As string 'f5015 130 '5015 Organ
 f5016 As string 'f5016 130 '5016 Name des Arztes (Briefempfänger)
 f5021 As string 'f5021 130 '5021 Datum letzte Krebsvorsorge
 f5026 As string 'f5026 130 '5026 Entlassungsdatum
 Faktor As string 'Faktor 130 '5062 Multiplikator für GOÄ-Rechnung
 f5098 As string 'f5098 130 '5098 0000000000
 LANR As string 'LANR 130 '5099 LANR
 letzVorg As date 'letzVorg 135 '5101 letzter Vorgang
 Ausn As string 'Ausn 130 '3677 Ausnahme/Begründung für abweichendes Geschlecht
 Beme As string 'Beme 130 '         Bemerkung
 absPos As long 'absPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 LANRid As long 'LANRid 3 'Bezug auf lanrpraxis.id
 Sachkbez As string 'Sachkbez 130 '5011 Sachkostenbezeichnung
 Sachkct As long 'Sachkct 3 '5012 Sach-/Materialkosten in ct
 Zone As string 'Zone 130 '5018 Zone bei Besuchen
end type

Public type medplan
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 MPNr As long 'MPNr 3 'Ordnungsziffer für Medikamentenplan
 ZeitPunkt As date 'ZeitPunkt 135 'Zeitpunkt, der Speicherung im Turbomed
 Datum As date 'Datum 135 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament As string 'Medikament 130 '
 MedAnfang As string 'MedAnfang 130 '
 FeldNr As integer 'FeldNr 2 '
 mo As string 'mo 130 '
 mi As string 'mi 130 '
 nm As string 'nm 130 '
 ab As string 'ab 130 '
 zn As string 'zn 130 '
 bBed As integer 'bBed 11 '
 Bemerkung As string 'Bemerkung 130 '
 AbsPos As long 'AbsPos 3 'Zeile in der BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type rezepteintraege
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '6200 + 6201
 Rezept As string 'Rezept 130 '6210, 3652(1), 6218(1)
 Rezeptklasse As string 'Rezeptklasse 130 '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Medikament As string 'Medikament 130 '3652(2), 6218(4)
 PZN As string 'PZN 130 '6210(2), 6218(3)
 absPos As long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 QS As string 'QS 130 'Quartal des Behandlungsfallbeginns sortiert
 QT As string 'QT 130 'Quartal des Behandlungsfallbeginns
 StByte As long 'StByte 3 'Statusbyte
 LANRid As long 'LANRid 3 'Bezug auf lanrpraxis.id
 id As long 'id 3 '
end type

Public type rr
 FID As long 'FID 3 'Fall-Bezug
 Pat_ID As long 'Pat_ID 3 '3000
 ZeitPunkt As date 'ZeitPunkt 135 '6200 + 6201
 RR As string 'RR 130 '6230
 Puls As long 'Puls 3 'Puls
 Quelle As string 'Quelle 130 'Informationsquelle
 Bemerkung As string 'Bemerkung 130 'Bemerkung
 absPos As long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type kvnrue
 lfdnr As long 'lfdnr 3 '
 Pat_ID As long 'Pat_ID 3 '
 KVNr As string 'KVNr 130 '
 absPos As long 'absPos 3 'Zeile in BDT-Datei
 AktZeit As date 'AktZeit 135 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
end type

Public type unbekannte_kennungen
 Kennung As string 'Kennung 130 '
 absPos As long 'absPos 3 '
 StByte As long 'StByte 3 '
 Pat_id As long 'Pat_id 3 'zugehöriger Patient für spätere Ermittlungen
 Inhalt As string 'Inhalt 130 'Inhalt Zeile zum Wiederauffinden
end type

Public type dmpreihe
 Abk As string 'Abk 129 'Abkürzung der DMP-Art
 Art As string 'Art 129 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As date 'KarteiDatum 133 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As date 'exportiert 135 'Datum des Exports
 DokuDatum As date 'DokuDatum 135 'Datum der Dokumentation
 obvoll As integer 'obvoll 11 'ob vollständig
 ok As integer 'ok 11 'ob "ok"
 ausgedruckt As integer 'ausgedruckt 11 'ob "ausgedruckt"
 NachName As string 'NachName 129 '
 VorName As string 'VorName 129 '
 GebDat As date 'GebDat 133 '
 Pat_id As long 'Pat_id 3 '
 StByte As long 'StByte 3 'Ordnungsnummer der Datenübertragung
 AktZeit As date 'AktZeit 135 'Aktualisierungzeit
 lanrid As long 'lanrid 3 'Bezug auf lanrpraxis.id
end type

Public type therarten
 id As long 'id 3 '
 pat_id As long 'pat_id 3 '
 zp As date 'zp 133 '
 mpnr As long 'mpnr 3 '
 therart As string 'therart 130 '
 absPos As long 'absPos 3 '
 AktZeit As date 'AktZeit 135 '
 StByte As long 'StByte 3 '
end type

Public type laborxsaetze
 SatzID As long 'SatzID 3 'zum Bezug für LaborUS
 DatID As long 'DatID 3 'Bezug zu LaborEingelesen
 Satzart As string 'Satzart 130 '8000 Satzart (Turbomed)
 Satzlänge As string 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 SatzlängeSchluss As string 'SatzlängeSchluss 130 '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb As string 'VersionSatzb 130 '9212 Version der Satzbeschreibung (Turbomed)
 Arztnr As string 'Arztnr 130 '201 Arztnummer (Turbomed)
 Arztname As string 'Arztname 130 '203 Arztname (Turbomed)
 StraßePraxis As string 'StraßePraxis 130 '205 Straße der Praxis (Turbomed)
 Arzt As string 'Arzt 130 ' 211 Ausführender Arzt
 LANR As string 'LANR 130 ' 212 LANR
 PLZPraxis As string 'PLZPraxis 130 '215 PLZ der Praxis (Turbomed)
 OrtPraxis As string 'OrtPraxis 130 '216 Ort der Praxis (Turbomed)
 Labor As string 'Labor 130 '8320 Labor
 StraßeLabor As string 'StraßeLabor 130 '8321 Straße der Laboradresse (Turbomed)
 PLZLabor As string 'PLZLabor 130 '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor As string 'OrtLabor 130 '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr As string 'KBVPrüfnr 130 '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz As string 'Zeichensatz 130 '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr As string 'Kundenarztnr 130 '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum As string 'Erstellungsdatum 130 '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge As string 'Gesamtlänge 130 '9202 Gesamtlänge des Datenpaketes (Turbomed)
end type

Public type laborxeingel
 DatID As long 'DatID 3 'Bezug auf LaborEingelesen
 Pfad As string 'Pfad 130 'Pfadname
 Name As string 'Name 130 'Name der eingelesenen Labordatei ohne Endung
 Zp As date 'Zp 135 'Einlesezeitpunkt
 fertig As integer 'fertig 11 'ob Einlesen fertig
end type

Public type laborxus
 RefNr As long 'RefNr 3 'Bezug auf LaborWert
 DatID As long 'DatID 3 'Bezug auf LaborEingelesen
 SatzID As long 'SatzID 3 'Bezug auf LaborXSätze
 Satzart As string 'Satzart 130 '8000 Satzart (Turbomed)
 Satzlänge As string 'Satzlänge 130 '8100 Satzlänge (Turbomed)
 Auftragsnummer As string 'Auftragsnummer 130 '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As string 'Auftragsschlüssel 130 '8311 Anforderungsnr d Labors (Turbomed)
 Eingang As date 'Eingang 135 '8301 Eingangsdatum in Datumsform
 Berichtsdatum As string 'Berichtsdatum 130 '8302 Berichtsdatum
 Pat_id As long 'Pat_id 3 '
 Nachname As string 'Nachname 130 '3101
 Vorname As string 'Vorname 130 '3102
 GebDat As string 'GebDat 130 '3103
 Titel As string 'Titel 130 '3104
 NVorsatz As string 'NVorsatz 130 '3100
 BefArt As string 'BefArt 130 '8401 Befundart (Turbomed) / Fertigstellungsgrad ("E"=Endbefund, "T" = Teilbefund)
 Abrechnungstyp As string 'Abrechnungstyp 130 '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd As string 'GebüOrd 130 '8403 Gebührenordnung (Turbomed)
 Auftraggeber As string 'Auftraggeber 130 '8615 Auftraggeber (LANR)
 Patienteninformation As string 'Patienteninformation 130 '8405 Patienteninformation (Turbomed)
 Geschlecht As string 'Geschlecht 130 '8407 Geschlecht (Turbomed)
 AuftrHinw As string 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp As string 'Pat_idUrsp 130 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG As string 'Pat_idErwVNG 130 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN As string 'Pat_idErwVN 130 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG As string 'Pat_idErwG 130 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB As string 'Pat_idErwGB 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL As string 'Pat_idErwGL 130 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu As string 'Pat_idLaborNeu 130 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu As date 'ZeitpunktLaborneu 135 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP As integer 'ZdüP 2 'Zahl der verglichenen Parameter
 ZdiP As long 'ZdiP 3 'Zahl der infragekommenden Patienten
 LWerte As string 'LWerte 130 'Laborwerte, die zur Zuordnung geführt haben
 verglichen As date 'verglichen 135 'Datum, zu dem Datensatz zuletzt verglichen wurde
 AfN As integer 'AfN 2 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
end type

Public type laborxbakt
 RefNr As long 'RefNr 3 '
 Verf As string 'Verf 130 '
 KuQu As string 'KuQu 130 '8428 Probenmaterial-Ident (Turbomed)
 Quelle As string 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As string 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Kommentar As string 'Kommentar 130 '8480 Ergebnistest (Turbomed)
 Erklärung As string 'Erklärung 130 '
 Keimzahl As string 'Keimzahl 130 '
end type

Public type laborxwert
 RefNr As long 'RefNr 3 'Bezug auf LaborUS
 Abkü As string 'Abkü 130 '8410 Test-Ident  (Turbomed)
 Langname As string 'Langname 130 '8411 Testbezeichnung (Turbomed)
 Quelle As string 'Quelle 130 '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As string 'QSpez 130 '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As date 'AbnDat 135 '8432 Abnahmedatum (Turbomed)
 Wert As string 'Wert 130 '8420 Ergebniswert (Turbomed)
 Einheit As string 'Einheit 130 '8421 Einheit (Turbomed)
 Grenzwerti As string 'Grenzwerti 130 '8422 Grenzwertindikator (Turbomed)
 Kommentar As string 'Kommentar 130 '8480 Ergebnistext (Turbomed)
 Teststatus As string 'Teststatus 130 '8418 Teststatus (Turbomed)
 Erklärung As string 'Erklärung 130 '8470 Testbezogene Hinweise (Turbomed)
 AuftrHinw As string 'AuftrHinw 130 '8490 Auftragsbezogene Hinweise (Turbomed)
 nbid As long 'nbid 3 'Bezug zu laborxplab.id
end type

Public type laborxleist
 RefNr As long 'RefNr 3 'Bezug auf LaborUS
 Abkü As string 'Abkü 130 '8410 Test-Ident (Turbomed)
 Verf As string 'Verf 130 '8434
 EBM As string 'EBM 130 '5001 GNR (Turbomed)
 goä As string 'goä 130 '8406
 Anzahl As string 'Anzahl 130 '5005
 abrd As string 'abrd 130 '8614 Abrechnung durch: 1 = Labor, 2 = Einweiser
end type

Public type anamnesebogen
 Prim As long 'Prim 3 'Primärschlüssel
 Pat_id As long 'Pat_id 3 '
 Nachname As string 'Nachname 130 '-
 Vorname As string 'Vorname 130 '
 NVorsatz As string 'NVorsatz 130 '
 Titel As string 'Titel 130 '
 Anrede As string 'Anrede 130 '
 GebDat As date 'GebDat 135 ', geb.
 Tkz As byte 'Tkz 16 'Tod-Kennzeichen
 Versicherungsart As string 'Versicherungsart 130 '
 Diabetestyp As string 'Diabetestyp 130 '^Diabetes Typ
 Diabetes_seit As string 'Diabetes_seit 130 '<seit
 Tabletten_seit As string 'Tabletten_seit 130 ', Tabletten seit
 Insulin_seit As string 'Insulin_seit 130 ', Insulin seit
 Grund_für_Vorstellung As string 'Grund_für_Vorstellung 130 '^:
 Familienanamnese As string 'Familienanamnese 130 '^:
 Größe As double 'Größe 5 '^:
 Gewicht As double 'Gewicht 5 ',:
 bmi As double 'bmi 131 '
 Tendenz As string 'Tendenz 130 '<, Tendenz
 DiabetesMedikament_1 As string 'DiabetesMedikament_1 130 '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge As string 'DiabetesMedikament_1_Menge 130 '<
 DiabetesMedikament_2 As string 'DiabetesMedikament_2 130 '<,
 DiabetesMedikament_2_Menge As string 'DiabetesMedikament_2_Menge 130 '<
 DiabetesMedikament_3 As string 'DiabetesMedikament_3 130 '<,
 DiabetesMedikament_3_Menge As string 'DiabetesMedikament_3_Menge 130 '<
 DiabetesMedikament_4 As string 'DiabetesMedikament_4 130 '<,
 DiabetesMedikament_4_Menge As string 'DiabetesMedikament_4_Menge 130 '<,
 Insulinpumpe As byte 'Insulinpumpe 16 '^:
 Insulinpumpe_seit As string 'Insulinpumpe_seit 130 '<seit
 Insulinpumpe_Marke As string 'Insulinpumpe_Marke 130 '<, Marke:
 Broteinheiten_gesamt As string 'Broteinheiten_gesamt 130 '^Broteinheiten:gesamt
 Broteinheiten_früh As string 'Broteinheiten_früh 130 '<, früh
 Broteinheiten_ZM_früh As string 'Broteinheiten_ZM_früh 130 '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags As string 'Broteinheiten_mittags 130 '<, mittags
 Broteinheiten_nachmittags As string 'Broteinheiten_nachmittags 130 '<, nachmittags
 Broteinheiten_abends As string 'Broteinheiten_abends 130 '<, abends
 Broteinheiten_nachts As string 'Broteinheiten_nachts 130 '<, nachts
 Essenszeit_früh As string 'Essenszeit_früh 130 '^Essenszeiten:früh
 Essenszeit_vormittags As string 'Essenszeit_vormittags 130 '<, vormittags
 Essenszeit_mittags As string 'Essenszeit_mittags 130 '<, mittags
 Essenszeit_nachmittags As string 'Essenszeit_nachmittags 130 '<, nachmittags
 Essenszeit_abends As string 'Essenszeit_abends 130 '<, abends
 Essenszeit_spät As string 'Essenszeit_spät 130 '<, spät
 Spritz_Eß_Abstand_früh As string 'Spritz_Eß_Abstand_früh 130 '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags As string 'Spritz_Eß_Abstand_mittags 130 '<, mittags
 Spritz_Eß_Abstand_abends As string 'Spritz_Eß_Abstand_abends 130 '<, abends
 Spritzstelle_früh As string 'Spritzstelle_früh 130 '^Spritzstellen:früh
 Spritzstelle_mittags As string 'Spritzstelle_mittags 130 '<, mittags
 Spritzstelle_abends As string 'Spritzstelle_abends 130 '<, abends
 Spritzstelle_nachts As string 'Spritzstelle_nachts 130 '<, nachts
 Jahr_letzte_Diabetesschulung As string 'Jahr_letzte_Diabetesschulung 130 '^Letzte Diabetesschulung:
 Ort_Schulung As string 'Ort_Schulung 130 '<in
 letztes_HbA1c As string 'letztes_HbA1c 130 '^Letztes HbA1c:
 gemessen_am As date 'gemessen_am 135 '<, gemessen
 vorherige_Werte As string 'vorherige_Werte 130 '<, vorher:
 BZMessungen_selbst As string 'BZMessungen_selbst 130 '^Blutzuckermessung:Selbstmessung?
 Gerät As string 'Gerät 130 '<:
 BZMessungen_pW As string 'BZMessungen_pW 130 '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE As string 'BZMessungen_pW_ndE 130 '<, davon nach dem Essen:
 BZMessungen_p_W_nachts As string 'BZMessungen_p_W_nachts 130 '<, nachts:
 Aufschreiben As string 'Aufschreiben 130 '<, Dokumentation:
 BZWerte_v_d_Essen As string 'BZWerte_v_d_Essen 130 '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen As string 'BZWerte_n_d_Essen 130 '<, nach dem Essen:
 UZ_Tageszeit As string 'UZ_Tageszeit 130 '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM As string 'Unterzucker_pM 130 '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig As string 'UZ_rechtzeitig 130 '<, rechtzeitig bemerkt:
 Fremde_Hilfe_pa As string 'Fremde_Hilfe_pa 130 '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa As string 'Bewußtlos_pa 130 '<, bewußtlos deshalb:
 Keto_pa As string 'Keto_pa 130 '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM As string 'BZgr300_pM 130 ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck As string 'Bluthochdruck 130 '^Bluthochdruck:
 BHD_seit As string 'BHD_seit 130 '<seit:
 BHD_beh_mit As string 'BHD_beh_mit 130 '<, behandelt mit:
 Blutdruckwerte As string 'Blutdruckwerte 130 '^Blutdruckwerte:
 BDselbst As string 'BDselbst 130 '^Blutdruckselbstmessung:
 Schwanger As string 'Schwanger 130 '^Aktuelle Schwangerschaft:
 Schwanger_seit As string 'Schwanger_seit 130 '<, seit:
 Augensp_zuletzt As string 'Augensp_zuletzt 130 '^Letzte Augenspiegelung:
 Augensp_Befund As string 'Augensp_Befund 130 '<, Befund:
 Netzhaut_gelasert As string 'Netzhaut_gelasert 130 ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar As string 'Sehminderung_unbehebbar 130 ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden As string 'Diabet_Nierenschaden 130 '^Diabetischer Nierenschaden:
 Albumin_zuletzt As string 'Albumin_zuletzt 130 ', letztes Albumin:
 erhöht As string 'erhöht 130 '<, Befund:
 Dialyse As byte 'Dialyse 16 ',:
 Dialyse_seit As string 'Dialyse_seit 130 '<seit
 andere_Nierenerkrankung As string 'andere_Nierenerkrankung 130 ', andere Nierenerkrankung:
 Herzkrankheit As string 'Herzkrankheit 130 '^Herzkrankheit:
 Angina_pectoris As string 'Angina_pectoris 130 ',:
 Herzinfarkt As string 'Herzinfarkt 130 ',:
 Herzinfarkt_wann As string 'Herzinfarkt_wann 130 '<, wann:
 PTCA_oder_Stent As string 'PTCA_oder_Stent 130 ',:
 Bypass_kardial As byte 'Bypass_kardial 16 ',:
 Bypass_wann As string 'Bypass_wann 130 '<, wann:
 Herzschwäche As string 'Herzschwäche 130 ',:
 Herzkrankheit_Beschreibung As string 'Herzkrankheit_Beschreibung 130 ', Beschreibung:
 Hirndurchblutungsstörung As string 'Hirndurchblutungsstörung 130 '^:
 Schlaganfall As string 'Schlaganfall 130 ',:
 Beindurchblutungsstörung As string 'Beindurchblutungsstörung 130 '^:
 Schaufensterkrankheit As string 'Schaufensterkrankheit 130 ',:
 Bypaß_peripher As byte 'Bypaß_peripher 16 ',:
 Geschwür As string 'Geschwür 130 ',:
 Amputation As string 'Amputation 130 ',:
 pAVK_Beschreibung As string 'pAVK_Beschreibung 130 ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen As string 'Ameisenlaufen 130 '^:
 Ameisen_Ausmaß As string 'Ameisen_Ausmaß 130 '<, Ausmaß:
 Druckstellen As string 'Druckstellen 130 ',:
 Verformungen As string 'Verformungen 130 ',:
 Verformungen_Beschreibung As string 'Verformungen_Beschreibung 130 '<Beschreibung:
 Fußpflege As string 'Fußpflege 130 '^:
 Podologie As string 'Podologie 130 ',:
 Einlagen As string 'Einlagen 130 ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen As string 'Neue_Fußkomplikationen 130 '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen As string 'Entleerungsstörungen_Magen 130 '^:
 Entleerungsstörungen_Harnblase As string 'Entleerungsstörungen_Harnblase 130 ',:
 Schwindel_Aufstehen As string 'Schwindel_Aufstehen 130 ',:
 Folgeerkrankungen_Haut As string 'Folgeerkrankungen_Haut 130 '^:
 Bewegungseinschränkungen As string 'Bewegungseinschränkungen 130 ',:
 Sexualstörung As string 'Sexualstörung 130 '^:
 Sexualstörung_seit As string 'Sexualstörung_seit 130 '<seit
 Weitere_Anamnese As string 'Weitere_Anamnese 130 '^:
 Alkohol As string 'Alkohol 130 '^:
 Tabak As string 'Tabak 130 ',:
 tabakex As string 'tabakex 130 '
 tabakbis As string 'tabakbis 130 '
 tabakakt As string 'tabakakt 130 '
 tabakmenge As string 'tabakmenge 130 '
 Weitere_Medikation As string 'Weitere_Medikation 130 '^:
 Liphypertrophien_Abdomen As string 'Liphypertrophien_Abdomen 130 '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine As string 'Liphypertrophien_Beine 130 '<, Beine:
 Liphypertrophien_Arme As string 'Liphypertrophien_Arme 130 '<, Arme:
 Beinbefund As string 'Beinbefund 130 '^:
 Hyperkeratosen As string 'Hyperkeratosen 130 ',:
 Ulcera As string 'Ulcera 130 ',:
 Kraft_Zehenheber As string 'Kraft_Zehenheber 130 '^Kraft:Zehenheber
 Kraft_Zehenbeuger As string 'Kraft_Zehenbeuger 130 '<, Zehenbeuger:
 Kraft_Knie As string 'Kraft_Knie 130 '<, Knie:
 ASR As string 'ASR 130 ',:
 PSR As string 'PSR 130 ',:
 Oberflächensensibilität As string 'Oberflächensensibilität 130 '^:
 Monofilamenttest As string 'Monofilamenttest 130 ',:
 Kalt_Warm As string 'Kalt_Warm 130 ', Kalt-Warm-Diskrimination:
 Vibration_IK As string 'Vibration_IK 130 ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe As string 'Vibration_Großzehe 130 '<, Großzehe:
 Puls_Leiste As string 'Puls_Leiste 130 '^Pulse:Leiste
 Puls_Kniekehle As string 'Puls_Kniekehle 130 '<,Kniekehle:
 Puls_Atp As string 'Puls_Atp 130 '<,Innenknöchel:
 Puls_Adp As string 'Puls_Adp 130 '<,Fußrücken:
 RR As string 'RR 130 '^Blutdruck:
 RRTurboMed As string 'RRTurboMed 130 '
 Herz As string 'Herz 130 '^:
 Lunge As string 'Lunge 130 ',:
 Bauch As string 'Bauch 130 ', Abdomen:
 WS As string 'WS 130 ', Wirbelsäule:
 NL As string 'NL 130 ', Nierenlager:
 SD As string 'SD 130 ', Schilddrüse:
 Carotiden As string 'Carotiden 130 ', Halsschlagadern:
 NNH As string 'NNH 130 ', Nasennebenhöhlen:
 Zähne As string 'Zähne 130 ',:
 Mundhöhle As string 'Mundhöhle 130 ',:
 LK As string 'LK 130 ', Lymphknoten:
 BeinödVen As string 'BeinödVen 130 ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst As string 'Neuro_sonst 130 '^Sonstige neurologische Befunde:
 Weitere_Befunde As string 'Weitere_Befunde 130 ', weitere Befunde:
 Schulung As string 'Schulung 130 'ob Schulungsbedarf
 DMP As string 'DMP 130 'ob Pat. bei HA im DMP
 DMSchulz As integer 'DMSchulz 2 'Zahl der DMP-Schulungen hier
 DMSchL As integer 'DMSchL 2 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz As integer 'RRSchulz 2 'Zahl der Hypertonie-Schulungen hier
 DMPhier As date 'DMPhier 135 'ob Pat hier im DMP
 HANr As string 'HANr 130 'mit "/"
 HANr2 As string 'HANr2 130 'mit "/"
 letzte_Änderung As date 'letzte_Änderung 135 'Datum der letzten Änderung
 Diagnosen As string 'Diagnosen 130 '
 Vorgestellt As date 'Vorgestellt 135 'Erstvorstellung
 Versicherung As string 'Versicherung 130 '
 AktZeit As date 'AktZeit 135 'Aktualisierungszeit
 Ther1 As string 'Ther1 130 'Diät, OAD, CT, Komb, ICT, CSII
 TherAkt As string 'TherAkt 130 'Diät, OAD, CT, Komb, ICT, CSII
 obAn1eing As byte 'obAn1eing 16 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As byte 'obAn2eing 16 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As byte 'obAnAeing 16 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck As byte 'obCheck 16 'ob Checkliste vorliegt
 obBZausgew As byte 'obBZausgew 16 'ob Blutzuckergerät ausgewechselt
 obOSaufgek As byte 'obOSaufgek 16 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As byte 'obPodAufgek 16 'ob über Podologie aufgeklärt
 obMBlAusgeh As byte 'obMBlAusgeh 16 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek As string 'obSchulaufgek 130 'ob über Podologie aufgeklärt
 obDMPaufgekl As string 'obDMPaufgekl 130 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz As byte 'obMedNetz 16 'ob von Med. Netz geschickt
 Hausarzt As string 'Hausarzt 130 'Hausarzt laut Anamnesebogen
 ob As byte 'ob 16 'für verschiedene Aktionen
 QS As string 'QS 130 'Quartal sortiert von vorgestellt
 QT As string 'QT 130 'Quartal sortiert von vorgestellt
end type

Public rNa() As namen
Public rFa() As faelle
Public rAu() As au
Public rBr() As briefe
Public rDi() As diagnosen
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
Public rRr() As rr
Public rKv() As kvnrue
Public rUn() As unbekannte_kennungen
Public rDm() As dmpreihe
Public rTh() As therarten
Public rLs() As laborxsaetze
Public rLg() As laborxeingel
Public rLu() As laborxus
Public rLo() As laborxbakt
Public rLw() As laborxwert
Public rLL() As laborxleist
Public rAna() As anamnesebogen

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
 ReDim rTh(0)
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

Public Function doEntleer(frm As lese, Tbl$)
 Dim rs As New ADODB.Recordset
 Set rs = DBCn.Execute("select count(0) As ct from `" & Tbl & "`")
 frm.Ausgeb "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)", True
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doEntleer

Public Function AllesLösch(frm As lese)
 Dim ct&, rs As new ADODB.recordset
 On Error GoTo fehler
 call ForeignNo0
 call ForeignNo1
 call doEntleer(frm, "therarten")
 call doEntleer(frm, "dmpreihe")
 call doEntleer(frm, "unbekannte kennungen")
 call doEntleer(frm, "kvnrue")
 call doEntleer(frm, "rr")
 call doEntleer(frm, "rezepteintraege")
 call doEntleer(frm, "medplan")
 call doEntleer(frm, "leistungen")
 call doEntleer(frm, "laborneu")
 call doEntleer(frm, "lbanforderungen")
 call doEntleer(frm, "kheinweis")
 call doEntleer(frm, "forminhfeld")
 call doEntleer(frm, "forminhkopf")
 call doEntleer(frm, "formulare")
 call doEntleer(frm, "forminhaltform_abk")
 call doEntleer(frm, "eintraege")
 call doEntleer(frm, "dokumente")
 call doEntleer(frm, "diagnosen")
 call doEntleer(frm, "briefe")
 call doEntleer(frm, "au")
 call doEntleer(frm, "faelle")
 call doEntleer(frm, "namen")
 call ForeignYes0
 call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' AllesLösch

Public Function LabLösch(frm As lese)
 Dim ct&, rs As new ADODB.recordset
 On Error GoTo fehler
 call ForeignNo0
 call ForeignNo1
 call doEntleer(frm, "laborxleist")
 call doEntleer(frm, "laborxwert")
 call doEntleer(frm, "laborxbakt")
 call doEntleer(frm, "laborxus")
 call doEntleer(frm, "laborxeingel")
 call doEntleer(frm, "laborxsaetze")
 call doentleer(frm, "laborxplab")
 call doentleer(frm, "laborxpneu")
 call doentleer(frm, "laborxpnb")
 call ForeignYes0
 call ForeignYes1
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end Select
End Function ' LabLösch

Function doBezFeh(csqlVal$, ErrDes$)
 Call ForeignNo0
 Call ForeignNo1
 obfor = True
 If lies.dlg.BeziehungsfehlerSpeichern <> 0 Then
  Open BezFeh For Append As #299
  Print #299, vbCrLf & vbCrLf & Now() & ": " & csqlVal
  Print #299, vbCrLf & ErrDes
  Close #299
 End If
End Function 'doBezFeh

Public Function namenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere namen"
 if not Allepat then
   sql = "delete from `namen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Postfach,Weggeldzone," & _
     "WeggzZahl,AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax," & _
     "DienstTel,PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit," & _
     "absPos,StByte,Cave,Notiz,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,Straße,KVKStatus,Geschlecht,Plz,Ort,Postfach,Weggeldzone," & _
     "WeggzZahl,AufnDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,PrivatTel_2,PrivatFax," & _
     "DienstTel,PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit," & _
     "absPos,StByte,Cave,Notiz,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM)       values"))
 For i = 0 to ubound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rNa(i).Pat_ID , "," , rNa(i).lfdnr , ",'" , rNa(i).NVorsatz , "','" , rNa(i).Nachname , "','" , rNa(i).Vorname , "'," , datformk( _
   rNa(i).GebDat) , ",'" , rNa(i).Straße , "','" , rNa(i).KVKStatus , "','" , rNa(i).Geschlecht , "','" , rNa(i).Plz , "','" , rNa(i).Ort , "','" ,  _
   rNa(i).Postfach , "','" , rNa(i).Weggeldzone , "'," , rNa(i).WeggzZahl , "," , datformk(rNa(i).AufnDat) , ",'" , rNa(i).LANR , "','" , rNa(i).BStNr , "','" ,  _
   rNa(i).Titel , "','" , rNa(i).Versichertennummer , "','" , rNa(i).PrivatTel , "','" , rNa(i).KVNr , "','" , rNa(i).PrivatTel_2 , "','" ,  _
   rNa(i).PrivatFax , "','" , rNa(i).DienstTel , "','" , rNa(i).PrivatMobil , "','" , rNa(i).Email , "','" , rNa(i).Arbeitgeber , "'," , cstr(-( _
   rNa(i).AnAllgda<>0)) , "," , cstr(-(rNa(i).An1da<>0)) , "," , cstr(-(rNa(i).An2da<>0)) , "," , cstr(-(rNa(i).Checkda<>0)) , ",'" ,  _
   rNa(i).DMTypaD , "'," , datformk( 0 ) , "," , rNa(i).absPos , "," , rNa(i).StByte , ",'" , rNa(i).Cave , "','" , rNa(i).Notiz , "'," , rNa(i).dmpklass , "," , datformk( _
   rNa(i).dmpbeg) , "," , rNa(i).dmpkhkklass , "," , datformk(rNa(i).dmpkhkbeg) , "," , rNa(i).dmpcopdklass , "," , datformk( _
   rNa(i).dmpcopdbeg) , "," , rNa(i).getHA0 , ",'" , rNa(i).fnHA0 , "'," , rNa(i).getHA1 , ",'" , rNa(i).fnHA1 , "'," , rNa(i).getHA2 , ",'" ,  _
   rNa(i).fnHA2 , "','" , rNa(i).zubenach , "','" , rNa(i).Verwandt , "','" , rNa(i).Sprache , "'," , datformk(rNa(i).lAktTM) , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rNa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rNa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(30)
 for k = iif(Lese.dlg.SammelInsert<>0,0,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rNa),i)
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
  If Len(rNa(k).LANR) > maxi(10) Then maxi(10) = Len(rNa(k).LANR)
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
  If Len(rNa(k).fnHA0) > maxi(25) Then maxi(25) = Len(rNa(k).fnHA0)
  If Len(rNa(k).fnHA1) > maxi(26) Then maxi(26) = Len(rNa(k).fnHA1)
  If Len(rNa(k).fnHA2) > maxi(27) Then maxi(27) = Len(rNa(k).fnHA2)
  If Len(rNa(k).zubenach) > maxi(28) Then maxi(28) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(29) Then maxi(29) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(30) Then maxi(30) = Len(rNa(k).Sprache)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,0, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rNa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & Left$(rNa(k).NVorsatz, maxL)  & "'",true : rNa(k).NVorsatz = Left$(rNa(k).NVorsatz, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & Left$(rNa(k).Nachname, maxL)  & "'",true : rNa(k).Nachname = Left$(rNa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & Left$(rNa(k).Vorname, maxL)  & "'",true : rNa(k).Vorname = Left$(rNa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & Left$(rNa(k).Straße, maxL)  & "'",true : rNa(k).Straße = Left$(rNa(k).Straße, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & Left$(rNa(k).KVKStatus, maxL)  & "'",true : rNa(k).KVKStatus = Left$(rNa(k).KVKStatus, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).Geschlecht & "' -> '" & Left$(rNa(k).Geschlecht, maxL)  & "'",true : rNa(k).Geschlecht = Left$(rNa(k).Geschlecht, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).Plz & "' -> '" & Left$(rNa(k).Plz, maxL)  & "'",true : rNa(k).Plz = Left$(rNa(k).Plz, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).Ort & "' -> '" & Left$(rNa(k).Ort, maxL)  & "'",true : rNa(k).Ort = Left$(rNa(k).Ort, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach: '" & rNa(k).Postfach & "' -> '" & Left$(rNa(k).Postfach, maxL)  & "'",true : rNa(k).Postfach = Left$(rNa(k).Postfach, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & Left$(rNa(k).Weggeldzone, maxL)  & "'",true : rNa(k).Weggeldzone = Left$(rNa(k).Weggeldzone, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).LANR & "' -> '" & Left$(rNa(k).LANR, maxL)  & "'",true : rNa(k).LANR = Left$(rNa(k).LANR, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rNa.BStNr: '" & rNa(k).BStNr & "' -> '" & Left$(rNa(k).BStNr, maxL)  & "'",true : rNa(k).BStNr = Left$(rNa(k).BStNr, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rNa.Titel: '" & rNa(k).Titel & "' -> '" & Left$(rNa(k).Titel, maxL)  & "'",true : rNa(k).Titel = Left$(rNa(k).Titel, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rNa.Versichertennummer: '" & rNa(k).Versichertennummer & "' -> '" & Left$(rNa(k).Versichertennummer, maxL)  & "'",true : rNa(k).Versichertennummer = Left$(rNa(k).Versichertennummer, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel: '" & rNa(k).PrivatTel & "' -> '" & Left$(rNa(k).PrivatTel, maxL)  & "'",true : rNa(k).PrivatTel = Left$(rNa(k).PrivatTel, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr: '" & rNa(k).KVNr & "' -> '" & Left$(rNa(k).KVNr, maxL)  & "'",true : rNa(k).KVNr = Left$(rNa(k).KVNr, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel_2: '" & rNa(k).PrivatTel_2 & "' -> '" & Left$(rNa(k).PrivatTel_2, maxL)  & "'",true : rNa(k).PrivatTel_2 = Left$(rNa(k).PrivatTel_2, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatFax: '" & rNa(k).PrivatFax & "' -> '" & Left$(rNa(k).PrivatFax, maxL)  & "'",true : rNa(k).PrivatFax = Left$(rNa(k).PrivatFax, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rNa.DienstTel: '" & rNa(k).DienstTel & "' -> '" & Left$(rNa(k).DienstTel, maxL)  & "'",true : rNa(k).DienstTel = Left$(rNa(k).DienstTel, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatMobil: '" & rNa(k).PrivatMobil & "' -> '" & Left$(rNa(k).PrivatMobil, maxL)  & "'",true : rNa(k).PrivatMobil = Left$(rNa(k).PrivatMobil, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rNa.Email: '" & rNa(k).Email & "' -> '" & Left$(rNa(k).Email, maxL)  & "'",true : rNa(k).Email = Left$(rNa(k).Email, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rNa.Arbeitgeber: '" & rNa(k).Arbeitgeber & "' -> '" & Left$(rNa(k).Arbeitgeber, maxL)  & "'",true : rNa(k).Arbeitgeber = Left$(rNa(k).Arbeitgeber, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rNa.DMTypaD: '" & rNa(k).DMTypaD & "' -> '" & Left$(rNa(k).DMTypaD, maxL)  & "'",true : rNa(k).DMTypaD = Left$(rNa(k).DMTypaD, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rNa.Cave: '" & rNa(k).Cave & "' -> '" & Left$(rNa(k).Cave, maxL)  & "'",true : rNa(k).Cave = Left$(rNa(k).Cave, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rNa.Notiz: '" & rNa(k).Notiz & "' -> '" & Left$(rNa(k).Notiz, maxL)  & "'",true : rNa(k).Notiz = Left$(rNa(k).Notiz, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & Left$(rNa(k).fnHA0, maxL)  & "'",true : rNa(k).fnHA0 = Left$(rNa(k).fnHA0, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & Left$(rNa(k).fnHA1, maxL)  & "'",true : rNa(k).fnHA1 = Left$(rNa(k).fnHA1, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & Left$(rNa(k).fnHA2, maxL)  & "'",true : rNa(k).fnHA2 = Left$(rNa(k).fnHA2, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL)  & "'",true : rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL)  & "'",true : rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL)  & "'",true : rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
ElseIf Err.Number = -2147217871 Or Err.Number = -2147217859 Or err.Number = -2147467259 Then
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
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function faelleSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere faelle"
 if not Allepat then
  sql = "select pat_id from `faelle` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `faelle` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
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
 For i = 1 to ubound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFa(i).Pat_ID , ",'" , rFa(i).Quartal , "','" , rFa(i).Nachname , "','" , rFa(i).Vorname , "'," , rFa(i).lfdnr , ",'" , rFa(i).TMFNr , "','" ,  _
   rFa(i).VKNr , "'," , datformk(rFa(i).BhFB) , "," , datformk(rFa(i).BhFE1) , "," , datformk(rFa(i).BhFE2) , ",'" , rFa(i).f4202 , "'," , datformk( _
   rFa(i).ausgst) , ",'" , rFa(i).KtrAbrB , "','" , rFa(i).AbrAr , "'," , datformk(rFa(i).lVorl) , ",'" , rFa(i).IK , "','" ,  _
   rFa(i).KVKs , "','" , rFa(i).KVKserg , "','" , rFa(i).Kasse , "','" , rFa(i).GebOr , "','" , rFa(i).AbrGb , "','" , rFa(i).PersKreis , "','" ,  _
   rFa(i).SKtZusatz , "','" , rFa(i).f4206 , "','" , rFa(i).ÜwText , "'," , rFa(i).f4210 , "," , rFa(i).AkfHAH , "," , rFa(i).AkfAB0 , "," , rFa(i).AkfAK , ",'" ,  _
   rFa(i).statNuller , "','" , rFa(i).ÜbwV , "','" , rFa(i).ÜbWVLANR , "','" , rFa(i).ÜbWVBSNR , "','" , rFa(i).ÜbWVKVNR , "','" ,  _
   rFa(i).AndÜw , "','" , rFa(i).Übwr , "','" , rFa(i).ÜbwLANR , "','" , rFa(i).ÜWZiel , "','" , rFa(i).ÜWNNr , "','" , rFa(i).ÜWNaN , "','" , rFa(i).ÜWTit , "','" ,  _
   rFa(i).ÜWVor , "','" , rFa(i).ÜWVsw , "'," , rFa(i).üwvid , ",'" , rFa(i).Auftrag , "','" , rFa(i).Verdacht , "','" , rFa(i).Befund , "','" ,  _
   rFa(i).statKlasse , "','" , rFa(i).f4237 , "'," , rFa(i).statBehTage , "," , rFa(i).SchGr , ",'" , rFa(i).Weiterbeh , "','" ,  _
   rFa(i).PGeb , "','" , rFa(i).PGebErg , "','" , rFa(i).Mahnfrist , "','" , rFa(i).GOÄKatNr , "','" , rFa(i).GOÄKatName , "','" , rFa(i).abrArzt , "','" ,  _
   rFa(i).privVers , "','" , rFa(i).AdNam , "','" , rFa(i).AdStr , "','" , rFa(i).AdPlz , "','" , rFa(i).AdOrt , "'," , datformk(rFa(i).BhFE) , ",'" ,  _
   rFa(i).s8000 , "','" , rFa(i).s8100 , "'," , datformk(rFa(i).AktZeit) , "," , datformk(rFa(i).Fanf) , ",'" , rFa(i).altQuart , "'," , datformk( _
   rFa(i).QAnf) , "," , datformk(rFa(i).QEnd) , ",'" , rFa(i).QS , "','" , rFa(i).QT , "'," , rFa(i).TherArt , "," , rFa(i).StByte , "," ,  _
   rFa(i).absPos , "," , rFa(i).LANRid , ",'" , rFa(i).f4108 , "','" , rFa(i).BGFallNr , "'," , rFa(i).lGewicht , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rFa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rFa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
  Set rs = DBCn.Execute("select * from faelle where pat_id = " & rFa(i).Pat_ID & " and quartal = '" & rFa(i).Quartal & "' and bhfb = " & DatFormk(rFa(i).BhFB) & " and bhfe1 = " & DatFormK(rFa(i).BhFE1) & " and ausgst = " & datformk(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & rs!FID, true
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
 redim maxi(55)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rFa),i)
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
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rFa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & Left$(rFa(k).Quartal, maxL)  & "'",true : rFa(k).Quartal = Left$(rFa(k).Quartal, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & Left$(rFa(k).Nachname, maxL)  & "'",true : rFa(k).Nachname = Left$(rFa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & Left$(rFa(k).Vorname, maxL)  & "'",true : rFa(k).Vorname = Left$(rFa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & Left$(rFa(k).TMFNr, maxL)  & "'",true : rFa(k).TMFNr = Left$(rFa(k).TMFNr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & Left$(rFa(k).VKNr, maxL)  & "'",true : rFa(k).VKNr = Left$(rFa(k).VKNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4202: '" & rFa(k).f4202 & "' -> '" & Left$(rFa(k).f4202, maxL)  & "'",true : rFa(k).f4202 = Left$(rFa(k).f4202, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & Left$(rFa(k).KtrAbrB, maxL)  & "'",true : rFa(k).KtrAbrB = Left$(rFa(k).KtrAbrB, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & Left$(rFa(k).AbrAr, maxL)  & "'",true : rFa(k).AbrAr = Left$(rFa(k).AbrAr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & Left$(rFa(k).IK, maxL)  & "'",true : rFa(k).IK = Left$(rFa(k).IK, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & Left$(rFa(k).KVKs, maxL)  & "'",true : rFa(k).KVKs = Left$(rFa(k).KVKs, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & Left$(rFa(k).KVKserg, maxL)  & "'",true : rFa(k).KVKserg = Left$(rFa(k).KVKserg, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.Kasse: '" & rFa(k).Kasse & "' -> '" & Left$(rFa(k).Kasse, maxL)  & "'",true : rFa(k).Kasse = Left$(rFa(k).Kasse, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & Left$(rFa(k).GebOr, maxL)  & "'",true : rFa(k).GebOr = Left$(rFa(k).GebOr, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & Left$(rFa(k).AbrGb, maxL)  & "'",true : rFa(k).AbrGb = Left$(rFa(k).AbrGb, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & Left$(rFa(k).PersKreis, maxL)  & "'",true : rFa(k).PersKreis = Left$(rFa(k).PersKreis, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & Left$(rFa(k).SKtZusatz, maxL)  & "'",true : rFa(k).SKtZusatz = Left$(rFa(k).SKtZusatz, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4206: '" & rFa(k).f4206 & "' -> '" & Left$(rFa(k).f4206, maxL)  & "'",true : rFa(k).f4206 = Left$(rFa(k).f4206, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & Left$(rFa(k).ÜwText, maxL)  & "'",true : rFa(k).ÜwText = Left$(rFa(k).ÜwText, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & Left$(rFa(k).statNuller, maxL)  & "'",true : rFa(k).statNuller = Left$(rFa(k).statNuller, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & Left$(rFa(k).ÜbwV, maxL)  & "'",true : rFa(k).ÜbwV = Left$(rFa(k).ÜbwV, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVLANR: '" & rFa(k).ÜbWVLANR & "' -> '" & Left$(rFa(k).ÜbWVLANR, maxL)  & "'",true : rFa(k).ÜbWVLANR = Left$(rFa(k).ÜbWVLANR, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVBSNR: '" & rFa(k).ÜbWVBSNR & "' -> '" & Left$(rFa(k).ÜbWVBSNR, maxL)  & "'",true : rFa(k).ÜbWVBSNR = Left$(rFa(k).ÜbWVBSNR, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVKVNR: '" & rFa(k).ÜbWVKVNR & "' -> '" & Left$(rFa(k).ÜbWVKVNR, maxL)  & "'",true : rFa(k).ÜbWVKVNR = Left$(rFa(k).ÜbWVKVNR, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & Left$(rFa(k).AndÜw, maxL)  & "'",true : rFa(k).AndÜw = Left$(rFa(k).AndÜw, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übwr: '" & rFa(k).Übwr & "' -> '" & Left$(rFa(k).Übwr, maxL)  & "'",true : rFa(k).Übwr = Left$(rFa(k).Übwr, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & Left$(rFa(k).ÜbwLANR, maxL)  & "'",true : rFa(k).ÜbwLANR = Left$(rFa(k).ÜbwLANR, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & Left$(rFa(k).ÜWZiel, maxL)  & "'",true : rFa(k).ÜWZiel = Left$(rFa(k).ÜWZiel, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & Left$(rFa(k).ÜWNNr, maxL)  & "'",true : rFa(k).ÜWNNr = Left$(rFa(k).ÜWNNr, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & Left$(rFa(k).ÜWNaN, maxL)  & "'",true : rFa(k).ÜWNaN = Left$(rFa(k).ÜWNaN, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & Left$(rFa(k).ÜWTit, maxL)  & "'",true : rFa(k).ÜWTit = Left$(rFa(k).ÜWTit, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & Left$(rFa(k).ÜWVor, maxL)  & "'",true : rFa(k).ÜWVor = Left$(rFa(k).ÜWVor, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & Left$(rFa(k).ÜWVsw, maxL)  & "'",true : rFa(k).ÜWVsw = Left$(rFa(k).ÜWVsw, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.Auftrag: '" & rFa(k).Auftrag & "' -> '" & Left$(rFa(k).Auftrag, maxL)  & "'",true : rFa(k).Auftrag = Left$(rFa(k).Auftrag, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.Verdacht: '" & rFa(k).Verdacht & "' -> '" & Left$(rFa(k).Verdacht, maxL)  & "'",true : rFa(k).Verdacht = Left$(rFa(k).Verdacht, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.Befund: '" & rFa(k).Befund & "' -> '" & Left$(rFa(k).Befund, maxL)  & "'",true : rFa(k).Befund = Left$(rFa(k).Befund, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & Left$(rFa(k).statKlasse, maxL)  & "'",true : rFa(k).statKlasse = Left$(rFa(k).statKlasse, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4237: '" & rFa(k).f4237 & "' -> '" & Left$(rFa(k).f4237, maxL)  & "'",true : rFa(k).f4237 = Left$(rFa(k).f4237, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & Left$(rFa(k).Weiterbeh, maxL)  & "'",true : rFa(k).Weiterbeh = Left$(rFa(k).Weiterbeh, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & Left$(rFa(k).PGeb, maxL)  & "'",true : rFa(k).PGeb = Left$(rFa(k).PGeb, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & Left$(rFa(k).PGebErg, maxL)  & "'",true : rFa(k).PGebErg = Left$(rFa(k).PGebErg, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & Left$(rFa(k).Mahnfrist, maxL)  & "'",true : rFa(k).Mahnfrist = Left$(rFa(k).Mahnfrist, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & Left$(rFa(k).GOÄKatNr, maxL)  & "'",true : rFa(k).GOÄKatNr = Left$(rFa(k).GOÄKatNr, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & Left$(rFa(k).GOÄKatName, maxL)  & "'",true : rFa(k).GOÄKatName = Left$(rFa(k).GOÄKatName, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & Left$(rFa(k).abrArzt, maxL)  & "'",true : rFa(k).abrArzt = Left$(rFa(k).abrArzt, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & Left$(rFa(k).privVers, maxL)  & "'",true : rFa(k).privVers = Left$(rFa(k).privVers, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & Left$(rFa(k).AdNam, maxL)  & "'",true : rFa(k).AdNam = Left$(rFa(k).AdNam, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & Left$(rFa(k).AdStr, maxL)  & "'",true : rFa(k).AdStr = Left$(rFa(k).AdStr, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & Left$(rFa(k).AdPlz, maxL)  & "'",true : rFa(k).AdPlz = Left$(rFa(k).AdPlz, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & Left$(rFa(k).AdOrt, maxL)  & "'",true : rFa(k).AdOrt = Left$(rFa(k).AdOrt, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & Left$(rFa(k).s8000, maxL)  & "'",true : rFa(k).s8000 = Left$(rFa(k).s8000, maxL)
       Case 50: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & Left$(rFa(k).s8100, maxL)  & "'",true : rFa(k).s8100 = Left$(rFa(k).s8100, maxL)
       Case 51: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & Left$(rFa(k).altQuart, maxL)  & "'",true : rFa(k).altQuart = Left$(rFa(k).altQuart, maxL)
       Case 52: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & Left$(rFa(k).QS, maxL)  & "'",true : rFa(k).QS = Left$(rFa(k).QS, maxL)
       Case 53: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & Left$(rFa(k).QT, maxL)  & "'",true : rFa(k).QT = Left$(rFa(k).QT, maxL)
       Case 54: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4108: '" & rFa(k).f4108 & "' -> '" & Left$(rFa(k).f4108, maxL)  & "'",true : rFa(k).f4108 = Left$(rFa(k).f4108, maxL)
       Case 55: Lese.Ausgeb "   Verkürze Inhalt von rFa.BGFallNr: '" & rFa(k).BGFallNr & "' -> '" & Left$(rFa(k).BGFallNr, maxL)  & "'",true : rFa(k).BGFallNr = Left$(rFa(k).BGFallNr, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "insert into `kassenliste`(`GO`,`VK`,`IK`) values (" & "'" & rFa(I).GOÄKatName & "', '" & rFa(I).VKNr & "', '" & rFa(I).IK & "')"
 InsKorr DBCn, DBCnS, sqlquer, rAF
 Resume
end if ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function auSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere au"
 if not Allepat then
  sql = "select pat_id from `au` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `au` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            values"))
 For i = 1 to ubound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rAu(i).FID , "," , rAu(i).Pat_ID , "," , datformk(rAu(i).ZeitPunkt) , ",'" , rAu(i).Beginn , "','" , rAu(i).Ende , "','" , rAu(i).ICDs , "'," ,  _
   rAu(i).absPos , "," , datformk(rAu(i).AktZeit) , "," , rAu(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rAu) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rAu) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(2)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rAu),i)
  If Len(rAu(k).Beginn) > maxi(0) Then maxi(0) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(1) Then maxi(1) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(2) Then maxi(2) = Len(rAu(k).ICDs)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rAu), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & Left$(rAu(k).Beginn, maxL)  & "'",true : rAu(k).Beginn = Left$(rAu(k).Beginn, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & Left$(rAu(k).Ende, maxL)  & "'",true : rAu(k).Ende = Left$(rAu(k).Ende, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & Left$(rAu(k).ICDs, maxL)  & "'",true : rAu(k).ICDs = Left$(rAu(k).ICDs, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function briefeSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere briefe"
 if not Allepat then
  sql = "select pat_id from `briefe` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `briefe` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,DokAenD,QS,QT,absPos," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Typ,AktZeit,DokGroe,DokAenD,QS,QT,absPos," & _
     "StByte)  values"))
 For i = 1 to ubound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rBr(i).FID , "," , rBr(i).Pat_ID , "," , datformk(rBr(i).ZeitPunkt) , ",'" , rBr(i).Pfad , "','" , rBr(i).Art , "','" , rBr(i).Name , "','" ,  _
   rBr(i).Typ , "'," , datformk(rBr(i).AktZeit) , "," , rBr(i).DokGroe , "," , datformk(rBr(i).DokAenD) , ",'" , rBr(i).QS , "','" ,  _
   rBr(i).QT , "'," , rBr(i).absPos , "," , rBr(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rBr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rBr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(5)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rBr),i)
  If Len(rBr(k).Pfad) > maxi(0) Then maxi(0) = Len(rBr(k).Pfad)
  If Len(rBr(k).Art) > maxi(1) Then maxi(1) = Len(rBr(k).Art)
  If Len(rBr(k).Name) > maxi(2) Then maxi(2) = Len(rBr(k).Name)
  If Len(rBr(k).Typ) > maxi(3) Then maxi(3) = Len(rBr(k).Typ)
  If Len(rBr(k).QS) > maxi(4) Then maxi(4) = Len(rBr(k).QS)
  If Len(rBr(k).QT) > maxi(5) Then maxi(5) = Len(rBr(k).QT)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rBr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL)  & "'",true : rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).Art & "' -> '" & Left$(rBr(k).Art, maxL)  & "'",true : rBr(k).Art = Left$(rBr(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).Name & "' -> '" & Left$(rBr(k).Name, maxL)  & "'",true : rBr(k).Name = Left$(rBr(k).Name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & Left$(rBr(k).Typ, maxL)  & "'",true : rBr(k).Typ = Left$(rBr(k).Typ, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & Left$(rBr(k).QS, maxL)  & "'",true : rBr(k).QS = Left$(rBr(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & Left$(rBr(k).QT, maxL)  & "'",true : rBr(k).QT = Left$(rBr(k).QT, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function diagnosenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere diagnosen"
 if not Allepat then
  sql = "select pat_id from `diagnosen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `diagnosen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit," & _
     "StByte,AusnBegr,f6010,f6011) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `diagnosen` (FID,Pat_id,GesName," & _
     "DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit," & _
     "StByte,AusnBegr,f6010,f6011)         values"))
 For i = 1 to ubound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDi(i).FID , "," , rDi(i).Pat_id , ",'" , rDi(i).GesName , "'," , datformk(rDi(i).DiagDatum) , ",'" , rDi(i).DiagSicherheit , "','" ,  _
   rDi(i).DiagText , "','" , rDi(i).DiagSeite , "','" , rDi(i).DiagAttr , "','" , rDi(i).ICD , "'," , rDi(i).obDauer , ",'" , rDi(i).intBemerk , "'," ,  _
   rDi(i).absPos , "," , datformk(rDi(i).AktZeit) , "," , rDi(i).StByte , ",'" , rDi(i).AusnBegr , "'," , rDi(i).f6010 , ",'" , rDi(i).f6011 , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rDi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rDi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(8)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rDi),i)
  If Len(rDi(k).GesName) > maxi(0) Then maxi(0) = Len(rDi(k).GesName)
  If Len(rDi(k).DiagSicherheit) > maxi(1) Then maxi(1) = Len(rDi(k).DiagSicherheit)
  If Len(rDi(k).DiagText) > maxi(2) Then maxi(2) = Len(rDi(k).DiagText)
  If Len(rDi(k).DiagSeite) > maxi(3) Then maxi(3) = Len(rDi(k).DiagSeite)
  If Len(rDi(k).DiagAttr) > maxi(4) Then maxi(4) = Len(rDi(k).DiagAttr)
  If Len(rDi(k).ICD) > maxi(5) Then maxi(5) = Len(rDi(k).ICD)
  If Len(rDi(k).intBemerk) > maxi(6) Then maxi(6) = Len(rDi(k).intBemerk)
  If Len(rDi(k).AusnBegr) > maxi(7) Then maxi(7) = Len(rDi(k).AusnBegr)
  If Len(rDi(k).f6011) > maxi(8) Then maxi(8) = Len(rDi(k).f6011)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rDi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDi.GesName: '" & rDi(k).GesName & "' -> '" & Left$(rDi(k).GesName, maxL)  & "'",true : rDi(k).GesName = Left$(rDi(k).GesName, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & Left$(rDi(k).DiagSicherheit, maxL)  & "'",true : rDi(k).DiagSicherheit = Left$(rDi(k).DiagSicherheit, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & Left$(rDi(k).DiagText, maxL)  & "'",true : rDi(k).DiagText = Left$(rDi(k).DiagText, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & Left$(rDi(k).DiagSeite, maxL)  & "'",true : rDi(k).DiagSeite = Left$(rDi(k).DiagSeite, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & Left$(rDi(k).DiagAttr, maxL)  & "'",true : rDi(k).DiagAttr = Left$(rDi(k).DiagAttr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & Left$(rDi(k).ICD, maxL)  & "'",true : rDi(k).ICD = Left$(rDi(k).ICD, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & Left$(rDi(k).intBemerk, maxL)  & "'",true : rDi(k).intBemerk = Left$(rDi(k).intBemerk, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.AusnBegr: '" & rDi(k).AusnBegr & "' -> '" & Left$(rDi(k).AusnBegr, maxL)  & "'",true : rDi(k).AusnBegr = Left$(rDi(k).AusnBegr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rDi.f6011: '" & rDi(k).f6011 & "' -> '" & Left$(rDi(k).f6011, maxL)  & "'",true : rDi(k).f6011 = Left$(rDi(k).f6011, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function dokumenteSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere dokumente"
 if not Allepat then
  sql = "select pat_id from `dokumente` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `dokumente` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte)  values"))
 For i = 1 to ubound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDo(i).FID , "," , rDo(i).Pat_ID , "," , datformk(rDo(i).ZeitPunkt) , ",'" , rDo(i).DokPfad , "','" , rDo(i).DokArt , "','" ,  _
   rDo(i).DokName , "'," , datformk(rDo(i).Quelldatum) , "," , rDo(i).absPos , "," , datformk(rDo(i).AktZeit) , "," , rDo(i).DokGroe , "," , datformk( _
   rDo(i).DokAenD) , ",'" , rDo(i).QS , "','" , rDo(i).QT , "'," , rDo(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rDo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rDo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(4)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rDo),i)
  If Len(rDo(k).DokPfad) > maxi(0) Then maxi(0) = Len(rDo(k).DokPfad)
  If Len(rDo(k).DokArt) > maxi(1) Then maxi(1) = Len(rDo(k).DokArt)
  If Len(rDo(k).DokName) > maxi(2) Then maxi(2) = Len(rDo(k).DokName)
  If Len(rDo(k).QS) > maxi(3) Then maxi(3) = Len(rDo(k).QS)
  If Len(rDo(k).QT) > maxi(4) Then maxi(4) = Len(rDo(k).QT)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rDo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & Left$(rDo(k).DokPfad, maxL)  & "'",true : rDo(k).DokPfad = Left$(rDo(k).DokPfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & Left$(rDo(k).DokArt, maxL)  & "'",true : rDo(k).DokArt = Left$(rDo(k).DokArt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & Left$(rDo(k).DokName, maxL)  & "'",true : rDo(k).DokName = Left$(rDo(k).DokName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & Left$(rDo(k).QS, maxL)  & "'",true : rDo(k).QS = Left$(rDo(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & Left$(rDo(k).QT, maxL)  & "'",true : rDo(k).QT = Left$(rDo(k).QT, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function eintraegeSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere eintraege"
 if not Allepat then
  sql = "select pat_id from `eintraege` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `eintraege` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum)     values"))
 For i = 1 to ubound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rEi(i).FID , "," , rEi(i).Pat_ID , "," , datformk(rEi(i).ZeitPunkt) , ",'" , rEi(i).Art , "','" , rEi(i).Inhalt , "'," , rEi(i).absPos , "," , datformk( _
   rEi(i).AktZeit) , ",'" , rEi(i).QS , "','" , rEi(i).QT , "'," , rEi(i).StByte , "," , rEi(i).inhNum , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rEi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rEi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(3)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rEi),i)
  If Len(rEi(k).Art) > maxi(0) Then maxi(0) = Len(rEi(k).Art)
  If Len(rEi(k).Inhalt) > maxi(1) Then maxi(1) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(2) Then maxi(2) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(3) Then maxi(3) = Len(rEi(k).QT)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rEi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).Art & "' -> '" & Left$(rEi(k).Art, maxL)  & "'",true : rEi(k).Art = Left$(rEi(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & Left$(rEi(k).Inhalt, maxL)  & "'",true : rEi(k).Inhalt = Left$(rEi(k).Inhalt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & Left$(rEi(k).QS, maxL)  & "'",true : rEi(k).QS = Left$(rEi(k).QS, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & Left$(rEi(k).QT, maxL)  & "'",true : rEi(k).QT = Left$(rEi(k).QT, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhaltform_abk"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhaltform_abk` (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhaltform_abk` (Form_Abk)             values"))
 For i = rFi1 + 1 to ubound(rFi)
'  rFi(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rFi(i).Form_Abk , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rFi) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rFi) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 DoEvents
 Next i
 rFi1 = ubound(rFi)
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
 redim maxi(0)
 for k = iif(Lese.dlg.SammelInsert<>0,rFi1 + 1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rFi),i)
  If Len(rFi(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFi(k).Form_Abk)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,rFi1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rFi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFi.Form_Abk: '" & rFi(k).Form_Abk & "' -> '" & Left$(rFi(k).Form_Abk, maxL)  & "'",true : rFi(k).Form_Abk = Left$(rFi(k).Form_Abk, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere formulare"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      values"))
 For i = rFo1 + 1 to ubound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFo(i).FormID , ",'" , rFo(i).Form_Abk , "','" , rFo(i).FormBez , "','" , rFo(i).FormVorl , "'," , datformk(rFo(i).AktZeit) , "," ,  _
   rFo(i).absPos , "," , rFo(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rFo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rFo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 DoEvents
 Next i
 rFo1 = ubound(rFo)
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
 redim maxi(2)
 for k = iif(Lese.dlg.SammelInsert<>0,rFo1 + 1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rFo),i)
  If Len(rFo(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFo(k).Form_Abk)
  If Len(rFo(k).FormBez) > maxi(1) Then maxi(1) = Len(rFo(k).FormBez)
  If Len(rFo(k).FormVorl) > maxi(2) Then maxi(2) = Len(rFo(k).FormVorl)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,rFo1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rFo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & Left$(rFo(k).Form_Abk, maxL)  & "'",true : rFo(k).Form_Abk = Left$(rFo(k).Form_Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & Left$(rFo(k).FormBez, maxL)  & "'",true : rFo(k).FormBez = Left$(rFo(k).FormBez, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & Left$(rFo(k).FormVorl, maxL)  & "'",true : rFo(k).FormVorl = Left$(rFo(k).FormVorl, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function forminhkopfSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhkopf"
 if not Allepat then
'  sql = "delete FROM `forminhfeld` where foid in (select foid from `forminhkopf` where pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "select pat_id from `forminhkopf` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `forminhkopf` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid)              values"))
 For i = 1 to ubound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFr(i).FoID , "," , rFr(i).FID , "," , rFr(i).Pat_ID , "," , rFr(i).Form_ID , "," , datformk(rFr(i).ZeitPunkt) , "," , rFr(i).AbsPos , "," , datformk( _
   rFr(i).AktZeit) , "," , rFr(i).StByte , ",'" , rFr(i).Satzart , "','" , rFr(i).Satzlänge , "'," , rFr(i).LANRid , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rFr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rFr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(1)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rFr),i)
  If Len(rFr(k).Satzart) > maxi(0) Then maxi(0) = Len(rFr(k).Satzart)
  If Len(rFr(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rFr(k).Satzlänge)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rFr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & Left$(rFr(k).Satzart, maxL)  & "'",true : rFr(k).Satzart = Left$(rFr(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & Left$(rFr(k).Satzlänge, maxL)  & "'",true : rFr(k).Satzlänge = Left$(rFr(k).Satzlänge, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhfeld"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      values"))
 For i = 1 to ubound(rFm)
'  rFm(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFm(i).FoID , "," , rFm(i).Nr , "," , rFm(i).FeldNr , "," , rFm(i).FeldVW , "," , rFm(i).FeldInhVW , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rFm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rFm) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(-1)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rFm),i)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhfeld", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rFm), i)
      Select Case m
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function kheinweisSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere kheinweis"
 if not Allepat then
  sql = "select pat_id from `kheinweis` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `kheinweis` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               values"))
 For i = 1 to ubound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKh(i).FID , "," , rKh(i).Pat_ID , "," , datformk(rKh(i).ZeitPunkt) , ",'" , rKh(i).Ziel , "','" , rKh(i).Diagnose , "'," , rKh(i).absPos , "," , datformk( _
   rKh(i).AktZeit) , "," , rKh(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rKh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rKh) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(1)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rKh),i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rKh), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & Left$(rKh(k).Ziel, maxL)  & "'",true : rKh(k).Ziel = Left$(rKh(k).Ziel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & Left$(rKh(k).Diagnose, maxL)  & "'",true : rKh(k).Diagnose = Left$(rKh(k).Diagnose, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function lbanforderungenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere lbanforderungen"
 if not Allepat then
  sql = "select pat_id from `lbanforderungen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `lbanforderungen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       values"))
 For i = 1 to ubound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLb(i).FID , "," , rLb(i).Pat_ID , "," , datformk(rLb(i).ZeitPunkt) , ",'" , rLb(i).AnfText , "'," , rLb(i).absPos , "," , datformk( _
   rLb(i).AktZeit) , "," , rLb(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLb) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLb) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(0)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLb),i)
  If Len(rLb(k).AnfText) > maxi(0) Then maxi(0) = Len(rLb(k).AnfText)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLb), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & Left$(rLb(k).AnfText, maxL)  & "'",true : rLb(k).AnfText = Left$(rLb(k).AnfText, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function laborneuSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborneu"
 if not Allepat then
  sql = "select pat_id from `laborneu` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `laborneu` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte)           values"))
 For i = 1 to ubound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLa(i).FID , "," , rLa(i).Pat_ID , "," , datformk(rLa(i).ZeitPunkt) , ",'" , rLa(i).FertigStGrad , "','" , rLa(i).Abkü , "'," ,  _
   rLa(i).LangtextVW , ",'" , rLa(i).Wert , "','" , rLa(i).Einheit , "'," , rLa(i).KommentarVW , "," , rLa(i).AbsPos , "," , datformk(rLa(i).AktZeit) , "," ,  _
   rLa(i).Refnr , "," , rLa(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLa) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLa) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(3)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLa),i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & Left$(rLa(k).FertigStGrad, maxL)  & "'",true : rLa(k).FertigStGrad = Left$(rLa(k).FertigStGrad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & Left$(rLa(k).Abkü, maxL)  & "'",true : rLa(k).Abkü = Left$(rLa(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & Left$(rLa(k).Wert, maxL)  & "'",true : rLa(k).Wert = Left$(rLa(k).Wert, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & Left$(rLa(k).Einheit, maxL)  & "'",true : rLa(k).Einheit = Left$(rLa(k).Einheit, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function leistungenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere leistungen"
 if not Allepat then
  sql = "select pat_id from `leistungen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `leistungen` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte,LANRid,Sachkbez,Sachkct,Zone) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT," & _
     "StByte,LANRid,Sachkbez,Sachkct,Zone)               values"))
 For i = 1 to ubound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLe(i).FID , "," , rLe(i).Pat_ID , "," , datformk(rLe(i).ZeitPunkt) , ",'" , rLe(i).Leistung , "','" , rLe(i).f5002 , "','" ,  _
   rLe(i).f5005 , "','" , rLe(i).f5006 , "','" , rLe(i).f5009 , "','" , rLe(i).Med , "','" , rLe(i).f5015 , "','" , rLe(i).f5016 , "','" , rLe(i).f5021 , "','" ,  _
   rLe(i).f5026 , "','" , rLe(i).Faktor , "','" , rLe(i).f5098 , "','" , rLe(i).LANR , "'," , datformk(rLe(i).letzVorg) , ",'" ,  _
   rLe(i).Ausn , "','" , rLe(i).Beme , "'," , rLe(i).absPos , "," , datformk(rLe(i).AktZeit) , ",'" , rLe(i).QS , "','" , rLe(i).QT , "'," , rLe(i).StByte , "," ,  _
   rLe(i).LANRid , ",'" , rLe(i).Sachkbez , "'," , rLe(i).Sachkct , ",'" , rLe(i).Zone , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(18)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLe),i)
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
  If Len(rLe(k).Sachkbez) > maxi(17) Then maxi(17) = Len(rLe(k).Sachkbez)
  If Len(rLe(k).Zone) > maxi(18) Then maxi(18) = Len(rLe(k).Zone)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLe.Leistung: '" & rLe(k).Leistung & "' -> '" & Left$(rLe(k).Leistung, maxL)  & "'",true : rLe(k).Leistung = Left$(rLe(k).Leistung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5002: '" & rLe(k).f5002 & "' -> '" & Left$(rLe(k).f5002, maxL)  & "'",true : rLe(k).f5002 = Left$(rLe(k).f5002, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5005: '" & rLe(k).f5005 & "' -> '" & Left$(rLe(k).f5005, maxL)  & "'",true : rLe(k).f5005 = Left$(rLe(k).f5005, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5006: '" & rLe(k).f5006 & "' -> '" & Left$(rLe(k).f5006, maxL)  & "'",true : rLe(k).f5006 = Left$(rLe(k).f5006, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5009: '" & rLe(k).f5009 & "' -> '" & Left$(rLe(k).f5009, maxL)  & "'",true : rLe(k).f5009 = Left$(rLe(k).f5009, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLe.Med: '" & rLe(k).Med & "' -> '" & Left$(rLe(k).Med, maxL)  & "'",true : rLe(k).Med = Left$(rLe(k).Med, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5015: '" & rLe(k).f5015 & "' -> '" & Left$(rLe(k).f5015, maxL)  & "'",true : rLe(k).f5015 = Left$(rLe(k).f5015, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5016: '" & rLe(k).f5016 & "' -> '" & Left$(rLe(k).f5016, maxL)  & "'",true : rLe(k).f5016 = Left$(rLe(k).f5016, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5021: '" & rLe(k).f5021 & "' -> '" & Left$(rLe(k).f5021, maxL)  & "'",true : rLe(k).f5021 = Left$(rLe(k).f5021, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5026: '" & rLe(k).f5026 & "' -> '" & Left$(rLe(k).f5026, maxL)  & "'",true : rLe(k).f5026 = Left$(rLe(k).f5026, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLe.Faktor: '" & rLe(k).Faktor & "' -> '" & Left$(rLe(k).Faktor, maxL)  & "'",true : rLe(k).Faktor = Left$(rLe(k).Faktor, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLe.f5098: '" & rLe(k).f5098 & "' -> '" & Left$(rLe(k).f5098, maxL)  & "'",true : rLe(k).f5098 = Left$(rLe(k).f5098, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).LANR & "' -> '" & Left$(rLe(k).LANR, maxL)  & "'",true : rLe(k).LANR = Left$(rLe(k).LANR, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL)  & "'",true : rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).Beme & "' -> '" & Left$(rLe(k).Beme, maxL)  & "'",true : rLe(k).Beme = Left$(rLe(k).Beme, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL)  & "'",true : rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL)  & "'",true : rLe(k).QT = Left$(rLe(k).QT, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & Left$(rLe(k).Sachkbez, maxL)  & "'",true : rLe(k).Sachkbez = Left$(rLe(k).Sachkbez, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & Left$(rLe(k).Zone, maxL)  & "'",true : rLe(k).Zone = Left$(rLe(k).Zone, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function medplanSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere medplan"
 if not Allepat then
  sql = "select pat_id from `medplan` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `medplan` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,FeldNr,mo,mi,nm,ab,zn," & _
     "bBed,Bemerkung,AbsPos,AktZeit,StByte)              values"))
 For i = 1 to ubound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rMe(i).FID , "," , rMe(i).Pat_ID , "," , rMe(i).MPNr , "," , datformk(rMe(i).ZeitPunkt) , "," , datformk(rMe(i).Datum) , ",'" ,  _
   rMe(i).Medikament , "','" , rMe(i).MedAnfang , "'," , rMe(i).FeldNr , ",'" , rMe(i).mo , "','" , rMe(i).mi , "','" , rMe(i).nm , "','" , rMe(i).ab , "','" ,  _
   rMe(i).zn , "'," , cstr(-(rMe(i).bBed<>0)) , ",'" , rMe(i).Bemerkung , "'," , rMe(i).AbsPos , "," , datformk(rMe(i).AktZeit) , "," ,  _
   rMe(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rMe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rMe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(7)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rMe),i)
  If Len(rMe(k).Medikament) > maxi(0) Then maxi(0) = Len(rMe(k).Medikament)
  If Len(rMe(k).MedAnfang) > maxi(1) Then maxi(1) = Len(rMe(k).MedAnfang)
  If Len(rMe(k).mo) > maxi(2) Then maxi(2) = Len(rMe(k).mo)
  If Len(rMe(k).mi) > maxi(3) Then maxi(3) = Len(rMe(k).mi)
  If Len(rMe(k).nm) > maxi(4) Then maxi(4) = Len(rMe(k).nm)
  If Len(rMe(k).ab) > maxi(5) Then maxi(5) = Len(rMe(k).ab)
  If Len(rMe(k).zn) > maxi(6) Then maxi(6) = Len(rMe(k).zn)
  If Len(rMe(k).Bemerkung) > maxi(7) Then maxi(7) = Len(rMe(k).Bemerkung)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rMe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & Left$(rMe(k).Medikament, maxL)  & "'",true : rMe(k).Medikament = Left$(rMe(k).Medikament, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & Left$(rMe(k).MedAnfang, maxL)  & "'",true : rMe(k).MedAnfang = Left$(rMe(k).MedAnfang, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & Left$(rMe(k).mo, maxL)  & "'",true : rMe(k).mo = Left$(rMe(k).mo, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & Left$(rMe(k).mi, maxL)  & "'",true : rMe(k).mi = Left$(rMe(k).mi, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & Left$(rMe(k).nm, maxL)  & "'",true : rMe(k).nm = Left$(rMe(k).nm, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & Left$(rMe(k).ab, maxL)  & "'",true : rMe(k).ab = Left$(rMe(k).ab, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).zn & "' -> '" & Left$(rMe(k).zn, maxL)  & "'",true : rMe(k).zn = Left$(rMe(k).zn, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & Left$(rMe(k).Bemerkung, maxL)  & "'",true : rMe(k).Bemerkung = Left$(rMe(k).Bemerkung, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function rezepteintraegeSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere rezepteintraege"
 if not Allepat then
  sql = "select pat_id from `rezepteintraege` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `rezepteintraege` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte,LANRid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,Rezeptklasse,Medikament,PZN,absPos,AktZeit,QS,QT,StByte,LANRid)         values"))
 For i = 1 to ubound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRe(i).FID , "," , rRe(i).Pat_ID , "," , datformk(rRe(i).ZeitPunkt) , ",'" , rRe(i).Rezept , "','" , rRe(i).Rezeptklasse , "','" ,  _
   rRe(i).Medikament , "','" , rRe(i).PZN , "'," , rRe(i).absPos , "," , datformk(rRe(i).AktZeit) , ",'" , rRe(i).QS , "','" , rRe(i).QT , "'," ,  _
   rRe(i).StByte , "," , rRe(i).LANRid , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rRe) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rRe) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(5)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rRe),i)
  If Len(rRe(k).Rezept) > maxi(0) Then maxi(0) = Len(rRe(k).Rezept)
  If Len(rRe(k).Rezeptklasse) > maxi(1) Then maxi(1) = Len(rRe(k).Rezeptklasse)
  If Len(rRe(k).Medikament) > maxi(2) Then maxi(2) = Len(rRe(k).Medikament)
  If Len(rRe(k).PZN) > maxi(3) Then maxi(3) = Len(rRe(k).PZN)
  If Len(rRe(k).QS) > maxi(4) Then maxi(4) = Len(rRe(k).QS)
  If Len(rRe(k).QT) > maxi(5) Then maxi(5) = Len(rRe(k).QT)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rRe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & Left$(rRe(k).Rezept, maxL)  & "'",true : rRe(k).Rezept = Left$(rRe(k).Rezept, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & Left$(rRe(k).Rezeptklasse, maxL)  & "'",true : rRe(k).Rezeptklasse = Left$(rRe(k).Rezeptklasse, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & Left$(rRe(k).Medikament, maxL)  & "'",true : rRe(k).Medikament = Left$(rRe(k).Medikament, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & Left$(rRe(k).PZN, maxL)  & "'",true : rRe(k).PZN = Left$(rRe(k).PZN, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & Left$(rRe(k).QS, maxL)  & "'",true : rRe(k).QS = Left$(rRe(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & Left$(rRe(k).QT, maxL)  & "'",true : rRe(k).QT = Left$(rRe(k).QT, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function rrSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere rr"
 if not Allepat then
  sql = "select pat_id from `rr` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `rr` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,Quelle,Bemerkung,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,Quelle,Bemerkung,absPos,AktZeit,StByte)    values"))
 For i = 1 to ubound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRr(i).FID , "," , rRr(i).Pat_ID , "," , datformk(rRr(i).ZeitPunkt) , ",'" , rRr(i).RR , "'," , rRr(i).Puls , ",'" , rRr(i).Quelle , "','" ,  _
   rRr(i).Bemerkung , "'," , rRr(i).absPos , "," , datformk(rRr(i).AktZeit) , "," , rRr(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rRr) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rRr) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(2)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rRr),i)
  If Len(rRr(k).RR) > maxi(0) Then maxi(0) = Len(rRr(k).RR)
  If Len(rRr(k).Quelle) > maxi(1) Then maxi(1) = Len(rRr(k).Quelle)
  If Len(rRr(k).Bemerkung) > maxi(2) Then maxi(2) = Len(rRr(k).Bemerkung)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rRr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & Left$(rRr(k).RR, maxL)  & "'",true : rRr(k).RR = Left$(rRr(k).RR, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.Quelle: '" & rRr(k).Quelle & "' -> '" & Left$(rRr(k).Quelle, maxL)  & "'",true : rRr(k).Quelle = Left$(rRr(k).Quelle, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRr.Bemerkung: '" & rRr(k).Bemerkung & "' -> '" & Left$(rRr(k).Bemerkung, maxL)  & "'",true : rRr(k).Bemerkung = Left$(rRr(k).Bemerkung, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere kvnrue"
 if not Allepat then
  sql = "select pat_id from `kvnrue` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `kvnrue` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        values"))
 For i = 1 to ubound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKv(i).Pat_ID , ",'" , rKv(i).KVNr , "'," , rKv(i).absPos , "," , datformk(rKv(i).AktZeit) , "," , rKv(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rKv) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rKv) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(0)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rKv),i)
  If Len(rKv(k).KVNr) > maxi(0) Then maxi(0) = Len(rKv(k).KVNr)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rKv), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & Left$(rKv(k).KVNr, maxL)  & "'",true : rKv(k).KVNr = Left$(rKv(k).KVNr, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 syscmd 4, pid & ": Speichere unbekannte_kennungen"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt)         values"))
 For i = rUn1 + 1 to ubound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rUn(i).Kennung , "'," , rUn(i).absPos , "," , rUn(i).StByte , "," , rUn(i).Pat_id , ",'" , rUn(i).Inhalt , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rUn) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rUn) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 DoEvents
 Next i
 rUn1 = ubound(rUn)
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
 redim maxi(1)
 for k = iif(Lese.dlg.SammelInsert<>0,rUn1 + 1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rUn),i)
  If Len(rUn(k).Kennung) > maxi(0) Then maxi(0) = Len(rUn(k).Kennung)
  If Len(rUn(k).Inhalt) > maxi(1) Then maxi(1) = Len(rUn(k).Inhalt)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,rUn1 + 1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rUn), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & Left$(rUn(k).Kennung, maxL)  & "'",true : rUn(k).Kennung = Left$(rUn(k).Kennung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & Left$(rUn(k).Inhalt, maxL)  & "'",true : rUn(k).Inhalt = Left$(rUn(k).Inhalt, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function dmpreiheSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere dmpreihe"
 if not Allepat then
  sql = "select pat_id from `dmpreihe` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `dmpreihe` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid)        values"))
 For i = 1 to ubound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rDm(i).Abk , "','" , rDm(i).Art , "'," , datformk(rDm(i).KarteiDatum) , "," , datformk(rDm(i).exportiert) , "," , datformk(rDm(i).DokuDatum) , "," , cstr(-( _
   rDm(i).obvoll<>0)) , "," , cstr(-(rDm(i).ok<>0)) , "," , cstr(-(rDm(i).ausgedruckt<>0)) , ",'" , rDm(i).NachName , "','" ,  _
   rDm(i).VorName , "'," , datformk(rDm(i).GebDat) , "," , rDm(i).Pat_id , "," , rDm(i).StByte , "," , datformk(rDm(i).AktZeit) , "," ,  _
   rDm(i).lanrid , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rDm) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rDm) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(3)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rDm),i)
  If Len(rDm(k).Abk) > maxi(0) Then maxi(0) = Len(rDm(k).Abk)
  If Len(rDm(k).Art) > maxi(1) Then maxi(1) = Len(rDm(k).Art)
  If Len(rDm(k).NachName) > maxi(2) Then maxi(2) = Len(rDm(k).NachName)
  If Len(rDm(k).VorName) > maxi(3) Then maxi(3) = Len(rDm(k).VorName)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rDm), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL)  & "'",true : rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).Art & "' -> '" & Left$(rDm(k).Art, maxL)  & "'",true : rDm(k).Art = Left$(rDm(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).NachName & "' -> '" & Left$(rDm(k).NachName, maxL)  & "'",true : rDm(k).NachName = Left$(rDm(k).NachName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).VorName & "' -> '" & Left$(rDm(k).VorName, maxL)  & "'",true : rDm(k).VorName = Left$(rDm(k).VorName, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function therartenSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere therarten"
 if not Allepat then
  sql = "select pat_id from `therarten` where Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "delete from `therarten` where Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `therarten` (pat_id,zp,mpnr," & _
     "therart,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `therarten` (pat_id,zp,mpnr," & _
     "therart,absPos,AktZeit,StByte)       values"))
 For i = 1 to ubound(rTh)
'  rTh(i).AktZeit = now()
  rTh(i).StByte = CStr(AktByte)
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rTh(i).pat_id , "," , datformk(rTh(i).zp) , "," , rTh(i).mpnr , ",'" , rTh(i).therart , "'," , rTh(i).absPos , "," , datformk( _
   rTh(i).AktZeit) , "," , rTh(i).StByte , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rTh) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rTh) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(0)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rTh),i)
  If Len(rTh(k).therart) > maxi(0) Then maxi(0) = Len(rTh(k).therart)
 next k
 DBCn.CommitTrans
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "therarten", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "therarten", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rTh), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rTh.therart: '" & rTh(k).therart & "' -> '" & Left$(rTh(k).therart, maxL)  & "'",true : rTh(k).therart = Left$(rTh(k).therart, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in therartenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' therartenSpeichern

Public Function laborxsaetzeSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             values"))
 For i = 0 to ubound(rLs)
'  rLs(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLs(i).DatID , ",'" , rLs(i).Satzart , "','" , rLs(i).Satzlänge , "','" , rLs(i).SatzlängeSchluss , "','" , rLs(i).VersionSatzb , "','" ,  _
   rLs(i).Arztnr , "','" , rLs(i).Arztname , "','" , rLs(i).StraßePraxis , "','" , rLs(i).Arzt , "','" , rLs(i).LANR , "','" , rLs(i).PLZPraxis , "','" ,  _
   rLs(i).OrtPraxis , "','" , rLs(i).Labor , "','" , rLs(i).StraßeLabor , "','" , rLs(i).PLZLabor , "','" , rLs(i).OrtLabor , "','" ,  _
   rLs(i).KBVPrüfnr , "','" , rLs(i).Zeichensatz , "','" , rLs(i).Kundenarztnr , "','" , rLs(i).Erstellungsdatum , "','" , rLs(i).Gesamtlänge , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLs) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLs) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(19)
 for k = iif(Lese.dlg.SammelInsert<>0,0,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLs),i)
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
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,0, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzart: '" & rLs(k).Satzart & "' -> '" & Left$(rLs(k).Satzart, maxL)  & "'",true : rLs(k).Satzart = Left$(rLs(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzlänge: '" & rLs(k).Satzlänge & "' -> '" & Left$(rLs(k).Satzlänge, maxL)  & "'",true : rLs(k).Satzlänge = Left$(rLs(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLs.SatzlängeSchluss: '" & rLs(k).SatzlängeSchluss & "' -> '" & Left$(rLs(k).SatzlängeSchluss, maxL)  & "'",true : rLs(k).SatzlängeSchluss = Left$(rLs(k).SatzlängeSchluss, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLs.VersionSatzb: '" & rLs(k).VersionSatzb & "' -> '" & Left$(rLs(k).VersionSatzb, maxL)  & "'",true : rLs(k).VersionSatzb = Left$(rLs(k).VersionSatzb, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztnr: '" & rLs(k).Arztnr & "' -> '" & Left$(rLs(k).Arztnr, maxL)  & "'",true : rLs(k).Arztnr = Left$(rLs(k).Arztnr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztname: '" & rLs(k).Arztname & "' -> '" & Left$(rLs(k).Arztname, maxL)  & "'",true : rLs(k).Arztname = Left$(rLs(k).Arztname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßePraxis: '" & rLs(k).StraßePraxis & "' -> '" & Left$(rLs(k).StraßePraxis, maxL)  & "'",true : rLs(k).StraßePraxis = Left$(rLs(k).StraßePraxis, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arzt: '" & rLs(k).Arzt & "' -> '" & Left$(rLs(k).Arzt, maxL)  & "'",true : rLs(k).Arzt = Left$(rLs(k).Arzt, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLs.LANR: '" & rLs(k).LANR & "' -> '" & Left$(rLs(k).LANR, maxL)  & "'",true : rLs(k).LANR = Left$(rLs(k).LANR, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZPraxis: '" & rLs(k).PLZPraxis & "' -> '" & Left$(rLs(k).PLZPraxis, maxL)  & "'",true : rLs(k).PLZPraxis = Left$(rLs(k).PLZPraxis, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtPraxis: '" & rLs(k).OrtPraxis & "' -> '" & Left$(rLs(k).OrtPraxis, maxL)  & "'",true : rLs(k).OrtPraxis = Left$(rLs(k).OrtPraxis, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLs.Labor: '" & rLs(k).Labor & "' -> '" & Left$(rLs(k).Labor, maxL)  & "'",true : rLs(k).Labor = Left$(rLs(k).Labor, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßeLabor: '" & rLs(k).StraßeLabor & "' -> '" & Left$(rLs(k).StraßeLabor, maxL)  & "'",true : rLs(k).StraßeLabor = Left$(rLs(k).StraßeLabor, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZLabor: '" & rLs(k).PLZLabor & "' -> '" & Left$(rLs(k).PLZLabor, maxL)  & "'",true : rLs(k).PLZLabor = Left$(rLs(k).PLZLabor, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtLabor: '" & rLs(k).OrtLabor & "' -> '" & Left$(rLs(k).OrtLabor, maxL)  & "'",true : rLs(k).OrtLabor = Left$(rLs(k).OrtLabor, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLs.KBVPrüfnr: '" & rLs(k).KBVPrüfnr & "' -> '" & Left$(rLs(k).KBVPrüfnr, maxL)  & "'",true : rLs(k).KBVPrüfnr = Left$(rLs(k).KBVPrüfnr, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLs.Zeichensatz: '" & rLs(k).Zeichensatz & "' -> '" & Left$(rLs(k).Zeichensatz, maxL)  & "'",true : rLs(k).Zeichensatz = Left$(rLs(k).Zeichensatz, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLs.Kundenarztnr: '" & rLs(k).Kundenarztnr & "' -> '" & Left$(rLs(k).Kundenarztnr, maxL)  & "'",true : rLs(k).Kundenarztnr = Left$(rLs(k).Kundenarztnr, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLs.Erstellungsdatum: '" & rLs(k).Erstellungsdatum & "' -> '" & Left$(rLs(k).Erstellungsdatum, maxL)  & "'",true : rLs(k).Erstellungsdatum = Left$(rLs(k).Erstellungsdatum, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLs.Gesamtlänge: '" & rLs(k).Gesamtlänge & "' -> '" & Left$(rLs(k).Gesamtlänge, maxL)  & "'",true : rLs(k).Gesamtlänge = Left$(rLs(k).Gesamtlänge, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  values"))
 For i = 1 to ubound(rLg)
'  rLg(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rLg(i).Pfad , "','" , rLg(i).Name , "'," , datformk(rLg(i).Zp) , "," , cstr(-(rLg(i).fertig<>0)) , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLg) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLg) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(1)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLg),i)
  If Len(rLg(k).Pfad) > maxi(0) Then maxi(0) = Len(rLg(k).Pfad)
  If Len(rLg(k).Name) > maxi(1) Then maxi(1) = Len(rLg(k).Name)
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLg), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL)  & "'",true : rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).Name & "' -> '" & Left$(rLg(k).Name, maxL)  & "'",true : rLg(k).Name = Left$(rLg(k).Name, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
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
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxus"
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
 For i = j to j
'  rLu(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLu(i).DatID , "," , rLu(i).SatzID , ",'" , rLu(i).Satzart , "','" , rLu(i).Satzlänge , "','" , rLu(i).Auftragsnummer , "','" ,  _
   rLu(i).Auftragsschlüssel , "'," , datformk(rLu(i).Eingang) , ",'" , rLu(i).Berichtsdatum , "'," , rLu(i).Pat_id , ",'" , rLu(i).Nachname , "','" ,  _
   rLu(i).Vorname , "','" , rLu(i).GebDat , "','" , rLu(i).Titel , "','" , rLu(i).NVorsatz , "','" , rLu(i).BefArt , "','" , rLu(i).Abrechnungstyp , "','" ,  _
   rLu(i).GebüOrd , "','" , rLu(i).Auftraggeber , "','" , rLu(i).Patienteninformation , "','" , rLu(i).Geschlecht , "','" , rLu(i).AuftrHinw , "','" ,  _
   rLu(i).Pat_idUrsp , "','" , rLu(i).Pat_idErwVNG , "','" , rLu(i).Pat_idErwVN , "','" , rLu(i).Pat_idErwG , "','" , rLu(i).Pat_idErwGB , "','" ,  _
   rLu(i).Pat_idErwGL , "','" , rLu(i).Pat_idLaborNeu , "'," , datformk(rLu(i).ZeitpunktLaborneu) , "," , rLu(i).ZdüP , "," ,  _
   rLu(i).ZdiP , ",'" , rLu(i).LWerte , "'," , datformk(rLu(i).verglichen) , "," , rLu(i).AfN , ")")
  If Lese.dlg.SammelInsert <> 0 And i < j Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = j Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(24)
 for k = iif(Lese.dlg.SammelInsert<>0,j,i) to iif(Lese.dlg.SammelInsert<>0,j,i)
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
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,j, i) To IIf(Lese.dlg.SammelInsert <> 0,j, i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzart: '" & rLu(k).Satzart & "' -> '" & Left$(rLu(k).Satzart, maxL)  & "'",true : rLu(k).Satzart = Left$(rLu(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzlänge: '" & rLu(k).Satzlänge & "' -> '" & Left$(rLu(k).Satzlänge, maxL)  & "'",true : rLu(k).Satzlänge = Left$(rLu(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsnummer: '" & rLu(k).Auftragsnummer & "' -> '" & Left$(rLu(k).Auftragsnummer, maxL)  & "'",true : rLu(k).Auftragsnummer = Left$(rLu(k).Auftragsnummer, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsschlüssel: '" & rLu(k).Auftragsschlüssel & "' -> '" & Left$(rLu(k).Auftragsschlüssel, maxL)  & "'",true : rLu(k).Auftragsschlüssel = Left$(rLu(k).Auftragsschlüssel, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLu.Berichtsdatum: '" & rLu(k).Berichtsdatum & "' -> '" & Left$(rLu(k).Berichtsdatum, maxL)  & "'",true : rLu(k).Berichtsdatum = Left$(rLu(k).Berichtsdatum, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLu.Nachname: '" & rLu(k).Nachname & "' -> '" & Left$(rLu(k).Nachname, maxL)  & "'",true : rLu(k).Nachname = Left$(rLu(k).Nachname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLu.Vorname: '" & rLu(k).Vorname & "' -> '" & Left$(rLu(k).Vorname, maxL)  & "'",true : rLu(k).Vorname = Left$(rLu(k).Vorname, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebDat: '" & rLu(k).GebDat & "' -> '" & Left$(rLu(k).GebDat, maxL)  & "'",true : rLu(k).GebDat = Left$(rLu(k).GebDat, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLu.Titel: '" & rLu(k).Titel & "' -> '" & Left$(rLu(k).Titel, maxL)  & "'",true : rLu(k).Titel = Left$(rLu(k).Titel, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLu.NVorsatz: '" & rLu(k).NVorsatz & "' -> '" & Left$(rLu(k).NVorsatz, maxL)  & "'",true : rLu(k).NVorsatz = Left$(rLu(k).NVorsatz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLu.BefArt: '" & rLu(k).BefArt & "' -> '" & Left$(rLu(k).BefArt, maxL)  & "'",true : rLu(k).BefArt = Left$(rLu(k).BefArt, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLu.Abrechnungstyp: '" & rLu(k).Abrechnungstyp & "' -> '" & Left$(rLu(k).Abrechnungstyp, maxL)  & "'",true : rLu(k).Abrechnungstyp = Left$(rLu(k).Abrechnungstyp, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebüOrd: '" & rLu(k).GebüOrd & "' -> '" & Left$(rLu(k).GebüOrd, maxL)  & "'",true : rLu(k).GebüOrd = Left$(rLu(k).GebüOrd, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftraggeber: '" & rLu(k).Auftraggeber & "' -> '" & Left$(rLu(k).Auftraggeber, maxL)  & "'",true : rLu(k).Auftraggeber = Left$(rLu(k).Auftraggeber, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLu.Patienteninformation: '" & rLu(k).Patienteninformation & "' -> '" & Left$(rLu(k).Patienteninformation, maxL)  & "'",true : rLu(k).Patienteninformation = Left$(rLu(k).Patienteninformation, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLu.Geschlecht: '" & rLu(k).Geschlecht & "' -> '" & Left$(rLu(k).Geschlecht, maxL)  & "'",true : rLu(k).Geschlecht = Left$(rLu(k).Geschlecht, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLu.AuftrHinw: '" & rLu(k).AuftrHinw & "' -> '" & Left$(rLu(k).AuftrHinw, maxL)  & "'",true : rLu(k).AuftrHinw = Left$(rLu(k).AuftrHinw, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idUrsp: '" & rLu(k).Pat_idUrsp & "' -> '" & Left$(rLu(k).Pat_idUrsp, maxL)  & "'",true : rLu(k).Pat_idUrsp = Left$(rLu(k).Pat_idUrsp, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVNG: '" & rLu(k).Pat_idErwVNG & "' -> '" & Left$(rLu(k).Pat_idErwVNG, maxL)  & "'",true : rLu(k).Pat_idErwVNG = Left$(rLu(k).Pat_idErwVNG, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVN: '" & rLu(k).Pat_idErwVN & "' -> '" & Left$(rLu(k).Pat_idErwVN, maxL)  & "'",true : rLu(k).Pat_idErwVN = Left$(rLu(k).Pat_idErwVN, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwG: '" & rLu(k).Pat_idErwG & "' -> '" & Left$(rLu(k).Pat_idErwG, maxL)  & "'",true : rLu(k).Pat_idErwG = Left$(rLu(k).Pat_idErwG, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGB: '" & rLu(k).Pat_idErwGB & "' -> '" & Left$(rLu(k).Pat_idErwGB, maxL)  & "'",true : rLu(k).Pat_idErwGB = Left$(rLu(k).Pat_idErwGB, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGL: '" & rLu(k).Pat_idErwGL & "' -> '" & Left$(rLu(k).Pat_idErwGL, maxL)  & "'",true : rLu(k).Pat_idErwGL = Left$(rLu(k).Pat_idErwGL, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idLaborNeu: '" & rLu(k).Pat_idLaborNeu & "' -> '" & Left$(rLu(k).Pat_idLaborNeu, maxL)  & "'",true : rLu(k).Pat_idLaborNeu = Left$(rLu(k).Pat_idLaborNeu, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLu.LWerte: '" & rLu(k).LWerte & "' -> '" & Left$(rLu(k).LWerte, maxL)  & "'",true : rLu(k).LWerte = Left$(rLu(k).LWerte, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  values"))
 For i = 1 to ubound(rLo)
'  rLo(i).AktZeit = now()
    Dim j&
    For j = 1 To i - 1
    If rLo(i).RefNr <> rLo(j).RefNr Then GoTo nextj
    If rLo(i).Verf <> rLo(j).Verf Then GoTo nextj
    If rLo(i).KuQu <> rLo(j).KuQu Then GoTo nextj
    If rLo(i).Quelle <> rLo(j).Quelle Then GoTo nextj
    If rLo(i).QSpez <> rLo(j).QSpez Then GoTo nextj
    If rLo(i).AbnDat <> rLo(j).AbnDat Then GoTo nextj
    If rLo(i).Kommentar <> rLo(j).Kommentar Then GoTo nextj
    If rLo(i).Erklärung <> rLo(j).Erklärung Then GoTo nextj
    If rLo(i).Keimzahl <> rLo(j).Keimzahl Then GoTo nextj
     GoTo nexti
nextj:
     Dim rsdop As New ADODB.Recordset
     Set rsdop = Nothing
     rsdop.Open "select 0 from `laborxbakt` where `RefNr` = " & rLo(i).RefNr & " and `Verf` = '" & rLo(i).Verf & "' and `KuQu` = '" & rLo(i).KuQu & "' and `Quelle` = '" &  _
   rLo(i).Quelle & "' and `QSpez` = '" & rLo(i).QSpez & "' and `AbnDat` = " & datformk(rLo(i).AbnDat) & " and `Kommentar` = '" &  _
   rLo(i).Kommentar & "' and `Erklärung` = '" & rLo(i).Erklärung & "' and `Keimzahl` = '" & rLo(i).Keimzahl & "'", DBCn, adOpenStatic, adLockReadOnly
     If Not rsdop.EOF Then GoTo nexti
    Next j
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLo(i).RefNr , ",'" , rLo(i).Verf , "','" , rLo(i).KuQu , "','" , rLo(i).Quelle , "','" , rLo(i).QSpez , "'," , datformk(rLo(i).AbnDat) , ",'" ,  _
   rLo(i).Kommentar , "','" , rLo(i).Erklärung , "','" , rLo(i).Keimzahl , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLo) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLo) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(6)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLo),i)
  If Len(rLo(k).Verf) > maxi(0) Then maxi(0) = Len(rLo(k).Verf)
  If Len(rLo(k).KuQu) > maxi(1) Then maxi(1) = Len(rLo(k).KuQu)
  If Len(rLo(k).Quelle) > maxi(2) Then maxi(2) = Len(rLo(k).Quelle)
  If Len(rLo(k).QSpez) > maxi(3) Then maxi(3) = Len(rLo(k).QSpez)
  If Len(rLo(k).Kommentar) > maxi(4) Then maxi(4) = Len(rLo(k).Kommentar)
  If Len(rLo(k).Erklärung) > maxi(5) Then maxi(5) = Len(rLo(k).Erklärung)
  If Len(rLo(k).Keimzahl) > maxi(6) Then maxi(6) = Len(rLo(k).Keimzahl)
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & Left$(rLo(k).Verf, maxL)  & "'",true : rLo(k).Verf = Left$(rLo(k).Verf, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & Left$(rLo(k).KuQu, maxL)  & "'",true : rLo(k).KuQu = Left$(rLo(k).KuQu, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & Left$(rLo(k).Quelle, maxL)  & "'",true : rLo(k).Quelle = Left$(rLo(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & Left$(rLo(k).QSpez, maxL)  & "'",true : rLo(k).QSpez = Left$(rLo(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & Left$(rLo(k).Kommentar, maxL)  & "'",true : rLo(k).Kommentar = Left$(rLo(k).Kommentar, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & Left$(rLo(k).Erklärung, maxL)  & "'",true : rLo(k).Erklärung = Left$(rLo(k).Erklärung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & Left$(rLo(k).Keimzahl, maxL)  & "'",true : rLo(k).Keimzahl = Left$(rLo(k).Keimzahl, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    values"))
 For i = 1 to ubound(rLw)
'  rLw(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLw(i).RefNr , ",'" , rLw(i).Abkü , "','" , rLw(i).Langname , "','" , rLw(i).Quelle , "','" , rLw(i).QSpez , "'," , datformk( _
   rLw(i).AbnDat) , ",'" , rLw(i).Wert , "','" , rLw(i).Einheit , "','" , rLw(i).Grenzwerti , "','" , rLw(i).Kommentar , "','" , rLw(i).Teststatus , "','" ,  _
   rLw(i).Erklärung , "','" , rLw(i).AuftrHinw , "'," , rLw(i).nbid , ")")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLw) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLw) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(10)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLw),i)
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
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLw), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLw.Abkü: '" & rLw(k).Abkü & "' -> '" & Left$(rLw(k).Abkü, maxL)  & "'",true : rLw(k).Abkü = Left$(rLw(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLw.Langname: '" & rLw(k).Langname & "' -> '" & Left$(rLw(k).Langname, maxL)  & "'",true : rLw(k).Langname = Left$(rLw(k).Langname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLw.Quelle: '" & rLw(k).Quelle & "' -> '" & Left$(rLw(k).Quelle, maxL)  & "'",true : rLw(k).Quelle = Left$(rLw(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLw.QSpez: '" & rLw(k).QSpez & "' -> '" & Left$(rLw(k).QSpez, maxL)  & "'",true : rLw(k).QSpez = Left$(rLw(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLw.Wert: '" & rLw(k).Wert & "' -> '" & Left$(rLw(k).Wert, maxL)  & "'",true : rLw(k).Wert = Left$(rLw(k).Wert, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLw.Einheit: '" & rLw(k).Einheit & "' -> '" & Left$(rLw(k).Einheit, maxL)  & "'",true : rLw(k).Einheit = Left$(rLw(k).Einheit, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLw.Grenzwerti: '" & rLw(k).Grenzwerti & "' -> '" & Left$(rLw(k).Grenzwerti, maxL)  & "'",true : rLw(k).Grenzwerti = Left$(rLw(k).Grenzwerti, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLw.Kommentar: '" & rLw(k).Kommentar & "' -> '" & Left$(rLw(k).Kommentar, maxL)  & "'",true : rLw(k).Kommentar = Left$(rLw(k).Kommentar, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLw.Teststatus: '" & rLw(k).Teststatus & "' -> '" & Left$(rLw(k).Teststatus, maxL)  & "'",true : rLw(k).Teststatus = Left$(rLw(k).Teststatus, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLw.Erklärung: '" & rLw(k).Erklärung & "' -> '" & Left$(rLw(k).Erklärung, maxL)  & "'",true : rLw(k).Erklärung = Left$(rLw(k).Erklärung, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLw.AuftrHinw: '" & rLw(k).AuftrHinw & "' -> '" & Left$(rLw(k).AuftrHinw, maxL)  & "'",true : rLw(k).AuftrHinw = Left$(rLw(k).AuftrHinw, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern
 Dim i%, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs as new ADODB.recordset
 T1 = Timer
 on error resume next
 Pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   values"))
 For i = 1 to ubound(rLL)
'  rLL(i).AktZeit = now()
  If Lese.dlg.SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLL(i).RefNr , ",'" , rLL(i).Abkü , "','" , rLL(i).Verf , "','" , rLL(i).EBM , "','" , rLL(i).goä , "','" , rLL(i).Anzahl , "','" ,  _
   rLL(i).abrd , "')")
  If Lese.dlg.SammelInsert <> 0 And i < ubound(rLL) Then csql.Append ","
  If Lese.dlg.SammelInsert = 0 Or i = ubound(rLL) Then
  Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
  If obfor Then
   Call ForeignYes0
   Call ForeignYes1
  End If
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
 redim maxi(5)
 for k = iif(Lese.dlg.SammelInsert<>0,1,i) to iif(Lese.dlg.SammelInsert<>0,ubound(rLL),i)
  If Len(rLL(k).Abkü) > maxi(0) Then maxi(0) = Len(rLL(k).Abkü)
  If Len(rLL(k).Verf) > maxi(1) Then maxi(1) = Len(rLL(k).Verf)
  If Len(rLL(k).EBM) > maxi(2) Then maxi(2) = Len(rLL(k).EBM)
  If Len(rLL(k).goä) > maxi(3) Then maxi(3) = Len(rLL(k).goä)
  If Len(rLL(k).Anzahl) > maxi(4) Then maxi(4) = Len(rLL(k).Anzahl)
  If Len(rLL(k).abrd) > maxi(5) Then maxi(5) = Len(rLL(k).abrd)
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    If maxL > 0 Then
     For k = IIf(Lese.dlg.SammelInsert <> 0,1, i) To IIf(Lese.dlg.SammelInsert <> 0,ubound(rLL), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & Left$(rLL(k).Abkü, maxL)  & "'",true : rLL(k).Abkü = Left$(rLL(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & Left$(rLL(k).Verf, maxL)  & "'",true : rLL(k).Verf = Left$(rLL(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & Left$(rLL(k).EBM, maxL)  & "'",true : rLL(k).EBM = Left$(rLL(k).EBM, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & Left$(rLL(k).goä, maxL)  & "'",true : rLL(k).goä = Left$(rLL(k).goä, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & Left$(rLL(k).Anzahl, maxL)  & "'",true : rLL(k).Anzahl = Left$(rLL(k).Anzahl, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.abrd: '" & rLL(k).abrd & "' -> '" & Left$(rLL(k).abrd, maxL)  & "'",true : rLL(k).abrd = Left$(rLL(k).abrd, maxL)
      End Select
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End Select
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lies.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
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
 Dim rAf&
 On Error GoTo fehler
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
 call therartenSpeichern
 Call DBCn.Execute("update `namen` set aktZeit = " & DatFormK(rNa(0).AktZeit) & " where pat_id = " & rNa(0).Pat_ID,rAf)
 If rAf <> 1 Then 
  frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_ID & " " & rNa(0).Nachname & " " & rNa(0).Vorname, true
 end if
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 Select Case MsgBox("FNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' doSpeichern

