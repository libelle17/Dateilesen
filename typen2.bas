Attribute VB_Name = "Typen"
Option Explicit
Public obForK%
Dim sql$, T1!, T2!, maxL%

Public Type namen
 Pat_id As Long 'Pat_ID int '3000
 lfdnr As Long 'lfdnr int 'laufende Patientennummer
 NVorsatz As String 'NVorsatz varchar '3100
 Nachname As String 'Nachname varchar '3101
 Vorname As String 'Vorname varchar '3102
 GebDat As Date 'GebDat datetime '3103
 f3004 As String 'f3004 varchar '3004 unbek.Feld, bis jetzt 0 oder 2
 f3006 As String 'f3006 varchar '3006 unbek.Feld, bis jetzt 5.1.0 oder 5.2.0
 Straße As String 'Straße varchar '3107
 KVKStatus As String 'KVKStatus varchar '3108
 Hausnr As String 'Hausnr varchar '3109
 geschlecht As String 'Geschlecht varchar '3110
 plz As String 'Plz varchar '3112
 ort As String 'Ort varchar '3113
 Lkz As String 'Lkz varchar '3114 Länderkennzeichen
 Anschrzus As String 'Anschrzus varchar '3115 Anschriftenzusatz
 NVors As String 'NVors varchar '3120 Namensvorsatz
 PFPlz As String 'PFPlz varchar '3121 Postfach-Postleitzahl
 PFOrt As String 'PFOrt varchar '3122 Postfach-Ort
 PFNr As String 'PFNr varchar '3123 Postfach-Nr.
 f3124 As String 'f3124 varchar '3124 unbek.Feld, bisher immer leer
 AnschrZus_2 As String 'AnschrZus_2 varchar '3215 Anschriftenzusatz, aufgeteiltes Feld
 Postfach_2 As String 'Postfach_2 varchar '3216 Postfach, aufgeteiltes Feld
 LK_2 As String 'LK_2 varchar '3217 Ländercode vermutlich Herkunftsland, aufgeteiltes Feld
 Postfach As String 'Postfach varchar '3216
 Beruf As String 'Beruf varchar '3620 Beruf
 Weggeldzone As String 'Weggeldzone varchar '3631 (1) Weggeldzone mit Z
 WeggzZahl As Double 'WeggzZahl decimal '3631 (2) Weggeldzone, Zahl in Feld 2
 AufnDat As Date 'AufnDat datetime '3610
 kAufDat As Date 'kAufDat date 'korrigiertes Aufnahmedatum: if(min(fanf)>min(bhfb),(min(fanf),min(bhfb))
 Lanr As String 'LANR varchar '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr As String 'BStNr varchar '3536 Betriebsstättennummer
 Titel As String 'Titel varchar '3104
 Versichertennummer As String 'Versichertennummer varchar '3105
 PrivatTel As String 'PrivatTel varchar '3629
 KVNr As String 'KVNr varchar '3630 Hausarzt
 KVNr2 As String 'KVNr2 varchar '3630 Hausarzt (2.Eintrag)
 KVNr3 As String 'KVNr3 varchar '3630 Hausarzt (3.Eintrag)
 KVNr4 As String 'KVNr4 varchar '3630 Hausarzt (4.Eintrag)
 PrivatTel_2 As String 'PrivatTel_2 varchar '3629
 PrivatFax As String 'PrivatFax varchar '3629
 DienstTel As String 'DienstTel varchar '3629
 PrivatMobil As String 'PrivatMobil varchar '3629
 email As String 'Email varchar 'Email
 Arbeitgeber As String 'Arbeitgeber varchar '3625
 AnAllgda As Integer 'AnAllgda bit 'Anamnese allgemein da
 An1da As Integer 'An1da bit 'Anamnese S.1 da
 An2da As Integer 'An2da bit 'Anamnese S.2 da
 Checkda As Integer 'Checkda bit 'Checkliste da
 DMTypaD As String 'DMTypaD varchar 'aus Diagnosen
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 StByteA As Long 'StByteA int 'Ordnungsnummer der Datenübertragung, Beginn der Übertragung
 Cave As String 'Cave varchar '3654
 Notiz As String 'Notiz varchar '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>, DMP ausgeschrieben <datum>
 obChk As String 'obChk varchar 'obChroniker (Feld 3800)
 dmpklass As Long 'dmpklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier, 4 = DMP ausgeschrieben
 dmpbeg As Date 'dmpbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass As Long 'dmpkhkklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg As Date 'dmpkhkbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass As Long 'dmpcopdklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg As Date 'dmpcopdbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpabklass As Long 'dmpabklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpabbeg As Date 'dmpabbeg date 'Datum der aktuellen DMP-Klassifikation
 dakab As Date 'dakab date 'DAK-Einverständnis-Datum
 HzV As Long 'HzV int '1 = HzV-Eintrag im Notizfeld
 HzVbeg As Date 'HzVbeg date 'HzV ab
 DS As Long 'DS int '1 = Datenschutzerklärung laut Notizfeld da
 DSbeg As Date 'DSbeg date 'DS ab
 getHA0 As Long 'getHA0 int 'KVNr aus getHausarzt -> Üw(12,0)
 fnHA0 As String 'fnHA0 varchar 'Funktion aus getHausarzt -> Üw(10,0)
 getHA1 As Long 'getHA1 int 'KVNr aus getHausarzt -> Üw(12,1)
 fnHA1 As String 'fnHA1 varchar 'Funktion aus getHausarzt -> Üw(10,1)
 getHA2 As Long 'getHA2 int 'KVNr aus getHausarzt -> Üw(12,2)
 fnHA2 As String 'fnHA2 varchar 'Funktion aus getHausarzt -> Üw(10,2)
 zubenach As String 'zubenach varchar '3633
 Verwandt As String 'Verwandt varchar '3632
 Sprache As String 'Sprache varchar '3628
 lAktTM As Date 'lAktTM datetime 'letzte Aktualisierung in Turbomed
 Mitarbeiter As Long 'Mitarbeiter int 'ob Pat. Mitarbeiter ist
End Type

Public Type Faelle
 FID As Long 'FID int '
 Pat_id As Long 'Pat_ID int '3000 Pat_id
 Quartal As String 'Quartal varchar '4101 Quartal der Ausstellung
 Nachname As String 'Nachname varchar '3101 Nachname
 Vorname As String 'Vorname varchar '3102 Vorname
 lfdnr As Long 'lfdnr int 'laufende Fallnummer
 TMFNr As String 'TMFNr varchar '4144 Fallnummer in Turbomed
 VKNr As String 'VKNr varchar '4104 VK-Nr.
 f4131 As String 'f4131 varchar '4131 unbekanntes  Feld, bis jetzt immer 0
 f4132 As String 'f4132 varchar '4132 unbekanntes  Feld, 0-6
 VschBeg As Date 'VschBeg datetime '4133 Versichternschutzbeginn
 KKasse_2 As String 'KKasse_2 varchar '4134 Krankenkasse aus Fall
 FaktPers As Single 'FaktPers float '4136 Faktor persönlich
 FaktTechn As Single 'FaktTechn float '4137 Faktor technisch
 FaktLabor As Single 'FaktLabor float '4138 Faktor Labor
 BhFB As Date 'BhFB datetime '4150 Behandlungsfall: Beginn
 BhFE1 As Date 'BhFE1 datetime '4151 Behandlungsfall: Ende (Musterwoman) / Wohl Ende des Behandlungsfallbeginnquartals
 BhFE2 As Date 'BhFE2 datetime '4152 Behandlungsfall: Ende (Musterwoman),bei offenem Fall 00000000, sonst z.B. 30092006 für 3/06
 f4202 As String 'f4202 varchar '4202 Unfall, Unfallfolgen nach 4152
 ausgst As Date 'ausgst datetime '4102 ('ausgestellt am')
 KtrAbrB As String 'KtrAbrB varchar '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr As String 'AbrAr varchar '4107, Abrechnungsart (1 = Primärkassen)
 lVorl As Date 'lVorl datetime '4109, letzte Vorlage
 IK As String 'IK varchar '4111 Krankenkassennummer (IK)
 KVKs As String 'KVKs varchar '4112 Versichertenstatus VK
 KVKserg As String 'KVKserg varchar '4113 Ost/West-Status VK
 Status As String 'Status varchar 'Feld Status in verschiedenen Formularen
 Kasse As String 'Kasse varchar '6299 Kasse (aus Formularen)
 KID As Long 'KID int 'Bezug auf kassenliste.id
 GebOr As String 'GebOr varchar '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As String 'AbrGb varchar '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis As String 'PersKreis varchar '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz As String 'SKtZusatz varchar '4124 SKT-Zusatzangaben
 letzteRegel As String 'letzteRegel varchar '4206, letzter Tag der Regel
 ÜwText As String 'ÜwText longtext '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 As Byte 'f4210 tinyint '4210, Ankreuzfeld LSR
 AkfHAH As Byte 'AkfHAH tinyint '4211 Ankreuzfeld HAH
 AkfAB0 As Byte 'AkfAB0 tinyint '4212 Ankreuzfeld AB0.RH
 AkfAK As Byte 'AkfAK tinyint '4213 Ankreuzfeld AK
 statNuller As String 'statNuller varchar '4216, nu bei Musterfrau 16 Nuller
 ÜbwV As String 'ÜbwV varchar '4218, überwiesen von Arztnummer
 ÜbWVLANR As String 'ÜbWVLANR varchar '4218(1) überwiesen von LANR
 ÜbWVBSNR As String 'ÜbWVBSNR varchar '4218(2) überwiesen von BSNR
 ÜbWVKVNR As String 'ÜbWVKVNR varchar '4218(3) überwiesen von KVNR
 AndÜw As String 'AndÜw varchar '4219, anderer Überweiser
 Übwr As String 'Übwr varchar 'resultierender Überweiser (BSNR): 4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR As String 'ÜbwLANR varchar '4242 LANR des Überweisers
 ÜWZiel As String 'ÜWZiel varchar '4220 Überweisung an
 ÜWNNr As String 'ÜWNNr varchar '4231(4): KV-Nummer des Überweisers
 ÜWNaN As String 'ÜWNaN varchar '4231(3): Nachname des Überweisers
 ÜWTit As String 'ÜWTit varchar '4231(3): Titel des Überweisers
 ÜWVor As String 'ÜWVor varchar '4231(2): Vorname des Überweisers
 ÜWVsw As String 'ÜWVsw varchar '4231(2b): Vorsatzwort des Überweisers
 üwvid As Long 'üwvid int '4247 Bezug auf ueberwvon
 Auftrag As String 'Auftrag varchar '4205 Auftrag bei Überweisung
 Verdacht As String 'Verdacht varchar '4207 Verdacht bei Überweisung

 Befund As String 'Befund varchar '4208 Befund bei Überweisung
 statKlasse As String 'statKlasse varchar '4236 Klasse bei Behandlung
 f4237 As String 'f4237 varchar '4237 Krankenhausname
 statBehTage As Long 'statBehTage int '4238 Krankenhausaufenthalt
 SchGr As Double 'SchGr decimal '4239, Schein(unter)gruppe
 Weiterbeh As String 'Weiterbeh varchar '4243, Weiterbehandelnder
 f4266 As Single 'f4266 float '4266 unbekanntes Datumsfeld, bisher nur bei Musterwoman
 PGeb As String 'PGeb varchar '4401, Praxisgebühr
 PGebErg As String 'PGebErg varchar '4402, Array
 Mahnfrist As String 'Mahnfrist varchar '4403, Mahnfrist bis
 Unfallort As String 'Unfallort varchar '4505 Unfallort
 BeschAls As String 'BeschAls varchar '4506 Beschäftigt als
 BeschSeit As Date 'BeschSeit datetime '4506 Beschäftigt seit
 Unfallbetrieb As String 'Unfallbetrieb varchar '4509 Unfallbetrieb
 f4570 As String 'f4570 varchar 'f4570 unbekanntes Feld, nur einamal 0
 GOÄKatNr As String 'GOÄKatNr varchar '4580 (1): Katalog-Nummer
 GOÄKatName As String 'GOÄKatName varchar '4580 (2): Privat-Abrechnungskatalog
 abrArzt As String 'abrArzt varchar '4585 abrechnender Arzt
 privVers As String 'privVers varchar '4586 private Versicherung
 AdNam As String 'AdNam varchar '4602(1) Name Rechnungsanschrift
 AdStr As String 'AdStr varchar '4602(2) Straße Rechnungsanschrift
 AdPlz As String 'AdPlz varchar '4602(3) PLZ Rechnungsanschrift
 AdOrt As String 'AdOrt varchar '4602(4) Ort Rechnungsanschrift
 ÜwBG As String 'ÜwBG varchar 'f4603 Überweiser BG
 BhFE As Date 'BhFE datetime '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As String 's8000 varchar '8000, Satzidentifikation
 s8100 As String 's8100 varchar '8100 Satzlänge
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 Fanf As Date 'Fanf datetime 'Fallanfang
 altQuart As String 'altQuart varchar '
 QAnf As Date 'QAnf datetime 'Quartalsanfang
 QEnd As Date 'QEnd datetime 'Quartalsende
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 f4108 As String 'f4108 varchar '4108
 BGFallNr As String 'BGFallNr varchar '3603 BG-Fall-Nummer
 lGewicht As Double 'lGewicht decimal 'letztes Gewicht in kg
 vorET As Date 'vorET date 'voraussichtlicher Entbindungstermin
 dmpVertret As String 'dmpVertret varchar 'DMP: Vertretung
 dmpArztw As String 'dmpArztw varchar 'DMP: Arztwechsel
 dmpHypos As String 'dmpHypos varchar 'DMP: Zahl der schweren Hypos im letzten Quartal
 dmpKhsA As String 'dmpKhsA varchar 'DMP: Zahl der Khs-Aufenthalte wg.Diabetes im letzten Quartal
 dmpDMSchulEmpf As String 'dmpDMSchulEmpf varchar 'DMP: Schulung D.m. empfohlen
 dmpDMSchulWahrg As String 'dmpDMSchulWahrg varchar 'DMP: Schulung D.m. wahrgenommen (bei letzter Doku)
 dmpHypertSchulEmpf As String 'dmpHypertSchulEmpf varchar 'DMP-Schulung zur Hypertonie empfohlen
 dmpHypertSchulWahrg As String 'dmpHypertSchulWahrg varchar 'DMP: Schulung Hypertonie wahrgenommen (bei letzter Doku)
 dmpKKTabakEmpf As String 'dmpKKTabakEmpf varchar 'DMP-Schulung zur Tabakentwöhnung empfohlen
 dmpKKErnEmpf As String 'dmpKKErnEmpf varchar 'DMP-Schulung zur Ernährung empfohlen
 dmpKKkTrainEmpf As String 'dmpKKkTrainEmpf varchar 'DMP-Schulung zum körperlichen Training empfohlen
 dmpHbA1cZiel As String 'dmpHbA1cZiel varchar 'DMP: HbA1c-Ziel erreicht
 dmpUewFuss As String 'dmpUewFuss varchar 'DMP: Überweisung Fußeinrichtung veranlasst
 dmpEinwDM As String 'dmpEinwDM varchar 'DMP: Einweisung wegen D.m. veranlasst
 dmphalbj As String 'dmphalbj varchar 'j=DMP nur halbjährlich dokumentieren
 dmpMA As String 'dmpMA varchar 'DMP: Mitarbeiter, der Makro eingegeben hat
End Type

Public Type au
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Beginn As String 'Beginn varchar '6285 1. Hälfte
 Ende As String 'Ende varchar '6285 2. Hälfte
 ICDs As String 'ICDs varchar '6286
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type briefe
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '
 Zeitpunkt As Date 'ZeitPunkt datetime '
 Pfad As String 'Pfad varchar '
 art As String 'Art varchar '
 name As String 'Name varchar '
 autor As String 'autor varchar 'Autor
 Quelldatum As Date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 Typ As String 'Typ varchar '
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe As Long 'DokGroe int 'Größe der Datei
 DokAenD As Date 'DokAenD datetime 'Dokument-letzte Änderung
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 id As Long 'ID int '
End Type

Public Type Diagnosen
 ID1 As Long 'ID1 int '
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_id int 'Bezug auf Anamneseblattt
 DiagDatum As Date 'DiagDatum datetime '5999 Datum, 6301 Uhrzeit
 DiagSicherheit As String 'DiagSicherheit varchar '3674, 6003 akut
 DiagText As String 'DiagText longtext '3650 Dauer, 6000 akut
 DiagSeite As String 'DiagSeite varchar '3675, 6004 akut
 DiagAttr As String 'DiagAttr varchar '6006 Diagnosenattribut (optionale Erläuterung)
 ICD As String 'ICD varchar '3673, 6001 akut
 obDauer As Byte 'obDauer tinyint 'ob Dauerdiagnose
 intBemerk As String 'intBemerk varchar '6009 interne Bemerkung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 AusnBegr As String 'AusnBegr varchar '6008 Ausnahmebegründung
 f6010 As Byte 'f6010 tinyint 'Falsch=dd-Eintrag, Wahr=bdd-Eintragurspr.: 6010 Diagnose gelöscht
 obKasse As Byte 'obKasse tinyint 'ob nach Kodierrichtlinien an Kasse zu übermitteln
 lKasse As Date 'lKasse datetime 'wann zuletzt Kassenübermittlung nach Kodierrichtlinien eingetragen (bd/bdd, f6010 Wahr)
 f6011 As String 'f6011 varchar '6011 8.12.10: bisher nur """"""""TM#?""""""""
End Type

Public Type dokumente
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '
 Zeitpunkt As Date 'ZeitPunkt datetime '
 DokPfad As String 'DokPfad varchar '
 DokArt As String 'DokArt varchar '
 DokName As String 'DokName varchar '
 Quelldatum As Date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe As Long 'DokGroe int 'Dokument-Größe
 DokAenD As Date 'DokAenD datetime 'Dokument-letzte Änderung
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type eintraege
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 art As String 'Art varchar '6330
 Inhalt As String 'Inhalt varchar '8480
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
 inhNum As Double 'inhNum double 'Inhalt numerisch
End Type

Public Type forminhaltform_abk
 Form_AbkVW As Long 'Form_AbkVW int '
 Form_Abk As String 'Form_Abk varchar '
End Type

Public Type formulare
 FormID As Long 'FormID int '
 Form_Abk As String 'Form_Abk varchar '
 FormBez As String 'FormBez longtext '
 FormVorl As String 'FormVorl varchar '
 AktZeit As Date 'AktZeit datetime 'Zeitpunkt der Aktualisierung
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhkopf
 Foid As Long 'FoID int '
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '
 Form_ID As Long 'Form_ID int '
 Zeitpunkt As Date 'ZeitPunkt datetime '
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 Satzart As String 'Satzart varchar '8000
 Satzlänge As String 'Satzlänge varchar '8100
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
End Type

Public Type forminhfeld
 Foid As Long 'FoID int '
 nr As Integer 'Nr smallint '
 FeldNr As Integer 'FeldNr smallint '
 FeldVW As Long 'FeldVW int '
 Feld As String
 FeldInhVW As Long 'FeldInhVW int '
 FeldInh As String
End Type

Public Type kheinweis
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 Ziel As String 'Ziel varchar '6291
 Diagnose As String 'Diagnose longtext '6230
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type lbanforderungen
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 AnfText As String 'AnfText longtext '6280
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Statusbyte
End Type

Public Type laborneu
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 FertigStGrad As String 'FertigStGrad varchar '8401
 Abkü As String 'Abkü varchar '8410
 LangtextVW As Long 'LangtextVW int '8411
 Langtext As String
 Wert As String 'Wert varchar '8420
 Einheit As String 'Einheit varchar '8421
 KommentarVW As Long 'KommentarVW int '8480
 Kommentar As String
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 RefNr As Long 'Refnr int 'Bezug auf LaborXUS
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 id As Long 'ID int '
End Type

Public Type leistungen
 id As Long 'id int 'eindeutige ID, hinzugefügt 26.3.11
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '5000 + 6201
 Leistung As String 'Leistung varchar '5001 Leistungsziffer
 f5002 As String 'f5002 varchar '5002 Art der Untersuchung
 f5005 As String 'f5005 varchar '5005 Anzahl
 f5006 As String 'f5006 varchar '5006 um Uhrzeit
 f5009 As String 'f5009 varchar '5009 freier Begründungstext
 Med As String 'Med varchar '5010 Medikament
 f5015 As String 'f5015 varchar '5015 Organ
 f5016 As String 'f5016 varchar '5016 Name des Arztes (Briefempfänger)
 f5021 As String 'f5021 varchar '5021 Datum letzte Krebsvorsorge
 f5026 As String 'f5026 varchar '5026 Entlassungsdatum
 Faktor As String 'Faktor varchar '5062 Multiplikator für GOÄ-Rechnung
 f5098 As String 'f5098 varchar '5098 0000000000
 Charge As String 'Charge varchar '9999 Charge
 Lanr As String 'LANR varchar '5099 LANR
 letzVorg As Date 'letzVorg datetime '5101 letzter Vorgang
 Ausn As String 'Ausn varchar '3677 Ausnahme/Begründung für abweichendes Geschlecht
 beme As String 'Beme varchar '         Bemerkung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 Sachkbez As String 'Sachkbez varchar '5011 Sachkostenbezeichnung
 Sachkct As Long 'Sachkct int '5012 Sach-/Materialkosten in ct
 Zone As String 'Zone varchar '5018 Zone bei Besuchen
End Type

Public Type medplan
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 MPNr As Long 'MPNr int 'Ordnungsziffer für Medikamentenplan
 Zeitpunkt As Date 'ZeitPunkt datetime 'Zeitpunkt, der Speicherung im Turbomed
 Datum As Date 'Datum datetime 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament As String 'Medikament varchar '
 MedAnfang As String 'MedAnfang varchar 'Referenz auf medarten.Medikament
 Wirkstoff As String 'Wirkstoff varchar 'Stärke aus BMP
 PZN As Long 'PZN int 'Pharmazentralnummer
 FeldNr As Integer 'FeldNr smallint '
 mo As String 'mo varchar '
 mi As String 'mi varchar '
 nm As String 'nm varchar '
 ab As String 'ab varchar '
 Zn As String 'zn varchar '
 bBed As Integer 'bBed bit '
 Bemerkung As String 'Bemerkung longtext '
 Grund As String 'Grund varchar 'Grund aus BMP
 Stärke As String 'Stärke varchar 'Stärke aus BMP
 Einheit As String 'Einheit varchar 'Stärke aus BMP
 Form As String 'Form varchar 'Stärke aus BMP
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 ergaenzt As Integer 'ergaenzt bit 'PZN ergaenzt von pznbdt (bdtnachw)
End Type

Public Type rezepteintraege
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Rezept As String 'Rezept varchar '6210, 3652(1), 6218(1)
 RKlnm As String 'RKlnm varchar 'Anfang des Rezeptklassennamens bei Langrezepten (HeilHilfsmittel, LangRezeptEintrag)
 Rezeptklasse As String 'Rezeptklasse varchar '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Rezklkurz As String 'Rezklkurz varchar 'letztes Split-Feld, z.B. 'rp' oder 'lar'
 Rezkllang As String 'Rezkllang varchar 'erstes Split-Feld, z.B. 'Sprechstundenbedarf': Langrezept, Sprechstundenbedarf, Heilmittel
 kbez As String 'kbez varchar 'vorletztes Split-Feld, z.B. 'DTronAnthra'
 Medikament As String 'Medikament varchar '3652(2), 6218(4)
 auti As Long 'auti int 'aut-idem-Eintrag im Rz. (1=kein Ausschluß, 0=Ausschluß)
 anzl As Long 'anzl int 'Anzahl
 PZN As String 'PZN varchar '6210(2), 6218(3)
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Statusbyte
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 id As Long 'id int '
End Type

Public Type RR
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 RR As String 'RR longtext '6230
 Puls As Long 'Puls int 'Puls
 RRsyst As Integer 'RRsyst smallint '
 RRdiast As Integer 'RRdiast smallint '
 RRzahl As Integer 'RRzahl smallint '
 Quelle As String 'Quelle varchar 'Informationsquelle
 Bemerkung As String 'Bemerkung varchar 'Bemerkung
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type kvnrue
 lfdnr As Long 'lfdnr int '
 Pat_id As Long 'Pat_ID int '
 KVNr As String 'KVNr varchar '
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type unbekannte_kennungen
 Kennung As String 'Kennung varchar '
 absPos As Long 'absPos int '
 StByte As Long 'StByte int '
 Pat_id As Long 'Pat_id int 'zugehöriger Patient für spätere Ermittlungen
 Inhalt As String 'Inhalt varchar 'Inhalt Zeile zum Wiederauffinden
End Type

Public Type dmpreihe
 Abk As String 'Abk varchar 'Abkürzung der DMP-Art
 art As String 'Art varchar 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As Date 'KarteiDatum date 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As Date 'exportiert datetime 'Datum des Exports
 DokuDatum As Date 'DokuDatum datetime 'Datum der Dokumentation
 obvoll As Integer 'obvoll bit 'ob vollständig
 Ok As Integer 'ok bit 'ob """"""""ok""""""""
 ausgedruckt As Integer 'ausgedruckt bit 'ob """"""""ausgedruckt""""""""
 Nachname As String 'NachName varchar '
 Vorname As String 'VorName varchar '
 GebDat As Date 'GebDat date '
 Pat_id As Long 'Pat_id int '
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 AktZeit As Date 'AktZeit datetime 'Aktualisierungzeit
 lanrid As Long 'lanrid int 'Bezug auf lanrpraxis.id
 Zusatzdaten As String 'Zusatzdaten varchar 'Zusatzdaten
End Type

Public Type desktop
 id As Long 'id int 'Primärschlüssel
 IDS As String 'IDS varchar 'id=
 Pat_id As Long 'Pat_ID int '3000
 erstZP As Date 'erstZP datetime 'erstellungsZeitpunkt
 exoL As String 'exoL varchar 'executeonLoad
 hideT As Byte 'hideT tinyint 'hideTitel
 iconPath As String 'iconPath varchar 'iconPath
 noteBkColor As Long 'noteBkColor int 'noteBkColor
 noteFgColor As Long 'noteFgColor int 'noteFgColor
 positionBottom As Long 'positionBottom int 'positionBottom
 positionLeft As Long 'positionLeft int 'positionLeft
 positionRight As Long 'positionRight int 'positionRight
 positionTop As Long 'positionTop int 'positionTop
 showAsNote As Byte 'showAsNote tinyint 'showAsNote
 syncInfoList As String 'syncInfoList varchar 'syncInfoList
 Titel As String 'titel varchar 'titel
 toolTipText As String 'toolTipText varchar 'toolTipText
 verankert As Byte 'verankert tinyint 'verankert
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type usdm
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 art As String 'Art varchar '6330
 Spritzst As String 'Spritzst text 'Spritzstellen/Liphyper./~atr.
 Fußbef_re As String 'Fußbef_re text 'Fußbefund rechts
 Fußbef_li As String 'Fußbef_li text 'Fußbefund links
 Hyperk_re As String 'Hyperk_re text 'Hyperkeratosen rechts
 Hyperk_li As String 'Hyperk_li text 'Hyperkeratosen links
 Ulcera_re As String 'Ulcera_re text 'Ulcera rechts
 Ulcera_li As String 'Ulcera_li text 'Ulcera links
 Kraft_Zh_re As String 'Kraft_Zh_re text 'Kraft Zehenheber re
 Kraft_Zh_li As String 'Kraft_Zh_li text 'Kraft Zehenheber li
 Kraft_Zb_re As String 'Kraft_Zb_re text 'Kraft Zehenbeuger re
 Kraft_Zb_li As String 'Kraft_Zb_li text 'Kraft Zehenbeuger li
 Kraft_Knie_re As String 'Kraft_Knie_re text 'Kraft Knie re
 Kraft_Knie_li As String 'Kraft_Knie_li text 'Kraft Knie li
 ASR_re As String 'ASR_re text 'Achillessehnenreflex rechts
 ASR_li As String 'ASR_li text 'Achillessehnenreflex links
 PSR_re As String 'PSR_re text 'Patellarsehnenreflex rechts
 PSR_li As String 'PSR_li text 'Patellarsehnenreflex links
 Oberfl_re As String 'Oberfl_re text 'Oberflächensensibilität rechts
 Oberfl_li As String 'Oberfl_li text 'Oberflächensensibilität links
 MF_re As String 'MF_re text 'Monofilament rechts
 MF_li As String 'MF_li text 'Monofilament links
 KW_re As String 'KW_re text 'Kalt-warm rechts
 KW_li As String 'KW_li text 'Kalt-warm links
 Vibr_IK_re As String 'Vibr_IK_re text 'Vibration Innenknöchel rechts
 Vibr_IK_li As String 'Vibr_IK_li text 'Vibration Innenknöchel links
 Vibr_GZ_re As String 'Vibr_GZ_re text 'Vibration Großzehe rechts
 Vibr_GZ_li As String 'Vibr_GZ_li text 'Vibration Großzehe links
 PulsL_re As String 'PulsL_re text 'Puls Leiste rechts
 PulsL_li As String 'PulsL_li text 'Puls Leiste links
 PulsKK_re As String 'PulsKK_re text 'Puls Kniekehle rechts
 PulsKK_li As String 'PulsKK_li text 'Puls Kniekehle links
 PulsAtp_re As String 'PulsAtp_re text 'Puls Arteria tibialis posterior rechts
 PulsAtp_li As String 'PulsAtp_li text 'Puls Arteria tibialis posterior links
 PulsAdp_re As String 'PulsAdp_re text 'Puls Arteria dorsalis pedis rechts
 PulsAdp_li As String 'PulsAdp_li text 'Puls Arteria dorsalis pedis links
 Mitarbeiter As String 'Mitarbeiter text 'Mitarbeiter
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
End Type

Public Type fuss
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 art As String 'Art varchar '6330
 Fußdeform As String 'Fußdeform varchar 'Fußdeformität re/li
 Hyper_mEin As String 'Hyper_mEin varchar 'Hyperkeratose mit Einblutung re/li
 Weiteres As String 'Weiteres varchar 'Hyperkeratose ohne Einblutung, Stauungszeichen, Ekzem, Nagelmykose usw.
 Zn_Ulcus As String 'Zn_Ulcus varchar 'Z.n. Ulcus re/li
 Zn_Amput As String 'Zn_Amput varchar 'Z.n. Amputation re/li
 Fuß_ang As String 'Fuß_ang varchar 'Füße genau angeschaut
 Ulcera As String 'Ulcera varchar 'Ulcera re/li
 Wundinfektion As String 'Wundinfektion varchar 'Wundinfektion re/li
 nae_US As String 'nae_US varchar 'nächste Untersuchung
 Mitarbeiter As String 'Mitarbeiter varchar 'Mitarbeiter
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
End Type

Public Type ulcus
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Lokalisation As String 'Lokalisation varchar 'Lokalisation des Ulcus am Fuß, ohne Seite
 Seite As String 'Seite varchar 're oder li
 Größe As String 'Größe varchar 'Größe in mm
 Beläge As String 'Beläge varchar 'Beläge auf Ulcus
 Exsudat As String 'Exsudat varchar 'Exsudat auf Ulcus
 Geruch As String 'Geruch varchar 'Geruch aus 3 cm Entfernung
 Wundrand As String 'Wundrand varchar 'Wundrand
 Wundumgebung As String 'Wundumgebung varchar 'Wundumgebung
 Temperatur As String 'Temperatur varchar 'Temperatur im Vergleich zu gesunden Stellen
 Fotodoku As String 'Fotodoku varchar 'Fotodoku durchgeführt (ja/nein)
 Wundversorgung As String 'Wundversorgung varchar 'Wundversorgung
 Mitarbeiter As String 'Mitarbeiter varchar 'Mitarbeiterkürzel
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type vkgd
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Wohlempfinden As String 'Wohlempfinden varchar 'Wohlempfinden
 Saettigung As String 'Saettigung varchar 'Sättigung
 Zielwerterreichung As String 'Zielwerterreichung varchar 'Zielwerterreichung
 Ketonkörper As String 'Ketonkörper varchar 'Ketonkörper
 Gynaekologenbefund As String 'Gynaekologenbefund varchar 'aktueller Gynäkologenbefund
 Gewichtsentwicklung As String 'Gewichtsentwicklung varchar 'Gewichtsentwicklung der Schwangeren
 HbA1c As String 'HbA1c varchar 'HbA1c (monatlich)+ ggf. TSH (dreimonatlich) abgenommen
 Bewegung As String 'Bewegung varchar 'Bewegung: Art
 Minuten As String 'Minuten varchar 'Minuten pro Woche
 Blutdruck As String 'Blutdruck varchar 'Blutdruck
 Puls As String 'Puls varchar 'Puls
 Mitarbeiter As String 'Mitarbeiter varchar 'Mitarbeiterkürzel
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type sws
 FID As Long 'FID int 'Fall-Bezug
 Pat_id As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 lR As Date 'LR date 'anzeigeText:>LR:
 vorET As Date 'vorET date 'voraussichtlicher ET:
 ET As Date 'ET date 'Entbindung ... Datum:
 efLR As Date 'efLR date 'EffektiveLetzteRegel
 erLR As Date 'erLR date 'ErfassteLetzteRegel
 kGT As Date 'kGT date 'KorrigierterGeburtstermin
 MB As Date 'MB date 'MutterschutzBeginn
 EndeArt As String 'EndeArt varchar 'EndeArt
 ED As Date 'ED date 'EndeDatum
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type laborxsaetze
 SatzID As Long 'SatzID int 'zum Bezug für LaborUS
 DatID As Long 'DatID int 'Bezug zu LaborEingelesen
 Satzart As String 'Satzart varchar '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge varchar '8100 Satzlänge (Turbomed)
 SatzlängeSchluss As String 'SatzlängeSchluss varchar '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb As String 'VersionSatzb varchar '9212 Version der Satzbeschreibung (Turbomed)
 Arztnr As String 'Arztnr varchar '201 Arztnummer (Turbomed)
 Arztname As String 'Arztname varchar '203 Arztname (Turbomed)
 StraßePraxis As String 'StraßePraxis varchar '205 Straße der Praxis (Turbomed)
 Arzt As String 'Arzt varchar ' 211 Ausführender Arzt
 Lanr As String 'LANR varchar ' 212 LANR
 PLZPraxis As String 'PLZPraxis varchar '215 PLZ der Praxis (Turbomed)
 OrtPraxis As String 'OrtPraxis varchar '216 Ort der Praxis (Turbomed)
 Labor As String 'Labor varchar '8320 Labor
 StraßeLabor As String 'StraßeLabor varchar '8321 Straße der Laboradresse (Turbomed)
 PLZLabor As String 'PLZLabor varchar '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor As String 'OrtLabor varchar '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr As String 'KBVPrüfnr varchar '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz As String 'Zeichensatz varchar '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr As String 'Kundenarztnr varchar '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum As String 'Erstellungsdatum varchar '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge As String 'Gesamtlänge varchar '9202 Gesamtlänge des Datenpaketes (Turbomed)
End Type

Public Type laborxeingel
 DatID As Long 'DatID int 'Bezug auf LaborEingelesen
 Pfad As String 'Pfad varchar 'Pfadname
 name As String 'Name varchar 'Name der eingelesenen Labordatei ohne Endung
 Zp As Date 'Zp datetime 'Einlesezeitpunkt
 Fertig As Integer 'fertig bit 'ob Einlesen fertig
End Type

Public Type laborxus
 RefNr As Long 'RefNr int 'Bezug auf LaborWert
 DatID As Long 'DatID int 'Bezug auf LaborEingelesen
 SatzID As Long 'SatzID int 'Bezug auf LaborXSätze
 Satzart As String 'Satzart varchar '8000 Satzart (Turbomed)
 Satzlänge As String 'Satzlänge varchar '8100 Satzlänge (Turbomed)
 Auftragsnummer As String 'Auftragsnummer varchar '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel As String 'Auftragsschlüssel varchar '8311 Anforderungsnr d Labors (Turbomed)
 Eingang As Date 'Eingang datetime '8301 Eingangsdatum in Datumsform
 Berichtsdatum As String 'Berichtsdatum varchar '8302 Berichtsdatum
 Pat_id As Long 'Pat_id int '
 Nachname As String 'Nachname varchar '3101
 Vorname As String 'Vorname varchar '3102
 GebDat As String 'GebDat varchar '3103
 Titel As String 'Titel varchar '3104
 NVorsatz As String 'NVorsatz varchar '3100
 BefArt As String 'BefArt varchar '8401 Befundart (Turbomed) / Fertigstellungsgrad (""""""""E""""""""=Endbefund, """"""""T"""""""" = Teilbefund)
 Abrechnungstyp As String 'Abrechnungstyp varchar '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd As String 'GebüOrd varchar '8403 Gebührenordnung (Turbomed)
 Auftraggeber As String 'Auftraggeber varchar '8615 Auftraggeber (LANR)
 Patienteninformation As String 'Patienteninformation varchar '8405 Patienteninformation (Turbomed)
 geschlecht As String 'Geschlecht varchar '8407 Geschlecht (Turbomed)
 AuftrHinw As String 'AuftrHinw varchar '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp As String 'Pat_idUrsp varchar 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG As String 'Pat_idErwVNG varchar 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN As String 'Pat_idErwVN varchar 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG As String 'Pat_idErwG varchar 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB As String 'Pat_idErwGB varchar 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL As String 'Pat_idErwGL varchar 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu As String 'Pat_idLaborNeu varchar 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu As Date 'ZeitpunktLaborneu datetime 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP As Integer 'ZdüP smallint 'Zahl der verglichenen Parameter
 ZdiP As Long 'ZdiP int 'Zahl der infragekommenden Patienten
 LWerte As String 'LWerte longtext 'Laborwerte, die zur Zuordnung geführt haben
 verglichen As Date 'verglichen datetime 'Datum, zu dem Datensatz zuletzt verglichen wurde
 AfN As Integer 'AfN smallint 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
End Type

Public Type laborxbakt
 RefNr As Long 'RefNr int '
 Verf As String 'Verf varchar '
 KuQu As String 'KuQu varchar '8428 Probenmaterial-Ident (Turbomed)
 Quelle As String 'Quelle varchar '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As String 'QSpez varchar '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat datetime '8432 Abnahmedatum (Turbomed)
 Kommentar As String 'Kommentar longtext '8480 Ergebnistest (Turbomed)
 Erklärung As String 'Erklärung longtext '
 Keimzahl As String 'Keimzahl varchar '
End Type

Public Type laborxwert
 RefNr As Long 'RefNr int 'Bezug auf LaborUS
 Abkü As String 'Abkü varchar '8410 Test-Ident  (Turbomed)
 Langname As String 'Langname varchar '8411 Testbezeichnung (Turbomed)
 Quelle As String 'Quelle varchar '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez As String 'QSpez varchar '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat As Date 'AbnDat datetime '8432 Abnahmedatum (Turbomed)
 Wert As String 'Wert varchar '8420 Ergebniswert (Turbomed)
 Einheit As String 'Einheit varchar '8421 Einheit (Turbomed)
 Grenzwerti As String 'Grenzwerti varchar '8422 Grenzwertindikator (Turbomed)
 Kommentar As String 'Kommentar varchar '8480 Ergebnistext (Turbomed)
 Teststatus As String 'Teststatus varchar '8418 Teststatus (Turbomed)
 Erklärung As String 'Erklärung varchar '8470 Testbezogene Hinweise (Turbomed)
 AuftrHinw As String 'AuftrHinw varchar '8490 Auftragsbezogene Hinweise (Turbomed)
 nbid As Long 'nbid int 'Bezug zu laborxplab.id
End Type

Public Type laborxleist
 RefNr As Long 'RefNr int 'Bezug auf LaborUS
 Abkü As String 'Abkü varchar '8410 Test-Ident (Turbomed)
 Verf As String 'Verf varchar '8434
 EBM As String 'EBM varchar '5001 GNR (Turbomed)
 goä As String 'goä varchar '8406
 Anzahl As String 'Anzahl varchar '5005
 abrd As String 'abrd varchar '8614 Abrechnung durch: 1 = Labor, 2 = Einweiser
End Type

Public Type LIUEZ
 name As String 'name varchar '
 Vorname As String 'vorname varchar '
 titelt As String 'titelt varchar '
 fachgruppe As String 'fachgruppe varchar '
 strasse As String 'strasse varchar '
 plz As String 'plz varchar '
 ort As String 'ort varchar '
 telefon As String 'telefon varchar '
 fax As String 'fax varchar '
 KVNr As String 'kvnr varchar '
 aktdat As Date 'aktdat date '
 id As Long 'id int '
 Überschrift As String 'überschrift varchar '
 DBNr As String 'dbnr varchar '
 BStelle As String 'bstelle varchar '
 anrede As String 'anrede varchar '
 tel1 As String 'tel1 varchar '
 tel2 As String 'tel2 varchar '
 tel3 As String 'tel3 varchar '
 tel4 As String 'tel4 varchar '
 fax1 As String 'fax1 varchar '
 fax2 As String 'fax2 varchar '
 fax3 As String 'fax3 varchar '
 email As String 'email varchar '
 zulg As String 'zulg varchar '
 arzttyp As String 'arzttyp varchar '
 gemmit As String 'gemmit longtext '
 beme As String 'beme longtext '
 dmpt2 As Byte 'dmpt2 tinyint '
 dmpt1 As Byte 'dmpt1 tinyint '
 geschlecht As String 'geschlecht varchar '
 Titel As String 'titel varchar '
 zusatz As String 'zusatz varchar '
 ursp As String 'ursp varchar '
 AktZeit As Date 'aktzeit datetime '
End Type

Public Type Anamnesebogen
 Prim As Long 'Prim int 'Primärschlüssel
 Pat_id As Long 'Pat_id int '
 Nachname As String 'Nachname varchar '-
 Vorname As String 'Vorname varchar '
 NVorsatz As String 'NVorsatz varchar '
 Titel As String 'Titel varchar '
 anrede As String 'Anrede varchar '
 GebDat As Date 'GebDat datetime ', geb.
 Tkz As Byte 'Tkz tinyint 'Tod-Kennzeichen
 Versicherungsart As String 'Versicherungsart varchar '
 Diabetestyp As String 'Diabetestyp varchar '^Diabetes Typ
 Diabetes_seit As String 'Diabetes_seit text '<seit
 Tabletten_seit As String 'Tabletten_seit text ', Tabletten seit
 Insulin_seit As String 'Insulin_seit text ', Insulin seit
 Grund_für_Vorstellung As String 'Grund_für_Vorstellung text '^:
 Familienanamnese As String 'Familienanamnese text '^:
 Größe As Double 'Größe double '^:
 Gewicht As Double 'Gewicht double ',:
 bmi As Double 'bmi decimal '
 Tendenz As String 'Tendenz varchar '<, Tendenz
 DiabetesMedikament_1 As String 'DiabetesMedikament_1 text '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge As String 'DiabetesMedikament_1_Menge text '<
 DiabetesMedikament_2 As String 'DiabetesMedikament_2 text '<,
 DiabetesMedikament_2_Menge As String 'DiabetesMedikament_2_Menge text '<
 DiabetesMedikament_3 As String 'DiabetesMedikament_3 text '<,
 DiabetesMedikament_3_Menge As String 'DiabetesMedikament_3_Menge text '<
 DiabetesMedikament_4 As String 'DiabetesMedikament_4 text '<,
 DiabetesMedikament_4_Menge As String 'DiabetesMedikament_4_Menge text '<,
 Insulinpumpe As Byte 'Insulinpumpe tinyint '^:
 Insulinpumpe_seit As String 'Insulinpumpe_seit text '<seit
 Insulinpumpe_Marke As String 'Insulinpumpe_Marke text '<, Marke:
 Broteinheiten_gesamt As String 'Broteinheiten_gesamt text '^Broteinheiten:gesamt
 Broteinheiten_früh As String 'Broteinheiten_früh text '<, früh
 Broteinheiten_ZM_früh As String 'Broteinheiten_ZM_früh text '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags As String 'Broteinheiten_mittags text '<, mittags
 Broteinheiten_nachmittags As String 'Broteinheiten_nachmittags text '<, nachmittags
 Broteinheiten_abends As String 'Broteinheiten_abends text '<, abends
 Broteinheiten_nachts As String 'Broteinheiten_nachts text '<, nachts
 Essenszeit_früh As String 'Essenszeit_früh text '^Essenszeiten:früh
 Essenszeit_vormittags As String 'Essenszeit_vormittags text '<, vormittags
 Essenszeit_mittags As String 'Essenszeit_mittags text '<, mittags
 Essenszeit_nachmittags As String 'Essenszeit_nachmittags text '<, nachmittags
 Essenszeit_abends As String 'Essenszeit_abends text '<, abends
 Essenszeit_spät As String 'Essenszeit_spät text '<, spät
 Spritz_Eß_Abstand_früh As String 'Spritz_Eß_Abstand_früh text '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags As String 'Spritz_Eß_Abstand_mittags text '<, mittags
 Spritz_Eß_Abstand_abends As String 'Spritz_Eß_Abstand_abends text '<, abends
 Spritzstelle_früh As String 'Spritzstelle_früh text '^Spritzstellen:früh
 Spritzstelle_mittags As String 'Spritzstelle_mittags text '<, mittags
 Spritzstelle_abends As String 'Spritzstelle_abends text '<, abends
 Spritzstelle_nachts As String 'Spritzstelle_nachts text '<, nachts
 Ernährung As String 'Ernährung text '^:
 Jahr_letzte_Diabetesschulung As String 'Jahr_letzte_Diabetesschulung varchar '^Letzte Diabetesschulung:
 Ort_Schulung As String 'Ort_Schulung text '<in
 letztes_HbA1c As String 'letztes_HbA1c text '^Letztes HbA1c:
 gemessen_am As String 'gemessen_am text '<, gemessen
 vorherige_Werte As String 'vorherige_Werte text '<, vorher:
 subcutane_Zuckermessung As String 'subcutane_Zuckermessung text '^:
 CGM_seit As String 'CGM_seit text '<
 BZMessungen_selbst As String 'BZMessungen_selbst text '^Blutzuckermessung:Selbstmessung?
 Gerät As String 'Gerät text '<:
 BZMessungen_pW As String 'BZMessungen_pW text '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE As String 'BZMessungen_pW_ndE text '<, davon nach dem Essen:
 BZMessungen_p_W_nachts As String 'BZMessungen_p_W_nachts text '<, nachts:
 Aufschreiben As String 'Aufschreiben text '<, Dokumentation:
 BZWerte_v_d_Essen As String 'BZWerte_v_d_Essen text '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen As String 'BZWerte_n_d_Essen text '<, nach dem Essen:
 UZ_Tageszeit As String 'UZ_Tageszeit text '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM As String 'Unterzucker_pM text '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig As String 'UZ_rechtzeitig text '<, rechtzeitig bemerkt:
 Schwere_Uzu As String 'Schwere_Uzu text 'Schwere Unterzucker
 Fremde_Hilfe_pa As String 'Fremde_Hilfe_pa text '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa As String 'Bewußtlos_pa text '<, bewußtlos deshalb:
 Keto As String 'Keto text '^Bisher Ketoazidosen mit Krankenhauseinweisung:
 Keto_pa As String 'Keto_pa text '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM As String 'BZgr300_pM text ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck As String 'Bluthochdruck text '^Bluthochdruck:
 BHD_seit As String 'BHD_seit text '<seit:
 BHD_beh_mit As String 'BHD_beh_mit text '<, behandelt mit:
 Blutdruckwerte As String 'Blutdruckwerte text '^Blutdruckwerte:
 BDselbst As String 'BDselbst text '^Blutdruckselbstmessung:
 Schwanger As String 'Schwanger text '^Aktuelle Schwangerschaft:
 Schwanger_seit As String 'Schwanger_seit text '<, seit:
 Augensp_zuletzt As String 'Augensp_zuletzt text '^Letzte Augenspiegelung:
 Augensp_Befund As String 'Augensp_Befund text '<, Befund:
 Netzhaut_gelasert As String 'Netzhaut_gelasert text ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar As String 'Sehminderung_unbehebbar text ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden As String 'Diabet_Nierenschaden text '^Diabetischer Nierenschaden:
 Albumin_zuletzt As String 'Albumin_zuletzt text ', letztes Albumin:
 erhöht As String 'erhöht text '<, Befund:
 Kreatinin As String 'Kreatinin text ',:
 Dialyse As Byte 'Dialyse tinyint ',:
 Dialyse_seit As String 'Dialyse_seit text '<seit
 andere_Nierenerkrankung As String 'andere_Nierenerkrankung text ', andere Nierenerkrankung:
 Herzkrankheit As String 'Herzkrankheit text '^Herzkrankheit:
 Angina_pectoris As String 'Angina_pectoris text ',:
 Herzinfarkt As String 'Herzinfarkt text ',:
 Herzinfarkt_wann As String 'Herzinfarkt_wann text '<, wann:
 PTCA_oder_Stent As String 'PTCA_oder_Stent text ',:
 Bypass_kardial As Byte 'Bypass_kardial tinyint ',:
 Bypass_wann As String 'Bypass_wann text '<, wann:
 Herzschwäche As String 'Herzschwäche text ',:
 Herzkrankheit_Beschreibung As String 'Herzkrankheit_Beschreibung text ', Beschreibung:
 Hirndurchblutungsstörung As String 'Hirndurchblutungsstörung text '^:
 Schlaganfall As String 'Schlaganfall text ',:
 Beindurchblutungsstörung As String 'Beindurchblutungsstörung text '^:
 Schaufensterkrankheit As String 'Schaufensterkrankheit text ',:
 Bypaß_peripher As Byte 'Bypaß_peripher tinyint ',:
 Geschwür As String 'Geschwür text ',:
 Amputation As String 'Amputation text ',:
 pAVK_Beschreibung As String 'pAVK_Beschreibung text ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen As String 'Ameisenlaufen text '^:
 Ameisen_Ausmaß As String 'Ameisen_Ausmaß text '<, Ausmaß:
 Druckstellen As String 'Druckstellen text ',:
 Verformungen As String 'Verformungen text ',:
 Verformungen_Beschreibung As String 'Verformungen_Beschreibung text '<Beschreibung:
 Fußpflege As String 'Fußpflege text '^:
 Podologie As String 'Podologie text ',:
 Einlagen As String 'Einlagen text ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen As String 'Neue_Fußkomplikationen text '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen As String 'Entleerungsstörungen_Magen text '^:
 Entleerungsstörungen_Harnblase As String 'Entleerungsstörungen_Harnblase text ',:
 Schwindel_Aufstehen As String 'Schwindel_Aufstehen text ',:
 Folgeerkrankungen_Haut As String 'Folgeerkrankungen_Haut text '^:
 Bewegungseinschränkungen As String 'Bewegungseinschränkungen text ',:
 Sexualstörung As String 'Sexualstörung text '^:
 Sexualstörung_seit As String 'Sexualstörung_seit text '<seit
 Weitere_Anamnese As String 'Weitere_Anamnese text '^:
 Tabak As String 'Tabak text '^Tabak:
 tabakex As String 'tabakex text '<, früher:
 tabakbis As String 'tabakbis text '<, bis:
 tabakakt As String 'tabakakt text '<, aktuell:
 tabakmenge As String 'tabakmenge text '<, Menge:
 Alkohol As String 'Alkohol text '^Alkohol pro Woche:
 Mitarbeiter As String 'Mitarbeiter text '<, Mitarbeiter:
 Weitere_Medikation As String 'Weitere_Medikation text '^:
 Liphypertrophien_Abdomen As String 'Liphypertrophien_Abdomen text '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine As String 'Liphypertrophien_Beine text '<, Beine:
 Liphypertrophien_Arme As String 'Liphypertrophien_Arme varchar '<, Arme:
 Beinbefund As String 'Beinbefund text '^:
 Hyperkeratosen As String 'Hyperkeratosen text ',:
 Ulcera As String 'Ulcera text ',:
 Kraft_Zehenheber As String 'Kraft_Zehenheber text '^Kraft:Zehenheber
 Kraft_Zehenbeuger As String 'Kraft_Zehenbeuger text '<, Zehenbeuger:
 Kraft_Knie As String 'Kraft_Knie text '<, Knie:
 ASR As String 'ASR text ',:
 PSR As String 'PSR text ',:
 Oberflächensensibilität As String 'Oberflächensensibilität text '^:
 Monofilamenttest As String 'Monofilamenttest text ',:
 Kalt_Warm As String 'Kalt_Warm text ', Kalt-Warm-Diskrimination:
 Vibration_IK As String 'Vibration_IK text ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe As String 'Vibration_Großzehe text '<, Großzehe:
 Puls_Leiste As String 'Puls_Leiste text '^Pulse:Leiste
 Puls_Kniekehle As String 'Puls_Kniekehle text '<,Kniekehle:
 Puls_Atp As String 'Puls_Atp text '<,Innenknöchel:
 Puls_Adp As String 'Puls_Adp text '<,Fußrücken:
 RR As String 'RR text '^Blutdruck:
 RRTurboMed As String 'RRTurboMed text '
 Herz As String 'Herz text '^:
 Lunge As String 'Lunge text ',:
 Bauch As String 'Bauch text ', Abdomen:
 WS As String 'WS text ', Wirbelsäule:
 NL As String 'NL text ', Nierenlager:
 SD As String 'SD text ', Schilddrüse:
 Carotiden As String 'Carotiden text ', Halsschlagadern:
 NNH As String 'NNH text ', Nasennebenhöhlen:
 Zähne As String 'Zähne text ',:
 Parodontitis As String 'Parodontitis text ',:
 Mundhöhle As String 'Mundhöhle text ',:
 LK As String 'LK text ', Lymphknoten:
 BeinödVen As String 'BeinödVen text ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst As String 'Neuro_sonst text '^Sonstige neurologische Befunde:
 Weitere_Befunde As String 'Weitere_Befunde text ', weitere Befunde:
 Schulung As String 'Schulung varchar 'ob Schulungsbedarf
 DMP As String 'DMP text 'ob Pat. bei HA im DMP
 DMSchulz As Integer 'DMSchulz smallint 'Zahl der DMP-Schulungen hier
 DMSchL As Integer 'DMSchL smallint 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz As Integer 'RRSchulz smallint 'Zahl der Hypertonie-Schulungen hier
 DMPhier As Date 'DMPhier datetime 'ob Pat hier im DMP
 HANr As String 'HANr varchar 'mit """"""""/""""""""
 HANr2 As String 'HANr2 varchar 'mit """"""""/""""""""
 letzte_Änderung As Date 'letzte_Änderung datetime 'Datum der letzten Änderung
 Diagnosen As String 'Diagnosen text '
 Vorgestellt As Date 'Vorgestellt datetime 'Erstvorstellung
 Versicherung As String 'Versicherung varchar '
 AktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 Ther1 As String 'Ther1 varchar 'Diät, OAD, CT, Komb, ICT, CSII
 TherAkt As String 'TherAkt varchar 'Diät, OAD, CT, Komb, ICT, CSII
 obAn1eing As Byte 'obAn1eing tinyint 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing As Byte 'obAn2eing tinyint 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing As Byte 'obAnAeing tinyint 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck As Byte 'obCheck tinyint 'ob Checkliste vorliegt
 obBZausgew As Byte 'obBZausgew tinyint 'ob Blutzuckergerät ausgewechselt
 obOSaufgek As Byte 'obOSaufgek tinyint 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek As Byte 'obPodAufgek tinyint 'ob über Podologie aufgeklärt
 obMBlAusgeh As Byte 'obMBlAusgeh tinyint 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek As String 'obSchulaufgek varchar 'ob über Podologie aufgeklärt
 obDMPaufgekl As String 'obDMPaufgekl varchar 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz As Byte 'obMedNetz tinyint 'ob von Med. Netz geschickt
 Hausarzt As String 'Hausarzt text 'Hausarzt laut Anamnesebogen
 ob As Byte 'ob tinyint 'für verschiedene Aktionen
 QS As String 'QS varchar 'Quartal sortiert von vorgestellt
 QT As String 'QT varchar 'Quartal sortiert von vorgestellt
End Type

Public rNa() As namen
Public rFa() As Faelle
Public rAu() As au
Public rBr() As briefe
Public rDi() As Diagnosen
Public rDo() As dokumente
Public rEi() As eintraege
Public rFi() As forminhaltform_abk
Public rFo() As formulare ' kommt vor in: formulareSpeichern, doTabVorb, dolies
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
Public rDe() As desktop
Public rUs() As usdm
Public rFu() As fuss
Public rUl() As ulcus
Public rVk() As vkgd
Public rSw() As sws
Public rLs() As laborxsaetze
Public rLg() As laborxeingel
Public rLu() As laborxus
Public rLo() As laborxbakt
Public rLw() As laborxwert
Public rLL() As laborxleist
Public rLi() As LIUEZ
Public rAna() As Anamnesebogen

Public roNa() As namen
Public roFa() As Faelle
Public roAu() As au
Public roBr() As briefe
Public roDi() As Diagnosen
Public roDo() As dokumente
Public roEi() As eintraege
Public roFi() As forminhaltform_abk
Public roFo() As formulare
Public roFr() As forminhkopf
Public roFm() As forminhfeld
Public roKh() As kheinweis
Public roLb() As lbanforderungen
Public roLa() As laborneu
Public roLe() As leistungen
Public roMe() As medplan
Public roRe() As rezepteintraege
Public roRr() As RR
Public roKv() As kvnrue
Public roUn() As unbekannte_kennungen
Public roDm() As dmpreihe
Public roDe() As desktop
Public roUs() As usdm
Public roFu() As fuss
Public roUl() As ulcus
Public roVk() As vkgd
Public roSw() As sws
Public roLs() As laborxsaetze
Public roLg() As laborxeingel
Public roLu() As laborxus
Public roLo() As laborxbakt
Public roLw() As laborxwert
Public roLL() As laborxleist
Public roLi() As LIUEZ
Public roAna() As Anamnesebogen

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
 ReDim rDe(0)
 ReDim rUs(0)
 ReDim rFu(0)
 ReDim rUl(0)
 ReDim rVk(0)
 ReDim rSw(0)
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
 ReDim rLi(0)
 wdh = -1
End Function ' Tinit

' in AllesLösch, LabLösch
Public Function doEntleer(frm As Lese, Tbl$)
 Dim rs As ADODB.Recordset
' SET rs = myEFrag("SELECT COUNT(0) ct FROM `" & Tbl & "`")
 myFrag rs, "SELECT COUNT(0) ct FROM `" & Tbl & "`"
 If Not rs.BOF Then
  frm.Ausgeb "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)", True
  sql = sqlDeletefrom & "`" & Tbl & "`"
 End If ' Not rs.BOF then
 Call myEFrag(sql) ' ,,adAsyncExecute
 DoEvents
End Function ' doEntleer

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doEntleer(frm, "sws")
 Call doEntleer(frm, "vkgd")
 Call doEntleer(frm, "ulcus")
 Call doEntleer(frm, "fuss")
 Call doEntleer(frm, "usdm")
 Call doEntleer(frm, "desktop")
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
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' AllesLösch

Public Function LabLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doEntleer(frm, "liuez")
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
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' LabLösch

Function doBezFeh(csqlVal$, obSpei%, ErrDes$)
 Call ForeignNo0
 Call ForeignNo1
 obForK = True
 If obSpei <> 0 Then
  Open BezFeh For Append As #299
  Print #299, vbCrLf & vbCrLf & Now() & ": " & csqlVal
  Print #299, vbCrLf & ErrDes
  Close #299
 End If
End Function 'doBezFeh

' aufgerufen in alleSpeichern
Function fidSetz()
 Dim i&, j&
 For i = 1 To UBound(rBr) ' Briefe
  For j = 1 To UBound(rFa)
   If rBr(i).FID = rFa(j).FID Then GoTo Brweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rBr(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rBr(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rBr(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rBr(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rBr(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Brweiter:
 Next i
 For i = 1 To UBound(rDi) ' Diagnosen
  For j = 1 To UBound(rFa)
   If rDi(i).FID = rFa(j).FID Then GoTo Diweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rDi(i).DiagDatum) >= Int(rFa(j).BhFB) And Int(rDi(i).DiagDatum) <= Int(rFa(j).BhFE1) Then
    rDi(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rDi(i).DiagDatum)) >= int(rFa(j).BhFB) AND int(rDi(i).diagDatum) <= int(rFa(j).BhFE1) Then
  Next j
Diweiter:
 Next i
 For i = 1 To UBound(rEi) ' eintraege
  For j = 1 To UBound(rFa)
   If rEi(i).FID = rFa(j).FID Then GoTo Eiweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rEi(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rEi(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rEi(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rEi(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rEi(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Eiweiter:
 Next i
 For i = 1 To UBound(rFr) ' FormInhKopf
  For j = 1 To UBound(rFa)
   If rFr(i).FID = rFa(j).FID Then GoTo Frweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rFr(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rFr(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rFr(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rFr(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rFr(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Frweiter:
 Next i
 For i = 1 To UBound(rLe) ' Leistungen
  For j = 1 To UBound(rFa)
   If rLe(i).FID = rFa(j).FID Then GoTo Leweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rLe(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rLe(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rLe(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rLe(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rLe(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Leweiter:
 Next i
 For i = 1 To UBound(rMe) ' MedPlan
  For j = 1 To UBound(rFa)
   If rMe(i).FID = rFa(j).FID Then GoTo Meweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rMe(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rMe(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rMe(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rMe(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rMe(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Meweiter:
 Next i
 For i = 1 To UBound(rRe) ' Rezepteintraege
  For j = 1 To UBound(rFa)
   If rRe(i).FID = rFa(j).FID Then GoTo Reweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rRe(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rRe(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rRe(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rRe(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rRe(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Reweiter:
 Next i
 For i = 1 To UBound(rFu) ' fuss
  For j = 1 To UBound(rFa)
   If rFu(i).FID = rFa(j).FID Then GoTo Fuweiter
  Next j
  For j = 1 To UBound(rFa)
   If Int(rFu(i).Zeitpunkt) >= Int(rFa(j).BhFB) And Int(rFu(i).Zeitpunkt) <= Int(rFa(j).BhFE1) Then
    rFu(i).FID = rFa(j).FID
    Exit For
   End If ' If int(rFu(i).Zeitpunkt)) >= int(rFa(j).BhFB) AND int(rFu(i).Zeitpunkt) <= int(rFa(j).BhFE1) Then
  Next j
Fuweiter:
 Next i
End Function ' FIDsetz

Public Function roNaZuw(i&, j&)
 roNa(i).Pat_id = rNa(j).Pat_id
 roNa(i).lfdnr = rNa(j).lfdnr
 roNa(i).NVorsatz = rNa(j).NVorsatz
 roNa(i).Nachname = rNa(j).Nachname
 roNa(i).Vorname = rNa(j).Vorname
 roNa(i).GebDat = rNa(j).GebDat
 roNa(i).f3004 = rNa(j).f3004
 roNa(i).f3006 = rNa(j).f3006
 roNa(i).Straße = rNa(j).Straße
 roNa(i).KVKStatus = rNa(j).KVKStatus
 roNa(i).Hausnr = rNa(j).Hausnr
 roNa(i).geschlecht = rNa(j).geschlecht
 roNa(i).plz = rNa(j).plz
 roNa(i).ort = rNa(j).ort
 roNa(i).Lkz = rNa(j).Lkz
 roNa(i).Anschrzus = rNa(j).Anschrzus
 roNa(i).NVors = rNa(j).NVors
 roNa(i).PFPlz = rNa(j).PFPlz
 roNa(i).PFOrt = rNa(j).PFOrt
 roNa(i).PFNr = rNa(j).PFNr
 roNa(i).f3124 = rNa(j).f3124
 roNa(i).AnschrZus_2 = rNa(j).AnschrZus_2
 roNa(i).Postfach_2 = rNa(j).Postfach_2
 roNa(i).LK_2 = rNa(j).LK_2
 roNa(i).Postfach = rNa(j).Postfach
 roNa(i).Beruf = rNa(j).Beruf
 roNa(i).Weggeldzone = rNa(j).Weggeldzone
 roNa(i).WeggzZahl = rNa(j).WeggzZahl
 roNa(i).AufnDat = rNa(j).AufnDat
 roNa(i).kAufDat = rNa(j).kAufDat
 roNa(i).Lanr = rNa(j).Lanr
 roNa(i).BStNr = rNa(j).BStNr
 roNa(i).Titel = rNa(j).Titel
 roNa(i).Versichertennummer = rNa(j).Versichertennummer
 roNa(i).PrivatTel = rNa(j).PrivatTel
 roNa(i).KVNr = rNa(j).KVNr
 roNa(i).KVNr2 = rNa(j).KVNr2
 roNa(i).KVNr3 = rNa(j).KVNr3
 roNa(i).KVNr4 = rNa(j).KVNr4
 roNa(i).PrivatTel_2 = rNa(j).PrivatTel_2
 roNa(i).PrivatFax = rNa(j).PrivatFax
 roNa(i).DienstTel = rNa(j).DienstTel
 roNa(i).PrivatMobil = rNa(j).PrivatMobil
 roNa(i).email = rNa(j).email
 roNa(i).Arbeitgeber = rNa(j).Arbeitgeber
 roNa(i).AnAllgda = rNa(j).AnAllgda
 roNa(i).An1da = rNa(j).An1da
 roNa(i).An2da = rNa(j).An2da
 roNa(i).Checkda = rNa(j).Checkda
 roNa(i).DMTypaD = rNa(j).DMTypaD
 roNa(i).AktZeit = rNa(j).AktZeit
 roNa(i).absPos = rNa(j).absPos
 roNa(i).StByte = rNa(j).StByte
 roNa(i).StByteA = rNa(j).StByteA
 roNa(i).Cave = rNa(j).Cave
 roNa(i).Notiz = rNa(j).Notiz
 roNa(i).obChk = rNa(j).obChk
 roNa(i).dmpklass = rNa(j).dmpklass
 roNa(i).dmpbeg = rNa(j).dmpbeg
 roNa(i).dmpkhkklass = rNa(j).dmpkhkklass
 roNa(i).dmpkhkbeg = rNa(j).dmpkhkbeg
 roNa(i).dmpcopdklass = rNa(j).dmpcopdklass
 roNa(i).dmpcopdbeg = rNa(j).dmpcopdbeg
 roNa(i).dmpabklass = rNa(j).dmpabklass
 roNa(i).dmpabbeg = rNa(j).dmpabbeg
 roNa(i).dakab = rNa(j).dakab
 roNa(i).HzV = rNa(j).HzV
 roNa(i).HzVbeg = rNa(j).HzVbeg
 roNa(i).DS = rNa(j).DS
 roNa(i).DSbeg = rNa(j).DSbeg
 roNa(i).getHA0 = rNa(j).getHA0
 roNa(i).fnHA0 = rNa(j).fnHA0
 roNa(i).getHA1 = rNa(j).getHA1
 roNa(i).fnHA1 = rNa(j).fnHA1
 roNa(i).getHA2 = rNa(j).getHA2
 roNa(i).fnHA2 = rNa(j).fnHA2
 roNa(i).zubenach = rNa(j).zubenach
 roNa(i).Verwandt = rNa(j).Verwandt
 roNa(i).Sprache = rNa(j).Sprache
 roNa(i).lAktTM = rNa(j).lAktTM
 roNa(i).Mitarbeiter = rNa(j).Mitarbeiter
End Function ' roNaZuw

Public Function NaZUnt%(i&, j&)
 If roNa(i).Pat_id <> rNa(j).Pat_id Then GoSub unter
 If roNa(i).lfdnr <> rNa(j).lfdnr Then GoSub unter
 If roNa(i).NVorsatz <> rNa(j).NVorsatz Then GoSub unter
 If roNa(i).Nachname <> rNa(j).Nachname Then GoSub unter
 If roNa(i).Vorname <> rNa(j).Vorname Then GoSub unter
 If roNa(i).GebDat <> rNa(j).GebDat Then GoSub unter
 If roNa(i).f3004 <> rNa(j).f3004 Then GoSub unter
 If roNa(i).f3006 <> rNa(j).f3006 Then GoSub unter
 If roNa(i).Straße <> rNa(j).Straße Then GoSub unter
 If roNa(i).KVKStatus <> rNa(j).KVKStatus Then GoSub unter
 If roNa(i).Hausnr <> rNa(j).Hausnr Then GoSub unter
 If roNa(i).geschlecht <> rNa(j).geschlecht Then GoSub unter
 If roNa(i).plz <> rNa(j).plz Then GoSub unter
 If roNa(i).ort <> rNa(j).ort Then GoSub unter
 If roNa(i).Lkz <> rNa(j).Lkz Then GoSub unter
 If roNa(i).Anschrzus <> rNa(j).Anschrzus Then GoSub unter
 If roNa(i).NVors <> rNa(j).NVors Then GoSub unter
 If roNa(i).PFPlz <> rNa(j).PFPlz Then GoSub unter
 If roNa(i).PFOrt <> rNa(j).PFOrt Then GoSub unter
 If roNa(i).PFNr <> rNa(j).PFNr Then GoSub unter
 If roNa(i).f3124 <> rNa(j).f3124 Then GoSub unter
 If roNa(i).AnschrZus_2 <> rNa(j).AnschrZus_2 Then GoSub unter
 If roNa(i).Postfach_2 <> rNa(j).Postfach_2 Then GoSub unter
 If roNa(i).LK_2 <> rNa(j).LK_2 Then GoSub unter
 If roNa(i).Postfach <> rNa(j).Postfach Then GoSub unter
 If roNa(i).Beruf <> rNa(j).Beruf Then GoSub unter
 If roNa(i).Weggeldzone <> rNa(j).Weggeldzone Then GoSub unter
 If roNa(i).WeggzZahl <> rNa(j).WeggzZahl Then GoSub unter
 If roNa(i).AufnDat <> rNa(j).AufnDat Then GoSub unter
 If roNa(i).kAufDat <> rNa(j).kAufDat Then GoSub unter
 If roNa(i).Lanr <> rNa(j).Lanr Then GoSub unter
 If roNa(i).BStNr <> rNa(j).BStNr Then GoSub unter
 If roNa(i).Titel <> rNa(j).Titel Then GoSub unter
 If roNa(i).Versichertennummer <> rNa(j).Versichertennummer Then GoSub unter
 If roNa(i).PrivatTel <> rNa(j).PrivatTel Then GoSub unter
 If roNa(i).KVNr <> rNa(j).KVNr Then GoSub unter
 If roNa(i).KVNr2 <> rNa(j).KVNr2 Then GoSub unter
 If roNa(i).KVNr3 <> rNa(j).KVNr3 Then GoSub unter
 If roNa(i).KVNr4 <> rNa(j).KVNr4 Then GoSub unter
 If roNa(i).PrivatTel_2 <> rNa(j).PrivatTel_2 Then GoSub unter
 If roNa(i).PrivatFax <> rNa(j).PrivatFax Then GoSub unter
 If roNa(i).DienstTel <> rNa(j).DienstTel Then GoSub unter
 If roNa(i).PrivatMobil <> rNa(j).PrivatMobil Then GoSub unter
 If roNa(i).email <> rNa(j).email Then GoSub unter
 If roNa(i).Arbeitgeber <> rNa(j).Arbeitgeber Then GoSub unter
 If roNa(i).AnAllgda <> rNa(j).AnAllgda Then GoSub unter
 If roNa(i).An1da <> rNa(j).An1da Then GoSub unter
 If roNa(i).An2da <> rNa(j).An2da Then GoSub unter
 If roNa(i).Checkda <> rNa(j).Checkda Then GoSub unter
 If roNa(i).DMTypaD <> rNa(j).DMTypaD Then GoSub unter
 If roNa(i).AktZeit <> rNa(j).AktZeit Then GoSub unter
 If roNa(i).absPos <> rNa(j).absPos Then GoSub unter
 If roNa(i).StByte <> rNa(j).StByte Then GoSub unter
 If roNa(i).StByteA <> rNa(j).StByteA Then GoSub unter
 If roNa(i).Cave <> rNa(j).Cave Then GoSub unter
 If roNa(i).Notiz <> rNa(j).Notiz Then GoSub unter
 If roNa(i).obChk <> rNa(j).obChk Then GoSub unter
 If roNa(i).dmpklass <> rNa(j).dmpklass Then GoSub unter
 If roNa(i).dmpbeg <> rNa(j).dmpbeg Then GoSub unter
 If roNa(i).dmpkhkklass <> rNa(j).dmpkhkklass Then GoSub unter
 If roNa(i).dmpkhkbeg <> rNa(j).dmpkhkbeg Then GoSub unter
 If roNa(i).dmpcopdklass <> rNa(j).dmpcopdklass Then GoSub unter
 If roNa(i).dmpcopdbeg <> rNa(j).dmpcopdbeg Then GoSub unter
 If roNa(i).dmpabklass <> rNa(j).dmpabklass Then GoSub unter
 If roNa(i).dmpabbeg <> rNa(j).dmpabbeg Then GoSub unter
 If roNa(i).dakab <> rNa(j).dakab Then GoSub unter
 If roNa(i).HzV <> rNa(j).HzV Then GoSub unter
 If roNa(i).HzVbeg <> rNa(j).HzVbeg Then GoSub unter
 If roNa(i).DS <> rNa(j).DS Then GoSub unter
 If roNa(i).DSbeg <> rNa(j).DSbeg Then GoSub unter
 If roNa(i).getHA0 <> rNa(j).getHA0 Then GoSub unter
 If roNa(i).fnHA0 <> rNa(j).fnHA0 Then GoSub unter
 If roNa(i).getHA1 <> rNa(j).getHA1 Then GoSub unter
 If roNa(i).fnHA1 <> rNa(j).fnHA1 Then GoSub unter
 If roNa(i).getHA2 <> rNa(j).getHA2 Then GoSub unter
 If roNa(i).fnHA2 <> rNa(j).fnHA2 Then GoSub unter
 If roNa(i).zubenach <> rNa(j).zubenach Then GoSub unter
 If roNa(i).Verwandt <> rNa(j).Verwandt Then GoSub unter
 If roNa(i).Sprache <> rNa(j).Sprache Then GoSub unter
 If roNa(i).lAktTM <> rNa(j).lAktTM Then GoSub unter
 If roNa(i).Mitarbeiter <> rNa(j).Mitarbeiter Then GoSub unter
 Exit Function
unter:
 NaZUnt = NaZUnt + 1
 Return
End Function ' NaZUnt

Public Function namenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(Pat_ID,0) Pat_ID,COALESCE(lfdnr,0) lfdnr,COALESCE(NVorsatz,'') NVorsatz,COALESCE(Nachname,'') Nachname" & _
",COALESCE(Vorname,'') Vorname,COALESCE(GebDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) GebDat,COALESCE(f3004,'') f3004,COALESCE(f3006,'') f3006" & _
",COALESCE(Straße,'') Straße,COALESCE(KVKStatus,'') KVKStatus,COALESCE(Hausnr,'') Hausnr,COALESCE(Geschlecht,'') Geschlecht" & _
",COALESCE(Plz,'') Plz,COALESCE(Ort,'') Ort,COALESCE(Lkz,'') Lkz,COALESCE(Anschrzus,'') Anschrzus" & _
",COALESCE(NVors,'') NVors,COALESCE(PFPlz,'') PFPlz,COALESCE(PFOrt,'') PFOrt,COALESCE(PFNr,'') PFNr" & _
",COALESCE(f3124,'') f3124,COALESCE(AnschrZus_2,'') AnschrZus_2,COALESCE(Postfach_2,'') Postfach_2,COALESCE(LK_2,'') LK_2" & _
",COALESCE(Postfach,'') Postfach,COALESCE(Beruf,'') Beruf,COALESCE(Weggeldzone,'') Weggeldzone,COALESCE(WeggzZahl,0) WeggzZahl" & _
",COALESCE(AufnDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AufnDat,COALESCE(kAufDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) kAufDat,COALESCE(LANR,'') LANR,COALESCE(BStNr,'') BStNr" & _
",COALESCE(Titel,'') Titel,COALESCE(Versichertennummer,'') Versichertennummer,COALESCE(PrivatTel,'') PrivatTel,COALESCE(KVNr,'') KVNr" & _
",COALESCE(KVNr2,'') KVNr2,COALESCE(KVNr3,'') KVNr3,COALESCE(KVNr4,'') KVNr4,COALESCE(PrivatTel_2,'') PrivatTel_2" & _
",COALESCE(PrivatFax,'') PrivatFax,COALESCE(DienstTel,'') DienstTel,COALESCE(PrivatMobil,'') PrivatMobil,COALESCE(Email,'') Email" & _
",COALESCE(Arbeitgeber,'') Arbeitgeber,COALESCE(AnAllgda,0) AnAllgda,COALESCE(An1da,0) An1da,COALESCE(An2da,0) An2da" & _
",COALESCE(Checkda,0) Checkda,COALESCE(DMTypaD,'') DMTypaD,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(absPos,0) absPos" & _
",COALESCE(StByte,0) StByte,COALESCE(StByteA,0) StByteA,COALESCE(Cave,'') Cave,COALESCE(Notiz,'') Notiz" & _
",COALESCE(obChk,'') obChk,COALESCE(dmpklass,0) dmpklass,COALESCE(dmpbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpbeg,COALESCE(dmpkhkklass,0) dmpkhkklass" & _
",COALESCE(dmpkhkbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpkhkbeg,COALESCE(dmpcopdklass,0) dmpcopdklass,COALESCE(dmpcopdbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpcopdbeg,COALESCE(dmpabklass,0) dmpabklass" & _
",COALESCE(dmpabbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpabbeg,COALESCE(dakab - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dakab,COALESCE(HzV,0) HzV,COALESCE(HzVbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) HzVbeg" & _
",COALESCE(DS,0) DS,COALESCE(DSbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DSbeg,COALESCE(getHA0,0) getHA0,COALESCE(fnHA0,'') fnHA0" & _
",COALESCE(getHA1,0) getHA1,COALESCE(fnHA1,'') fnHA1,COALESCE(getHA2,0) getHA2,COALESCE(fnHA2,'') fnHA2" & _
",COALESCE(zubenach,'') zubenach,COALESCE(Verwandt,'') Verwandt,COALESCE(Sprache,'') Sprache,COALESCE(lAktTM - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lAktTM" & _
",COALESCE(Mitarbeiter,0) Mitarbeiter FROM `namen` WHERE Pat_ID=" & pid & " ORDER BY `kAufDat`"
 myFrag rs, sql
 ReDim roNa(1)
 If Not rs.EOF Then
  Do While Not rs.EOF
   akt = UBound(roNa)
   roNa(akt).Pat_id = rs!Pat_id
   roNa(akt).lfdnr = rs!lfdnr
   roNa(akt).NVorsatz = doUmwfSQL(rs!NVorsatz, lies.obMySQL, False)
   roNa(akt).Nachname = doUmwfSQL(rs!Nachname, lies.obMySQL, False)
   roNa(akt).Vorname = doUmwfSQL(rs!Vorname, lies.obMySQL, False)
   roNa(akt).GebDat = rs!GebDat
   roNa(akt).f3004 = doUmwfSQL(rs!f3004, lies.obMySQL, False)
   roNa(akt).f3006 = doUmwfSQL(rs!f3006, lies.obMySQL, False)
   roNa(akt).Straße = doUmwfSQL(rs!Straße, lies.obMySQL, False)
   roNa(akt).KVKStatus = doUmwfSQL(rs!KVKStatus, lies.obMySQL, False)
   roNa(akt).Hausnr = doUmwfSQL(rs!Hausnr, lies.obMySQL, False)
   roNa(akt).geschlecht = doUmwfSQL(rs!geschlecht, lies.obMySQL, False)
   roNa(akt).plz = doUmwfSQL(rs!plz, lies.obMySQL, False)
   roNa(akt).ort = doUmwfSQL(rs!ort, lies.obMySQL, False)
   roNa(akt).Lkz = doUmwfSQL(rs!Lkz, lies.obMySQL, False)
   roNa(akt).Anschrzus = doUmwfSQL(rs!Anschrzus, lies.obMySQL, False)
   roNa(akt).NVors = doUmwfSQL(rs!NVors, lies.obMySQL, False)
   roNa(akt).PFPlz = doUmwfSQL(rs!PFPlz, lies.obMySQL, False)
   roNa(akt).PFOrt = doUmwfSQL(rs!PFOrt, lies.obMySQL, False)
   roNa(akt).PFNr = doUmwfSQL(rs!PFNr, lies.obMySQL, False)
   roNa(akt).f3124 = doUmwfSQL(rs!f3124, lies.obMySQL, False)
   roNa(akt).AnschrZus_2 = doUmwfSQL(rs!AnschrZus_2, lies.obMySQL, False)
   roNa(akt).Postfach_2 = doUmwfSQL(rs!Postfach_2, lies.obMySQL, False)
   roNa(akt).LK_2 = doUmwfSQL(rs!LK_2, lies.obMySQL, False)
   roNa(akt).Postfach = doUmwfSQL(rs!Postfach, lies.obMySQL, False)
   roNa(akt).Beruf = doUmwfSQL(rs!Beruf, lies.obMySQL, False)
   roNa(akt).Weggeldzone = doUmwfSQL(rs!Weggeldzone, lies.obMySQL, False)
   roNa(akt).WeggzZahl = rs!WeggzZahl
   roNa(akt).AufnDat = rs!AufnDat
   roNa(akt).kAufDat = rs!kAufDat
   roNa(akt).Lanr = doUmwfSQL(rs!Lanr, lies.obMySQL, False)
   roNa(akt).BStNr = doUmwfSQL(rs!BStNr, lies.obMySQL, False)
   roNa(akt).Titel = doUmwfSQL(rs!Titel, lies.obMySQL, False)
   roNa(akt).Versichertennummer = doUmwfSQL(rs!Versichertennummer, lies.obMySQL, False)
   roNa(akt).PrivatTel = doUmwfSQL(rs!PrivatTel, lies.obMySQL, False)
   roNa(akt).KVNr = doUmwfSQL(rs!KVNr, lies.obMySQL, False)
   roNa(akt).KVNr2 = doUmwfSQL(rs!KVNr2, lies.obMySQL, False)
   roNa(akt).KVNr3 = doUmwfSQL(rs!KVNr3, lies.obMySQL, False)
   roNa(akt).KVNr4 = doUmwfSQL(rs!KVNr4, lies.obMySQL, False)
   roNa(akt).PrivatTel_2 = doUmwfSQL(rs!PrivatTel_2, lies.obMySQL, False)
   roNa(akt).PrivatFax = doUmwfSQL(rs!PrivatFax, lies.obMySQL, False)
   roNa(akt).DienstTel = doUmwfSQL(rs!DienstTel, lies.obMySQL, False)
   roNa(akt).PrivatMobil = doUmwfSQL(rs!PrivatMobil, lies.obMySQL, False)
   roNa(akt).email = doUmwfSQL(rs!email, lies.obMySQL, False)
   roNa(akt).Arbeitgeber = doUmwfSQL(rs!Arbeitgeber, lies.obMySQL, False)
   roNa(akt).AnAllgda = rs!AnAllgda
   roNa(akt).An1da = rs!An1da
   roNa(akt).An2da = rs!An2da
   roNa(akt).Checkda = rs!Checkda
   roNa(akt).DMTypaD = doUmwfSQL(rs!DMTypaD, lies.obMySQL, False)
   roNa(akt).AktZeit = rs!AktZeit
   roNa(akt).absPos = rs!absPos
   roNa(akt).StByte = rs!StByte
   roNa(akt).StByteA = rs!StByteA
   roNa(akt).Cave = doUmwfSQL(rs!Cave, lies.obMySQL, False)
   roNa(akt).Notiz = doUmwfSQL(rs!Notiz, lies.obMySQL, False)
   roNa(akt).obChk = doUmwfSQL(rs!obChk, lies.obMySQL, False)
   roNa(akt).dmpklass = rs!dmpklass
   roNa(akt).dmpbeg = rs!dmpbeg
   roNa(akt).dmpkhkklass = rs!dmpkhkklass
   roNa(akt).dmpkhkbeg = rs!dmpkhkbeg
   roNa(akt).dmpcopdklass = rs!dmpcopdklass
   roNa(akt).dmpcopdbeg = rs!dmpcopdbeg
   roNa(akt).dmpabklass = rs!dmpabklass
   roNa(akt).dmpabbeg = rs!dmpabbeg
   roNa(akt).dakab = rs!dakab
   roNa(akt).HzV = rs!HzV
   roNa(akt).HzVbeg = rs!HzVbeg
   roNa(akt).DS = rs!DS
   roNa(akt).DSbeg = rs!DSbeg
   roNa(akt).getHA0 = rs!getHA0
   roNa(akt).fnHA0 = doUmwfSQL(rs!fnHA0, lies.obMySQL, False)
   roNa(akt).getHA1 = rs!getHA1
   roNa(akt).fnHA1 = doUmwfSQL(rs!fnHA1, lies.obMySQL, False)
   roNa(akt).getHA2 = rs!getHA2
   roNa(akt).fnHA2 = doUmwfSQL(rs!fnHA2, lies.obMySQL, False)
   roNa(akt).zubenach = doUmwfSQL(rs!zubenach, lies.obMySQL, False)
   roNa(akt).Verwandt = doUmwfSQL(rs!Verwandt, lies.obMySQL, False)
   roNa(akt).Sprache = doUmwfSQL(rs!Sprache, lies.obMySQL, False)
   roNa(akt).lAktTM = rs!lAktTM
   roNa(akt).Mitarbeiter = rs!Mitarbeiter
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roNa(UBound(roNa) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenLaden

Function namenEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rNa) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rNa)
   If rNa(ri).kAufDat >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roNa)
    If roNa(robeg).kAufDat >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roNa(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roNa(robeg + UBound(rNa) - rbeg)
   For ri = rbeg To UBound(rNa)
    Call roNaZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rNa = roNa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenEinf

Public Function namenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rNa) + 1 & " Sätze in `namen`"
' sql0 = " INSERT " & sqlignore &  "INTO `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,obChk,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,HzV,HzVbeg,DS,DSbeg,getHA0,fnHA0,getHA1," & _
     "fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM,Mitarbeiter) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,obChk,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,HzV,HzVbeg,DS,DSbeg,getHA0,fnHA0,getHA1," & _
     "fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM,Mitarbeiter)               VALUES"))
 For i = 0 To UBound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rNa(i).Pat_id, ",", rNa(i).lfdnr, ",'", rNa(i).NVorsatz, "','", rNa(i).Nachname, "','", rNa(i).Vorname, "',", DatFor_k(rNa(i).GebDat), ",'", rNa(i).f3004, "','", rNa(i).f3006, "','", _
   rNa(i).Straße, "','", rNa(i).KVKStatus, "','", rNa(i).Hausnr, "','", rNa(i).geschlecht, "','", rNa(i).plz, "','", rNa(i).ort, "','", rNa(i).Lkz, "','", rNa(i).Anschrzus, "','", _
   rNa(i).NVors, "','", rNa(i).PFPlz, "','", rNa(i).PFOrt, "','", rNa(i).PFNr, "','", rNa(i).f3124, "','", rNa(i).AnschrZus_2, "','", rNa(i).Postfach_2, "','", rNa(i).LK_2, "','", rNa(i).Postfach, "','", _
   rNa(i).Beruf, "','", rNa(i).Weggeldzone, "',", rNa(i).WeggzZahl, ",", DatFor_k(rNa(i).AufnDat), ",", DatFor_k(rNa(i).kAufDat), ",'", rNa(i).Lanr, "','", rNa(i).BStNr, "','", _
   rNa(i).Titel, "','", rNa(i).Versichertennummer, "','", rNa(i).PrivatTel, "','", rNa(i).KVNr, "','", rNa(i).KVNr2, "','", rNa(i).KVNr3, "','", rNa(i).KVNr4, "','", rNa(i).PrivatTel_2, "','", _
   rNa(i).PrivatFax, "','", rNa(i).DienstTel, "','", rNa(i).PrivatMobil, "','", rNa(i).email, "','", rNa(i).Arbeitgeber, "',", CStr(-(rNa(i).AnAllgda <> 0)), ",", CStr(-(rNa(i).An1da <> 0)), ",", CStr(-( _
   rNa(i).An2da <> 0)), ",", CStr(-(rNa(i).Checkda <> 0)), ",'", rNa(i).DMTypaD, "',", DatFor_k(0), ",", rNa(i).absPos, ",", rNa(i).StByte, ",", rNa(i).StByteA, ",'", _
   rNa(i).Cave, "','", rNa(i).Notiz, "','", rNa(i).obChk, "',", rNa(i).dmpklass, ",", DatFor_k(rNa(i).dmpbeg), ",", rNa(i).dmpkhkklass, ",", DatFor_k(rNa(i).dmpkhkbeg), ",", rNa(i).dmpcopdklass, ",", DatFor_k( _
   rNa(i).dmpcopdbeg), ",", rNa(i).dmpabklass, ",", DatFor_k(rNa(i).dmpabbeg), ",", DatFor_k(rNa(i).dakab), ",", rNa(i).HzV, ",", DatFor_k(rNa(i).HzVbeg), ",", rNa(i).DS, ",", DatFor_k( _
   rNa(i).DSbeg), ",", rNa(i).getHA0, ",'", rNa(i).fnHA0, "',", rNa(i).getHA1, ",'", rNa(i).fnHA1, "',", rNa(i).getHA2, ",'", rNa(i).fnHA2, "','", rNa(i).zubenach, "','", _
   rNa(i).Verwandt, "','", rNa(i).Sprache, "',", DatFor_k(rNa(i).lAktTM), ",", rNa(i).Mitarbeiter, ")")
  If SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rNa) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `namen` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rNa)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""namenSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(48)
 For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rNa), i)
  If Len(rNa(k).NVorsatz) > maxi(0) Then maxi(0) = Len(rNa(k).NVorsatz)
  If Len(rNa(k).Nachname) > maxi(1) Then maxi(1) = Len(rNa(k).Nachname)
  If Len(rNa(k).Vorname) > maxi(2) Then maxi(2) = Len(rNa(k).Vorname)
  If Len(rNa(k).f3004) > maxi(3) Then maxi(3) = Len(rNa(k).f3004)
  If Len(rNa(k).f3006) > maxi(4) Then maxi(4) = Len(rNa(k).f3006)
  If Len(rNa(k).Straße) > maxi(5) Then maxi(5) = Len(rNa(k).Straße)
  If Len(rNa(k).KVKStatus) > maxi(6) Then maxi(6) = Len(rNa(k).KVKStatus)
  If Len(rNa(k).Hausnr) > maxi(7) Then maxi(7) = Len(rNa(k).Hausnr)
  If Len(rNa(k).geschlecht) > maxi(8) Then maxi(8) = Len(rNa(k).geschlecht)
  If Len(rNa(k).plz) > maxi(9) Then maxi(9) = Len(rNa(k).plz)
  If Len(rNa(k).ort) > maxi(10) Then maxi(10) = Len(rNa(k).ort)
  If Len(rNa(k).Lkz) > maxi(11) Then maxi(11) = Len(rNa(k).Lkz)
  If Len(rNa(k).Anschrzus) > maxi(12) Then maxi(12) = Len(rNa(k).Anschrzus)
  If Len(rNa(k).NVors) > maxi(13) Then maxi(13) = Len(rNa(k).NVors)
  If Len(rNa(k).PFPlz) > maxi(14) Then maxi(14) = Len(rNa(k).PFPlz)
  If Len(rNa(k).PFOrt) > maxi(15) Then maxi(15) = Len(rNa(k).PFOrt)
  If Len(rNa(k).PFNr) > maxi(16) Then maxi(16) = Len(rNa(k).PFNr)
  If Len(rNa(k).f3124) > maxi(17) Then maxi(17) = Len(rNa(k).f3124)
  If Len(rNa(k).AnschrZus_2) > maxi(18) Then maxi(18) = Len(rNa(k).AnschrZus_2)
  If Len(rNa(k).Postfach_2) > maxi(19) Then maxi(19) = Len(rNa(k).Postfach_2)
  If Len(rNa(k).LK_2) > maxi(20) Then maxi(20) = Len(rNa(k).LK_2)
  If Len(rNa(k).Postfach) > maxi(21) Then maxi(21) = Len(rNa(k).Postfach)
  If Len(rNa(k).Beruf) > maxi(22) Then maxi(22) = Len(rNa(k).Beruf)
  If Len(rNa(k).Weggeldzone) > maxi(23) Then maxi(23) = Len(rNa(k).Weggeldzone)
  If Len(rNa(k).Lanr) > maxi(24) Then maxi(24) = Len(rNa(k).Lanr)
  If Len(rNa(k).BStNr) > maxi(25) Then maxi(25) = Len(rNa(k).BStNr)
  If Len(rNa(k).Titel) > maxi(26) Then maxi(26) = Len(rNa(k).Titel)
  If Len(rNa(k).Versichertennummer) > maxi(27) Then maxi(27) = Len(rNa(k).Versichertennummer)
  If Len(rNa(k).PrivatTel) > maxi(28) Then maxi(28) = Len(rNa(k).PrivatTel)
  If Len(rNa(k).KVNr) > maxi(29) Then maxi(29) = Len(rNa(k).KVNr)
  If Len(rNa(k).KVNr2) > maxi(30) Then maxi(30) = Len(rNa(k).KVNr2)
  If Len(rNa(k).KVNr3) > maxi(31) Then maxi(31) = Len(rNa(k).KVNr3)
  If Len(rNa(k).KVNr4) > maxi(32) Then maxi(32) = Len(rNa(k).KVNr4)
  If Len(rNa(k).PrivatTel_2) > maxi(33) Then maxi(33) = Len(rNa(k).PrivatTel_2)
  If Len(rNa(k).PrivatFax) > maxi(34) Then maxi(34) = Len(rNa(k).PrivatFax)
  If Len(rNa(k).DienstTel) > maxi(35) Then maxi(35) = Len(rNa(k).DienstTel)
  If Len(rNa(k).PrivatMobil) > maxi(36) Then maxi(36) = Len(rNa(k).PrivatMobil)
  If Len(rNa(k).email) > maxi(37) Then maxi(37) = Len(rNa(k).email)
  If Len(rNa(k).Arbeitgeber) > maxi(38) Then maxi(38) = Len(rNa(k).Arbeitgeber)
  If Len(rNa(k).DMTypaD) > maxi(39) Then maxi(39) = Len(rNa(k).DMTypaD)
  If Len(rNa(k).Cave) > maxi(40) Then maxi(40) = Len(rNa(k).Cave)
  If Len(rNa(k).Notiz) > maxi(41) Then maxi(41) = Len(rNa(k).Notiz)
  If Len(rNa(k).obChk) > maxi(42) Then maxi(42) = Len(rNa(k).obChk)
  If Len(rNa(k).fnHA0) > maxi(43) Then maxi(43) = Len(rNa(k).fnHA0)
  If Len(rNa(k).fnHA1) > maxi(44) Then maxi(44) = Len(rNa(k).fnHA1)
  If Len(rNa(k).fnHA2) > maxi(45) Then maxi(45) = Len(rNa(k).fnHA2)
  If Len(rNa(k).zubenach) > maxi(46) Then maxi(46) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(47) Then maxi(47) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(48) Then maxi(48) = Len(rNa(k).Sprache)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rNa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & Left$(rNa(k).NVorsatz, maxL) & "'", True: rNa(k).NVorsatz = Left$(rNa(k).NVorsatz, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & Left$(rNa(k).Nachname, maxL) & "'", True: rNa(k).Nachname = Left$(rNa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & Left$(rNa(k).Vorname, maxL) & "'", True: rNa(k).Vorname = Left$(rNa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3004: '" & rNa(k).f3004 & "' -> '" & Left$(rNa(k).f3004, maxL) & "'", True: rNa(k).f3004 = Left$(rNa(k).f3004, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3006: '" & rNa(k).f3006 & "' -> '" & Left$(rNa(k).f3006, maxL) & "'", True: rNa(k).f3006 = Left$(rNa(k).f3006, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & Left$(rNa(k).Straße, maxL) & "'", True: rNa(k).Straße = Left$(rNa(k).Straße, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & Left$(rNa(k).KVKStatus, maxL) & "'", True: rNa(k).KVKStatus = Left$(rNa(k).KVKStatus, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Hausnr: '" & rNa(k).Hausnr & "' -> '" & Left$(rNa(k).Hausnr, maxL) & "'", True: rNa(k).Hausnr = Left$(rNa(k).Hausnr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).geschlecht & "' -> '" & Left$(rNa(k).geschlecht, maxL) & "'", True: rNa(k).geschlecht = Left$(rNa(k).geschlecht, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).plz & "' -> '" & Left$(rNa(k).plz, maxL) & "'", True: rNa(k).plz = Left$(rNa(k).plz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).ort & "' -> '" & Left$(rNa(k).ort, maxL) & "'", True: rNa(k).ort = Left$(rNa(k).ort, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rNa.Lkz: '" & rNa(k).Lkz & "' -> '" & Left$(rNa(k).Lkz, maxL) & "'", True: rNa(k).Lkz = Left$(rNa(k).Lkz, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rNa.Anschrzus: '" & rNa(k).Anschrzus & "' -> '" & Left$(rNa(k).Anschrzus, maxL) & "'", True: rNa(k).Anschrzus = Left$(rNa(k).Anschrzus, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVors: '" & rNa(k).NVors & "' -> '" & Left$(rNa(k).NVors, maxL) & "'", True: rNa(k).NVors = Left$(rNa(k).NVors, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFPlz: '" & rNa(k).PFPlz & "' -> '" & Left$(rNa(k).PFPlz, maxL) & "'", True: rNa(k).PFPlz = Left$(rNa(k).PFPlz, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFOrt: '" & rNa(k).PFOrt & "' -> '" & Left$(rNa(k).PFOrt, maxL) & "'", True: rNa(k).PFOrt = Left$(rNa(k).PFOrt, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFNr: '" & rNa(k).PFNr & "' -> '" & Left$(rNa(k).PFNr, maxL) & "'", True: rNa(k).PFNr = Left$(rNa(k).PFNr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3124: '" & rNa(k).f3124 & "' -> '" & Left$(rNa(k).f3124, maxL) & "'", True: rNa(k).f3124 = Left$(rNa(k).f3124, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rNa.AnschrZus_2: '" & rNa(k).AnschrZus_2 & "' -> '" & Left$(rNa(k).AnschrZus_2, maxL) & "'", True: rNa(k).AnschrZus_2 = Left$(rNa(k).AnschrZus_2, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach_2: '" & rNa(k).Postfach_2 & "' -> '" & Left$(rNa(k).Postfach_2, maxL) & "'", True: rNa(k).Postfach_2 = Left$(rNa(k).Postfach_2, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rNa.LK_2: '" & rNa(k).LK_2 & "' -> '" & Left$(rNa(k).LK_2, maxL) & "'", True: rNa(k).LK_2 = Left$(rNa(k).LK_2, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach: '" & rNa(k).Postfach & "' -> '" & Left$(rNa(k).Postfach, maxL) & "'", True: rNa(k).Postfach = Left$(rNa(k).Postfach, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rNa.Beruf: '" & rNa(k).Beruf & "' -> '" & Left$(rNa(k).Beruf, maxL) & "'", True: rNa(k).Beruf = Left$(rNa(k).Beruf, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & Left$(rNa(k).Weggeldzone, maxL) & "'", True: rNa(k).Weggeldzone = Left$(rNa(k).Weggeldzone, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).Lanr & "' -> '" & Left$(rNa(k).Lanr, maxL) & "'", True: rNa(k).Lanr = Left$(rNa(k).Lanr, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.BStNr: '" & rNa(k).BStNr & "' -> '" & Left$(rNa(k).BStNr, maxL) & "'", True: rNa(k).BStNr = Left$(rNa(k).BStNr, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.Titel: '" & rNa(k).Titel & "' -> '" & Left$(rNa(k).Titel, maxL) & "'", True: rNa(k).Titel = Left$(rNa(k).Titel, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.Versichertennummer: '" & rNa(k).Versichertennummer & "' -> '" & Left$(rNa(k).Versichertennummer, maxL) & "'", True: rNa(k).Versichertennummer = Left$(rNa(k).Versichertennummer, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel: '" & rNa(k).PrivatTel & "' -> '" & Left$(rNa(k).PrivatTel, maxL) & "'", True: rNa(k).PrivatTel = Left$(rNa(k).PrivatTel, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr: '" & rNa(k).KVNr & "' -> '" & Left$(rNa(k).KVNr, maxL) & "'", True: rNa(k).KVNr = Left$(rNa(k).KVNr, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr2: '" & rNa(k).KVNr2 & "' -> '" & Left$(rNa(k).KVNr2, maxL) & "'", True: rNa(k).KVNr2 = Left$(rNa(k).KVNr2, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr3: '" & rNa(k).KVNr3 & "' -> '" & Left$(rNa(k).KVNr3, maxL) & "'", True: rNa(k).KVNr3 = Left$(rNa(k).KVNr3, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr4: '" & rNa(k).KVNr4 & "' -> '" & Left$(rNa(k).KVNr4, maxL) & "'", True: rNa(k).KVNr4 = Left$(rNa(k).KVNr4, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel_2: '" & rNa(k).PrivatTel_2 & "' -> '" & Left$(rNa(k).PrivatTel_2, maxL) & "'", True: rNa(k).PrivatTel_2 = Left$(rNa(k).PrivatTel_2, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatFax: '" & rNa(k).PrivatFax & "' -> '" & Left$(rNa(k).PrivatFax, maxL) & "'", True: rNa(k).PrivatFax = Left$(rNa(k).PrivatFax, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rNa.DienstTel: '" & rNa(k).DienstTel & "' -> '" & Left$(rNa(k).DienstTel, maxL) & "'", True: rNa(k).DienstTel = Left$(rNa(k).DienstTel, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatMobil: '" & rNa(k).PrivatMobil & "' -> '" & Left$(rNa(k).PrivatMobil, maxL) & "'", True: rNa(k).PrivatMobil = Left$(rNa(k).PrivatMobil, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rNa.Email: '" & rNa(k).email & "' -> '" & Left$(rNa(k).email, maxL) & "'", True: rNa(k).email = Left$(rNa(k).email, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rNa.Arbeitgeber: '" & rNa(k).Arbeitgeber & "' -> '" & Left$(rNa(k).Arbeitgeber, maxL) & "'", True: rNa(k).Arbeitgeber = Left$(rNa(k).Arbeitgeber, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rNa.DMTypaD: '" & rNa(k).DMTypaD & "' -> '" & Left$(rNa(k).DMTypaD, maxL) & "'", True: rNa(k).DMTypaD = Left$(rNa(k).DMTypaD, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rNa.Cave: '" & rNa(k).Cave & "' -> '" & Left$(rNa(k).Cave, maxL) & "'", True: rNa(k).Cave = Left$(rNa(k).Cave, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rNa.Notiz: '" & rNa(k).Notiz & "' -> '" & Left$(rNa(k).Notiz, maxL) & "'", True: rNa(k).Notiz = Left$(rNa(k).Notiz, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rNa.obChk: '" & rNa(k).obChk & "' -> '" & Left$(rNa(k).obChk, maxL) & "'", True: rNa(k).obChk = Left$(rNa(k).obChk, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & Left$(rNa(k).fnHA0, maxL) & "'", True: rNa(k).fnHA0 = Left$(rNa(k).fnHA0, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & Left$(rNa(k).fnHA1, maxL) & "'", True: rNa(k).fnHA1 = Left$(rNa(k).fnHA1, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & Left$(rNa(k).fnHA2, maxL) & "'", True: rNa(k).fnHA2 = Left$(rNa(k).fnHA2, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL) & "'", True: rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL) & "'", True: rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL) & "'", True: rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
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
 Resume
ElseIf ErrNumber = -2147217871 Or ErrNumber = -2147217859 Or ErrNumber = -2147467259 Then
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function roFaZuw(i&, j&)
 roFa(i).FID = rFa(j).FID
 roFa(i).Pat_id = rFa(j).Pat_id
 roFa(i).Quartal = rFa(j).Quartal
 roFa(i).Nachname = rFa(j).Nachname
 roFa(i).Vorname = rFa(j).Vorname
 roFa(i).lfdnr = rFa(j).lfdnr
 roFa(i).TMFNr = rFa(j).TMFNr
 roFa(i).VKNr = rFa(j).VKNr
 roFa(i).f4131 = rFa(j).f4131
 roFa(i).f4132 = rFa(j).f4132
 roFa(i).VschBeg = rFa(j).VschBeg
 roFa(i).KKasse_2 = rFa(j).KKasse_2
 roFa(i).FaktPers = rFa(j).FaktPers
 roFa(i).FaktTechn = rFa(j).FaktTechn
 roFa(i).FaktLabor = rFa(j).FaktLabor
 roFa(i).BhFB = rFa(j).BhFB
 roFa(i).BhFE1 = rFa(j).BhFE1
 roFa(i).BhFE2 = rFa(j).BhFE2
 roFa(i).f4202 = rFa(j).f4202
 roFa(i).ausgst = rFa(j).ausgst
 roFa(i).KtrAbrB = rFa(j).KtrAbrB
 roFa(i).AbrAr = rFa(j).AbrAr
 roFa(i).lVorl = rFa(j).lVorl
 roFa(i).IK = rFa(j).IK
 roFa(i).KVKs = rFa(j).KVKs
 roFa(i).KVKserg = rFa(j).KVKserg
 roFa(i).Status = rFa(j).Status
 roFa(i).Kasse = rFa(j).Kasse
 roFa(i).KID = rFa(j).KID
 roFa(i).GebOr = rFa(j).GebOr
 roFa(i).AbrGb = rFa(j).AbrGb
 roFa(i).PersKreis = rFa(j).PersKreis
 roFa(i).SKtZusatz = rFa(j).SKtZusatz
 roFa(i).letzteRegel = rFa(j).letzteRegel
 roFa(i).ÜwText = rFa(j).ÜwText
 roFa(i).f4210 = rFa(j).f4210
 roFa(i).AkfHAH = rFa(j).AkfHAH
 roFa(i).AkfAB0 = rFa(j).AkfAB0
 roFa(i).AkfAK = rFa(j).AkfAK
 roFa(i).statNuller = rFa(j).statNuller
 roFa(i).ÜbwV = rFa(j).ÜbwV
 roFa(i).ÜbWVLANR = rFa(j).ÜbWVLANR
 roFa(i).ÜbWVBSNR = rFa(j).ÜbWVBSNR
 roFa(i).ÜbWVKVNR = rFa(j).ÜbWVKVNR
 roFa(i).AndÜw = rFa(j).AndÜw
 roFa(i).Übwr = rFa(j).Übwr
 roFa(i).ÜbwLANR = rFa(j).ÜbwLANR
 roFa(i).ÜWZiel = rFa(j).ÜWZiel
 roFa(i).ÜWNNr = rFa(j).ÜWNNr
 roFa(i).ÜWNaN = rFa(j).ÜWNaN
 roFa(i).ÜWTit = rFa(j).ÜWTit
 roFa(i).ÜWVor = rFa(j).ÜWVor
 roFa(i).ÜWVsw = rFa(j).ÜWVsw
 roFa(i).üwvid = rFa(j).üwvid
 roFa(i).Auftrag = rFa(j).Auftrag
 roFa(i).Verdacht = rFa(j).Verdacht
 roFa(i).Befund = rFa(j).Befund
 roFa(i).statKlasse = rFa(j).statKlasse
 roFa(i).f4237 = rFa(j).f4237
 roFa(i).statBehTage = rFa(j).statBehTage
 roFa(i).SchGr = rFa(j).SchGr
 roFa(i).Weiterbeh = rFa(j).Weiterbeh
 roFa(i).f4266 = rFa(j).f4266
 roFa(i).PGeb = rFa(j).PGeb
 roFa(i).PGebErg = rFa(j).PGebErg
 roFa(i).Mahnfrist = rFa(j).Mahnfrist
 roFa(i).Unfallort = rFa(j).Unfallort
 roFa(i).BeschAls = rFa(j).BeschAls
 roFa(i).BeschSeit = rFa(j).BeschSeit
 roFa(i).Unfallbetrieb = rFa(j).Unfallbetrieb
 roFa(i).f4570 = rFa(j).f4570
 roFa(i).GOÄKatNr = rFa(j).GOÄKatNr
 roFa(i).GOÄKatName = rFa(j).GOÄKatName
 roFa(i).abrArzt = rFa(j).abrArzt
 roFa(i).privVers = rFa(j).privVers
 roFa(i).AdNam = rFa(j).AdNam
 roFa(i).AdStr = rFa(j).AdStr
 roFa(i).AdPlz = rFa(j).AdPlz
 roFa(i).AdOrt = rFa(j).AdOrt
 roFa(i).ÜwBG = rFa(j).ÜwBG
 roFa(i).BhFE = rFa(j).BhFE
 roFa(i).s8000 = rFa(j).s8000
 roFa(i).s8100 = rFa(j).s8100
 roFa(i).AktZeit = rFa(j).AktZeit
 roFa(i).Fanf = rFa(j).Fanf
 roFa(i).altQuart = rFa(j).altQuart
 roFa(i).QAnf = rFa(j).QAnf
 roFa(i).QEnd = rFa(j).QEnd
 roFa(i).QS = rFa(j).QS
 roFa(i).QT = rFa(j).QT
 roFa(i).StByte = rFa(j).StByte
 roFa(i).absPos = rFa(j).absPos
 roFa(i).lanrid = rFa(j).lanrid
 roFa(i).f4108 = rFa(j).f4108
 roFa(i).BGFallNr = rFa(j).BGFallNr
 roFa(i).lGewicht = rFa(j).lGewicht
 roFa(i).vorET = rFa(j).vorET
 roFa(i).dmpVertret = rFa(j).dmpVertret
 roFa(i).dmpArztw = rFa(j).dmpArztw
 roFa(i).dmpHypos = rFa(j).dmpHypos
 roFa(i).dmpKhsA = rFa(j).dmpKhsA
 roFa(i).dmpDMSchulEmpf = rFa(j).dmpDMSchulEmpf
 roFa(i).dmpDMSchulWahrg = rFa(j).dmpDMSchulWahrg
 roFa(i).dmpHypertSchulEmpf = rFa(j).dmpHypertSchulEmpf
 roFa(i).dmpHypertSchulWahrg = rFa(j).dmpHypertSchulWahrg
 roFa(i).dmpKKTabakEmpf = rFa(j).dmpKKTabakEmpf
 roFa(i).dmpKKErnEmpf = rFa(j).dmpKKErnEmpf
 roFa(i).dmpKKkTrainEmpf = rFa(j).dmpKKkTrainEmpf
 roFa(i).dmpHbA1cZiel = rFa(j).dmpHbA1cZiel
 roFa(i).dmpUewFuss = rFa(j).dmpUewFuss
 roFa(i).dmpEinwDM = rFa(j).dmpEinwDM
 roFa(i).dmphalbj = rFa(j).dmphalbj
 roFa(i).dmpMA = rFa(j).dmpMA
End Function ' roFaZuw

Public Function FaZUnt%(i&, j&)
 If roFa(i).FID <> rFa(j).FID Then GoSub unter
 If roFa(i).Pat_id <> rFa(j).Pat_id Then GoSub unter
 If roFa(i).Quartal <> rFa(j).Quartal Then GoSub unter
 If roFa(i).Nachname <> rFa(j).Nachname Then GoSub unter
 If roFa(i).Vorname <> rFa(j).Vorname Then GoSub unter
 If roFa(i).lfdnr <> rFa(j).lfdnr Then GoSub unter
 If roFa(i).TMFNr <> rFa(j).TMFNr Then GoSub unter
 If roFa(i).VKNr <> rFa(j).VKNr Then GoSub unter
 If roFa(i).f4131 <> rFa(j).f4131 Then GoSub unter
 If roFa(i).f4132 <> rFa(j).f4132 Then GoSub unter
 If roFa(i).VschBeg <> rFa(j).VschBeg Then GoSub unter
 If roFa(i).KKasse_2 <> rFa(j).KKasse_2 Then GoSub unter
 If roFa(i).FaktPers <> rFa(j).FaktPers Then GoSub unter
 If roFa(i).FaktTechn <> rFa(j).FaktTechn Then GoSub unter
 If roFa(i).FaktLabor <> rFa(j).FaktLabor Then GoSub unter
 If roFa(i).BhFB <> rFa(j).BhFB Then GoSub unter
 If roFa(i).BhFE1 <> rFa(j).BhFE1 Then GoSub unter
 If roFa(i).BhFE2 <> rFa(j).BhFE2 Then GoSub unter
 If roFa(i).f4202 <> rFa(j).f4202 Then GoSub unter
 If roFa(i).ausgst <> rFa(j).ausgst Then GoSub unter
 If roFa(i).KtrAbrB <> rFa(j).KtrAbrB Then GoSub unter
 If roFa(i).AbrAr <> rFa(j).AbrAr Then GoSub unter
 If roFa(i).lVorl <> rFa(j).lVorl Then GoSub unter
 If roFa(i).IK <> rFa(j).IK Then GoSub unter
 If roFa(i).KVKs <> rFa(j).KVKs Then GoSub unter
 If roFa(i).KVKserg <> rFa(j).KVKserg Then GoSub unter
 If roFa(i).Status <> rFa(j).Status Then GoSub unter
 If roFa(i).Kasse <> rFa(j).Kasse Then GoSub unter
 If roFa(i).KID <> rFa(j).KID Then GoSub unter
 If roFa(i).GebOr <> rFa(j).GebOr Then GoSub unter
 If roFa(i).AbrGb <> rFa(j).AbrGb Then GoSub unter
 If roFa(i).PersKreis <> rFa(j).PersKreis Then GoSub unter
 If roFa(i).SKtZusatz <> rFa(j).SKtZusatz Then GoSub unter
 If roFa(i).letzteRegel <> rFa(j).letzteRegel Then GoSub unter
 If roFa(i).ÜwText <> rFa(j).ÜwText Then GoSub unter
 If roFa(i).f4210 <> rFa(j).f4210 Then GoSub unter
 If roFa(i).AkfHAH <> rFa(j).AkfHAH Then GoSub unter
 If roFa(i).AkfAB0 <> rFa(j).AkfAB0 Then GoSub unter
 If roFa(i).AkfAK <> rFa(j).AkfAK Then GoSub unter
 If roFa(i).statNuller <> rFa(j).statNuller Then GoSub unter
 If roFa(i).ÜbwV <> rFa(j).ÜbwV Then GoSub unter
 If roFa(i).ÜbWVLANR <> rFa(j).ÜbWVLANR Then GoSub unter
 If roFa(i).ÜbWVBSNR <> rFa(j).ÜbWVBSNR Then GoSub unter
 If roFa(i).ÜbWVKVNR <> rFa(j).ÜbWVKVNR Then GoSub unter
 If roFa(i).AndÜw <> rFa(j).AndÜw Then GoSub unter
 If roFa(i).Übwr <> rFa(j).Übwr Then GoSub unter
 If roFa(i).ÜbwLANR <> rFa(j).ÜbwLANR Then GoSub unter
 If roFa(i).ÜWZiel <> rFa(j).ÜWZiel Then GoSub unter
 If roFa(i).ÜWNNr <> rFa(j).ÜWNNr Then GoSub unter
 If roFa(i).ÜWNaN <> rFa(j).ÜWNaN Then GoSub unter
 If roFa(i).ÜWTit <> rFa(j).ÜWTit Then GoSub unter
 If roFa(i).ÜWVor <> rFa(j).ÜWVor Then GoSub unter
 If roFa(i).ÜWVsw <> rFa(j).ÜWVsw Then GoSub unter
 If roFa(i).üwvid <> rFa(j).üwvid Then GoSub unter
 If roFa(i).Auftrag <> rFa(j).Auftrag Then GoSub unter
 If roFa(i).Verdacht <> rFa(j).Verdacht Then GoSub unter
 If roFa(i).Befund <> rFa(j).Befund Then GoSub unter
 If roFa(i).statKlasse <> rFa(j).statKlasse Then GoSub unter
 If roFa(i).f4237 <> rFa(j).f4237 Then GoSub unter
 If roFa(i).statBehTage <> rFa(j).statBehTage Then GoSub unter
 If roFa(i).SchGr <> rFa(j).SchGr Then GoSub unter
 If roFa(i).Weiterbeh <> rFa(j).Weiterbeh Then GoSub unter
 If roFa(i).f4266 <> rFa(j).f4266 Then GoSub unter
 If roFa(i).PGeb <> rFa(j).PGeb Then GoSub unter
 If roFa(i).PGebErg <> rFa(j).PGebErg Then GoSub unter
 If roFa(i).Mahnfrist <> rFa(j).Mahnfrist Then GoSub unter
 If roFa(i).Unfallort <> rFa(j).Unfallort Then GoSub unter
 If roFa(i).BeschAls <> rFa(j).BeschAls Then GoSub unter
 If roFa(i).BeschSeit <> rFa(j).BeschSeit Then GoSub unter
 If roFa(i).Unfallbetrieb <> rFa(j).Unfallbetrieb Then GoSub unter
 If roFa(i).f4570 <> rFa(j).f4570 Then GoSub unter
 If roFa(i).GOÄKatNr <> rFa(j).GOÄKatNr Then GoSub unter
 If roFa(i).GOÄKatName <> rFa(j).GOÄKatName Then GoSub unter
 If roFa(i).abrArzt <> rFa(j).abrArzt Then GoSub unter
 If roFa(i).privVers <> rFa(j).privVers Then GoSub unter
 If roFa(i).AdNam <> rFa(j).AdNam Then GoSub unter
 If roFa(i).AdStr <> rFa(j).AdStr Then GoSub unter
 If roFa(i).AdPlz <> rFa(j).AdPlz Then GoSub unter
 If roFa(i).AdOrt <> rFa(j).AdOrt Then GoSub unter
 If roFa(i).ÜwBG <> rFa(j).ÜwBG Then GoSub unter
 If roFa(i).BhFE <> rFa(j).BhFE Then GoSub unter
 If roFa(i).s8000 <> rFa(j).s8000 Then GoSub unter
 If roFa(i).s8100 <> rFa(j).s8100 Then GoSub unter
 If roFa(i).AktZeit <> rFa(j).AktZeit Then GoSub unter
 If roFa(i).Fanf <> rFa(j).Fanf Then GoSub unter
 If roFa(i).altQuart <> rFa(j).altQuart Then GoSub unter
 If roFa(i).QAnf <> rFa(j).QAnf Then GoSub unter
 If roFa(i).QEnd <> rFa(j).QEnd Then GoSub unter
 If roFa(i).QS <> rFa(j).QS Then GoSub unter
 If roFa(i).QT <> rFa(j).QT Then GoSub unter
 If roFa(i).StByte <> rFa(j).StByte Then GoSub unter
 If roFa(i).absPos <> rFa(j).absPos Then GoSub unter
 If roFa(i).lanrid <> rFa(j).lanrid Then GoSub unter
 If roFa(i).f4108 <> rFa(j).f4108 Then GoSub unter
 If roFa(i).BGFallNr <> rFa(j).BGFallNr Then GoSub unter
 If roFa(i).lGewicht <> rFa(j).lGewicht Then GoSub unter
 If roFa(i).vorET <> rFa(j).vorET Then GoSub unter
 If roFa(i).dmpVertret <> rFa(j).dmpVertret Then GoSub unter
 If roFa(i).dmpArztw <> rFa(j).dmpArztw Then GoSub unter
 If roFa(i).dmpHypos <> rFa(j).dmpHypos Then GoSub unter
 If roFa(i).dmpKhsA <> rFa(j).dmpKhsA Then GoSub unter
 If roFa(i).dmpDMSchulEmpf <> rFa(j).dmpDMSchulEmpf Then GoSub unter
 If roFa(i).dmpDMSchulWahrg <> rFa(j).dmpDMSchulWahrg Then GoSub unter
 If roFa(i).dmpHypertSchulEmpf <> rFa(j).dmpHypertSchulEmpf Then GoSub unter
 If roFa(i).dmpHypertSchulWahrg <> rFa(j).dmpHypertSchulWahrg Then GoSub unter
 If roFa(i).dmpKKTabakEmpf <> rFa(j).dmpKKTabakEmpf Then GoSub unter
 If roFa(i).dmpKKErnEmpf <> rFa(j).dmpKKErnEmpf Then GoSub unter
 If roFa(i).dmpKKkTrainEmpf <> rFa(j).dmpKKkTrainEmpf Then GoSub unter
 If roFa(i).dmpHbA1cZiel <> rFa(j).dmpHbA1cZiel Then GoSub unter
 If roFa(i).dmpUewFuss <> rFa(j).dmpUewFuss Then GoSub unter
 If roFa(i).dmpEinwDM <> rFa(j).dmpEinwDM Then GoSub unter
 If roFa(i).dmphalbj <> rFa(j).dmphalbj Then GoSub unter
 If roFa(i).dmpMA <> rFa(j).dmpMA Then GoSub unter
 Exit Function
unter:
 FaZUnt = FaZUnt + 1
 Return
End Function ' FaZUnt

Public Function faelleLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(Quartal,'') Quartal,COALESCE(Nachname,'') Nachname" & _
",COALESCE(Vorname,'') Vorname,COALESCE(lfdnr,0) lfdnr,COALESCE(TMFNr,'') TMFNr,COALESCE(VKNr,'') VKNr" & _
",COALESCE(f4131,'') f4131,COALESCE(f4132,'') f4132,COALESCE(VschBeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) VschBeg,COALESCE(KKasse_2,'') KKasse_2" & _
",COALESCE(FaktPers,0) FaktPers,COALESCE(FaktTechn,0) FaktTechn,COALESCE(FaktLabor,0) FaktLabor,COALESCE(BhFB - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFB" & _
",COALESCE(BhFE1 - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE1,COALESCE(BhFE2 - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE2,COALESCE(f4202,'') f4202,COALESCE(ausgst - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ausgst" & _
",COALESCE(KtrAbrB,'') KtrAbrB,COALESCE(AbrAr,'') AbrAr,COALESCE(lVorl - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lVorl,COALESCE(IK,'') IK" & _
",COALESCE(KVKs,'') KVKs,COALESCE(KVKserg,'') KVKserg,COALESCE(Status,'') Status,COALESCE(Kasse,'') Kasse" & _
",COALESCE(KID,0) KID,COALESCE(GebOr,'') GebOr,COALESCE(AbrGb,'') AbrGb,COALESCE(PersKreis,'') PersKreis" & _
",COALESCE(SKtZusatz,'') SKtZusatz,COALESCE(letzteRegel,'') letzteRegel,COALESCE(ÜwText,'') ÜwText,COALESCE(f4210,0) f4210" & _
",COALESCE(AkfHAH,0) AkfHAH,COALESCE(AkfAB0,0) AkfAB0,COALESCE(AkfAK,0) AkfAK,COALESCE(statNuller,'') statNuller" & _
",COALESCE(ÜbwV,'') ÜbwV,COALESCE(ÜbWVLANR,'') ÜbWVLANR,COALESCE(ÜbWVBSNR,'') ÜbWVBSNR,COALESCE(ÜbWVKVNR,'') ÜbWVKVNR" & _
",COALESCE(AndÜw,'') AndÜw,COALESCE(Übwr,'') Übwr,COALESCE(ÜbwLANR,'') ÜbwLANR,COALESCE(ÜWZiel,'') ÜWZiel" & _
",COALESCE(ÜWNNr,'') ÜWNNr,COALESCE(ÜWNaN,'') ÜWNaN,COALESCE(ÜWTit,'') ÜWTit,COALESCE(ÜWVor,'') ÜWVor" & _
",COALESCE(ÜWVsw,'') ÜWVsw,COALESCE(üwvid,0) üwvid,COALESCE(Auftrag,'') Auftrag,COALESCE(Verdacht,'') Verdacht" & _
",COALESCE(Befund,'') Befund,COALESCE(statKlasse,'') statKlasse,COALESCE(f4237,'') f4237,COALESCE(statBehTage,0) statBehTage" & _
",COALESCE(SchGr,0) SchGr,COALESCE(Weiterbeh,'') Weiterbeh,COALESCE(f4266,0) f4266,COALESCE(PGeb,'') PGeb" & _
",COALESCE(PGebErg,'') PGebErg,COALESCE(Mahnfrist,'') Mahnfrist,COALESCE(Unfallort,'') Unfallort,COALESCE(BeschAls,'') BeschAls" & _
",COALESCE(BeschSeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BeschSeit,COALESCE(Unfallbetrieb,'') Unfallbetrieb,COALESCE(f4570,'') f4570,COALESCE(GOÄKatNr,'') GOÄKatNr" & _
",COALESCE(GOÄKatName,'') GOÄKatName,COALESCE(abrArzt,'') abrArzt,COALESCE(privVers,'') privVers,COALESCE(AdNam,'') AdNam" & _
",COALESCE(AdStr,'') AdStr,COALESCE(AdPlz,'') AdPlz,COALESCE(AdOrt,'') AdOrt,COALESCE(ÜwBG,'') ÜwBG" & _
",COALESCE(BhFE - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE,COALESCE(s8000,'') s8000,COALESCE(s8100,'') s8100,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(Fanf - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Fanf,COALESCE(altQuart,'') altQuart,COALESCE(QAnf - INTERVAL 0 DAY,CONVERT('18991230',DATE)) QAnf,COALESCE(QEnd - INTERVAL 0 DAY,CONVERT('18991230',DATE)) QEnd" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(absPos,0) absPos" & _
",COALESCE(LANRid,0) LANRid,COALESCE(f4108,'') f4108,COALESCE(BGFallNr,'') BGFallNr,COALESCE(lGewicht,0) lGewicht" & _
",COALESCE(vorET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) vorET,COALESCE(dmpVertret,'') dmpVertret,COALESCE(dmpArztw,'') dmpArztw,COALESCE(dmpHypos,'') dmpHypos"
sql = sql & ",COALESCE(dmpKhsA,'') dmpKhsA,COALESCE(dmpDMSchulEmpf,'') dmpDMSchulEmpf,COALESCE(dmpDMSchulWahrg,'') dmpDMSchulWahrg,COALESCE(dmpHypertSchulEmpf,'') dmpHypertSchulEmpf" & _
",COALESCE(dmpHypertSchulWahrg,'') dmpHypertSchulWahrg,COALESCE(dmpKKTabakEmpf,'') dmpKKTabakEmpf,COALESCE(dmpKKErnEmpf,'') dmpKKErnEmpf,COALESCE(dmpKKkTrainEmpf,'') dmpKKkTrainEmpf" & _
",COALESCE(dmpHbA1cZiel,'') dmpHbA1cZiel,COALESCE(dmpUewFuss,'') dmpUewFuss,COALESCE(dmpEinwDM,'') dmpEinwDM,COALESCE(dmphalbj,'') dmphalbj" & _
",COALESCE(dmpMA,'') dmpMA FROM `faelle` WHERE Pat_ID=" & pid & " ORDER BY `fanf`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roFa(0)
 Else ' rs.EOF Then
  ReDim roFa(1)
  Do While Not rs.EOF
   akt = UBound(roFa)
   roFa(akt).FID = rs!FID
   roFa(akt).Pat_id = rs!Pat_id
   roFa(akt).Quartal = doUmwfSQL(rs!Quartal, lies.obMySQL, False)
   roFa(akt).Nachname = doUmwfSQL(rs!Nachname, lies.obMySQL, False)
   roFa(akt).Vorname = doUmwfSQL(rs!Vorname, lies.obMySQL, False)
   roFa(akt).lfdnr = rs!lfdnr
   roFa(akt).TMFNr = doUmwfSQL(rs!TMFNr, lies.obMySQL, False)
   roFa(akt).VKNr = doUmwfSQL(rs!VKNr, lies.obMySQL, False)
   roFa(akt).f4131 = doUmwfSQL(rs!f4131, lies.obMySQL, False)
   roFa(akt).f4132 = doUmwfSQL(rs!f4132, lies.obMySQL, False)
   roFa(akt).VschBeg = rs!VschBeg
   roFa(akt).KKasse_2 = doUmwfSQL(rs!KKasse_2, lies.obMySQL, False)
   roFa(akt).FaktPers = rs!FaktPers
   roFa(akt).FaktTechn = rs!FaktTechn
   roFa(akt).FaktLabor = rs!FaktLabor
   roFa(akt).BhFB = rs!BhFB
   roFa(akt).BhFE1 = rs!BhFE1
   roFa(akt).BhFE2 = rs!BhFE2
   roFa(akt).f4202 = doUmwfSQL(rs!f4202, lies.obMySQL, False)
   roFa(akt).ausgst = rs!ausgst
   roFa(akt).KtrAbrB = doUmwfSQL(rs!KtrAbrB, lies.obMySQL, False)
   roFa(akt).AbrAr = doUmwfSQL(rs!AbrAr, lies.obMySQL, False)
   roFa(akt).lVorl = rs!lVorl
   roFa(akt).IK = doUmwfSQL(rs!IK, lies.obMySQL, False)
   roFa(akt).KVKs = doUmwfSQL(rs!KVKs, lies.obMySQL, False)
   roFa(akt).KVKserg = doUmwfSQL(rs!KVKserg, lies.obMySQL, False)
   roFa(akt).Status = doUmwfSQL(rs!Status, lies.obMySQL, False)
   roFa(akt).Kasse = doUmwfSQL(rs!Kasse, lies.obMySQL, False)
   roFa(akt).KID = rs!KID
   roFa(akt).GebOr = doUmwfSQL(rs!GebOr, lies.obMySQL, False)
   roFa(akt).AbrGb = doUmwfSQL(rs!AbrGb, lies.obMySQL, False)
   roFa(akt).PersKreis = doUmwfSQL(rs!PersKreis, lies.obMySQL, False)
   roFa(akt).SKtZusatz = doUmwfSQL(rs!SKtZusatz, lies.obMySQL, False)
   roFa(akt).letzteRegel = doUmwfSQL(rs!letzteRegel, lies.obMySQL, False)
   roFa(akt).ÜwText = doUmwfSQL(rs!ÜwText, lies.obMySQL, False)
   roFa(akt).f4210 = rs!f4210
   roFa(akt).AkfHAH = rs!AkfHAH
   roFa(akt).AkfAB0 = rs!AkfAB0
   roFa(akt).AkfAK = rs!AkfAK
   roFa(akt).statNuller = doUmwfSQL(rs!statNuller, lies.obMySQL, False)
   roFa(akt).ÜbwV = doUmwfSQL(rs!ÜbwV, lies.obMySQL, False)
   roFa(akt).ÜbWVLANR = doUmwfSQL(rs!ÜbWVLANR, lies.obMySQL, False)
   roFa(akt).ÜbWVBSNR = doUmwfSQL(rs!ÜbWVBSNR, lies.obMySQL, False)
   roFa(akt).ÜbWVKVNR = doUmwfSQL(rs!ÜbWVKVNR, lies.obMySQL, False)
   roFa(akt).AndÜw = doUmwfSQL(rs!AndÜw, lies.obMySQL, False)
   roFa(akt).Übwr = doUmwfSQL(rs!Übwr, lies.obMySQL, False)
   roFa(akt).ÜbwLANR = doUmwfSQL(rs!ÜbwLANR, lies.obMySQL, False)
   roFa(akt).ÜWZiel = doUmwfSQL(rs!ÜWZiel, lies.obMySQL, False)
   roFa(akt).ÜWNNr = doUmwfSQL(rs!ÜWNNr, lies.obMySQL, False)
   roFa(akt).ÜWNaN = doUmwfSQL(rs!ÜWNaN, lies.obMySQL, False)
   roFa(akt).ÜWTit = doUmwfSQL(rs!ÜWTit, lies.obMySQL, False)
   roFa(akt).ÜWVor = doUmwfSQL(rs!ÜWVor, lies.obMySQL, False)
   roFa(akt).ÜWVsw = doUmwfSQL(rs!ÜWVsw, lies.obMySQL, False)
   roFa(akt).üwvid = rs!üwvid
   roFa(akt).Auftrag = doUmwfSQL(rs!Auftrag, lies.obMySQL, False)
   roFa(akt).Verdacht = doUmwfSQL(rs!Verdacht, lies.obMySQL, False)
   roFa(akt).Befund = doUmwfSQL(rs!Befund, lies.obMySQL, False)
   roFa(akt).statKlasse = doUmwfSQL(rs!statKlasse, lies.obMySQL, False)
   roFa(akt).f4237 = doUmwfSQL(rs!f4237, lies.obMySQL, False)
   roFa(akt).statBehTage = rs!statBehTage
   roFa(akt).SchGr = rs!SchGr
   roFa(akt).Weiterbeh = doUmwfSQL(rs!Weiterbeh, lies.obMySQL, False)
   roFa(akt).f4266 = rs!f4266
   roFa(akt).PGeb = doUmwfSQL(rs!PGeb, lies.obMySQL, False)
   roFa(akt).PGebErg = doUmwfSQL(rs!PGebErg, lies.obMySQL, False)
   roFa(akt).Mahnfrist = doUmwfSQL(rs!Mahnfrist, lies.obMySQL, False)
   roFa(akt).Unfallort = doUmwfSQL(rs!Unfallort, lies.obMySQL, False)
   roFa(akt).BeschAls = doUmwfSQL(rs!BeschAls, lies.obMySQL, False)
   roFa(akt).BeschSeit = rs!BeschSeit
   roFa(akt).Unfallbetrieb = doUmwfSQL(rs!Unfallbetrieb, lies.obMySQL, False)
   roFa(akt).f4570 = doUmwfSQL(rs!f4570, lies.obMySQL, False)
   roFa(akt).GOÄKatNr = doUmwfSQL(rs!GOÄKatNr, lies.obMySQL, False)
   roFa(akt).GOÄKatName = doUmwfSQL(rs!GOÄKatName, lies.obMySQL, False)
   roFa(akt).abrArzt = doUmwfSQL(rs!abrArzt, lies.obMySQL, False)
   roFa(akt).privVers = doUmwfSQL(rs!privVers, lies.obMySQL, False)
   roFa(akt).AdNam = doUmwfSQL(rs!AdNam, lies.obMySQL, False)
   roFa(akt).AdStr = doUmwfSQL(rs!AdStr, lies.obMySQL, False)
   roFa(akt).AdPlz = doUmwfSQL(rs!AdPlz, lies.obMySQL, False)
   roFa(akt).AdOrt = doUmwfSQL(rs!AdOrt, lies.obMySQL, False)
   roFa(akt).ÜwBG = doUmwfSQL(rs!ÜwBG, lies.obMySQL, False)
   roFa(akt).BhFE = rs!BhFE
   roFa(akt).s8000 = doUmwfSQL(rs!s8000, lies.obMySQL, False)
   roFa(akt).s8100 = doUmwfSQL(rs!s8100, lies.obMySQL, False)
   roFa(akt).AktZeit = rs!AktZeit
   roFa(akt).Fanf = rs!Fanf
   roFa(akt).altQuart = doUmwfSQL(rs!altQuart, lies.obMySQL, False)
   roFa(akt).QAnf = rs!QAnf
   roFa(akt).QEnd = rs!QEnd
   roFa(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roFa(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roFa(akt).StByte = rs!StByte
   roFa(akt).absPos = rs!absPos
   roFa(akt).lanrid = rs!lanrid
   roFa(akt).f4108 = doUmwfSQL(rs!f4108, lies.obMySQL, False)
   roFa(akt).BGFallNr = doUmwfSQL(rs!BGFallNr, lies.obMySQL, False)
   roFa(akt).lGewicht = rs!lGewicht
   roFa(akt).vorET = rs!vorET
   roFa(akt).dmpVertret = doUmwfSQL(rs!dmpVertret, lies.obMySQL, False)
   roFa(akt).dmpArztw = doUmwfSQL(rs!dmpArztw, lies.obMySQL, False)
   roFa(akt).dmpHypos = doUmwfSQL(rs!dmpHypos, lies.obMySQL, False)
   roFa(akt).dmpKhsA = doUmwfSQL(rs!dmpKhsA, lies.obMySQL, False)
   roFa(akt).dmpDMSchulEmpf = doUmwfSQL(rs!dmpDMSchulEmpf, lies.obMySQL, False)
   roFa(akt).dmpDMSchulWahrg = doUmwfSQL(rs!dmpDMSchulWahrg, lies.obMySQL, False)
   roFa(akt).dmpHypertSchulEmpf = doUmwfSQL(rs!dmpHypertSchulEmpf, lies.obMySQL, False)
   roFa(akt).dmpHypertSchulWahrg = doUmwfSQL(rs!dmpHypertSchulWahrg, lies.obMySQL, False)
   roFa(akt).dmpKKTabakEmpf = doUmwfSQL(rs!dmpKKTabakEmpf, lies.obMySQL, False)
   roFa(akt).dmpKKErnEmpf = doUmwfSQL(rs!dmpKKErnEmpf, lies.obMySQL, False)
   roFa(akt).dmpKKkTrainEmpf = doUmwfSQL(rs!dmpKKkTrainEmpf, lies.obMySQL, False)
   roFa(akt).dmpHbA1cZiel = doUmwfSQL(rs!dmpHbA1cZiel, lies.obMySQL, False)
   roFa(akt).dmpUewFuss = doUmwfSQL(rs!dmpUewFuss, lies.obMySQL, False)
   roFa(akt).dmpEinwDM = doUmwfSQL(rs!dmpEinwDM, lies.obMySQL, False)
   roFa(akt).dmphalbj = doUmwfSQL(rs!dmphalbj, lies.obMySQL, False)
   roFa(akt).dmpMA = doUmwfSQL(rs!dmpMA, lies.obMySQL, False)
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roFa(UBound(roFa) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleLaden

Function faelleEinf()
 Dim rbeg&, robeg&, ri&, roi&, jj&, fazu%()
 On Error GoTo fehler
 If UBound(rFa) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rFa)
   If rFa(ri).Fanf >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roFa)
    If roFa(robeg).Fanf >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim fazu(robeg To UBound(roFa))
    For roi = robeg To UBound(roFa)
     For ri = rbeg To UBound(rFa) - 1
      If roFa(roi).Fanf >= rFa(ri).Fanf And roFa(roi).Fanf < rFa(ri).BhFE Then
       fazu(roi) = ri
       GoTo Fertig
      End If
     Next ri
     fazu(roi) = UBound(rFa)
Fertig:
    Next roi
    For roi = robeg To UBound(roFa)
     For jj = 1 To UBound(roAu)
      If roAu(jj).FID = roFa(roi).FID Then roAu(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roBr)
      If roBr(jj).FID = roFa(roi).FID Then roBr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roDi)
      If roDi(jj).FID = roFa(roi).FID Then roDi(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roDo)
      If roDo(jj).FID = roFa(roi).FID Then roDo(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roEi)
      If roEi(jj).FID = roFa(roi).FID Then roEi(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roFr)
      If roFr(jj).FID = roFa(roi).FID Then roFr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roKh)
      If roKh(jj).FID = roFa(roi).FID Then roKh(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLb)
      If roLb(jj).FID = roFa(roi).FID Then roLb(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLa)
      If roLa(jj).FID = roFa(roi).FID Then roLa(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLe)
      If roLe(jj).FID = roFa(roi).FID Then roLe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roMe)
      If roMe(jj).FID = roFa(roi).FID Then roMe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roRe)
      If roRe(jj).FID = roFa(roi).FID Then roRe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roRr)
      If roRr(jj).FID = roFa(roi).FID Then roRr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roUs)
      If roUs(jj).FID = roFa(roi).FID Then roUs(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roFu)
      If roFu(jj).FID = roFa(roi).FID Then roFu(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roUl)
      If roUl(jj).FID = roFa(roi).FID Then roUl(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roVk)
      If roVk(jj).FID = roFa(roi).FID Then roVk(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roSw)
      If roSw(jj).FID = roFa(roi).FID Then roSw(jj).FID = rFa(fazu(roi)).FID
     Next jj
    Next roi
    ReDim Preserve roFa(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFa(robeg + UBound(rFa) - rbeg)
   For ri = rbeg To UBound(rFa)
    Call roFaZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rFa = roFa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleEinf

Public Function faelleSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim j%
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFa) + 0 & " Sätze in `faelle`"
' sql0 = " INSERT " & sqlignore &  "INTO `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn," & _
     "FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK," & _
     "KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel," & _
     "ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR," & _
     "AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid," & _
     "Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb," & _
     "PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt," & _
     "privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit," & _
     "Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108," & _
     "BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf," & _
     "dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `faelle` (Pat_ID,Quartal,Nachname," & _
     "Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn," & _
     "FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK," & _
     "KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel," & _
     "ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR," & _
     "AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid," & _
     "Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb," & _
     "PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt," & _
     "privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit," & _
     "Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108," & _
     "BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf," & _
     "dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)        VALUES"))
 For i = 1 To UBound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
abermals:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFa(i).Pat_id, ",'", rFa(i).Quartal, "','", rFa(i).Nachname, "','", rFa(i).Vorname, "',", rFa(i).lfdnr, ",'", rFa(i).TMFNr, "','", rFa(i).VKNr, "','", rFa(i).f4131, "','", _
   rFa(i).f4132, "',", DatFor_k(rFa(i).VschBeg), ",'", rFa(i).KKasse_2, "',", REPLACE$(rFa(i).FaktPers, ",", "."), ",", REPLACE$(rFa(i).FaktTechn, ",", "."), ",", REPLACE$(rFa(i).FaktLabor, ",", "."), ",", DatFor_k( _
   rFa(i).BhFB), ",", DatFor_k(rFa(i).BhFE1), ",", DatFor_k(rFa(i).BhFE2), ",'", rFa(i).f4202, "',", DatFor_k(rFa(i).ausgst), ",'", rFa(i).KtrAbrB, "','", rFa(i).AbrAr, "',", DatFor_k( _
   rFa(i).lVorl), ",'", rFa(i).IK, "','", rFa(i).KVKs, "','", rFa(i).KVKserg, "','", rFa(i).Status, "','", rFa(i).Kasse, "',", rFa(i).KID, ",'", rFa(i).GebOr, "','", rFa(i).AbrGb, "','", _
   rFa(i).PersKreis, "','", rFa(i).SKtZusatz, "','", rFa(i).letzteRegel, "','", rFa(i).ÜwText, "',", rFa(i).f4210, ",", rFa(i).AkfHAH, ",", rFa(i).AkfAB0, ",", rFa(i).AkfAK, ",'", _
   rFa(i).statNuller, "','", rFa(i).ÜbwV, "','", rFa(i).ÜbWVLANR, "','", rFa(i).ÜbWVBSNR, "','", rFa(i).ÜbWVKVNR, "','", rFa(i).AndÜw, "','", rFa(i).Übwr, "','", rFa(i).ÜbwLANR, "','", _
   rFa(i).ÜWZiel, "','", rFa(i).ÜWNNr, "','", rFa(i).ÜWNaN, "','", rFa(i).ÜWTit, "','", rFa(i).ÜWVor, "','", rFa(i).ÜWVsw, "',", rFa(i).üwvid, ",'", rFa(i).Auftrag, "','", rFa(i).Verdacht, "','", _
   rFa(i).Befund, "','", rFa(i).statKlasse, "','", rFa(i).f4237, "',", rFa(i).statBehTage, ",", rFa(i).SchGr, ",'", rFa(i).Weiterbeh, "',", REPLACE$(rFa(i).f4266, ",", "."), ",'", _
   rFa(i).PGeb, "','", rFa(i).PGebErg, "','", rFa(i).Mahnfrist, "','", rFa(i).Unfallort, "','", rFa(i).BeschAls, "',", DatFor_k(rFa(i).BeschSeit), ",'", rFa(i).Unfallbetrieb, "','", rFa(i).f4570, "','", _
   rFa(i).GOÄKatNr, "','", rFa(i).GOÄKatName, "','", rFa(i).abrArzt, "','", rFa(i).privVers, "','", rFa(i).AdNam, "','", rFa(i).AdStr, "','", rFa(i).AdPlz, "','", rFa(i).AdOrt, "','", _
   rFa(i).ÜwBG, "',", DatFor_k(rFa(i).BhFE), ",'", rFa(i).s8000, "','", rFa(i).s8100, "',", DatFor_k(rFa(i).AktZeit), ",", DatFor_k(rFa(i).Fanf), ",'", rFa(i).altQuart, "',", DatFor_k(rFa(i).QAnf), ",", DatFor_k( _
   rFa(i).QEnd), ",'", rFa(i).QS, "','", rFa(i).QT, "',", rFa(i).StByte, ",", rFa(i).absPos, ",", rFa(i).lanrid, ",'", rFa(i).f4108, "','", rFa(i).BGFallNr, "',", _
   rFa(i).lGewicht, ",", DatFor_k(rFa(i).vorET), ",'", rFa(i).dmpVertret, "','", rFa(i).dmpArztw, "','", rFa(i).dmpHypos, "','", rFa(i).dmpKhsA, "','", rFa(i).dmpDMSchulEmpf, "','", _
   rFa(i).dmpDMSchulWahrg, "','", rFa(i).dmpHypertSchulEmpf, "','", rFa(i).dmpHypertSchulWahrg, "','", rFa(i).dmpKKTabakEmpf, "','", rFa(i).dmpKKErnEmpf, "','", rFa(i).dmpKKkTrainEmpf, "','", _
   rFa(i).dmpHbA1cZiel, "','", rFa(i).dmpUewFuss, "','", rFa(i).dmpEinwDM, "','", rFa(i).dmphalbj, "','", rFa(i).dmpMA, "')")
  If SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFa) Then
'   IF Not obForK THEN ForeignNo0
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `faelle` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Dim altmode$
   altmode = DBCn.Execute("SELECT @@global.sql_mode").Fields(0)
   Call DBCn.Execute("SET GLOBAL sql_mode='STRICT_TRANS_TABLES'")
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If rAf = 0 Then
    Err.Raise 998, , "Fehler in faelleSpeichern b.Pat. " & rFa(i).Pat_id & ", Err.Number " & Err.Number & ", err.description: " & Err.Description
   End If ' rAF = 0 THEN
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
'   If Lese.obMySQL And obMitAlterTab Then
''    Set rs = myEFrag("SHOW WARNINGS")
'    Set rs = DBCn.Execute("show warnings", rAf)
'    If Not rs.BOF() Then
'     If rs!code = 1265 Then
'      Err.Raise -2147217833
'     End If ' rs!code = 1265 THEN
'    End If ' lese.obmysql AND obmitAlterTab THEN
'   End If ' lese.obmysql AND obmitAlterTab
   If SammelInsert = 0 Then
   'Hier gibts mit Sammelins noch ein Problem ...
    Set rs = Nothing
    For j = 1 To 2
     If j = 1 Then
      Set rs = myEFrag("SELECT LAST_INSERT_ID() FID") ' session-spezifisch
     Else
      Set rs = myEFrag("SELECT MAX(fid) FID FROM `faelle` WHERE pat_id = " & rFa(i).Pat_id & " AND quartal = '" & rFa(i).Quartal & "' AND bhfb = " & DatFor_k(rFa(i).BhFB) & " AND bhfe1 = " & DatFor_k(rFa(i).BhFE1) & " AND ausgst = " & DatFor_k(rFa(i).ausgst))
     End If
     If Not rs.BOF Then If rs.Fields(0) <> 0 Then Exit For
    Next j
    If rs.BOF Then
     Err.Raise 999, , "Fehler bei der Fallaktualisierung b.Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID
    ElseIf rs!FID = 0 Then
     MsgBox "Fehler in faellespeichern:" & vbCrLf & rs.source
     GoTo sql
    ElseIf rs!FID <> rFa(i).FID Then
      Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_id & ", FID " & rFa(i).FID & " -> " & rs!FID & " in zu speichernden Tabellen mit fallid", True
      Dim jjj&
      For jjj = 1 To UBound(rAu)
       If rAu(jjj).FID = rFa(i).FID Then
        rAu(jjj).FID = rs!FID
       End If ' rAu(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rBr)
       If rBr(jjj).FID = rFa(i).FID Then
        rBr(jjj).FID = rs!FID
       End If ' rBr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rDi)
       If rDi(jjj).FID = rFa(i).FID Then
        rDi(jjj).FID = rs!FID
       End If ' rDi(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rDo)
       If rDo(jjj).FID = rFa(i).FID Then
        rDo(jjj).FID = rs!FID
       End If ' rDo(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rEi)
       If rEi(jjj).FID = rFa(i).FID Then
        rEi(jjj).FID = rs!FID
       End If ' rEi(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rFr)
       If rFr(jjj).FID = rFa(i).FID Then
        rFr(jjj).FID = rs!FID
       End If ' rFr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rKh)
       If rKh(jjj).FID = rFa(i).FID Then
        rKh(jjj).FID = rs!FID
       End If ' rKh(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLb)
       If rLb(jjj).FID = rFa(i).FID Then
        rLb(jjj).FID = rs!FID
       End If ' rLb(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLa)
       If rLa(jjj).FID = rFa(i).FID Then
        rLa(jjj).FID = rs!FID
       End If ' rLa(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLe)
       If rLe(jjj).FID = rFa(i).FID Then
        rLe(jjj).FID = rs!FID
       End If ' rLe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rMe)
       If rMe(jjj).FID = rFa(i).FID Then
        rMe(jjj).FID = rs!FID
       End If ' rMe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rRe)
       If rRe(jjj).FID = rFa(i).FID Then
        rRe(jjj).FID = rs!FID
       End If ' rRe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rRr)
       If rRr(jjj).FID = rFa(i).FID Then
        rRr(jjj).FID = rs!FID
       End If ' rRr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rUs)
       If rUs(jjj).FID = rFa(i).FID Then
        rUs(jjj).FID = rs!FID
       End If ' rUs(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rFu)
       If rFu(jjj).FID = rFa(i).FID Then
        rFu(jjj).FID = rs!FID
       End If ' rFu(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rUl)
       If rUl(jjj).FID = rFa(i).FID Then
        rUl(jjj).FID = rs!FID
       End If ' rUl(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rVk)
       If rVk(jjj).FID = rFa(i).FID Then
        rVk(jjj).FID = rs!FID
       End If ' rVk(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rSw)
       If rSw(jjj).FID = rFa(i).FID Then
        rSw(jjj).FID = rs!FID
       End If ' rSw(jjj).FID = rFa(i).FID THEN
      Next jjj
    End If ' rs.BOF Then
    Set rs = Nothing
    csql.m_Len = 0
   End If ' IF Sammelinsert = 0 Then
  End If ' SammelInsert = 0 OR i = ubound(rFa)
  DoEvents
 Next i
 syscmd 5
 Call DBCn.Execute("SET GLOBAL sql_mode='" & altmode & "'")
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""faelleSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(80)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFa), i)
  If Len(rFa(k).Quartal) > maxi(0) Then maxi(0) = Len(rFa(k).Quartal)
  If Len(rFa(k).Nachname) > maxi(1) Then maxi(1) = Len(rFa(k).Nachname)
  If Len(rFa(k).Vorname) > maxi(2) Then maxi(2) = Len(rFa(k).Vorname)
  If Len(rFa(k).TMFNr) > maxi(3) Then maxi(3) = Len(rFa(k).TMFNr)
  If Len(rFa(k).VKNr) > maxi(4) Then maxi(4) = Len(rFa(k).VKNr)
  If Len(rFa(k).f4131) > maxi(5) Then maxi(5) = Len(rFa(k).f4131)
  If Len(rFa(k).f4132) > maxi(6) Then maxi(6) = Len(rFa(k).f4132)
  If Len(rFa(k).KKasse_2) > maxi(7) Then maxi(7) = Len(rFa(k).KKasse_2)
  If Len(rFa(k).f4202) > maxi(8) Then maxi(8) = Len(rFa(k).f4202)
  If Len(rFa(k).KtrAbrB) > maxi(9) Then maxi(9) = Len(rFa(k).KtrAbrB)
  If Len(rFa(k).AbrAr) > maxi(10) Then maxi(10) = Len(rFa(k).AbrAr)
  If Len(rFa(k).IK) > maxi(11) Then maxi(11) = Len(rFa(k).IK)
  If Len(rFa(k).KVKs) > maxi(12) Then maxi(12) = Len(rFa(k).KVKs)
  If Len(rFa(k).KVKserg) > maxi(13) Then maxi(13) = Len(rFa(k).KVKserg)
  If Len(rFa(k).Status) > maxi(14) Then maxi(14) = Len(rFa(k).Status)
  If Len(rFa(k).Kasse) > maxi(15) Then maxi(15) = Len(rFa(k).Kasse)
  If Len(rFa(k).GebOr) > maxi(16) Then maxi(16) = Len(rFa(k).GebOr)
  If Len(rFa(k).AbrGb) > maxi(17) Then maxi(17) = Len(rFa(k).AbrGb)
  If Len(rFa(k).PersKreis) > maxi(18) Then maxi(18) = Len(rFa(k).PersKreis)
  If Len(rFa(k).SKtZusatz) > maxi(19) Then maxi(19) = Len(rFa(k).SKtZusatz)
  If Len(rFa(k).letzteRegel) > maxi(20) Then maxi(20) = Len(rFa(k).letzteRegel)
  If Len(rFa(k).ÜwText) > maxi(21) Then maxi(21) = Len(rFa(k).ÜwText)
  If Len(rFa(k).statNuller) > maxi(22) Then maxi(22) = Len(rFa(k).statNuller)
  If Len(rFa(k).ÜbwV) > maxi(23) Then maxi(23) = Len(rFa(k).ÜbwV)
  If Len(rFa(k).ÜbWVLANR) > maxi(24) Then maxi(24) = Len(rFa(k).ÜbWVLANR)
  If Len(rFa(k).ÜbWVBSNR) > maxi(25) Then maxi(25) = Len(rFa(k).ÜbWVBSNR)
  If Len(rFa(k).ÜbWVKVNR) > maxi(26) Then maxi(26) = Len(rFa(k).ÜbWVKVNR)
  If Len(rFa(k).AndÜw) > maxi(27) Then maxi(27) = Len(rFa(k).AndÜw)
  If Len(rFa(k).Übwr) > maxi(28) Then maxi(28) = Len(rFa(k).Übwr)
  If Len(rFa(k).ÜbwLANR) > maxi(29) Then maxi(29) = Len(rFa(k).ÜbwLANR)
  If Len(rFa(k).ÜWZiel) > maxi(30) Then maxi(30) = Len(rFa(k).ÜWZiel)
  If Len(rFa(k).ÜWNNr) > maxi(31) Then maxi(31) = Len(rFa(k).ÜWNNr)
  If Len(rFa(k).ÜWNaN) > maxi(32) Then maxi(32) = Len(rFa(k).ÜWNaN)
  If Len(rFa(k).ÜWTit) > maxi(33) Then maxi(33) = Len(rFa(k).ÜWTit)
  If Len(rFa(k).ÜWVor) > maxi(34) Then maxi(34) = Len(rFa(k).ÜWVor)
  If Len(rFa(k).ÜWVsw) > maxi(35) Then maxi(35) = Len(rFa(k).ÜWVsw)
  If Len(rFa(k).Auftrag) > maxi(36) Then maxi(36) = Len(rFa(k).Auftrag)
  If Len(rFa(k).Verdacht) > maxi(37) Then maxi(37) = Len(rFa(k).Verdacht)
  If Len(rFa(k).Befund) > maxi(38) Then maxi(38) = Len(rFa(k).Befund)
  If Len(rFa(k).statKlasse) > maxi(39) Then maxi(39) = Len(rFa(k).statKlasse)
  If Len(rFa(k).f4237) > maxi(40) Then maxi(40) = Len(rFa(k).f4237)
  If Len(rFa(k).Weiterbeh) > maxi(41) Then maxi(41) = Len(rFa(k).Weiterbeh)
  If Len(rFa(k).PGeb) > maxi(42) Then maxi(42) = Len(rFa(k).PGeb)
  If Len(rFa(k).PGebErg) > maxi(43) Then maxi(43) = Len(rFa(k).PGebErg)
  If Len(rFa(k).Mahnfrist) > maxi(44) Then maxi(44) = Len(rFa(k).Mahnfrist)
  If Len(rFa(k).Unfallort) > maxi(45) Then maxi(45) = Len(rFa(k).Unfallort)
  If Len(rFa(k).BeschAls) > maxi(46) Then maxi(46) = Len(rFa(k).BeschAls)
  If Len(rFa(k).Unfallbetrieb) > maxi(47) Then maxi(47) = Len(rFa(k).Unfallbetrieb)
  If Len(rFa(k).f4570) > maxi(48) Then maxi(48) = Len(rFa(k).f4570)
  If Len(rFa(k).GOÄKatNr) > maxi(49) Then maxi(49) = Len(rFa(k).GOÄKatNr)
  If Len(rFa(k).GOÄKatName) > maxi(50) Then maxi(50) = Len(rFa(k).GOÄKatName)
  If Len(rFa(k).abrArzt) > maxi(51) Then maxi(51) = Len(rFa(k).abrArzt)
  If Len(rFa(k).privVers) > maxi(52) Then maxi(52) = Len(rFa(k).privVers)
  If Len(rFa(k).AdNam) > maxi(53) Then maxi(53) = Len(rFa(k).AdNam)
  If Len(rFa(k).AdStr) > maxi(54) Then maxi(54) = Len(rFa(k).AdStr)
  If Len(rFa(k).AdPlz) > maxi(55) Then maxi(55) = Len(rFa(k).AdPlz)
  If Len(rFa(k).AdOrt) > maxi(56) Then maxi(56) = Len(rFa(k).AdOrt)
  If Len(rFa(k).ÜwBG) > maxi(57) Then maxi(57) = Len(rFa(k).ÜwBG)
  If Len(rFa(k).s8000) > maxi(58) Then maxi(58) = Len(rFa(k).s8000)
  If Len(rFa(k).s8100) > maxi(59) Then maxi(59) = Len(rFa(k).s8100)
  If Len(rFa(k).altQuart) > maxi(60) Then maxi(60) = Len(rFa(k).altQuart)
  If Len(rFa(k).QS) > maxi(61) Then maxi(61) = Len(rFa(k).QS)
  If Len(rFa(k).QT) > maxi(62) Then maxi(62) = Len(rFa(k).QT)
  If Len(rFa(k).f4108) > maxi(63) Then maxi(63) = Len(rFa(k).f4108)
  If Len(rFa(k).BGFallNr) > maxi(64) Then maxi(64) = Len(rFa(k).BGFallNr)
  If Len(rFa(k).dmpVertret) > maxi(65) Then maxi(65) = Len(rFa(k).dmpVertret)
  If Len(rFa(k).dmpArztw) > maxi(66) Then maxi(66) = Len(rFa(k).dmpArztw)
  If Len(rFa(k).dmpHypos) > maxi(67) Then maxi(67) = Len(rFa(k).dmpHypos)
  If Len(rFa(k).dmpKhsA) > maxi(68) Then maxi(68) = Len(rFa(k).dmpKhsA)
  If Len(rFa(k).dmpDMSchulEmpf) > maxi(69) Then maxi(69) = Len(rFa(k).dmpDMSchulEmpf)
  If Len(rFa(k).dmpDMSchulWahrg) > maxi(70) Then maxi(70) = Len(rFa(k).dmpDMSchulWahrg)
  If Len(rFa(k).dmpHypertSchulEmpf) > maxi(71) Then maxi(71) = Len(rFa(k).dmpHypertSchulEmpf)
  If Len(rFa(k).dmpHypertSchulWahrg) > maxi(72) Then maxi(72) = Len(rFa(k).dmpHypertSchulWahrg)
  If Len(rFa(k).dmpKKTabakEmpf) > maxi(73) Then maxi(73) = Len(rFa(k).dmpKKTabakEmpf)
  If Len(rFa(k).dmpKKErnEmpf) > maxi(74) Then maxi(74) = Len(rFa(k).dmpKKErnEmpf)
  If Len(rFa(k).dmpKKkTrainEmpf) > maxi(75) Then maxi(75) = Len(rFa(k).dmpKKkTrainEmpf)
  If Len(rFa(k).dmpHbA1cZiel) > maxi(76) Then maxi(76) = Len(rFa(k).dmpHbA1cZiel)
  If Len(rFa(k).dmpUewFuss) > maxi(77) Then maxi(77) = Len(rFa(k).dmpUewFuss)
  If Len(rFa(k).dmpEinwDM) > maxi(78) Then maxi(78) = Len(rFa(k).dmpEinwDM)
  If Len(rFa(k).dmphalbj) > maxi(79) Then maxi(79) = Len(rFa(k).dmphalbj)
  If Len(rFa(k).dmpMA) > maxi(80) Then maxi(80) = Len(rFa(k).dmpMA)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & Left$(rFa(k).Quartal, maxL) & "'", True: rFa(k).Quartal = Left$(rFa(k).Quartal, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & Left$(rFa(k).Nachname, maxL) & "'", True: rFa(k).Nachname = Left$(rFa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & Left$(rFa(k).Vorname, maxL) & "'", True: rFa(k).Vorname = Left$(rFa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & Left$(rFa(k).TMFNr, maxL) & "'", True: rFa(k).TMFNr = Left$(rFa(k).TMFNr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & Left$(rFa(k).VKNr, maxL) & "'", True: rFa(k).VKNr = Left$(rFa(k).VKNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4131: '" & rFa(k).f4131 & "' -> '" & Left$(rFa(k).f4131, maxL) & "'", True: rFa(k).f4131 = Left$(rFa(k).f4131, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4132: '" & rFa(k).f4132 & "' -> '" & Left$(rFa(k).f4132, maxL) & "'", True: rFa(k).f4132 = Left$(rFa(k).f4132, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.KKasse_2: '" & rFa(k).KKasse_2 & "' -> '" & Left$(rFa(k).KKasse_2, maxL) & "'", True: rFa(k).KKasse_2 = Left$(rFa(k).KKasse_2, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4202: '" & rFa(k).f4202 & "' -> '" & Left$(rFa(k).f4202, maxL) & "'", True: rFa(k).f4202 = Left$(rFa(k).f4202, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & Left$(rFa(k).KtrAbrB, maxL) & "'", True: rFa(k).KtrAbrB = Left$(rFa(k).KtrAbrB, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & Left$(rFa(k).AbrAr, maxL) & "'", True: rFa(k).AbrAr = Left$(rFa(k).AbrAr, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & Left$(rFa(k).IK, maxL) & "'", True: rFa(k).IK = Left$(rFa(k).IK, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & Left$(rFa(k).KVKs, maxL) & "'", True: rFa(k).KVKs = Left$(rFa(k).KVKs, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & Left$(rFa(k).KVKserg, maxL) & "'", True: rFa(k).KVKserg = Left$(rFa(k).KVKserg, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.Status: '" & rFa(k).Status & "' -> '" & Left$(rFa(k).Status, maxL) & "'", True: rFa(k).Status = Left$(rFa(k).Status, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.Kasse: '" & rFa(k).Kasse & "' -> '" & Left$(rFa(k).Kasse, maxL) & "'", True: rFa(k).Kasse = Left$(rFa(k).Kasse, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & Left$(rFa(k).GebOr, maxL) & "'", True: rFa(k).GebOr = Left$(rFa(k).GebOr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & Left$(rFa(k).AbrGb, maxL) & "'", True: rFa(k).AbrGb = Left$(rFa(k).AbrGb, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & Left$(rFa(k).PersKreis, maxL) & "'", True: rFa(k).PersKreis = Left$(rFa(k).PersKreis, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & Left$(rFa(k).SKtZusatz, maxL) & "'", True: rFa(k).SKtZusatz = Left$(rFa(k).SKtZusatz, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.letzteRegel: '" & rFa(k).letzteRegel & "' -> '" & Left$(rFa(k).letzteRegel, maxL) & "'", True: rFa(k).letzteRegel = Left$(rFa(k).letzteRegel, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & Left$(rFa(k).ÜwText, maxL) & "'", True: rFa(k).ÜwText = Left$(rFa(k).ÜwText, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & Left$(rFa(k).statNuller, maxL) & "'", True: rFa(k).statNuller = Left$(rFa(k).statNuller, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & Left$(rFa(k).ÜbwV, maxL) & "'", True: rFa(k).ÜbwV = Left$(rFa(k).ÜbwV, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVLANR: '" & rFa(k).ÜbWVLANR & "' -> '" & Left$(rFa(k).ÜbWVLANR, maxL) & "'", True: rFa(k).ÜbWVLANR = Left$(rFa(k).ÜbWVLANR, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVBSNR: '" & rFa(k).ÜbWVBSNR & "' -> '" & Left$(rFa(k).ÜbWVBSNR, maxL) & "'", True: rFa(k).ÜbWVBSNR = Left$(rFa(k).ÜbWVBSNR, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVKVNR: '" & rFa(k).ÜbWVKVNR & "' -> '" & Left$(rFa(k).ÜbWVKVNR, maxL) & "'", True: rFa(k).ÜbWVKVNR = Left$(rFa(k).ÜbWVKVNR, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & Left$(rFa(k).AndÜw, maxL) & "'", True: rFa(k).AndÜw = Left$(rFa(k).AndÜw, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übwr: '" & rFa(k).Übwr & "' -> '" & Left$(rFa(k).Übwr, maxL) & "'", True: rFa(k).Übwr = Left$(rFa(k).Übwr, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & Left$(rFa(k).ÜbwLANR, maxL) & "'", True: rFa(k).ÜbwLANR = Left$(rFa(k).ÜbwLANR, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & Left$(rFa(k).ÜWZiel, maxL) & "'", True: rFa(k).ÜWZiel = Left$(rFa(k).ÜWZiel, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & Left$(rFa(k).ÜWNNr, maxL) & "'", True: rFa(k).ÜWNNr = Left$(rFa(k).ÜWNNr, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & Left$(rFa(k).ÜWNaN, maxL) & "'", True: rFa(k).ÜWNaN = Left$(rFa(k).ÜWNaN, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & Left$(rFa(k).ÜWTit, maxL) & "'", True: rFa(k).ÜWTit = Left$(rFa(k).ÜWTit, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & Left$(rFa(k).ÜWVor, maxL) & "'", True: rFa(k).ÜWVor = Left$(rFa(k).ÜWVor, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & Left$(rFa(k).ÜWVsw, maxL) & "'", True: rFa(k).ÜWVsw = Left$(rFa(k).ÜWVsw, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.Auftrag: '" & rFa(k).Auftrag & "' -> '" & Left$(rFa(k).Auftrag, maxL) & "'", True: rFa(k).Auftrag = Left$(rFa(k).Auftrag, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.Verdacht: '" & rFa(k).Verdacht & "' -> '" & Left$(rFa(k).Verdacht, maxL) & "'", True: rFa(k).Verdacht = Left$(rFa(k).Verdacht, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.Befund: '" & rFa(k).Befund & "' -> '" & Left$(rFa(k).Befund, maxL) & "'", True: rFa(k).Befund = Left$(rFa(k).Befund, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & Left$(rFa(k).statKlasse, maxL) & "'", True: rFa(k).statKlasse = Left$(rFa(k).statKlasse, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4237: '" & rFa(k).f4237 & "' -> '" & Left$(rFa(k).f4237, maxL) & "'", True: rFa(k).f4237 = Left$(rFa(k).f4237, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & Left$(rFa(k).Weiterbeh, maxL) & "'", True: rFa(k).Weiterbeh = Left$(rFa(k).Weiterbeh, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & Left$(rFa(k).PGeb, maxL) & "'", True: rFa(k).PGeb = Left$(rFa(k).PGeb, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & Left$(rFa(k).PGebErg, maxL) & "'", True: rFa(k).PGebErg = Left$(rFa(k).PGebErg, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & Left$(rFa(k).Mahnfrist, maxL) & "'", True: rFa(k).Mahnfrist = Left$(rFa(k).Mahnfrist, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallort: '" & rFa(k).Unfallort & "' -> '" & Left$(rFa(k).Unfallort, maxL) & "'", True: rFa(k).Unfallort = Left$(rFa(k).Unfallort, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.BeschAls: '" & rFa(k).BeschAls & "' -> '" & Left$(rFa(k).BeschAls, maxL) & "'", True: rFa(k).BeschAls = Left$(rFa(k).BeschAls, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallbetrieb: '" & rFa(k).Unfallbetrieb & "' -> '" & Left$(rFa(k).Unfallbetrieb, maxL) & "'", True: rFa(k).Unfallbetrieb = Left$(rFa(k).Unfallbetrieb, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4570: '" & rFa(k).f4570 & "' -> '" & Left$(rFa(k).f4570, maxL) & "'", True: rFa(k).f4570 = Left$(rFa(k).f4570, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & Left$(rFa(k).GOÄKatNr, maxL) & "'", True: rFa(k).GOÄKatNr = Left$(rFa(k).GOÄKatNr, maxL)
       Case 50: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & Left$(rFa(k).GOÄKatName, maxL) & "'", True: rFa(k).GOÄKatName = Left$(rFa(k).GOÄKatName, maxL)
       Case 51: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & Left$(rFa(k).abrArzt, maxL) & "'", True: rFa(k).abrArzt = Left$(rFa(k).abrArzt, maxL)
       Case 52: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & Left$(rFa(k).privVers, maxL) & "'", True: rFa(k).privVers = Left$(rFa(k).privVers, maxL)
       Case 53: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & Left$(rFa(k).AdNam, maxL) & "'", True: rFa(k).AdNam = Left$(rFa(k).AdNam, maxL)
       Case 54: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & Left$(rFa(k).AdStr, maxL) & "'", True: rFa(k).AdStr = Left$(rFa(k).AdStr, maxL)
       Case 55: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & Left$(rFa(k).AdPlz, maxL) & "'", True: rFa(k).AdPlz = Left$(rFa(k).AdPlz, maxL)
       Case 56: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & Left$(rFa(k).AdOrt, maxL) & "'", True: rFa(k).AdOrt = Left$(rFa(k).AdOrt, maxL)
       Case 57: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwBG: '" & rFa(k).ÜwBG & "' -> '" & Left$(rFa(k).ÜwBG, maxL) & "'", True: rFa(k).ÜwBG = Left$(rFa(k).ÜwBG, maxL)
       Case 58: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & Left$(rFa(k).s8000, maxL) & "'", True: rFa(k).s8000 = Left$(rFa(k).s8000, maxL)
       Case 59: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & Left$(rFa(k).s8100, maxL) & "'", True: rFa(k).s8100 = Left$(rFa(k).s8100, maxL)
       Case 60: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & Left$(rFa(k).altQuart, maxL) & "'", True: rFa(k).altQuart = Left$(rFa(k).altQuart, maxL)
       Case 61: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & Left$(rFa(k).QS, maxL) & "'", True: rFa(k).QS = Left$(rFa(k).QS, maxL)
       Case 62: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & Left$(rFa(k).QT, maxL) & "'", True: rFa(k).QT = Left$(rFa(k).QT, maxL)
       Case 63: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4108: '" & rFa(k).f4108 & "' -> '" & Left$(rFa(k).f4108, maxL) & "'", True: rFa(k).f4108 = Left$(rFa(k).f4108, maxL)
       Case 64: Lese.Ausgeb "   Verkürze Inhalt von rFa.BGFallNr: '" & rFa(k).BGFallNr & "' -> '" & Left$(rFa(k).BGFallNr, maxL) & "'", True: rFa(k).BGFallNr = Left$(rFa(k).BGFallNr, maxL)
       Case 65: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpVertret: '" & rFa(k).dmpVertret & "' -> '" & Left$(rFa(k).dmpVertret, maxL) & "'", True: rFa(k).dmpVertret = Left$(rFa(k).dmpVertret, maxL)
       Case 66: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpArztw: '" & rFa(k).dmpArztw & "' -> '" & Left$(rFa(k).dmpArztw, maxL) & "'", True: rFa(k).dmpArztw = Left$(rFa(k).dmpArztw, maxL)
       Case 67: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypos: '" & rFa(k).dmpHypos & "' -> '" & Left$(rFa(k).dmpHypos, maxL) & "'", True: rFa(k).dmpHypos = Left$(rFa(k).dmpHypos, maxL)
       Case 68: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKhsA: '" & rFa(k).dmpKhsA & "' -> '" & Left$(rFa(k).dmpKhsA, maxL) & "'", True: rFa(k).dmpKhsA = Left$(rFa(k).dmpKhsA, maxL)
       Case 69: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulEmpf: '" & rFa(k).dmpDMSchulEmpf & "' -> '" & Left$(rFa(k).dmpDMSchulEmpf, maxL) & "'", True: rFa(k).dmpDMSchulEmpf = Left$(rFa(k).dmpDMSchulEmpf, maxL)
       Case 70: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulWahrg: '" & rFa(k).dmpDMSchulWahrg & "' -> '" & Left$(rFa(k).dmpDMSchulWahrg, maxL) & "'", True: rFa(k).dmpDMSchulWahrg = Left$(rFa(k).dmpDMSchulWahrg, maxL)
       Case 71: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulEmpf: '" & rFa(k).dmpHypertSchulEmpf & "' -> '" & Left$(rFa(k).dmpHypertSchulEmpf, maxL) & "'", True: rFa(k).dmpHypertSchulEmpf = Left$(rFa(k).dmpHypertSchulEmpf, maxL)
       Case 72: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulWahrg: '" & rFa(k).dmpHypertSchulWahrg & "' -> '" & Left$(rFa(k).dmpHypertSchulWahrg, maxL) & "'", True: rFa(k).dmpHypertSchulWahrg = Left$(rFa(k).dmpHypertSchulWahrg, maxL)
       Case 73: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKTabakEmpf: '" & rFa(k).dmpKKTabakEmpf & "' -> '" & Left$(rFa(k).dmpKKTabakEmpf, maxL) & "'", True: rFa(k).dmpKKTabakEmpf = Left$(rFa(k).dmpKKTabakEmpf, maxL)
       Case 74: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKErnEmpf: '" & rFa(k).dmpKKErnEmpf & "' -> '" & Left$(rFa(k).dmpKKErnEmpf, maxL) & "'", True: rFa(k).dmpKKErnEmpf = Left$(rFa(k).dmpKKErnEmpf, maxL)
       Case 75: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKkTrainEmpf: '" & rFa(k).dmpKKkTrainEmpf & "' -> '" & Left$(rFa(k).dmpKKkTrainEmpf, maxL) & "'", True: rFa(k).dmpKKkTrainEmpf = Left$(rFa(k).dmpKKkTrainEmpf, maxL)
       Case 76: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHbA1cZiel: '" & rFa(k).dmpHbA1cZiel & "' -> '" & Left$(rFa(k).dmpHbA1cZiel, maxL) & "'", True: rFa(k).dmpHbA1cZiel = Left$(rFa(k).dmpHbA1cZiel, maxL)
       Case 77: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpUewFuss: '" & rFa(k).dmpUewFuss & "' -> '" & Left$(rFa(k).dmpUewFuss, maxL) & "'", True: rFa(k).dmpUewFuss = Left$(rFa(k).dmpUewFuss, maxL)
       Case 78: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpEinwDM: '" & rFa(k).dmpEinwDM & "' -> '" & Left$(rFa(k).dmpEinwDM, maxL) & "'", True: rFa(k).dmpEinwDM = Left$(rFa(k).dmpEinwDM, maxL)
       Case 79: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmphalbj: '" & rFa(k).dmphalbj & "' -> '" & Left$(rFa(k).dmphalbj, maxL) & "'", True: rFa(k).dmphalbj = Left$(rFa(k).dmphalbj, maxL)
       Case 80: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpMA: '" & rFa(k).dmpMA & "' -> '" & Left$(rFa(k).dmpMA, maxL) & "'", True: rFa(k).dmpMA = Left$(rFa(k).dmpMA, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
If ErrNumber = -2147467259 Then
 Dim sqlquer$
 sqlquer = "INSERT INTO `kassenliste`(name,kurzname,`GO`,`VKNR`,`IK`,`eingef`,pid) values (" & "'" & rFa(i).Kasse & "', '" & rFa(i).KKasse_2 & "', '" & rFa(i).GOÄKatName & "', '" & rFa(i).VKNr & "', '" & rFa(i).IK & "'," & Format(Now(), "yyyymmddHHMMSS") & "," & rFa(i).Pat_id & ")"
 InsKorr DBCn, DBCnS, sqlquer, rAf
 Resume
End If ' ErrNumber = -2147467259 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function roAuZuw(i&, j&)
 roAu(i).FID = rAu(j).FID
 roAu(i).Pat_id = rAu(j).Pat_id
 roAu(i).Zeitpunkt = rAu(j).Zeitpunkt
 roAu(i).Beginn = rAu(j).Beginn
 roAu(i).Ende = rAu(j).Ende
 roAu(i).ICDs = rAu(j).ICDs
 roAu(i).absPos = rAu(j).absPos
 roAu(i).AktZeit = rAu(j).AktZeit
 roAu(i).StByte = rAu(j).StByte
End Function ' roAuZuw

Public Function AuZUnt%(i&, j&)
 If roAu(i).FID <> rAu(j).FID Then GoSub unter
 If roAu(i).Pat_id <> rAu(j).Pat_id Then GoSub unter
 If roAu(i).Zeitpunkt <> rAu(j).Zeitpunkt Then GoSub unter
 If roAu(i).Beginn <> rAu(j).Beginn Then GoSub unter
 If roAu(i).Ende <> rAu(j).Ende Then GoSub unter
 If roAu(i).ICDs <> rAu(j).ICDs Then GoSub unter
 If roAu(i).absPos <> rAu(j).absPos Then GoSub unter
 If roAu(i).AktZeit <> rAu(j).AktZeit Then GoSub unter
 If roAu(i).StByte <> rAu(j).StByte Then GoSub unter
 Exit Function
unter:
 AuZUnt = AuZUnt + 1
 Return
End Function ' AuZUnt

Public Function auLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Beginn,'') Beginn" & _
",COALESCE(Ende,'') Ende,COALESCE(ICDs,'') ICDs,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `au` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roAu(0)
 Else ' rs.EOF Then
  ReDim roAu(1)
  Do While Not rs.EOF
   akt = UBound(roAu)
   roAu(akt).FID = rs!FID
   roAu(akt).Pat_id = rs!Pat_id
   roAu(akt).Zeitpunkt = rs!Zeitpunkt
   roAu(akt).Beginn = doUmwfSQL(rs!Beginn, lies.obMySQL, False)
   roAu(akt).Ende = doUmwfSQL(rs!Ende, lies.obMySQL, False)
   roAu(akt).ICDs = doUmwfSQL(rs!ICDs, lies.obMySQL, False)
   roAu(akt).absPos = rs!absPos
   roAu(akt).AktZeit = rs!AktZeit
   roAu(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roAu(UBound(roAu) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auLaden

Function auEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rAu) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rAu)
   If rAu(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roAu)
    If roAu(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roAu(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roAu(robeg + UBound(rAu) - rbeg)
   For ri = rbeg To UBound(rAu)
    Call roAuZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rAu = roAu
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auEinf

Public Function auSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rAu) + 0 & " Sätze in `au`"
' sql0 = " INSERT " & sqlignore &  "INTO `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            VALUES"))
 For i = 1 To UBound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rAu(i).FID, ",", rAu(i).Pat_id, ",", DatFor_k(rAu(i).Zeitpunkt), ",'", rAu(i).Beginn, "','", rAu(i).Ende, "','", rAu(i).ICDs, "',", rAu(i).absPos, ",", DatFor_k(rAu(i).AktZeit), ",", _
   rAu(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rAu) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `au` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rAu)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""auSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rAu), i)
  If Len(rAu(k).Beginn) > maxi(0) Then maxi(0) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(1) Then maxi(1) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(2) Then maxi(2) = Len(rAu(k).ICDs)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rAu), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function roBrZuw(i&, j&)
 roBr(i).FID = rBr(j).FID
 roBr(i).Pat_id = rBr(j).Pat_id
 roBr(i).Zeitpunkt = rBr(j).Zeitpunkt
 roBr(i).Pfad = rBr(j).Pfad
 roBr(i).art = rBr(j).art
 roBr(i).name = rBr(j).name
 roBr(i).autor = rBr(j).autor
 roBr(i).Quelldatum = rBr(j).Quelldatum
 roBr(i).Typ = rBr(j).Typ
 roBr(i).AktZeit = rBr(j).AktZeit
 roBr(i).DokGroe = rBr(j).DokGroe
 roBr(i).DokAenD = rBr(j).DokAenD
 roBr(i).QS = rBr(j).QS
 roBr(i).QT = rBr(j).QT
 roBr(i).absPos = rBr(j).absPos
 roBr(i).StByte = rBr(j).StByte
 roBr(i).id = rBr(j).id
End Function ' roBrZuw

Public Function BrZUnt%(i&, j&)
 If roBr(i).FID <> rBr(j).FID Then GoSub unter
 If roBr(i).Pat_id <> rBr(j).Pat_id Then GoSub unter
 If roBr(i).Zeitpunkt <> rBr(j).Zeitpunkt Then GoSub unter
 If roBr(i).Pfad <> rBr(j).Pfad Then GoSub unter
 If roBr(i).art <> rBr(j).art Then GoSub unter
 If roBr(i).name <> rBr(j).name Then GoSub unter
 If roBr(i).autor <> rBr(j).autor Then GoSub unter
 If roBr(i).Quelldatum <> rBr(j).Quelldatum Then GoSub unter
 If roBr(i).Typ <> rBr(j).Typ Then GoSub unter
 If roBr(i).AktZeit <> rBr(j).AktZeit Then GoSub unter
 If roBr(i).DokGroe <> rBr(j).DokGroe Then GoSub unter
 If roBr(i).DokAenD <> rBr(j).DokAenD Then GoSub unter
 If roBr(i).QS <> rBr(j).QS Then GoSub unter
 If roBr(i).QT <> rBr(j).QT Then GoSub unter
 If roBr(i).absPos <> rBr(j).absPos Then GoSub unter
 If roBr(i).StByte <> rBr(j).StByte Then GoSub unter
 If roBr(i).id <> rBr(j).id Then GoSub unter
 Exit Function
unter:
 BrZUnt = BrZUnt + 1
 Return
End Function ' BrZUnt

Public Function briefeLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Pfad,'') Pfad" & _
",COALESCE(Art,'') Art,COALESCE(Name,'') Name,COALESCE(autor,'') autor,COALESCE(Quelldatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Quelldatum" & _
",COALESCE(Typ,'') Typ,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(DokGroe,0) DokGroe,COALESCE(DokAenD - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DokAenD" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(absPos,0) absPos,COALESCE(StByte,0) StByte" & _
",COALESCE(ID,0) ID FROM `briefe` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roBr(0)
 Else ' rs.EOF Then
  ReDim roBr(1)
  Do While Not rs.EOF
   akt = UBound(roBr)
   roBr(akt).FID = rs!FID
   roBr(akt).Pat_id = rs!Pat_id
   roBr(akt).Zeitpunkt = rs!Zeitpunkt
   roBr(akt).Pfad = doUmwfSQL(rs!Pfad, lies.obMySQL, False)
   roBr(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roBr(akt).name = doUmwfSQL(rs!name, lies.obMySQL, False)
   roBr(akt).autor = doUmwfSQL(rs!autor, lies.obMySQL, False)
   roBr(akt).Quelldatum = rs!Quelldatum
   roBr(akt).Typ = doUmwfSQL(rs!Typ, lies.obMySQL, False)
   roBr(akt).AktZeit = rs!AktZeit
   roBr(akt).DokGroe = rs!DokGroe
   roBr(akt).DokAenD = rs!DokAenD
   roBr(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roBr(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roBr(akt).absPos = rs!absPos
   roBr(akt).StByte = rs!StByte
   roBr(akt).id = rs!id
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roBr(UBound(roBr) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeLaden

Function briefeEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rBr) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rBr)
   If rBr(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roBr)
    If roBr(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roBr(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roBr(robeg + UBound(rBr) - rbeg)
   For ri = rbeg To UBound(rBr)
    Call roBrZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rBr = roBr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeEinf

Public Function briefeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rBr) + 0 & " Sätze in `briefe`"
' sql0 = " INSERT " & sqlignore &  "INTO `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,autor,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS," & _
     "QT,absPos,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,autor,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS," & _
     "QT,absPos,StByte)      VALUES"))
 For i = 1 To UBound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rBr(i).FID, ",", rBr(i).Pat_id, ",", DatFor_k(rBr(i).Zeitpunkt), ",'", rBr(i).Pfad, "','", rBr(i).art, "','", rBr(i).name, "','", rBr(i).autor, "',", DatFor_k(rBr(i).Quelldatum), ",'", _
   rBr(i).Typ, "',", DatFor_k(rBr(i).AktZeit), ",", rBr(i).DokGroe, ",", DatFor_k(rBr(i).DokAenD), ",'", rBr(i).QS, "','", rBr(i).QT, "',", rBr(i).absPos, ",", rBr(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rBr) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `briefe` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rBr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""briefeSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(6)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rBr), i)
  If Len(rBr(k).Pfad) > maxi(0) Then maxi(0) = Len(rBr(k).Pfad)
  If Len(rBr(k).art) > maxi(1) Then maxi(1) = Len(rBr(k).art)
  If Len(rBr(k).name) > maxi(2) Then maxi(2) = Len(rBr(k).name)
  If Len(rBr(k).autor) > maxi(3) Then maxi(3) = Len(rBr(k).autor)
  If Len(rBr(k).Typ) > maxi(4) Then maxi(4) = Len(rBr(k).Typ)
  If Len(rBr(k).QS) > maxi(5) Then maxi(5) = Len(rBr(k).QS)
  If Len(rBr(k).QT) > maxi(6) Then maxi(6) = Len(rBr(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rBr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL) & "'", True: rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).art & "' -> '" & Left$(rBr(k).art, maxL) & "'", True: rBr(k).art = Left$(rBr(k).art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).name & "' -> '" & Left$(rBr(k).name, maxL) & "'", True: rBr(k).name = Left$(rBr(k).name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.autor: '" & rBr(k).autor & "' -> '" & Left$(rBr(k).autor, maxL) & "'", True: rBr(k).autor = Left$(rBr(k).autor, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function roDiZuw(i&, j&)
 roDi(i).ID1 = rDi(j).ID1
 roDi(i).FID = rDi(j).FID
 roDi(i).Pat_id = rDi(j).Pat_id
 roDi(i).DiagDatum = rDi(j).DiagDatum
 roDi(i).DiagSicherheit = rDi(j).DiagSicherheit
 roDi(i).DiagText = rDi(j).DiagText
 roDi(i).DiagSeite = rDi(j).DiagSeite
 roDi(i).DiagAttr = rDi(j).DiagAttr
 roDi(i).ICD = rDi(j).ICD
 roDi(i).obDauer = rDi(j).obDauer
 roDi(i).intBemerk = rDi(j).intBemerk
 roDi(i).absPos = rDi(j).absPos
 roDi(i).AktZeit = rDi(j).AktZeit
 roDi(i).StByte = rDi(j).StByte
 roDi(i).AusnBegr = rDi(j).AusnBegr
 roDi(i).f6010 = rDi(j).f6010
 roDi(i).obKasse = rDi(j).obKasse
 roDi(i).lKasse = rDi(j).lKasse
 roDi(i).f6011 = rDi(j).f6011
End Function ' roDiZuw

Public Function DiZUnt%(i&, j&)
 If roDi(i).ID1 <> rDi(j).ID1 Then GoSub unter
 If roDi(i).FID <> rDi(j).FID Then GoSub unter
 If roDi(i).Pat_id <> rDi(j).Pat_id Then GoSub unter
 If roDi(i).DiagDatum <> rDi(j).DiagDatum Then GoSub unter
 If roDi(i).DiagSicherheit <> rDi(j).DiagSicherheit Then GoSub unter
 If roDi(i).DiagText <> rDi(j).DiagText Then GoSub unter
 If roDi(i).DiagSeite <> rDi(j).DiagSeite Then GoSub unter
 If roDi(i).DiagAttr <> rDi(j).DiagAttr Then GoSub unter
 If roDi(i).ICD <> rDi(j).ICD Then GoSub unter
 If roDi(i).obDauer <> rDi(j).obDauer Then GoSub unter
 If roDi(i).intBemerk <> rDi(j).intBemerk Then GoSub unter
 If roDi(i).absPos <> rDi(j).absPos Then GoSub unter
 If roDi(i).AktZeit <> rDi(j).AktZeit Then GoSub unter
 If roDi(i).StByte <> rDi(j).StByte Then GoSub unter
 If roDi(i).AusnBegr <> rDi(j).AusnBegr Then GoSub unter
 If roDi(i).f6010 <> rDi(j).f6010 Then GoSub unter
 If roDi(i).obKasse <> rDi(j).obKasse Then GoSub unter
 If roDi(i).lKasse <> rDi(j).lKasse Then GoSub unter
 If roDi(i).f6011 <> rDi(j).f6011 Then GoSub unter
 Exit Function
unter:
 DiZUnt = DiZUnt + 1
 Return
End Function ' DiZUnt

Public Function diagnosenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(ID1,0) ID1,COALESCE(FID,0) FID,COALESCE(Pat_id,0) Pat_id,COALESCE(DiagDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DiagDatum" & _
",COALESCE(DiagSicherheit,'') DiagSicherheit,COALESCE(DiagText,'') DiagText,COALESCE(DiagSeite,'') DiagSeite,COALESCE(DiagAttr,'') DiagAttr" & _
",COALESCE(ICD,'') ICD,COALESCE(obDauer,0) obDauer,COALESCE(intBemerk,'') intBemerk,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte,COALESCE(AusnBegr,'') AusnBegr,COALESCE(f6010,0) f6010" & _
",COALESCE(obKasse,0) obKasse,COALESCE(lKasse - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lKasse,COALESCE(f6011,'') f6011 FROM `diagnosen` WHERE Pat_ID=" & pid & " ORDER BY `DiagDatum`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDi(0)
 Else ' rs.EOF Then
  ReDim roDi(1)
  Do While Not rs.EOF
   akt = UBound(roDi)
   roDi(akt).ID1 = rs!ID1
   roDi(akt).FID = rs!FID
   roDi(akt).Pat_id = rs!Pat_id
   roDi(akt).DiagDatum = rs!DiagDatum
   roDi(akt).DiagSicherheit = doUmwfSQL(rs!DiagSicherheit, lies.obMySQL, False)
   roDi(akt).DiagText = doUmwfSQL(rs!DiagText, lies.obMySQL, False)
   roDi(akt).DiagSeite = doUmwfSQL(rs!DiagSeite, lies.obMySQL, False)
   roDi(akt).DiagAttr = doUmwfSQL(rs!DiagAttr, lies.obMySQL, False)
   roDi(akt).ICD = doUmwfSQL(rs!ICD, lies.obMySQL, False)
   roDi(akt).obDauer = rs!obDauer
   roDi(akt).intBemerk = doUmwfSQL(rs!intBemerk, lies.obMySQL, False)
   roDi(akt).absPos = rs!absPos
   roDi(akt).AktZeit = rs!AktZeit
   roDi(akt).StByte = rs!StByte
   roDi(akt).AusnBegr = doUmwfSQL(rs!AusnBegr, lies.obMySQL, False)
   roDi(akt).f6010 = rs!f6010
   roDi(akt).obKasse = rs!obKasse
   roDi(akt).lKasse = rs!lKasse
   roDi(akt).f6011 = doUmwfSQL(rs!f6011, lies.obMySQL, False)
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roDi(UBound(roDi) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenLaden

Function diagnosenEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDi) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rDi)
   If rDi(ri).DiagDatum >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roDi)
    If roDi(robeg).DiagDatum >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roDi(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDi(robeg + UBound(rDi) - rbeg)
   For ri = rbeg To UBound(rDi)
    Call roDiZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rDi = roDi
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenEinf

Public Function diagnosenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDi) + 0 & " Sätze in `diagnosen`"
' sql0 = " INSERT " & sqlignore &  "INTO `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,obKasse,lKasse,f6011) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,obKasse,lKasse,f6011)               VALUES"))
 For i = 1 To UBound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDi(i).FID, ",", rDi(i).Pat_id, ",", DatFor_k(rDi(i).DiagDatum), ",'", rDi(i).DiagSicherheit, "','", rDi(i).DiagText, "','", rDi(i).DiagSeite, "','", rDi(i).DiagAttr, "','", _
   rDi(i).ICD, "',", rDi(i).obDauer, ",'", rDi(i).intBemerk, "',", rDi(i).absPos, ",", DatFor_k(rDi(i).AktZeit), ",", rDi(i).StByte, ",'", rDi(i).AusnBegr, "',", rDi(i).f6010, ",", rDi(i).obKasse, ",", DatFor_k( _
   rDi(i).lKasse), ",'", rDi(i).f6011, "')")
  If SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDi) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `diagnosen` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rDi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""diagnosenSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(7)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDi), i)
  If Len(rDi(k).DiagSicherheit) > maxi(0) Then maxi(0) = Len(rDi(k).DiagSicherheit)
  If Len(rDi(k).DiagText) > maxi(1) Then maxi(1) = Len(rDi(k).DiagText)
  If Len(rDi(k).DiagSeite) > maxi(2) Then maxi(2) = Len(rDi(k).DiagSeite)
  If Len(rDi(k).DiagAttr) > maxi(3) Then maxi(3) = Len(rDi(k).DiagAttr)
  If Len(rDi(k).ICD) > maxi(4) Then maxi(4) = Len(rDi(k).ICD)
  If Len(rDi(k).intBemerk) > maxi(5) Then maxi(5) = Len(rDi(k).intBemerk)
  If Len(rDi(k).AusnBegr) > maxi(6) Then maxi(6) = Len(rDi(k).AusnBegr)
  If Len(rDi(k).f6011) > maxi(7) Then maxi(7) = Len(rDi(k).f6011)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & Left$(rDi(k).DiagSicherheit, maxL) & "'", True: rDi(k).DiagSicherheit = Left$(rDi(k).DiagSicherheit, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & Left$(rDi(k).DiagText, maxL) & "'", True: rDi(k).DiagText = Left$(rDi(k).DiagText, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & Left$(rDi(k).DiagSeite, maxL) & "'", True: rDi(k).DiagSeite = Left$(rDi(k).DiagSeite, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & Left$(rDi(k).DiagAttr, maxL) & "'", True: rDi(k).DiagAttr = Left$(rDi(k).DiagAttr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & Left$(rDi(k).ICD, maxL) & "'", True: rDi(k).ICD = Left$(rDi(k).ICD, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & Left$(rDi(k).intBemerk, maxL) & "'", True: rDi(k).intBemerk = Left$(rDi(k).intBemerk, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.AusnBegr: '" & rDi(k).AusnBegr & "' -> '" & Left$(rDi(k).AusnBegr, maxL) & "'", True: rDi(k).AusnBegr = Left$(rDi(k).AusnBegr, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.f6011: '" & rDi(k).f6011 & "' -> '" & Left$(rDi(k).f6011, maxL) & "'", True: rDi(k).f6011 = Left$(rDi(k).f6011, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function roDoZuw(i&, j&)
 roDo(i).FID = rDo(j).FID
 roDo(i).Pat_id = rDo(j).Pat_id
 roDo(i).Zeitpunkt = rDo(j).Zeitpunkt
 roDo(i).DokPfad = rDo(j).DokPfad
 roDo(i).DokArt = rDo(j).DokArt
 roDo(i).DokName = rDo(j).DokName
 roDo(i).Quelldatum = rDo(j).Quelldatum
 roDo(i).absPos = rDo(j).absPos
 roDo(i).AktZeit = rDo(j).AktZeit
 roDo(i).DokGroe = rDo(j).DokGroe
 roDo(i).DokAenD = rDo(j).DokAenD
 roDo(i).QS = rDo(j).QS
 roDo(i).QT = rDo(j).QT
 roDo(i).StByte = rDo(j).StByte
End Function ' roDoZuw

Public Function DoZUnt%(i&, j&)
 If roDo(i).FID <> rDo(j).FID Then GoSub unter
 If roDo(i).Pat_id <> rDo(j).Pat_id Then GoSub unter
 If roDo(i).Zeitpunkt <> rDo(j).Zeitpunkt Then GoSub unter
 If roDo(i).DokPfad <> rDo(j).DokPfad Then GoSub unter
 If roDo(i).DokArt <> rDo(j).DokArt Then GoSub unter
 If roDo(i).DokName <> rDo(j).DokName Then GoSub unter
 If roDo(i).Quelldatum <> rDo(j).Quelldatum Then GoSub unter
 If roDo(i).absPos <> rDo(j).absPos Then GoSub unter
 If roDo(i).AktZeit <> rDo(j).AktZeit Then GoSub unter
 If roDo(i).DokGroe <> rDo(j).DokGroe Then GoSub unter
 If roDo(i).DokAenD <> rDo(j).DokAenD Then GoSub unter
 If roDo(i).QS <> rDo(j).QS Then GoSub unter
 If roDo(i).QT <> rDo(j).QT Then GoSub unter
 If roDo(i).StByte <> rDo(j).StByte Then GoSub unter
 Exit Function
unter:
 DoZUnt = DoZUnt + 1
 Return
End Function ' DoZUnt

Public Function dokumenteLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(DokPfad,'') DokPfad" & _
",COALESCE(DokArt,'') DokArt,COALESCE(DokName,'') DokName,COALESCE(Quelldatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Quelldatum,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(DokGroe,0) DokGroe,COALESCE(DokAenD - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DokAenD,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte FROM `dokumente` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDo(0)
 Else ' rs.EOF Then
  ReDim roDo(1)
  Do While Not rs.EOF
   akt = UBound(roDo)
   roDo(akt).FID = rs!FID
   roDo(akt).Pat_id = rs!Pat_id
   roDo(akt).Zeitpunkt = rs!Zeitpunkt
   roDo(akt).DokPfad = doUmwfSQL(rs!DokPfad, lies.obMySQL, False)
   roDo(akt).DokArt = doUmwfSQL(rs!DokArt, lies.obMySQL, False)
   roDo(akt).DokName = doUmwfSQL(rs!DokName, lies.obMySQL, False)
   roDo(akt).Quelldatum = rs!Quelldatum
   roDo(akt).absPos = rs!absPos
   roDo(akt).AktZeit = rs!AktZeit
   roDo(akt).DokGroe = rs!DokGroe
   roDo(akt).DokAenD = rs!DokAenD
   roDo(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roDo(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roDo(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roDo(UBound(roDo) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteLaden

Function dokumenteEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDo) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rDo)
   If rDo(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roDo)
    If roDo(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roDo(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDo(robeg + UBound(rDo) - rbeg)
   For ri = rbeg To UBound(rDo)
    Call roDoZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rDo = roDo
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteEinf

Public Function dokumenteSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDo) + 0 & " Sätze in `dokumente`"
' sql0 = " INSERT " & sqlignore &  "INTO `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte)  VALUES"))
 For i = 1 To UBound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rDo(i).FID, ",", rDo(i).Pat_id, ",", DatFor_k(rDo(i).Zeitpunkt), ",'", rDo(i).DokPfad, "','", rDo(i).DokArt, "','", rDo(i).DokName, "',", DatFor_k(rDo(i).Quelldatum), ",", _
   rDo(i).absPos, ",", DatFor_k(rDo(i).AktZeit), ",", rDo(i).DokGroe, ",", DatFor_k(rDo(i).DokAenD), ",'", rDo(i).QS, "','", rDo(i).QT, "',", rDo(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDo) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `dokumente` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rDo)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""dokumenteSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDo), i)
  If Len(rDo(k).DokPfad) > maxi(0) Then maxi(0) = Len(rDo(k).DokPfad)
  If Len(rDo(k).DokArt) > maxi(1) Then maxi(1) = Len(rDo(k).DokArt)
  If Len(rDo(k).DokName) > maxi(2) Then maxi(2) = Len(rDo(k).DokName)
  If Len(rDo(k).QS) > maxi(3) Then maxi(3) = Len(rDo(k).QS)
  If Len(rDo(k).QT) > maxi(4) Then maxi(4) = Len(rDo(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDo), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function roEiZuw(i&, j&)
 roEi(i).FID = rEi(j).FID
 roEi(i).Pat_id = rEi(j).Pat_id
 roEi(i).Zeitpunkt = rEi(j).Zeitpunkt
 roEi(i).art = rEi(j).art
 roEi(i).Inhalt = rEi(j).Inhalt
 roEi(i).absPos = rEi(j).absPos
 roEi(i).AktZeit = rEi(j).AktZeit
 roEi(i).QS = rEi(j).QS
 roEi(i).QT = rEi(j).QT
 roEi(i).StByte = rEi(j).StByte
 roEi(i).id = rEi(j).id
 roEi(i).inhNum = rEi(j).inhNum
End Function ' roEiZuw

Public Function EiZUnt%(i&, j&)
 If roEi(i).FID <> rEi(j).FID Then GoSub unter
 If roEi(i).Pat_id <> rEi(j).Pat_id Then GoSub unter
 If roEi(i).Zeitpunkt <> rEi(j).Zeitpunkt Then GoSub unter
 If roEi(i).art <> rEi(j).art Then GoSub unter
 If roEi(i).Inhalt <> rEi(j).Inhalt Then GoSub unter
 If roEi(i).absPos <> rEi(j).absPos Then GoSub unter
 If roEi(i).AktZeit <> rEi(j).AktZeit Then GoSub unter
 If roEi(i).QS <> rEi(j).QS Then GoSub unter
 If roEi(i).QT <> rEi(j).QT Then GoSub unter
 If roEi(i).StByte <> rEi(j).StByte Then GoSub unter
 If roEi(i).id <> rEi(j).id Then GoSub unter
 If roEi(i).inhNum <> rEi(j).inhNum Then GoSub unter
 Exit Function
unter:
 EiZUnt = EiZUnt + 1
 Return
End Function ' EiZUnt

Public Function eintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Inhalt,'') Inhalt,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(id,0) id,COALESCE(inhNum,0) inhNum" & _
" FROM `eintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roEi(0)
 Else ' rs.EOF Then
  ReDim roEi(1)
  Do While Not rs.EOF
   akt = UBound(roEi)
   roEi(akt).FID = rs!FID
   roEi(akt).Pat_id = rs!Pat_id
   roEi(akt).Zeitpunkt = rs!Zeitpunkt
   roEi(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roEi(akt).Inhalt = doUmwfSQL(rs!Inhalt, lies.obMySQL, False)
   roEi(akt).absPos = rs!absPos
   roEi(akt).AktZeit = rs!AktZeit
   roEi(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roEi(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roEi(akt).StByte = rs!StByte
   roEi(akt).id = rs!id
   roEi(akt).inhNum = rs!inhNum
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roEi(UBound(roEi) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeLaden

Function eintraegeEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rEi) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rEi)
   If rEi(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roEi)
    If roEi(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roEi(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roEi(robeg + UBound(rEi) - rbeg)
   For ri = rbeg To UBound(rEi)
    Call roEiZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rEi = roEi
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeEinf

Public Function eintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rEi) + 0 & " Sätze in `eintraege`"
' sql0 = " INSERT " & sqlignore &  "INTO `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum)     VALUES"))
 For i = 1 To UBound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rEi(i).FID, ",", rEi(i).Pat_id, ",", DatFor_k(rEi(i).Zeitpunkt), ",'", rEi(i).art, "','", rEi(i).Inhalt, "',", rEi(i).absPos, ",", DatFor_k(rEi(i).AktZeit), ",'", rEi(i).QS, "','", _
   rEi(i).QT, "',", rEi(i).StByte, ",", REPLACE$(rEi(i).inhNum, ",", "."), ")")
  If SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rEi) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `eintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rEi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""eintraegeSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rEi), i)
  If Len(rEi(k).art) > maxi(0) Then maxi(0) = Len(rEi(k).art)
  If Len(rEi(k).Inhalt) > maxi(1) Then maxi(1) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(2) Then maxi(2) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(3) Then maxi(3) = Len(rEi(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rEi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).art & "' -> '" & Left$(rEi(k).art, maxL) & "'", True: rEi(k).art = Left$(rEi(k).art, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function forminhaltform_abkSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFi) + 0 & " Sätze in `forminhaltform_abk`"
' sql0 = " INSERT " & sqlignore &  "INTO `forminhaltform_abk` (Form_Abk) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `forminhaltform_abk` (Form_Abk)             VALUES"))
 For i = rFi1 + 1 To UBound(rFi)
'  rFi(i).AktZeit = now()
  If SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rFi(i).Form_Abk, "')")
  If SammelInsert <> 0 And i < UBound(rFi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFi) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rFi)
  DoEvents
 Next i
 rFi1 = UBound(rFi)
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""forminhaltform_abkSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(SammelInsert <> 0, rFi1 + 1, i) To IIf(SammelInsert <> 0, UBound(rFi), i)
  If Len(rFi(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFi(k).Form_Abk)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, rFi1 + 1, i) To IIf(SammelInsert <> 0, UBound(rFi), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhaltform_abkSpeichern

Public Function formulareSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFo) + 0 & " Sätze in `formulare`"
' sql0 = " INSERT " & sqlignore &  "INTO `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      VALUES"))
 For i = rFo1 + 1 To UBound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFo(i).FormID, ",'", rFo(i).Form_Abk, "','", rFo(i).FormBez, "','", rFo(i).FormVorl, "',", DatFor_k(rFo(i).AktZeit), ",", rFo(i).absPos, ",", rFo(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rFo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFo) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rFo)
  DoEvents
 Next i
 rFo1 = UBound(rFo)
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""formulareSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(SammelInsert <> 0, rFo1 + 1, i) To IIf(SammelInsert <> 0, UBound(rFo), i)
  If Len(rFo(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFo(k).Form_Abk)
  If Len(rFo(k).FormBez) > maxi(1) Then maxi(1) = Len(rFo(k).FormBez)
  If Len(rFo(k).FormVorl) > maxi(2) Then maxi(2) = Len(rFo(k).FormVorl)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, rFo1 + 1, i) To IIf(SammelInsert <> 0, UBound(rFo), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function roFrZuw(i&, j&)
 roFr(i).Foid = rFr(j).Foid
 roFr(i).FID = rFr(j).FID
 roFr(i).Pat_id = rFr(j).Pat_id
 roFr(i).Form_ID = rFr(j).Form_ID
 roFr(i).Zeitpunkt = rFr(j).Zeitpunkt
 roFr(i).absPos = rFr(j).absPos
 roFr(i).AktZeit = rFr(j).AktZeit
 roFr(i).StByte = rFr(j).StByte
 roFr(i).Satzart = rFr(j).Satzart
 roFr(i).Satzlänge = rFr(j).Satzlänge
 roFr(i).lanrid = rFr(j).lanrid
End Function ' roFrZuw

Public Function FrZUnt%(i&, j&)
 If roFr(i).Foid <> rFr(j).Foid Then GoSub unter
 If roFr(i).FID <> rFr(j).FID Then GoSub unter
 If roFr(i).Pat_id <> rFr(j).Pat_id Then GoSub unter
 If roFr(i).Form_ID <> rFr(j).Form_ID Then GoSub unter
 If roFr(i).Zeitpunkt <> rFr(j).Zeitpunkt Then GoSub unter
 If roFr(i).absPos <> rFr(j).absPos Then GoSub unter
 If roFr(i).AktZeit <> rFr(j).AktZeit Then GoSub unter
 If roFr(i).StByte <> rFr(j).StByte Then GoSub unter
 If roFr(i).Satzart <> rFr(j).Satzart Then GoSub unter
 If roFr(i).Satzlänge <> rFr(j).Satzlänge Then GoSub unter
 If roFr(i).lanrid <> rFr(j).lanrid Then GoSub unter
 Exit Function
unter:
 FrZUnt = FrZUnt + 1
 Return
End Function ' FrZUnt

Public Function forminhkopfLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FoID,0) FoID,COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(Form_ID,0) Form_ID" & _
",COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(AbsPos,0) AbsPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
",COALESCE(Satzart,'') Satzart,COALESCE(Satzlänge,'') Satzlänge,COALESCE(LANRid,0) LANRid FROM `forminhkopf` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roFr(0)
 Else ' rs.EOF Then
  ReDim roFr(1)
  Do While Not rs.EOF
   akt = UBound(roFr)
   roFr(akt).Foid = rs!Foid
   roFr(akt).FID = rs!FID
   roFr(akt).Pat_id = rs!Pat_id
   roFr(akt).Form_ID = rs!Form_ID
   roFr(akt).Zeitpunkt = rs!Zeitpunkt
   roFr(akt).absPos = rs!absPos
   roFr(akt).AktZeit = rs!AktZeit
   roFr(akt).StByte = rs!StByte
   roFr(akt).Satzart = doUmwfSQL(rs!Satzart, lies.obMySQL, False)
   roFr(akt).Satzlänge = doUmwfSQL(rs!Satzlänge, lies.obMySQL, False)
   roFr(akt).lanrid = rs!lanrid
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roFr(UBound(roFr) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfLaden

Function forminhkopfEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rFr) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rFr)
   If rFr(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roFr)
    If roFr(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roFr(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFr(robeg + UBound(rFr) - rbeg)
   For ri = rbeg To UBound(rFr)
    Call roFrZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rFr = roFr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfEinf

Public Function forminhkopfSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFr) + 0 & " Sätze in `forminhkopf`"
' sql0 = " INSERT " & sqlignore &  "INTO `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid)              VALUES"))
 FoIDv = 0
erneut:
 For i = 1 To UBound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFr(i).Foid, ",", rFr(i).FID, ",", rFr(i).Pat_id, ",", rFr(i).Form_ID, ",", DatFor_k(rFr(i).Zeitpunkt), ",", rFr(i).absPos, ",", DatFor_k(rFr(i).AktZeit), ",", rFr(i).StByte, ",'", _
   rFr(i).Satzart, "','", rFr(i).Satzlänge, "',", rFr(i).lanrid, ")")
  If SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFr) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
'    sql = "DELETE FROM `forminhfeld` WHERE foid IN (SELECT foid FROM `forminhkopf` WHERE pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'    Call myEFrag(sql)
     sql = "DELETE FROM `forminhkopf` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rFr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
If ErrNumber = -2147217900 And ErrDescr Like "*Duplicate entry * for key 'PRIMARY'" Then
 Dim schlüssel$, pos&, iiru&, jjru&
 pos = InStr(ErrDescr, "'")
 schlüssel = Mid$(ErrDescr, pos + 1, InStr(pos + 2, ErrDescr, "'") - pos - 1)
' Debug.Print schlüssel
 For iiru = 1 To UBound(rFr)
  If rFr(iiru).Foid = schlüssel Then
   If FoIDv = 0 Then FoIDv = DBCn.Execute("SELECT (MAX(foid)+1) FROM forminhkopf").Fields(0) Else FoIDv = FoIDv + 1
   rFr(iiru).Foid = FoIDv
   For jjru = 1 To UBound(rFm)
    If rFm(jjru).Foid = schlüssel Then rFm(jjru).Foid = FoIDv
   Next jjru
   Exit For
  End If ' rFr(iiru).Foid = schlüssel Then
 Next iiru
 csql = ""
 Resume erneut
End If ' ErrNumber = -2147217900 And ErrDescr Like "*Duplicate entry * for key 'PRIMARY'" Then
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""forminhkopfSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFr), i)
  If Len(rFr(k).Satzart) > maxi(0) Then maxi(0) = Len(rFr(k).Satzart)
  If Len(rFr(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rFr(k).Satzlänge)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFr), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function forminhfeldSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFm) + 0 & " Sätze in `forminhfeld`"
' sql0 = " INSERT " & sqlignore &  "INTO `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      VALUES"))
 For i = 1 To UBound(rFm)
'  rFm(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFm(i).Foid, ",", rFm(i).nr, ",", rFm(i).FeldNr, ",", rFm(i).FeldVW, ",", rFm(i).FeldInhVW, ")")
  If SammelInsert <> 0 And i < UBound(rFm) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFm) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rFm)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""forminhfeldSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(-1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFm), i)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhfeld", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFm), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function roKhZuw(i&, j&)
 roKh(i).FID = rKh(j).FID
 roKh(i).Pat_id = rKh(j).Pat_id
 roKh(i).Zeitpunkt = rKh(j).Zeitpunkt
 roKh(i).Ziel = rKh(j).Ziel
 roKh(i).Diagnose = rKh(j).Diagnose
 roKh(i).absPos = rKh(j).absPos
 roKh(i).AktZeit = rKh(j).AktZeit
 roKh(i).StByte = rKh(j).StByte
End Function ' roKhZuw

Public Function KhZUnt%(i&, j&)
 If roKh(i).FID <> rKh(j).FID Then GoSub unter
 If roKh(i).Pat_id <> rKh(j).Pat_id Then GoSub unter
 If roKh(i).Zeitpunkt <> rKh(j).Zeitpunkt Then GoSub unter
 If roKh(i).Ziel <> rKh(j).Ziel Then GoSub unter
 If roKh(i).Diagnose <> rKh(j).Diagnose Then GoSub unter
 If roKh(i).absPos <> rKh(j).absPos Then GoSub unter
 If roKh(i).AktZeit <> rKh(j).AktZeit Then GoSub unter
 If roKh(i).StByte <> rKh(j).StByte Then GoSub unter
 Exit Function
unter:
 KhZUnt = KhZUnt + 1
 Return
End Function ' KhZUnt

Public Function kheinweisLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Ziel,'') Ziel" & _
",COALESCE(Diagnose,'') Diagnose,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
" FROM `kheinweis` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roKh(0)
 Else ' rs.EOF Then
  ReDim roKh(1)
  Do While Not rs.EOF
   akt = UBound(roKh)
   roKh(akt).FID = rs!FID
   roKh(akt).Pat_id = rs!Pat_id
   roKh(akt).Zeitpunkt = rs!Zeitpunkt
   roKh(akt).Ziel = doUmwfSQL(rs!Ziel, lies.obMySQL, False)
   roKh(akt).Diagnose = doUmwfSQL(rs!Diagnose, lies.obMySQL, False)
   roKh(akt).absPos = rs!absPos
   roKh(akt).AktZeit = rs!AktZeit
   roKh(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roKh(UBound(roKh) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisLaden

Function kheinweisEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rKh) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rKh)
   If rKh(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roKh)
    If roKh(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roKh(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roKh(robeg + UBound(rKh) - rbeg)
   For ri = rbeg To UBound(rKh)
    Call roKhZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rKh = roKh
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisEinf

Public Function kheinweisSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rKh) + 0 & " Sätze in `kheinweis`"
' sql0 = " INSERT " & sqlignore &  "INTO `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               VALUES"))
 For i = 1 To UBound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKh(i).FID, ",", rKh(i).Pat_id, ",", DatFor_k(rKh(i).Zeitpunkt), ",'", rKh(i).Ziel, "','", rKh(i).Diagnose, "',", rKh(i).absPos, ",", DatFor_k(rKh(i).AktZeit), ",", rKh(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rKh) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `kheinweis` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rKh)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""kheinweisSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKh), i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKh), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function roLbZuw(i&, j&)
 roLb(i).FID = rLb(j).FID
 roLb(i).Pat_id = rLb(j).Pat_id
 roLb(i).Zeitpunkt = rLb(j).Zeitpunkt
 roLb(i).AnfText = rLb(j).AnfText
 roLb(i).absPos = rLb(j).absPos
 roLb(i).AktZeit = rLb(j).AktZeit
 roLb(i).StByte = rLb(j).StByte
End Function ' roLbZuw

Public Function LbZUnt%(i&, j&)
 If roLb(i).FID <> rLb(j).FID Then GoSub unter
 If roLb(i).Pat_id <> rLb(j).Pat_id Then GoSub unter
 If roLb(i).Zeitpunkt <> rLb(j).Zeitpunkt Then GoSub unter
 If roLb(i).AnfText <> rLb(j).AnfText Then GoSub unter
 If roLb(i).absPos <> rLb(j).absPos Then GoSub unter
 If roLb(i).AktZeit <> rLb(j).AktZeit Then GoSub unter
 If roLb(i).StByte <> rLb(j).StByte Then GoSub unter
 Exit Function
unter:
 LbZUnt = LbZUnt + 1
 Return
End Function ' LbZUnt

Public Function lbanforderungenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(AnfText,'') AnfText" & _
",COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `lbanforderungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roLb(0)
 Else ' rs.EOF Then
  ReDim roLb(1)
  Do While Not rs.EOF
   akt = UBound(roLb)
   roLb(akt).FID = rs!FID
   roLb(akt).Pat_id = rs!Pat_id
   roLb(akt).Zeitpunkt = rs!Zeitpunkt
   roLb(akt).AnfText = doUmwfSQL(rs!AnfText, lies.obMySQL, False)
   roLb(akt).absPos = rs!absPos
   roLb(akt).AktZeit = rs!AktZeit
   roLb(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roLb(UBound(roLb) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenLaden

Function lbanforderungenEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLb) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rLb)
   If rLb(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roLb)
    If roLb(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roLb(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLb(robeg + UBound(rLb) - rbeg)
   For ri = rbeg To UBound(rLb)
    Call roLbZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rLb = roLb
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenEinf

Public Function lbanforderungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLb) + 0 & " Sätze in `lbanforderungen`"
' sql0 = " INSERT " & sqlignore &  "INTO `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       VALUES"))
 For i = 1 To UBound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLb(i).FID, ",", rLb(i).Pat_id, ",", DatFor_k(rLb(i).Zeitpunkt), ",'", rLb(i).AnfText, "',", rLb(i).absPos, ",", DatFor_k(rLb(i).AktZeit), ",", rLb(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLb) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `lbanforderungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLb)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""lbanforderungenSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLb), i)
  If Len(rLb(k).AnfText) > maxi(0) Then maxi(0) = Len(rLb(k).AnfText)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLb), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function roLaZuw(i&, j&)
 roLa(i).FID = rLa(j).FID
 roLa(i).Pat_id = rLa(j).Pat_id
 roLa(i).Zeitpunkt = rLa(j).Zeitpunkt
 roLa(i).FertigStGrad = rLa(j).FertigStGrad
 roLa(i).Abkü = rLa(j).Abkü
 roLa(i).LangtextVW = rLa(j).LangtextVW
 roLa(i).Wert = rLa(j).Wert
 roLa(i).Einheit = rLa(j).Einheit
 roLa(i).KommentarVW = rLa(j).KommentarVW
 roLa(i).absPos = rLa(j).absPos
 roLa(i).AktZeit = rLa(j).AktZeit
 roLa(i).RefNr = rLa(j).RefNr
 roLa(i).StByte = rLa(j).StByte
 roLa(i).id = rLa(j).id
End Function ' roLaZuw

Public Function LaZUnt%(i&, j&)
 If roLa(i).FID <> rLa(j).FID Then GoSub unter
 If roLa(i).Pat_id <> rLa(j).Pat_id Then GoSub unter
 If roLa(i).Zeitpunkt <> rLa(j).Zeitpunkt Then GoSub unter
 If roLa(i).FertigStGrad <> rLa(j).FertigStGrad Then GoSub unter
 If roLa(i).Abkü <> rLa(j).Abkü Then GoSub unter
 If roLa(i).LangtextVW <> rLa(j).LangtextVW Then GoSub unter
 If roLa(i).Wert <> rLa(j).Wert Then GoSub unter
 If roLa(i).Einheit <> rLa(j).Einheit Then GoSub unter
 If roLa(i).KommentarVW <> rLa(j).KommentarVW Then GoSub unter
 If roLa(i).absPos <> rLa(j).absPos Then GoSub unter
 If roLa(i).AktZeit <> rLa(j).AktZeit Then GoSub unter
 If roLa(i).RefNr <> rLa(j).RefNr Then GoSub unter
 If roLa(i).StByte <> rLa(j).StByte Then GoSub unter
 If roLa(i).id <> rLa(j).id Then GoSub unter
 Exit Function
unter:
 LaZUnt = LaZUnt + 1
 Return
End Function ' LaZUnt

Public Function laborneuLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(FertigStGrad,'') FertigStGrad" & _
",COALESCE(Abkü,'') Abkü,COALESCE(LangtextVW,0) LangtextVW,COALESCE(Wert,'') Wert,COALESCE(Einheit,'') Einheit" & _
",COALESCE(KommentarVW,0) KommentarVW,COALESCE(AbsPos,0) AbsPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(Refnr,0) Refnr" & _
",COALESCE(StByte,0) StByte,COALESCE(ID,0) ID FROM `laborneu` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roLa(0)
 Else ' rs.EOF Then
  ReDim roLa(1)
  Do While Not rs.EOF
   akt = UBound(roLa)
   roLa(akt).FID = rs!FID
   roLa(akt).Pat_id = rs!Pat_id
   roLa(akt).Zeitpunkt = rs!Zeitpunkt
   roLa(akt).FertigStGrad = doUmwfSQL(rs!FertigStGrad, lies.obMySQL, False)
   roLa(akt).Abkü = doUmwfSQL(rs!Abkü, lies.obMySQL, False)
   roLa(akt).LangtextVW = rs!LangtextVW
   roLa(akt).Wert = doUmwfSQL(rs!Wert, lies.obMySQL, False)
   roLa(akt).Einheit = doUmwfSQL(rs!Einheit, lies.obMySQL, False)
   roLa(akt).KommentarVW = rs!KommentarVW
   roLa(akt).absPos = rs!absPos
   roLa(akt).AktZeit = rs!AktZeit
   roLa(akt).RefNr = rs!RefNr
   roLa(akt).StByte = rs!StByte
   roLa(akt).id = rs!id
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roLa(UBound(roLa) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuLaden

Function laborneuEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLa) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rLa)
   If rLa(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roLa)
    If roLa(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roLa(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLa(robeg + UBound(rLa) - rbeg)
   For ri = rbeg To UBound(rLa)
    Call roLaZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rLa = roLa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuEinf

Public Function laborneuSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLa) + 0 & " Sätze in `laborneu`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte)           VALUES"))
 For i = 1 To UBound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLa(i).FID, ",", rLa(i).Pat_id, ",", DatFor_k(rLa(i).Zeitpunkt), ",'", rLa(i).FertigStGrad, "','", rLa(i).Abkü, "',", rLa(i).LangtextVW, ",'", rLa(i).Wert, "','", rLa(i).Einheit, "',", _
   rLa(i).KommentarVW, ",", rLa(i).absPos, ",", DatFor_k(rLa(i).AktZeit), ",", rLa(i).RefNr, ",", rLa(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLa) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `laborneu` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLa)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborneuSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(3)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLa), i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLa), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function roLeZuw(i&, j&)
 roLe(i).id = rLe(j).id
 roLe(i).FID = rLe(j).FID
 roLe(i).Pat_id = rLe(j).Pat_id
 roLe(i).Zeitpunkt = rLe(j).Zeitpunkt
 roLe(i).Leistung = rLe(j).Leistung
 roLe(i).f5002 = rLe(j).f5002
 roLe(i).f5005 = rLe(j).f5005
 roLe(i).f5006 = rLe(j).f5006
 roLe(i).f5009 = rLe(j).f5009
 roLe(i).Med = rLe(j).Med
 roLe(i).f5015 = rLe(j).f5015
 roLe(i).f5016 = rLe(j).f5016
 roLe(i).f5021 = rLe(j).f5021
 roLe(i).f5026 = rLe(j).f5026
 roLe(i).Faktor = rLe(j).Faktor
 roLe(i).f5098 = rLe(j).f5098
 roLe(i).Charge = rLe(j).Charge
 roLe(i).Lanr = rLe(j).Lanr
 roLe(i).letzVorg = rLe(j).letzVorg
 roLe(i).Ausn = rLe(j).Ausn
 roLe(i).beme = rLe(j).beme
 roLe(i).absPos = rLe(j).absPos
 roLe(i).AktZeit = rLe(j).AktZeit
 roLe(i).QS = rLe(j).QS
 roLe(i).QT = rLe(j).QT
 roLe(i).StByte = rLe(j).StByte
 roLe(i).lanrid = rLe(j).lanrid
 roLe(i).Sachkbez = rLe(j).Sachkbez
 roLe(i).Sachkct = rLe(j).Sachkct
 roLe(i).Zone = rLe(j).Zone
End Function ' roLeZuw

Public Function LeZUnt%(i&, j&)
 If roLe(i).id <> rLe(j).id Then GoSub unter
 If roLe(i).FID <> rLe(j).FID Then GoSub unter
 If roLe(i).Pat_id <> rLe(j).Pat_id Then GoSub unter
 If roLe(i).Zeitpunkt <> rLe(j).Zeitpunkt Then GoSub unter
 If roLe(i).Leistung <> rLe(j).Leistung Then GoSub unter
 If roLe(i).f5002 <> rLe(j).f5002 Then GoSub unter
 If roLe(i).f5005 <> rLe(j).f5005 Then GoSub unter
 If roLe(i).f5006 <> rLe(j).f5006 Then GoSub unter
 If roLe(i).f5009 <> rLe(j).f5009 Then GoSub unter
 If roLe(i).Med <> rLe(j).Med Then GoSub unter
 If roLe(i).f5015 <> rLe(j).f5015 Then GoSub unter
 If roLe(i).f5016 <> rLe(j).f5016 Then GoSub unter
 If roLe(i).f5021 <> rLe(j).f5021 Then GoSub unter
 If roLe(i).f5026 <> rLe(j).f5026 Then GoSub unter
 If roLe(i).Faktor <> rLe(j).Faktor Then GoSub unter
 If roLe(i).f5098 <> rLe(j).f5098 Then GoSub unter
 If roLe(i).Charge <> rLe(j).Charge Then GoSub unter
 If roLe(i).Lanr <> rLe(j).Lanr Then GoSub unter
 If roLe(i).letzVorg <> rLe(j).letzVorg Then GoSub unter
 If roLe(i).Ausn <> rLe(j).Ausn Then GoSub unter
 If roLe(i).beme <> rLe(j).beme Then GoSub unter
 If roLe(i).absPos <> rLe(j).absPos Then GoSub unter
 If roLe(i).AktZeit <> rLe(j).AktZeit Then GoSub unter
 If roLe(i).QS <> rLe(j).QS Then GoSub unter
 If roLe(i).QT <> rLe(j).QT Then GoSub unter
 If roLe(i).StByte <> rLe(j).StByte Then GoSub unter
 If roLe(i).lanrid <> rLe(j).lanrid Then GoSub unter
 If roLe(i).Sachkbez <> rLe(j).Sachkbez Then GoSub unter
 If roLe(i).Sachkct <> rLe(j).Sachkct Then GoSub unter
 If roLe(i).Zone <> rLe(j).Zone Then GoSub unter
 Exit Function
unter:
 LeZUnt = LeZUnt + 1
 Return
End Function ' LeZUnt

Public Function leistungenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(id,0) id,COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt" & _
",COALESCE(Leistung,'') Leistung,COALESCE(f5002,'') f5002,COALESCE(f5005,'') f5005,COALESCE(f5006,'') f5006" & _
",COALESCE(f5009,'') f5009,COALESCE(Med,'') Med,COALESCE(f5015,'') f5015,COALESCE(f5016,'') f5016" & _
",COALESCE(f5021,'') f5021,COALESCE(f5026,'') f5026,COALESCE(Faktor,'') Faktor,COALESCE(f5098,'') f5098" & _
",COALESCE(Charge,'') Charge,COALESCE(LANR,'') LANR,COALESCE(letzVorg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) letzVorg,COALESCE(Ausn,'') Ausn" & _
",COALESCE(Beme,'') Beme,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid,COALESCE(Sachkbez,'') Sachkbez" & _
",COALESCE(Sachkct,0) Sachkct,COALESCE(Zone,'') Zone FROM `leistungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roLe(0)
 Else ' rs.EOF Then
  ReDim roLe(1)
  Do While Not rs.EOF
   akt = UBound(roLe)
   roLe(akt).id = rs!id
   roLe(akt).FID = rs!FID
   roLe(akt).Pat_id = rs!Pat_id
   roLe(akt).Zeitpunkt = rs!Zeitpunkt
   roLe(akt).Leistung = doUmwfSQL(rs!Leistung, lies.obMySQL, False)
   roLe(akt).f5002 = doUmwfSQL(rs!f5002, lies.obMySQL, False)
   roLe(akt).f5005 = doUmwfSQL(rs!f5005, lies.obMySQL, False)
   roLe(akt).f5006 = doUmwfSQL(rs!f5006, lies.obMySQL, False)
   roLe(akt).f5009 = doUmwfSQL(rs!f5009, lies.obMySQL, False)
   roLe(akt).Med = doUmwfSQL(rs!Med, lies.obMySQL, False)
   roLe(akt).f5015 = doUmwfSQL(rs!f5015, lies.obMySQL, False)
   roLe(akt).f5016 = doUmwfSQL(rs!f5016, lies.obMySQL, False)
   roLe(akt).f5021 = doUmwfSQL(rs!f5021, lies.obMySQL, False)
   roLe(akt).f5026 = doUmwfSQL(rs!f5026, lies.obMySQL, False)
   roLe(akt).Faktor = doUmwfSQL(rs!Faktor, lies.obMySQL, False)
   roLe(akt).f5098 = doUmwfSQL(rs!f5098, lies.obMySQL, False)
   roLe(akt).Charge = doUmwfSQL(rs!Charge, lies.obMySQL, False)
   roLe(akt).Lanr = doUmwfSQL(rs!Lanr, lies.obMySQL, False)
   roLe(akt).letzVorg = rs!letzVorg
   roLe(akt).Ausn = doUmwfSQL(rs!Ausn, lies.obMySQL, False)
   roLe(akt).beme = doUmwfSQL(rs!beme, lies.obMySQL, False)
   roLe(akt).absPos = rs!absPos
   roLe(akt).AktZeit = rs!AktZeit
   roLe(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roLe(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roLe(akt).StByte = rs!StByte
   roLe(akt).lanrid = rs!lanrid
   roLe(akt).Sachkbez = doUmwfSQL(rs!Sachkbez, lies.obMySQL, False)
   roLe(akt).Sachkct = rs!Sachkct
   roLe(akt).Zone = doUmwfSQL(rs!Zone, lies.obMySQL, False)
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roLe(UBound(roLe) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenLaden

Function leistungenEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLe) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rLe)
   If rLe(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roLe)
    If roLe(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roLe(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLe(robeg + UBound(rLe) - rbeg)
   For ri = rbeg To UBound(rLe)
    Call roLeZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rLe = roLe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenEinf

Public Function leistungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLe) + 0 & " Sätze in `leistungen`"
' sql0 = " INSERT " & sqlignore &  "INTO `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,Charge,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS," & _
     "QT,StByte,LANRid,Sachkbez,Sachkct,Zone) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,Charge,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS," & _
     "QT,StByte,LANRid,Sachkbez,Sachkct,Zone)            VALUES"))
 For i = 1 To UBound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLe(i).FID, ",", rLe(i).Pat_id, ",", DatFor_k(rLe(i).Zeitpunkt), ",'", rLe(i).Leistung, "','", rLe(i).f5002, "','", rLe(i).f5005, "','", rLe(i).f5006, "','", rLe(i).f5009, "','", _
   rLe(i).Med, "','", rLe(i).f5015, "','", rLe(i).f5016, "','", rLe(i).f5021, "','", rLe(i).f5026, "','", rLe(i).Faktor, "','", rLe(i).f5098, "','", rLe(i).Charge, "','", rLe(i).Lanr, "',", DatFor_k( _
   rLe(i).letzVorg), ",'", rLe(i).Ausn, "','", rLe(i).beme, "',", rLe(i).absPos, ",", DatFor_k(rLe(i).AktZeit), ",'", rLe(i).QS, "','", rLe(i).QT, "',", rLe(i).StByte, ",", _
   rLe(i).lanrid, ",'", rLe(i).Sachkbez, "',", rLe(i).Sachkct, ",'", rLe(i).Zone, "')")
  If SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLe) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `leistungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""leistungenSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(19)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLe), i)
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
  If Len(rLe(k).Charge) > maxi(12) Then maxi(12) = Len(rLe(k).Charge)
  If Len(rLe(k).Lanr) > maxi(13) Then maxi(13) = Len(rLe(k).Lanr)
  If Len(rLe(k).Ausn) > maxi(14) Then maxi(14) = Len(rLe(k).Ausn)
  If Len(rLe(k).beme) > maxi(15) Then maxi(15) = Len(rLe(k).beme)
  If Len(rLe(k).QS) > maxi(16) Then maxi(16) = Len(rLe(k).QS)
  If Len(rLe(k).QT) > maxi(17) Then maxi(17) = Len(rLe(k).QT)
  If Len(rLe(k).Sachkbez) > maxi(18) Then maxi(18) = Len(rLe(k).Sachkbez)
  If Len(rLe(k).Zone) > maxi(19) Then maxi(19) = Len(rLe(k).Zone)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLe), i)
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
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.Charge: '" & rLe(k).Charge & "' -> '" & Left$(rLe(k).Charge, maxL) & "'", True: rLe(k).Charge = Left$(rLe(k).Charge, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).Lanr & "' -> '" & Left$(rLe(k).Lanr, maxL) & "'", True: rLe(k).Lanr = Left$(rLe(k).Lanr, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL) & "'", True: rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).beme & "' -> '" & Left$(rLe(k).beme, maxL) & "'", True: rLe(k).beme = Left$(rLe(k).beme, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL) & "'", True: rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL) & "'", True: rLe(k).QT = Left$(rLe(k).QT, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & Left$(rLe(k).Sachkbez, maxL) & "'", True: rLe(k).Sachkbez = Left$(rLe(k).Sachkbez, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & Left$(rLe(k).Zone, maxL) & "'", True: rLe(k).Zone = Left$(rLe(k).Zone, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function roMeZuw(i&, j&)
 roMe(i).FID = rMe(j).FID
 roMe(i).Pat_id = rMe(j).Pat_id
 roMe(i).MPNr = rMe(j).MPNr
 roMe(i).Zeitpunkt = rMe(j).Zeitpunkt
 roMe(i).Datum = rMe(j).Datum
 roMe(i).Medikament = rMe(j).Medikament
 roMe(i).MedAnfang = rMe(j).MedAnfang
 roMe(i).Wirkstoff = rMe(j).Wirkstoff
 roMe(i).PZN = rMe(j).PZN
 roMe(i).FeldNr = rMe(j).FeldNr
 roMe(i).mo = rMe(j).mo
 roMe(i).mi = rMe(j).mi
 roMe(i).nm = rMe(j).nm
 roMe(i).ab = rMe(j).ab
 roMe(i).Zn = rMe(j).Zn
 roMe(i).bBed = rMe(j).bBed
 roMe(i).Bemerkung = rMe(j).Bemerkung
 roMe(i).Grund = rMe(j).Grund
 roMe(i).Stärke = rMe(j).Stärke
 roMe(i).Einheit = rMe(j).Einheit
 roMe(i).Form = rMe(j).Form
 roMe(i).absPos = rMe(j).absPos
 roMe(i).AktZeit = rMe(j).AktZeit
 roMe(i).StByte = rMe(j).StByte
 roMe(i).ergaenzt = rMe(j).ergaenzt
End Function ' roMeZuw

Public Function MeZUnt%(i&, j&)
 If roMe(i).FID <> rMe(j).FID Then GoSub unter
 If roMe(i).Pat_id <> rMe(j).Pat_id Then GoSub unter
 If roMe(i).MPNr <> rMe(j).MPNr Then GoSub unter
 If roMe(i).Zeitpunkt <> rMe(j).Zeitpunkt Then GoSub unter
 If roMe(i).Datum <> rMe(j).Datum Then GoSub unter
 If roMe(i).Medikament <> rMe(j).Medikament Then GoSub unter
 If roMe(i).MedAnfang <> rMe(j).MedAnfang Then GoSub unter
 If roMe(i).Wirkstoff <> rMe(j).Wirkstoff Then GoSub unter
 If roMe(i).PZN <> rMe(j).PZN Then GoSub unter
 If roMe(i).FeldNr <> rMe(j).FeldNr Then GoSub unter
 If roMe(i).mo <> rMe(j).mo Then GoSub unter
 If roMe(i).mi <> rMe(j).mi Then GoSub unter
 If roMe(i).nm <> rMe(j).nm Then GoSub unter
 If roMe(i).ab <> rMe(j).ab Then GoSub unter
 If roMe(i).Zn <> rMe(j).Zn Then GoSub unter
 If roMe(i).bBed <> rMe(j).bBed Then GoSub unter
 If roMe(i).Bemerkung <> rMe(j).Bemerkung Then GoSub unter
 If roMe(i).Grund <> rMe(j).Grund Then GoSub unter
 If roMe(i).Stärke <> rMe(j).Stärke Then GoSub unter
 If roMe(i).Einheit <> rMe(j).Einheit Then GoSub unter
 If roMe(i).Form <> rMe(j).Form Then GoSub unter
 If roMe(i).absPos <> rMe(j).absPos Then GoSub unter
 If roMe(i).AktZeit <> rMe(j).AktZeit Then GoSub unter
 If roMe(i).StByte <> rMe(j).StByte Then GoSub unter
 If roMe(i).ergaenzt <> rMe(j).ergaenzt Then GoSub unter
 Exit Function
unter:
 MeZUnt = MeZUnt + 1
 Return
End Function ' MeZUnt

Public Function medplanLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(MPNr,0) MPNr,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt" & _
",COALESCE(Datum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Datum,COALESCE(Medikament,'') Medikament,COALESCE(MedAnfang,'') MedAnfang,COALESCE(Wirkstoff,'') Wirkstoff" & _
",COALESCE(PZN,0) PZN,COALESCE(FeldNr,0) FeldNr,COALESCE(mo,'') mo,COALESCE(mi,'') mi" & _
",COALESCE(nm,'') nm,COALESCE(ab,'') ab,COALESCE(zn,'') zn,COALESCE(bBed,0) bBed" & _
",COALESCE(Bemerkung,'') Bemerkung,COALESCE(Grund,'') Grund,COALESCE(Stärke,'') Stärke,COALESCE(Einheit,'') Einheit" & _
",COALESCE(Form,'') Form,COALESCE(AbsPos,0) AbsPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
",COALESCE(ergaenzt,0) ergaenzt FROM `medplan` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roMe(0)
 Else ' rs.EOF Then
  ReDim roMe(1)
  Do While Not rs.EOF
   akt = UBound(roMe)
   roMe(akt).FID = rs!FID
   roMe(akt).Pat_id = rs!Pat_id
   roMe(akt).MPNr = rs!MPNr
   roMe(akt).Zeitpunkt = rs!Zeitpunkt
   roMe(akt).Datum = rs!Datum
   roMe(akt).Medikament = doUmwfSQL(rs!Medikament, lies.obMySQL, False)
   roMe(akt).MedAnfang = doUmwfSQL(rs!MedAnfang, lies.obMySQL, False)
   roMe(akt).Wirkstoff = doUmwfSQL(rs!Wirkstoff, lies.obMySQL, False)
   roMe(akt).PZN = rs!PZN
   roMe(akt).FeldNr = rs!FeldNr
   roMe(akt).mo = doUmwfSQL(rs!mo, lies.obMySQL, False)
   roMe(akt).mi = doUmwfSQL(rs!mi, lies.obMySQL, False)
   roMe(akt).nm = doUmwfSQL(rs!nm, lies.obMySQL, False)
   roMe(akt).ab = doUmwfSQL(rs!ab, lies.obMySQL, False)
   roMe(akt).Zn = doUmwfSQL(rs!Zn, lies.obMySQL, False)
   roMe(akt).bBed = rs!bBed
   roMe(akt).Bemerkung = doUmwfSQL(rs!Bemerkung, lies.obMySQL, False)
   roMe(akt).Grund = doUmwfSQL(rs!Grund, lies.obMySQL, False)
   roMe(akt).Stärke = doUmwfSQL(rs!Stärke, lies.obMySQL, False)
   roMe(akt).Einheit = doUmwfSQL(rs!Einheit, lies.obMySQL, False)
   roMe(akt).Form = doUmwfSQL(rs!Form, lies.obMySQL, False)
   roMe(akt).absPos = rs!absPos
   roMe(akt).AktZeit = rs!AktZeit
   roMe(akt).StByte = rs!StByte
   roMe(akt).ergaenzt = rs!ergaenzt
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roMe(UBound(roMe) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanLaden

Function medplanEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rMe) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rMe)
   If rMe(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roMe)
    If roMe(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roMe(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roMe(robeg + UBound(rMe) - rbeg)
   For ri = rbeg To UBound(rMe)
    Call roMeZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rMe = roMe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanEinf

Public Function medplanSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rMe) + 0 & " Sätze in `medplan`"
' sql0 = " INSERT " & sqlignore &  "INTO `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte,ergaenzt) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte,ergaenzt)       VALUES"))
 For i = 1 To UBound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rMe(i).FID, ",", rMe(i).Pat_id, ",", rMe(i).MPNr, ",", DatFor_k(rMe(i).Zeitpunkt), ",", DatFor_k(rMe(i).Datum), ",'", rMe(i).Medikament, "','", rMe(i).MedAnfang, "','", rMe(i).Wirkstoff, "',", _
   rMe(i).PZN, ",", rMe(i).FeldNr, ",'", rMe(i).mo, "','", rMe(i).mi, "','", rMe(i).nm, "','", rMe(i).ab, "','", rMe(i).Zn, "',", CStr(-(rMe(i).bBed <> 0)), ",'", rMe(i).Bemerkung, "','", _
   rMe(i).Grund, "','", rMe(i).Stärke, "','", rMe(i).Einheit, "','", rMe(i).Form, "',", rMe(i).absPos, ",", DatFor_k(rMe(i).AktZeit), ",", rMe(i).StByte, ",", CStr(-(rMe(i).ergaenzt <> 0)), ")")
  If SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rMe) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `medplan` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rMe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""medplanSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(12)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rMe), i)
  If Len(rMe(k).Medikament) > maxi(0) Then maxi(0) = Len(rMe(k).Medikament)
  If Len(rMe(k).MedAnfang) > maxi(1) Then maxi(1) = Len(rMe(k).MedAnfang)
  If Len(rMe(k).Wirkstoff) > maxi(2) Then maxi(2) = Len(rMe(k).Wirkstoff)
  If Len(rMe(k).mo) > maxi(3) Then maxi(3) = Len(rMe(k).mo)
  If Len(rMe(k).mi) > maxi(4) Then maxi(4) = Len(rMe(k).mi)
  If Len(rMe(k).nm) > maxi(5) Then maxi(5) = Len(rMe(k).nm)
  If Len(rMe(k).ab) > maxi(6) Then maxi(6) = Len(rMe(k).ab)
  If Len(rMe(k).Zn) > maxi(7) Then maxi(7) = Len(rMe(k).Zn)
  If Len(rMe(k).Bemerkung) > maxi(8) Then maxi(8) = Len(rMe(k).Bemerkung)
  If Len(rMe(k).Grund) > maxi(9) Then maxi(9) = Len(rMe(k).Grund)
  If Len(rMe(k).Stärke) > maxi(10) Then maxi(10) = Len(rMe(k).Stärke)
  If Len(rMe(k).Einheit) > maxi(11) Then maxi(11) = Len(rMe(k).Einheit)
  If Len(rMe(k).Form) > maxi(12) Then maxi(12) = Len(rMe(k).Form)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rMe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & Left$(rMe(k).Medikament, maxL) & "'", True: rMe(k).Medikament = Left$(rMe(k).Medikament, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & Left$(rMe(k).MedAnfang, maxL) & "'", True: rMe(k).MedAnfang = Left$(rMe(k).MedAnfang, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.Wirkstoff: '" & rMe(k).Wirkstoff & "' -> '" & Left$(rMe(k).Wirkstoff, maxL) & "'", True: rMe(k).Wirkstoff = Left$(rMe(k).Wirkstoff, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & Left$(rMe(k).mo, maxL) & "'", True: rMe(k).mo = Left$(rMe(k).mo, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & Left$(rMe(k).mi, maxL) & "'", True: rMe(k).mi = Left$(rMe(k).mi, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & Left$(rMe(k).nm, maxL) & "'", True: rMe(k).nm = Left$(rMe(k).nm, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & Left$(rMe(k).ab, maxL) & "'", True: rMe(k).ab = Left$(rMe(k).ab, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).Zn & "' -> '" & Left$(rMe(k).Zn, maxL) & "'", True: rMe(k).Zn = Left$(rMe(k).Zn, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & Left$(rMe(k).Bemerkung, maxL) & "'", True: rMe(k).Bemerkung = Left$(rMe(k).Bemerkung, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rMe.Grund: '" & rMe(k).Grund & "' -> '" & Left$(rMe(k).Grund, maxL) & "'", True: rMe(k).Grund = Left$(rMe(k).Grund, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rMe.Stärke: '" & rMe(k).Stärke & "' -> '" & Left$(rMe(k).Stärke, maxL) & "'", True: rMe(k).Stärke = Left$(rMe(k).Stärke, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rMe.Einheit: '" & rMe(k).Einheit & "' -> '" & Left$(rMe(k).Einheit, maxL) & "'", True: rMe(k).Einheit = Left$(rMe(k).Einheit, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rMe.Form: '" & rMe(k).Form & "' -> '" & Left$(rMe(k).Form, maxL) & "'", True: rMe(k).Form = Left$(rMe(k).Form, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function roReZuw(i&, j&)
 roRe(i).FID = rRe(j).FID
 roRe(i).Pat_id = rRe(j).Pat_id
 roRe(i).Zeitpunkt = rRe(j).Zeitpunkt
 roRe(i).Rezept = rRe(j).Rezept
 roRe(i).RKlnm = rRe(j).RKlnm
 roRe(i).Rezeptklasse = rRe(j).Rezeptklasse
 roRe(i).Rezklkurz = rRe(j).Rezklkurz
 roRe(i).Rezkllang = rRe(j).Rezkllang
 roRe(i).kbez = rRe(j).kbez
 roRe(i).Medikament = rRe(j).Medikament
 roRe(i).auti = rRe(j).auti
 roRe(i).anzl = rRe(j).anzl
 roRe(i).PZN = rRe(j).PZN
 roRe(i).absPos = rRe(j).absPos
 roRe(i).AktZeit = rRe(j).AktZeit
 roRe(i).QS = rRe(j).QS
 roRe(i).QT = rRe(j).QT
 roRe(i).StByte = rRe(j).StByte
 roRe(i).lanrid = rRe(j).lanrid
 roRe(i).id = rRe(j).id
End Function ' roReZuw

Public Function ReZUnt%(i&, j&)
 If roRe(i).FID <> rRe(j).FID Then GoSub unter
 If roRe(i).Pat_id <> rRe(j).Pat_id Then GoSub unter
 If roRe(i).Zeitpunkt <> rRe(j).Zeitpunkt Then GoSub unter
 If roRe(i).Rezept <> rRe(j).Rezept Then GoSub unter
 If roRe(i).RKlnm <> rRe(j).RKlnm Then GoSub unter
 If roRe(i).Rezeptklasse <> rRe(j).Rezeptklasse Then GoSub unter
 If roRe(i).Rezklkurz <> rRe(j).Rezklkurz Then GoSub unter
 If roRe(i).Rezkllang <> rRe(j).Rezkllang Then GoSub unter
 If roRe(i).kbez <> rRe(j).kbez Then GoSub unter
 If roRe(i).Medikament <> rRe(j).Medikament Then GoSub unter
 If roRe(i).auti <> rRe(j).auti Then GoSub unter
 If roRe(i).anzl <> rRe(j).anzl Then GoSub unter
 If roRe(i).PZN <> rRe(j).PZN Then GoSub unter
 If roRe(i).absPos <> rRe(j).absPos Then GoSub unter
 If roRe(i).AktZeit <> rRe(j).AktZeit Then GoSub unter
 If roRe(i).QS <> rRe(j).QS Then GoSub unter
 If roRe(i).QT <> rRe(j).QT Then GoSub unter
 If roRe(i).StByte <> rRe(j).StByte Then GoSub unter
 If roRe(i).lanrid <> rRe(j).lanrid Then GoSub unter
 If roRe(i).id <> rRe(j).id Then GoSub unter
 Exit Function
unter:
 ReZUnt = ReZUnt + 1
 Return
End Function ' ReZUnt

Public Function rezepteintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Rezept,'') Rezept" & _
",COALESCE(RKlnm,'') RKlnm,COALESCE(Rezeptklasse,'') Rezeptklasse,COALESCE(Rezklkurz,'') Rezklkurz,COALESCE(Rezkllang,'') Rezkllang" & _
",COALESCE(kbez,'') kbez,COALESCE(Medikament,'') Medikament,COALESCE(auti,0) auti,COALESCE(anzl,0) anzl" & _
",COALESCE(PZN,'') PZN,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid,COALESCE(id,0) id" & _
" FROM `rezepteintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roRe(0)
 Else ' rs.EOF Then
  ReDim roRe(1)
  Do While Not rs.EOF
   akt = UBound(roRe)
   roRe(akt).FID = rs!FID
   roRe(akt).Pat_id = rs!Pat_id
   roRe(akt).Zeitpunkt = rs!Zeitpunkt
   roRe(akt).Rezept = doUmwfSQL(rs!Rezept, lies.obMySQL, False)
   roRe(akt).RKlnm = doUmwfSQL(rs!RKlnm, lies.obMySQL, False)
   roRe(akt).Rezeptklasse = doUmwfSQL(rs!Rezeptklasse, lies.obMySQL, False)
   roRe(akt).Rezklkurz = doUmwfSQL(rs!Rezklkurz, lies.obMySQL, False)
   roRe(akt).Rezkllang = doUmwfSQL(rs!Rezkllang, lies.obMySQL, False)
   roRe(akt).kbez = doUmwfSQL(rs!kbez, lies.obMySQL, False)
   roRe(akt).Medikament = doUmwfSQL(rs!Medikament, lies.obMySQL, False)
   roRe(akt).auti = rs!auti
   roRe(akt).anzl = rs!anzl
   roRe(akt).PZN = doUmwfSQL(rs!PZN, lies.obMySQL, False)
   roRe(akt).absPos = rs!absPos
   roRe(akt).AktZeit = rs!AktZeit
   roRe(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roRe(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roRe(akt).StByte = rs!StByte
   roRe(akt).lanrid = rs!lanrid
   roRe(akt).id = rs!id
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roRe(UBound(roRe) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeLaden

Function rezepteintraegeEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rRe) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rRe)
   If rRe(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roRe)
    If roRe(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roRe(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roRe(robeg + UBound(rRe) - rbeg)
   For ri = rbeg To UBound(rRe)
    Call roReZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rRe = roRe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeEinf

Public Function rezepteintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rRe) + 0 & " Sätze in `rezepteintraege`"
' sql0 = " INSERT " & sqlignore &  "INTO `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid)  VALUES"))
 For i = 1 To UBound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRe(i).FID, ",", rRe(i).Pat_id, ",", DatFor_k(rRe(i).Zeitpunkt), ",'", rRe(i).Rezept, "','", rRe(i).RKlnm, "','", rRe(i).Rezeptklasse, "','", rRe(i).Rezklkurz, "','", rRe(i).Rezkllang, "','", _
   rRe(i).kbez, "','", rRe(i).Medikament, "',", rRe(i).auti, ",", rRe(i).anzl, ",'", rRe(i).PZN, "',", rRe(i).absPos, ",", DatFor_k(rRe(i).AktZeit), ",'", rRe(i).QS, "','", _
   rRe(i).QT, "',", rRe(i).StByte, ",", rRe(i).lanrid, ")")
  If SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rRe) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `rezepteintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rRe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""rezepteintraegeSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(9)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRe), i)
  If Len(rRe(k).Rezept) > maxi(0) Then maxi(0) = Len(rRe(k).Rezept)
  If Len(rRe(k).RKlnm) > maxi(1) Then maxi(1) = Len(rRe(k).RKlnm)
  If Len(rRe(k).Rezeptklasse) > maxi(2) Then maxi(2) = Len(rRe(k).Rezeptklasse)
  If Len(rRe(k).Rezklkurz) > maxi(3) Then maxi(3) = Len(rRe(k).Rezklkurz)
  If Len(rRe(k).Rezkllang) > maxi(4) Then maxi(4) = Len(rRe(k).Rezkllang)
  If Len(rRe(k).kbez) > maxi(5) Then maxi(5) = Len(rRe(k).kbez)
  If Len(rRe(k).Medikament) > maxi(6) Then maxi(6) = Len(rRe(k).Medikament)
  If Len(rRe(k).PZN) > maxi(7) Then maxi(7) = Len(rRe(k).PZN)
  If Len(rRe(k).QS) > maxi(8) Then maxi(8) = Len(rRe(k).QS)
  If Len(rRe(k).QT) > maxi(9) Then maxi(9) = Len(rRe(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & Left$(rRe(k).Rezept, maxL) & "'", True: rRe(k).Rezept = Left$(rRe(k).Rezept, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.RKlnm: '" & rRe(k).RKlnm & "' -> '" & Left$(rRe(k).RKlnm, maxL) & "'", True: rRe(k).RKlnm = Left$(rRe(k).RKlnm, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & Left$(rRe(k).Rezeptklasse, maxL) & "'", True: rRe(k).Rezeptklasse = Left$(rRe(k).Rezeptklasse, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezklkurz: '" & rRe(k).Rezklkurz & "' -> '" & Left$(rRe(k).Rezklkurz, maxL) & "'", True: rRe(k).Rezklkurz = Left$(rRe(k).Rezklkurz, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezkllang: '" & rRe(k).Rezkllang & "' -> '" & Left$(rRe(k).Rezkllang, maxL) & "'", True: rRe(k).Rezkllang = Left$(rRe(k).Rezkllang, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.kbez: '" & rRe(k).kbez & "' -> '" & Left$(rRe(k).kbez, maxL) & "'", True: rRe(k).kbez = Left$(rRe(k).kbez, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & Left$(rRe(k).Medikament, maxL) & "'", True: rRe(k).Medikament = Left$(rRe(k).Medikament, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & Left$(rRe(k).PZN, maxL) & "'", True: rRe(k).PZN = Left$(rRe(k).PZN, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & Left$(rRe(k).QS, maxL) & "'", True: rRe(k).QS = Left$(rRe(k).QS, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & Left$(rRe(k).QT, maxL) & "'", True: rRe(k).QT = Left$(rRe(k).QT, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function roRrZuw(i&, j&)
 roRr(i).FID = rRr(j).FID
 roRr(i).Pat_id = rRr(j).Pat_id
 roRr(i).Zeitpunkt = rRr(j).Zeitpunkt
 roRr(i).RR = rRr(j).RR
 roRr(i).Puls = rRr(j).Puls
 roRr(i).RRsyst = rRr(j).RRsyst
 roRr(i).RRdiast = rRr(j).RRdiast
 roRr(i).RRzahl = rRr(j).RRzahl
 roRr(i).Quelle = rRr(j).Quelle
 roRr(i).Bemerkung = rRr(j).Bemerkung
 roRr(i).absPos = rRr(j).absPos
 roRr(i).AktZeit = rRr(j).AktZeit
 roRr(i).StByte = rRr(j).StByte
End Function ' roRrZuw

Public Function RrZUnt%(i&, j&)
 If roRr(i).FID <> rRr(j).FID Then GoSub unter
 If roRr(i).Pat_id <> rRr(j).Pat_id Then GoSub unter
 If roRr(i).Zeitpunkt <> rRr(j).Zeitpunkt Then GoSub unter
 If roRr(i).RR <> rRr(j).RR Then GoSub unter
 If roRr(i).Puls <> rRr(j).Puls Then GoSub unter
 If roRr(i).RRsyst <> rRr(j).RRsyst Then GoSub unter
 If roRr(i).RRdiast <> rRr(j).RRdiast Then GoSub unter
 If roRr(i).RRzahl <> rRr(j).RRzahl Then GoSub unter
 If roRr(i).Quelle <> rRr(j).Quelle Then GoSub unter
 If roRr(i).Bemerkung <> rRr(j).Bemerkung Then GoSub unter
 If roRr(i).absPos <> rRr(j).absPos Then GoSub unter
 If roRr(i).AktZeit <> rRr(j).AktZeit Then GoSub unter
 If roRr(i).StByte <> rRr(j).StByte Then GoSub unter
 Exit Function
unter:
 RrZUnt = RrZUnt + 1
 Return
End Function ' RrZUnt

Public Function rrLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(RR,'') RR" & _
",COALESCE(Puls,0) Puls,COALESCE(RRsyst,0) RRsyst,COALESCE(RRdiast,0) RRdiast,COALESCE(RRzahl,0) RRzahl" & _
",COALESCE(Quelle,'') Quelle,COALESCE(Bemerkung,'') Bemerkung,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `rr` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roRr(0)
 Else ' rs.EOF Then
  ReDim roRr(1)
  Do While Not rs.EOF
   akt = UBound(roRr)
   roRr(akt).FID = rs!FID
   roRr(akt).Pat_id = rs!Pat_id
   roRr(akt).Zeitpunkt = rs!Zeitpunkt
   roRr(akt).RR = doUmwfSQL(rs!RR, lies.obMySQL, False)
   roRr(akt).Puls = rs!Puls
   roRr(akt).RRsyst = rs!RRsyst
   roRr(akt).RRdiast = rs!RRdiast
   roRr(akt).RRzahl = rs!RRzahl
   roRr(akt).Quelle = doUmwfSQL(rs!Quelle, lies.obMySQL, False)
   roRr(akt).Bemerkung = doUmwfSQL(rs!Bemerkung, lies.obMySQL, False)
   roRr(akt).absPos = rs!absPos
   roRr(akt).AktZeit = rs!AktZeit
   roRr(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roRr(UBound(roRr) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrLaden

Function rrEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rRr) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rRr)
   If rRr(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roRr)
    If roRr(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roRr(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roRr(robeg + UBound(rRr) - rbeg)
   For ri = rbeg To UBound(rRr)
    Call roRrZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rRr = roRr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrEinf

Public Function rrSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rRr) + 0 & " Sätze in `rr`"
' sql0 = " INSERT " & sqlignore &  "INTO `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte)          VALUES"))
 For i = 1 To UBound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rRr(i).FID, ",", rRr(i).Pat_id, ",", DatFor_k(rRr(i).Zeitpunkt), ",'", rRr(i).RR, "',", rRr(i).Puls, ",", rRr(i).RRsyst, ",", rRr(i).RRdiast, ",", rRr(i).RRzahl, ",'", rRr(i).Quelle, "','", _
   rRr(i).Bemerkung, "',", rRr(i).absPos, ",", DatFor_k(rRr(i).AktZeit), ",", rRr(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rRr) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `rr` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rRr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""rrSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(2)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRr), i)
  If Len(rRr(k).RR) > maxi(0) Then maxi(0) = Len(rRr(k).RR)
  If Len(rRr(k).Quelle) > maxi(1) Then maxi(1) = Len(rRr(k).Quelle)
  If Len(rRr(k).Bemerkung) > maxi(2) Then maxi(2) = Len(rRr(k).Bemerkung)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRr), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function kvnrueSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rKv) + 0 & " Sätze in `kvnrue`"
' sql0 = " INSERT " & sqlignore &  "INTO `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        VALUES"))
 For i = 1 To UBound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rKv(i).Pat_id, ",'", rKv(i).KVNr, "',", rKv(i).absPos, ",", DatFor_k(rKv(i).AktZeit), ",", rKv(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rKv) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `kvnrue` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rKv)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""kvnrueSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKv), i)
  If Len(rKv(k).KVNr) > maxi(0) Then maxi(0) = Len(rKv(k).KVNr)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKv), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function unbekannte_kennungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 syscmd 4, pid & ": Speichere " & UBound(rUn) + 0 & " Sätze in `unbekannte_kennungen`"
' sql0 = " INSERT " & sqlignore &  "INTO `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt)         VALUES"))
 For i = rUn1 + 1 To UBound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rUn(i).Kennung, "',", rUn(i).absPos, ",", rUn(i).StByte, ",", rUn(i).Pat_id, ",'", rUn(i).Inhalt, "')")
  If SammelInsert <> 0 And i < UBound(rUn) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUn) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rUn)
  DoEvents
 Next i
 rUn1 = UBound(rUn)
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""unbekannte_kennungenSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, rUn1 + 1, i) To IIf(SammelInsert <> 0, UBound(rUn), i)
  If Len(rUn(k).Kennung) > maxi(0) Then maxi(0) = Len(rUn(k).Kennung)
  If Len(rUn(k).Inhalt) > maxi(1) Then maxi(1) = Len(rUn(k).Inhalt)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, rUn1 + 1, i) To IIf(SammelInsert <> 0, UBound(rUn), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbekannte_kennungenSpeichern

Public Function roDmZuw(i&, j&)
 roDm(i).Abk = rDm(j).Abk
 roDm(i).art = rDm(j).art
 roDm(i).KarteiDatum = rDm(j).KarteiDatum
 roDm(i).exportiert = rDm(j).exportiert
 roDm(i).DokuDatum = rDm(j).DokuDatum
 roDm(i).obvoll = rDm(j).obvoll
 roDm(i).Ok = rDm(j).Ok
 roDm(i).ausgedruckt = rDm(j).ausgedruckt
 roDm(i).Nachname = rDm(j).Nachname
 roDm(i).Vorname = rDm(j).Vorname
 roDm(i).GebDat = rDm(j).GebDat
 roDm(i).Pat_id = rDm(j).Pat_id
 roDm(i).StByte = rDm(j).StByte
 roDm(i).AktZeit = rDm(j).AktZeit
 roDm(i).lanrid = rDm(j).lanrid
 roDm(i).Zusatzdaten = rDm(j).Zusatzdaten
End Function ' roDmZuw

Public Function DmZUnt%(i&, j&)
 If roDm(i).Abk <> rDm(j).Abk Then GoSub unter
 If roDm(i).art <> rDm(j).art Then GoSub unter
 If roDm(i).KarteiDatum <> rDm(j).KarteiDatum Then GoSub unter
 If roDm(i).exportiert <> rDm(j).exportiert Then GoSub unter
 If roDm(i).DokuDatum <> rDm(j).DokuDatum Then GoSub unter
 If roDm(i).obvoll <> rDm(j).obvoll Then GoSub unter
 If roDm(i).Ok <> rDm(j).Ok Then GoSub unter
 If roDm(i).ausgedruckt <> rDm(j).ausgedruckt Then GoSub unter
 If roDm(i).Nachname <> rDm(j).Nachname Then GoSub unter
 If roDm(i).Vorname <> rDm(j).Vorname Then GoSub unter
 If roDm(i).GebDat <> rDm(j).GebDat Then GoSub unter
 If roDm(i).Pat_id <> rDm(j).Pat_id Then GoSub unter
 If roDm(i).StByte <> rDm(j).StByte Then GoSub unter
 If roDm(i).AktZeit <> rDm(j).AktZeit Then GoSub unter
 If roDm(i).lanrid <> rDm(j).lanrid Then GoSub unter
 If roDm(i).Zusatzdaten <> rDm(j).Zusatzdaten Then GoSub unter
 Exit Function
unter:
 DmZUnt = DmZUnt + 1
 Return
End Function ' DmZUnt

Public Function dmpreiheLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(Abk,'') Abk,COALESCE(Art,'') Art,COALESCE(KarteiDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) KarteiDatum,COALESCE(exportiert - INTERVAL 0 DAY,CONVERT('18991230',DATE)) exportiert" & _
",COALESCE(DokuDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DokuDatum,COALESCE(obvoll,0) obvoll,COALESCE(ok,0) ok,COALESCE(ausgedruckt,0) ausgedruckt" & _
",COALESCE(NachName,'') NachName,COALESCE(VorName,'') VorName,COALESCE(GebDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) GebDat,COALESCE(Pat_id,0) Pat_id" & _
",COALESCE(StByte,0) StByte,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(lanrid,0) lanrid,COALESCE(Zusatzdaten,'') Zusatzdaten" & _
" FROM `dmpreihe` WHERE Pat_ID=" & pid & " ORDER BY `Dokudatum`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDm(0)
 Else ' rs.EOF Then
  ReDim roDm(1)
  Do While Not rs.EOF
   akt = UBound(roDm)
   roDm(akt).Abk = doUmwfSQL(rs!Abk, lies.obMySQL, False)
   roDm(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roDm(akt).KarteiDatum = rs!KarteiDatum
   roDm(akt).exportiert = rs!exportiert
   roDm(akt).DokuDatum = rs!DokuDatum
   roDm(akt).obvoll = rs!obvoll
   roDm(akt).Ok = rs!Ok
   roDm(akt).ausgedruckt = rs!ausgedruckt
   roDm(akt).Nachname = doUmwfSQL(rs!Nachname, lies.obMySQL, False)
   roDm(akt).Vorname = doUmwfSQL(rs!Vorname, lies.obMySQL, False)
   roDm(akt).GebDat = rs!GebDat
   roDm(akt).Pat_id = rs!Pat_id
   roDm(akt).StByte = rs!StByte
   roDm(akt).AktZeit = rs!AktZeit
   roDm(akt).lanrid = rs!lanrid
   roDm(akt).Zusatzdaten = doUmwfSQL(rs!Zusatzdaten, lies.obMySQL, False)
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roDm(UBound(roDm) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheLaden

Function dmpreiheEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDm) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rDm)
   If rDm(ri).DokuDatum >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roDm)
    If roDm(robeg).DokuDatum >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roDm(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDm(robeg + UBound(rDm) - rbeg)
   For ri = rbeg To UBound(rDm)
    Call roDmZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rDm = roDm
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheEinf

Public Function dmpreiheSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDm) + 0 & " Sätze in `dmpreihe`"
' sql0 = " INSERT " & sqlignore &  "INTO `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten)          VALUES"))
 For i = 1 To UBound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rDm(i).Abk, "','", rDm(i).art, "',", DatFor_k(rDm(i).KarteiDatum), ",", DatFor_k(rDm(i).exportiert), ",", DatFor_k(rDm(i).DokuDatum), ",", CStr(-(rDm(i).obvoll <> 0)), ",", CStr(-( _
   rDm(i).Ok <> 0)), ",", CStr(-(rDm(i).ausgedruckt <> 0)), ",'", rDm(i).Nachname, "','", rDm(i).Vorname, "',", DatFor_k(rDm(i).GebDat), ",", rDm(i).Pat_id, ",", rDm(i).StByte, ",", DatFor_k(rDm(i).AktZeit), ",", _
   rDm(i).lanrid, ",'", rDm(i).Zusatzdaten, "')")
  If SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDm) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `dmpreihe` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rDm)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""dmpreiheSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDm), i)
  If Len(rDm(k).Abk) > maxi(0) Then maxi(0) = Len(rDm(k).Abk)
  If Len(rDm(k).art) > maxi(1) Then maxi(1) = Len(rDm(k).art)
  If Len(rDm(k).Nachname) > maxi(2) Then maxi(2) = Len(rDm(k).Nachname)
  If Len(rDm(k).Vorname) > maxi(3) Then maxi(3) = Len(rDm(k).Vorname)
  If Len(rDm(k).Zusatzdaten) > maxi(4) Then maxi(4) = Len(rDm(k).Zusatzdaten)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDm), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL) & "'", True: rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).art & "' -> '" & Left$(rDm(k).art, maxL) & "'", True: rDm(k).art = Left$(rDm(k).art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).Nachname & "' -> '" & Left$(rDm(k).Nachname, maxL) & "'", True: rDm(k).Nachname = Left$(rDm(k).Nachname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).Vorname & "' -> '" & Left$(rDm(k).Vorname, maxL) & "'", True: rDm(k).Vorname = Left$(rDm(k).Vorname, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDm.Zusatzdaten: '" & rDm(k).Zusatzdaten & "' -> '" & Left$(rDm(k).Zusatzdaten, maxL) & "'", True: rDm(k).Zusatzdaten = Left$(rDm(k).Zusatzdaten, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dmpreiheSpeichern

Public Function roDeZuw(i&, j&)
 roDe(i).id = rDe(j).id
 roDe(i).IDS = rDe(j).IDS
 roDe(i).Pat_id = rDe(j).Pat_id
 roDe(i).erstZP = rDe(j).erstZP
 roDe(i).exoL = rDe(j).exoL
 roDe(i).hideT = rDe(j).hideT
 roDe(i).iconPath = rDe(j).iconPath
 roDe(i).noteBkColor = rDe(j).noteBkColor
 roDe(i).noteFgColor = rDe(j).noteFgColor
 roDe(i).positionBottom = rDe(j).positionBottom
 roDe(i).positionLeft = rDe(j).positionLeft
 roDe(i).positionRight = rDe(j).positionRight
 roDe(i).positionTop = rDe(j).positionTop
 roDe(i).showAsNote = rDe(j).showAsNote
 roDe(i).syncInfoList = rDe(j).syncInfoList
 roDe(i).Titel = rDe(j).Titel
 roDe(i).toolTipText = rDe(j).toolTipText
 roDe(i).verankert = rDe(j).verankert
 roDe(i).absPos = rDe(j).absPos
 roDe(i).AktZeit = rDe(j).AktZeit
 roDe(i).StByte = rDe(j).StByte
End Function ' roDeZuw

Public Function DeZUnt%(i&, j&)
 If roDe(i).id <> rDe(j).id Then GoSub unter
 If roDe(i).IDS <> rDe(j).IDS Then GoSub unter
 If roDe(i).Pat_id <> rDe(j).Pat_id Then GoSub unter
 If roDe(i).erstZP <> rDe(j).erstZP Then GoSub unter
 If roDe(i).exoL <> rDe(j).exoL Then GoSub unter
 If roDe(i).hideT <> rDe(j).hideT Then GoSub unter
 If roDe(i).iconPath <> rDe(j).iconPath Then GoSub unter
 If roDe(i).noteBkColor <> rDe(j).noteBkColor Then GoSub unter
 If roDe(i).noteFgColor <> rDe(j).noteFgColor Then GoSub unter
 If roDe(i).positionBottom <> rDe(j).positionBottom Then GoSub unter
 If roDe(i).positionLeft <> rDe(j).positionLeft Then GoSub unter
 If roDe(i).positionRight <> rDe(j).positionRight Then GoSub unter
 If roDe(i).positionTop <> rDe(j).positionTop Then GoSub unter
 If roDe(i).showAsNote <> rDe(j).showAsNote Then GoSub unter
 If roDe(i).syncInfoList <> rDe(j).syncInfoList Then GoSub unter
 If roDe(i).Titel <> rDe(j).Titel Then GoSub unter
 If roDe(i).toolTipText <> rDe(j).toolTipText Then GoSub unter
 If roDe(i).verankert <> rDe(j).verankert Then GoSub unter
 If roDe(i).absPos <> rDe(j).absPos Then GoSub unter
 If roDe(i).AktZeit <> rDe(j).AktZeit Then GoSub unter
 If roDe(i).StByte <> rDe(j).StByte Then GoSub unter
 Exit Function
unter:
 DeZUnt = DeZUnt + 1
 Return
End Function ' DeZUnt

Public Function desktopLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(id,0) id,COALESCE(IDS,'') IDS,COALESCE(Pat_ID,0) Pat_ID,COALESCE(erstZP - INTERVAL 0 DAY,CONVERT('18991230',DATE)) erstZP" & _
",COALESCE(exoL,'') exoL,COALESCE(hideT,0) hideT,COALESCE(iconPath,'') iconPath,COALESCE(noteBkColor,0) noteBkColor" & _
",COALESCE(noteFgColor,0) noteFgColor,COALESCE(positionBottom,0) positionBottom,COALESCE(positionLeft,0) positionLeft,COALESCE(positionRight,0) positionRight" & _
",COALESCE(positionTop,0) positionTop,COALESCE(showAsNote,0) showAsNote,COALESCE(syncInfoList,'') syncInfoList,COALESCE(titel,'') titel" & _
",COALESCE(toolTipText,'') toolTipText,COALESCE(verankert,0) verankert,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `desktop` WHERE Pat_ID=" & pid & " ORDER BY `erstZP`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDe(0)
 Else ' rs.EOF Then
  ReDim roDe(1)
  Do While Not rs.EOF
   akt = UBound(roDe)
   roDe(akt).id = rs!id
   roDe(akt).IDS = doUmwfSQL(rs!IDS, lies.obMySQL, False)
   roDe(akt).Pat_id = rs!Pat_id
   roDe(akt).erstZP = rs!erstZP
   roDe(akt).exoL = doUmwfSQL(rs!exoL, lies.obMySQL, False)
   roDe(akt).hideT = rs!hideT
   roDe(akt).iconPath = doUmwfSQL(rs!iconPath, lies.obMySQL, False)
   roDe(akt).noteBkColor = rs!noteBkColor
   roDe(akt).noteFgColor = rs!noteFgColor
   roDe(akt).positionBottom = rs!positionBottom
   roDe(akt).positionLeft = rs!positionLeft
   roDe(akt).positionRight = rs!positionRight
   roDe(akt).positionTop = rs!positionTop
   roDe(akt).showAsNote = rs!showAsNote
   roDe(akt).syncInfoList = doUmwfSQL(rs!syncInfoList, lies.obMySQL, False)
   roDe(akt).Titel = doUmwfSQL(rs!Titel, lies.obMySQL, False)
   roDe(akt).toolTipText = doUmwfSQL(rs!toolTipText, lies.obMySQL, False)
   roDe(akt).verankert = rs!verankert
   roDe(akt).absPos = rs!absPos
   roDe(akt).AktZeit = rs!AktZeit
   roDe(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roDe(UBound(roDe) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' desktopLaden

Function desktopEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDe) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rDe)
   If rDe(ri).erstZP >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roDe)
    If roDe(robeg).erstZP >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roDe(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDe(robeg + UBound(rDe) - rbeg)
   For ri = rbeg To UBound(rDe)
    Call roDeZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rDe = roDe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' desktopEinf

Public Function desktopSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDe) + 0 & " Sätze in `desktop`"
' sql0 = " INSERT " & sqlignore &  "INTO `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte)  VALUES"))
 For i = 1 To UBound(rDe)
'  rDe(i).AktZeit = now()
  rDe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rDe(i).IDS, "',", rDe(i).Pat_id, ",", DatFor_k(rDe(i).erstZP), ",'", rDe(i).exoL, "',", rDe(i).hideT, ",'", rDe(i).iconPath, "',", rDe(i).noteBkColor, ",", rDe(i).noteFgColor, ",", _
   rDe(i).positionBottom, ",", rDe(i).positionLeft, ",", rDe(i).positionRight, ",", rDe(i).positionTop, ",", rDe(i).showAsNote, ",'", rDe(i).syncInfoList, "','", rDe(i).Titel, "','", _
   rDe(i).toolTipText, "',", rDe(i).verankert, ",", rDe(i).absPos, ",", DatFor_k(rDe(i).AktZeit), ",", rDe(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rDe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDe) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `desktop` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rDe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""desktopSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDe), i)
  If Len(rDe(k).IDS) > maxi(0) Then maxi(0) = Len(rDe(k).IDS)
  If Len(rDe(k).exoL) > maxi(1) Then maxi(1) = Len(rDe(k).exoL)
  If Len(rDe(k).iconPath) > maxi(2) Then maxi(2) = Len(rDe(k).iconPath)
  If Len(rDe(k).syncInfoList) > maxi(3) Then maxi(3) = Len(rDe(k).syncInfoList)
  If Len(rDe(k).Titel) > maxi(4) Then maxi(4) = Len(rDe(k).Titel)
  If Len(rDe(k).toolTipText) > maxi(5) Then maxi(5) = Len(rDe(k).toolTipText)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "desktop", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "desktop", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDe.IDS: '" & rDe(k).IDS & "' -> '" & Left$(rDe(k).IDS, maxL) & "'", True: rDe(k).IDS = Left$(rDe(k).IDS, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' desktopSpeichern

Public Function roUsZuw(i&, j&)
 roUs(i).FID = rUs(j).FID
 roUs(i).Pat_id = rUs(j).Pat_id
 roUs(i).Zeitpunkt = rUs(j).Zeitpunkt
 roUs(i).art = rUs(j).art
 roUs(i).Spritzst = rUs(j).Spritzst
 roUs(i).Fußbef_re = rUs(j).Fußbef_re
 roUs(i).Fußbef_li = rUs(j).Fußbef_li
 roUs(i).Hyperk_re = rUs(j).Hyperk_re
 roUs(i).Hyperk_li = rUs(j).Hyperk_li
 roUs(i).Ulcera_re = rUs(j).Ulcera_re
 roUs(i).Ulcera_li = rUs(j).Ulcera_li
 roUs(i).Kraft_Zh_re = rUs(j).Kraft_Zh_re
 roUs(i).Kraft_Zh_li = rUs(j).Kraft_Zh_li
 roUs(i).Kraft_Zb_re = rUs(j).Kraft_Zb_re
 roUs(i).Kraft_Zb_li = rUs(j).Kraft_Zb_li
 roUs(i).Kraft_Knie_re = rUs(j).Kraft_Knie_re
 roUs(i).Kraft_Knie_li = rUs(j).Kraft_Knie_li
 roUs(i).ASR_re = rUs(j).ASR_re
 roUs(i).ASR_li = rUs(j).ASR_li
 roUs(i).PSR_re = rUs(j).PSR_re
 roUs(i).PSR_li = rUs(j).PSR_li
 roUs(i).Oberfl_re = rUs(j).Oberfl_re
 roUs(i).Oberfl_li = rUs(j).Oberfl_li
 roUs(i).MF_re = rUs(j).MF_re
 roUs(i).MF_li = rUs(j).MF_li
 roUs(i).KW_re = rUs(j).KW_re
 roUs(i).KW_li = rUs(j).KW_li
 roUs(i).Vibr_IK_re = rUs(j).Vibr_IK_re
 roUs(i).Vibr_IK_li = rUs(j).Vibr_IK_li
 roUs(i).Vibr_GZ_re = rUs(j).Vibr_GZ_re
 roUs(i).Vibr_GZ_li = rUs(j).Vibr_GZ_li
 roUs(i).PulsL_re = rUs(j).PulsL_re
 roUs(i).PulsL_li = rUs(j).PulsL_li
 roUs(i).PulsKK_re = rUs(j).PulsKK_re
 roUs(i).PulsKK_li = rUs(j).PulsKK_li
 roUs(i).PulsAtp_re = rUs(j).PulsAtp_re
 roUs(i).PulsAtp_li = rUs(j).PulsAtp_li
 roUs(i).PulsAdp_re = rUs(j).PulsAdp_re
 roUs(i).PulsAdp_li = rUs(j).PulsAdp_li
 roUs(i).Mitarbeiter = rUs(j).Mitarbeiter
 roUs(i).absPos = rUs(j).absPos
 roUs(i).AktZeit = rUs(j).AktZeit
 roUs(i).QS = rUs(j).QS
 roUs(i).QT = rUs(j).QT
 roUs(i).StByte = rUs(j).StByte
 roUs(i).id = rUs(j).id
End Function ' roUsZuw

Public Function UsZUnt%(i&, j&)
 If roUs(i).FID <> rUs(j).FID Then GoSub unter
 If roUs(i).Pat_id <> rUs(j).Pat_id Then GoSub unter
 If roUs(i).Zeitpunkt <> rUs(j).Zeitpunkt Then GoSub unter
 If roUs(i).art <> rUs(j).art Then GoSub unter
 If roUs(i).Spritzst <> rUs(j).Spritzst Then GoSub unter
 If roUs(i).Fußbef_re <> rUs(j).Fußbef_re Then GoSub unter
 If roUs(i).Fußbef_li <> rUs(j).Fußbef_li Then GoSub unter
 If roUs(i).Hyperk_re <> rUs(j).Hyperk_re Then GoSub unter
 If roUs(i).Hyperk_li <> rUs(j).Hyperk_li Then GoSub unter
 If roUs(i).Ulcera_re <> rUs(j).Ulcera_re Then GoSub unter
 If roUs(i).Ulcera_li <> rUs(j).Ulcera_li Then GoSub unter
 If roUs(i).Kraft_Zh_re <> rUs(j).Kraft_Zh_re Then GoSub unter
 If roUs(i).Kraft_Zh_li <> rUs(j).Kraft_Zh_li Then GoSub unter
 If roUs(i).Kraft_Zb_re <> rUs(j).Kraft_Zb_re Then GoSub unter
 If roUs(i).Kraft_Zb_li <> rUs(j).Kraft_Zb_li Then GoSub unter
 If roUs(i).Kraft_Knie_re <> rUs(j).Kraft_Knie_re Then GoSub unter
 If roUs(i).Kraft_Knie_li <> rUs(j).Kraft_Knie_li Then GoSub unter
 If roUs(i).ASR_re <> rUs(j).ASR_re Then GoSub unter
 If roUs(i).ASR_li <> rUs(j).ASR_li Then GoSub unter
 If roUs(i).PSR_re <> rUs(j).PSR_re Then GoSub unter
 If roUs(i).PSR_li <> rUs(j).PSR_li Then GoSub unter
 If roUs(i).Oberfl_re <> rUs(j).Oberfl_re Then GoSub unter
 If roUs(i).Oberfl_li <> rUs(j).Oberfl_li Then GoSub unter
 If roUs(i).MF_re <> rUs(j).MF_re Then GoSub unter
 If roUs(i).MF_li <> rUs(j).MF_li Then GoSub unter
 If roUs(i).KW_re <> rUs(j).KW_re Then GoSub unter
 If roUs(i).KW_li <> rUs(j).KW_li Then GoSub unter
 If roUs(i).Vibr_IK_re <> rUs(j).Vibr_IK_re Then GoSub unter
 If roUs(i).Vibr_IK_li <> rUs(j).Vibr_IK_li Then GoSub unter
 If roUs(i).Vibr_GZ_re <> rUs(j).Vibr_GZ_re Then GoSub unter
 If roUs(i).Vibr_GZ_li <> rUs(j).Vibr_GZ_li Then GoSub unter
 If roUs(i).PulsL_re <> rUs(j).PulsL_re Then GoSub unter
 If roUs(i).PulsL_li <> rUs(j).PulsL_li Then GoSub unter
 If roUs(i).PulsKK_re <> rUs(j).PulsKK_re Then GoSub unter
 If roUs(i).PulsKK_li <> rUs(j).PulsKK_li Then GoSub unter
 If roUs(i).PulsAtp_re <> rUs(j).PulsAtp_re Then GoSub unter
 If roUs(i).PulsAtp_li <> rUs(j).PulsAtp_li Then GoSub unter
 If roUs(i).PulsAdp_re <> rUs(j).PulsAdp_re Then GoSub unter
 If roUs(i).PulsAdp_li <> rUs(j).PulsAdp_li Then GoSub unter
 If roUs(i).Mitarbeiter <> rUs(j).Mitarbeiter Then GoSub unter
 If roUs(i).absPos <> rUs(j).absPos Then GoSub unter
 If roUs(i).AktZeit <> rUs(j).AktZeit Then GoSub unter
 If roUs(i).QS <> rUs(j).QS Then GoSub unter
 If roUs(i).QT <> rUs(j).QT Then GoSub unter
 If roUs(i).StByte <> rUs(j).StByte Then GoSub unter
 If roUs(i).id <> rUs(j).id Then GoSub unter
 Exit Function
unter:
 UsZUnt = UsZUnt + 1
 Return
End Function ' UsZUnt

Public Function usdmLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Spritzst,0) Spritzst,COALESCE(Fußbef_re,0) Fußbef_re,COALESCE(Fußbef_li,0) Fußbef_li,COALESCE(Hyperk_re,0) Hyperk_re" & _
",COALESCE(Hyperk_li,0) Hyperk_li,COALESCE(Ulcera_re,0) Ulcera_re,COALESCE(Ulcera_li,0) Ulcera_li,COALESCE(Kraft_Zh_re,0) Kraft_Zh_re" & _
",COALESCE(Kraft_Zh_li,0) Kraft_Zh_li,COALESCE(Kraft_Zb_re,0) Kraft_Zb_re,COALESCE(Kraft_Zb_li,0) Kraft_Zb_li,COALESCE(Kraft_Knie_re,0) Kraft_Knie_re" & _
",COALESCE(Kraft_Knie_li,0) Kraft_Knie_li,COALESCE(ASR_re,0) ASR_re,COALESCE(ASR_li,0) ASR_li,COALESCE(PSR_re,0) PSR_re" & _
",COALESCE(PSR_li,0) PSR_li,COALESCE(Oberfl_re,0) Oberfl_re,COALESCE(Oberfl_li,0) Oberfl_li,COALESCE(MF_re,0) MF_re" & _
",COALESCE(MF_li,0) MF_li,COALESCE(KW_re,0) KW_re,COALESCE(KW_li,0) KW_li,COALESCE(Vibr_IK_re,0) Vibr_IK_re" & _
",COALESCE(Vibr_IK_li,0) Vibr_IK_li,COALESCE(Vibr_GZ_re,0) Vibr_GZ_re,COALESCE(Vibr_GZ_li,0) Vibr_GZ_li,COALESCE(PulsL_re,0) PulsL_re" & _
",COALESCE(PulsL_li,0) PulsL_li,COALESCE(PulsKK_re,0) PulsKK_re,COALESCE(PulsKK_li,0) PulsKK_li,COALESCE(PulsAtp_re,0) PulsAtp_re" & _
",COALESCE(PulsAtp_li,0) PulsAtp_li,COALESCE(PulsAdp_re,0) PulsAdp_re,COALESCE(PulsAdp_li,0) PulsAdp_li,COALESCE(Mitarbeiter,0) Mitarbeiter" & _
",COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS,COALESCE(QT,'') QT" & _
",COALESCE(StByte,0) StByte,COALESCE(id,0) id FROM `usdm` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roUs(0)
 Else ' rs.EOF Then
  ReDim roUs(1)
  Do While Not rs.EOF
   akt = UBound(roUs)
   roUs(akt).FID = rs!FID
   roUs(akt).Pat_id = rs!Pat_id
   roUs(akt).Zeitpunkt = rs!Zeitpunkt
   roUs(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roUs(akt).Spritzst = doUmwfSQL(rs!Spritzst, lies.obMySQL, False)
   roUs(akt).Fußbef_re = doUmwfSQL(rs!Fußbef_re, lies.obMySQL, False)
   roUs(akt).Fußbef_li = doUmwfSQL(rs!Fußbef_li, lies.obMySQL, False)
   roUs(akt).Hyperk_re = doUmwfSQL(rs!Hyperk_re, lies.obMySQL, False)
   roUs(akt).Hyperk_li = doUmwfSQL(rs!Hyperk_li, lies.obMySQL, False)
   roUs(akt).Ulcera_re = doUmwfSQL(rs!Ulcera_re, lies.obMySQL, False)
   roUs(akt).Ulcera_li = doUmwfSQL(rs!Ulcera_li, lies.obMySQL, False)
   roUs(akt).Kraft_Zh_re = doUmwfSQL(rs!Kraft_Zh_re, lies.obMySQL, False)
   roUs(akt).Kraft_Zh_li = doUmwfSQL(rs!Kraft_Zh_li, lies.obMySQL, False)
   roUs(akt).Kraft_Zb_re = doUmwfSQL(rs!Kraft_Zb_re, lies.obMySQL, False)
   roUs(akt).Kraft_Zb_li = doUmwfSQL(rs!Kraft_Zb_li, lies.obMySQL, False)
   roUs(akt).Kraft_Knie_re = doUmwfSQL(rs!Kraft_Knie_re, lies.obMySQL, False)
   roUs(akt).Kraft_Knie_li = doUmwfSQL(rs!Kraft_Knie_li, lies.obMySQL, False)
   roUs(akt).ASR_re = doUmwfSQL(rs!ASR_re, lies.obMySQL, False)
   roUs(akt).ASR_li = doUmwfSQL(rs!ASR_li, lies.obMySQL, False)
   roUs(akt).PSR_re = doUmwfSQL(rs!PSR_re, lies.obMySQL, False)
   roUs(akt).PSR_li = doUmwfSQL(rs!PSR_li, lies.obMySQL, False)
   roUs(akt).Oberfl_re = doUmwfSQL(rs!Oberfl_re, lies.obMySQL, False)
   roUs(akt).Oberfl_li = doUmwfSQL(rs!Oberfl_li, lies.obMySQL, False)
   roUs(akt).MF_re = doUmwfSQL(rs!MF_re, lies.obMySQL, False)
   roUs(akt).MF_li = doUmwfSQL(rs!MF_li, lies.obMySQL, False)
   roUs(akt).KW_re = doUmwfSQL(rs!KW_re, lies.obMySQL, False)
   roUs(akt).KW_li = doUmwfSQL(rs!KW_li, lies.obMySQL, False)
   roUs(akt).Vibr_IK_re = doUmwfSQL(rs!Vibr_IK_re, lies.obMySQL, False)
   roUs(akt).Vibr_IK_li = doUmwfSQL(rs!Vibr_IK_li, lies.obMySQL, False)
   roUs(akt).Vibr_GZ_re = doUmwfSQL(rs!Vibr_GZ_re, lies.obMySQL, False)
   roUs(akt).Vibr_GZ_li = doUmwfSQL(rs!Vibr_GZ_li, lies.obMySQL, False)
   roUs(akt).PulsL_re = doUmwfSQL(rs!PulsL_re, lies.obMySQL, False)
   roUs(akt).PulsL_li = doUmwfSQL(rs!PulsL_li, lies.obMySQL, False)
   roUs(akt).PulsKK_re = doUmwfSQL(rs!PulsKK_re, lies.obMySQL, False)
   roUs(akt).PulsKK_li = doUmwfSQL(rs!PulsKK_li, lies.obMySQL, False)
   roUs(akt).PulsAtp_re = doUmwfSQL(rs!PulsAtp_re, lies.obMySQL, False)
   roUs(akt).PulsAtp_li = doUmwfSQL(rs!PulsAtp_li, lies.obMySQL, False)
   roUs(akt).PulsAdp_re = doUmwfSQL(rs!PulsAdp_re, lies.obMySQL, False)
   roUs(akt).PulsAdp_li = doUmwfSQL(rs!PulsAdp_li, lies.obMySQL, False)
   roUs(akt).Mitarbeiter = doUmwfSQL(rs!Mitarbeiter, lies.obMySQL, False)
   roUs(akt).absPos = rs!absPos
   roUs(akt).AktZeit = rs!AktZeit
   roUs(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roUs(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roUs(akt).StByte = rs!StByte
   roUs(akt).id = rs!id
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roUs(UBound(roUs) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' usdmLaden

Function usdmEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rUs) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rUs)
   If rUs(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roUs)
    If roUs(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roUs(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roUs(robeg + UBound(rUs) - rbeg)
   For ri = rbeg To UBound(rUs)
    Call roUsZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rUs = roUs
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' usdmEinf

Public Function usdmSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rUs) + 0 & " Sätze in `usdm`"
' sql0 = " INSERT " & sqlignore &  "INTO `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte)             VALUES"))
 For i = 1 To UBound(rUs)
'  rUs(i).AktZeit = now()
  rUs(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rUs(i).FID, ",", rUs(i).Pat_id, ",", DatFor_k(rUs(i).Zeitpunkt), ",'", rUs(i).art, "','", rUs(i).Spritzst, "','", rUs(i).Fußbef_re, "','", rUs(i).Fußbef_li, "','", rUs(i).Hyperk_re, "','", _
   rUs(i).Hyperk_li, "','", rUs(i).Ulcera_re, "','", rUs(i).Ulcera_li, "','", rUs(i).Kraft_Zh_re, "','", rUs(i).Kraft_Zh_li, "','", rUs(i).Kraft_Zb_re, "','", rUs(i).Kraft_Zb_li, "','", _
   rUs(i).Kraft_Knie_re, "','", rUs(i).Kraft_Knie_li, "','", rUs(i).ASR_re, "','", rUs(i).ASR_li, "','", rUs(i).PSR_re, "','", rUs(i).PSR_li, "','", rUs(i).Oberfl_re, "','", rUs(i).Oberfl_li, "','", _
   rUs(i).MF_re, "','", rUs(i).MF_li, "','", rUs(i).KW_re, "','", rUs(i).KW_li, "','", rUs(i).Vibr_IK_re, "','", rUs(i).Vibr_IK_li, "','", rUs(i).Vibr_GZ_re, "','", _
   rUs(i).Vibr_GZ_li, "','", rUs(i).PulsL_re, "','", rUs(i).PulsL_li, "','", rUs(i).PulsKK_re, "','", rUs(i).PulsKK_li, "','", rUs(i).PulsAtp_re, "','", rUs(i).PulsAtp_li, "','", rUs(i).PulsAdp_re, "','", _
   rUs(i).PulsAdp_li, "','", rUs(i).Mitarbeiter, "',", rUs(i).absPos, ",", DatFor_k(rUs(i).AktZeit), ",'", rUs(i).QS, "','", rUs(i).QT, "',", rUs(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rUs) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUs) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `usdm` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rUs)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""usdmSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(38)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUs), i)
  If Len(rUs(k).art) > maxi(0) Then maxi(0) = Len(rUs(k).art)
  If Len(rUs(k).Spritzst) > maxi(1) Then maxi(1) = Len(rUs(k).Spritzst)
  If Len(rUs(k).Fußbef_re) > maxi(2) Then maxi(2) = Len(rUs(k).Fußbef_re)
  If Len(rUs(k).Fußbef_li) > maxi(3) Then maxi(3) = Len(rUs(k).Fußbef_li)
  If Len(rUs(k).Hyperk_re) > maxi(4) Then maxi(4) = Len(rUs(k).Hyperk_re)
  If Len(rUs(k).Hyperk_li) > maxi(5) Then maxi(5) = Len(rUs(k).Hyperk_li)
  If Len(rUs(k).Ulcera_re) > maxi(6) Then maxi(6) = Len(rUs(k).Ulcera_re)
  If Len(rUs(k).Ulcera_li) > maxi(7) Then maxi(7) = Len(rUs(k).Ulcera_li)
  If Len(rUs(k).Kraft_Zh_re) > maxi(8) Then maxi(8) = Len(rUs(k).Kraft_Zh_re)
  If Len(rUs(k).Kraft_Zh_li) > maxi(9) Then maxi(9) = Len(rUs(k).Kraft_Zh_li)
  If Len(rUs(k).Kraft_Zb_re) > maxi(10) Then maxi(10) = Len(rUs(k).Kraft_Zb_re)
  If Len(rUs(k).Kraft_Zb_li) > maxi(11) Then maxi(11) = Len(rUs(k).Kraft_Zb_li)
  If Len(rUs(k).Kraft_Knie_re) > maxi(12) Then maxi(12) = Len(rUs(k).Kraft_Knie_re)
  If Len(rUs(k).Kraft_Knie_li) > maxi(13) Then maxi(13) = Len(rUs(k).Kraft_Knie_li)
  If Len(rUs(k).ASR_re) > maxi(14) Then maxi(14) = Len(rUs(k).ASR_re)
  If Len(rUs(k).ASR_li) > maxi(15) Then maxi(15) = Len(rUs(k).ASR_li)
  If Len(rUs(k).PSR_re) > maxi(16) Then maxi(16) = Len(rUs(k).PSR_re)
  If Len(rUs(k).PSR_li) > maxi(17) Then maxi(17) = Len(rUs(k).PSR_li)
  If Len(rUs(k).Oberfl_re) > maxi(18) Then maxi(18) = Len(rUs(k).Oberfl_re)
  If Len(rUs(k).Oberfl_li) > maxi(19) Then maxi(19) = Len(rUs(k).Oberfl_li)
  If Len(rUs(k).MF_re) > maxi(20) Then maxi(20) = Len(rUs(k).MF_re)
  If Len(rUs(k).MF_li) > maxi(21) Then maxi(21) = Len(rUs(k).MF_li)
  If Len(rUs(k).KW_re) > maxi(22) Then maxi(22) = Len(rUs(k).KW_re)
  If Len(rUs(k).KW_li) > maxi(23) Then maxi(23) = Len(rUs(k).KW_li)
  If Len(rUs(k).Vibr_IK_re) > maxi(24) Then maxi(24) = Len(rUs(k).Vibr_IK_re)
  If Len(rUs(k).Vibr_IK_li) > maxi(25) Then maxi(25) = Len(rUs(k).Vibr_IK_li)
  If Len(rUs(k).Vibr_GZ_re) > maxi(26) Then maxi(26) = Len(rUs(k).Vibr_GZ_re)
  If Len(rUs(k).Vibr_GZ_li) > maxi(27) Then maxi(27) = Len(rUs(k).Vibr_GZ_li)
  If Len(rUs(k).PulsL_re) > maxi(28) Then maxi(28) = Len(rUs(k).PulsL_re)
  If Len(rUs(k).PulsL_li) > maxi(29) Then maxi(29) = Len(rUs(k).PulsL_li)
  If Len(rUs(k).PulsKK_re) > maxi(30) Then maxi(30) = Len(rUs(k).PulsKK_re)
  If Len(rUs(k).PulsKK_li) > maxi(31) Then maxi(31) = Len(rUs(k).PulsKK_li)
  If Len(rUs(k).PulsAtp_re) > maxi(32) Then maxi(32) = Len(rUs(k).PulsAtp_re)
  If Len(rUs(k).PulsAtp_li) > maxi(33) Then maxi(33) = Len(rUs(k).PulsAtp_li)
  If Len(rUs(k).PulsAdp_re) > maxi(34) Then maxi(34) = Len(rUs(k).PulsAdp_re)
  If Len(rUs(k).PulsAdp_li) > maxi(35) Then maxi(35) = Len(rUs(k).PulsAdp_li)
  If Len(rUs(k).Mitarbeiter) > maxi(36) Then maxi(36) = Len(rUs(k).Mitarbeiter)
  If Len(rUs(k).QS) > maxi(37) Then maxi(37) = Len(rUs(k).QS)
  If Len(rUs(k).QT) > maxi(38) Then maxi(38) = Len(rUs(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "usdm", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "usdm", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUs.Art: '" & rUs(k).art & "' -> '" & Left$(rUs(k).art, maxL) & "'", True: rUs(k).art = Left$(rUs(k).art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUs.Spritzst: '" & rUs(k).Spritzst & "' -> '" & Left$(rUs(k).Spritzst, maxL) & "'", True: rUs(k).Spritzst = Left$(rUs(k).Spritzst, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_re: '" & rUs(k).Fußbef_re & "' -> '" & Left$(rUs(k).Fußbef_re, maxL) & "'", True: rUs(k).Fußbef_re = Left$(rUs(k).Fußbef_re, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_li: '" & rUs(k).Fußbef_li & "' -> '" & Left$(rUs(k).Fußbef_li, maxL) & "'", True: rUs(k).Fußbef_li = Left$(rUs(k).Fußbef_li, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_re: '" & rUs(k).Hyperk_re & "' -> '" & Left$(rUs(k).Hyperk_re, maxL) & "'", True: rUs(k).Hyperk_re = Left$(rUs(k).Hyperk_re, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_li: '" & rUs(k).Hyperk_li & "' -> '" & Left$(rUs(k).Hyperk_li, maxL) & "'", True: rUs(k).Hyperk_li = Left$(rUs(k).Hyperk_li, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_re: '" & rUs(k).Ulcera_re & "' -> '" & Left$(rUs(k).Ulcera_re, maxL) & "'", True: rUs(k).Ulcera_re = Left$(rUs(k).Ulcera_re, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_li: '" & rUs(k).Ulcera_li & "' -> '" & Left$(rUs(k).Ulcera_li, maxL) & "'", True: rUs(k).Ulcera_li = Left$(rUs(k).Ulcera_li, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_re: '" & rUs(k).Kraft_Zh_re & "' -> '" & Left$(rUs(k).Kraft_Zh_re, maxL) & "'", True: rUs(k).Kraft_Zh_re = Left$(rUs(k).Kraft_Zh_re, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_li: '" & rUs(k).Kraft_Zh_li & "' -> '" & Left$(rUs(k).Kraft_Zh_li, maxL) & "'", True: rUs(k).Kraft_Zh_li = Left$(rUs(k).Kraft_Zh_li, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_re: '" & rUs(k).Kraft_Zb_re & "' -> '" & Left$(rUs(k).Kraft_Zb_re, maxL) & "'", True: rUs(k).Kraft_Zb_re = Left$(rUs(k).Kraft_Zb_re, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_li: '" & rUs(k).Kraft_Zb_li & "' -> '" & Left$(rUs(k).Kraft_Zb_li, maxL) & "'", True: rUs(k).Kraft_Zb_li = Left$(rUs(k).Kraft_Zb_li, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_re: '" & rUs(k).Kraft_Knie_re & "' -> '" & Left$(rUs(k).Kraft_Knie_re, maxL) & "'", True: rUs(k).Kraft_Knie_re = Left$(rUs(k).Kraft_Knie_re, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_li: '" & rUs(k).Kraft_Knie_li & "' -> '" & Left$(rUs(k).Kraft_Knie_li, maxL) & "'", True: rUs(k).Kraft_Knie_li = Left$(rUs(k).Kraft_Knie_li, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_re: '" & rUs(k).ASR_re & "' -> '" & Left$(rUs(k).ASR_re, maxL) & "'", True: rUs(k).ASR_re = Left$(rUs(k).ASR_re, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_li: '" & rUs(k).ASR_li & "' -> '" & Left$(rUs(k).ASR_li, maxL) & "'", True: rUs(k).ASR_li = Left$(rUs(k).ASR_li, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_re: '" & rUs(k).PSR_re & "' -> '" & Left$(rUs(k).PSR_re, maxL) & "'", True: rUs(k).PSR_re = Left$(rUs(k).PSR_re, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_li: '" & rUs(k).PSR_li & "' -> '" & Left$(rUs(k).PSR_li, maxL) & "'", True: rUs(k).PSR_li = Left$(rUs(k).PSR_li, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_re: '" & rUs(k).Oberfl_re & "' -> '" & Left$(rUs(k).Oberfl_re, maxL) & "'", True: rUs(k).Oberfl_re = Left$(rUs(k).Oberfl_re, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_li: '" & rUs(k).Oberfl_li & "' -> '" & Left$(rUs(k).Oberfl_li, maxL) & "'", True: rUs(k).Oberfl_li = Left$(rUs(k).Oberfl_li, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_re: '" & rUs(k).MF_re & "' -> '" & Left$(rUs(k).MF_re, maxL) & "'", True: rUs(k).MF_re = Left$(rUs(k).MF_re, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_li: '" & rUs(k).MF_li & "' -> '" & Left$(rUs(k).MF_li, maxL) & "'", True: rUs(k).MF_li = Left$(rUs(k).MF_li, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_re: '" & rUs(k).KW_re & "' -> '" & Left$(rUs(k).KW_re, maxL) & "'", True: rUs(k).KW_re = Left$(rUs(k).KW_re, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_li: '" & rUs(k).KW_li & "' -> '" & Left$(rUs(k).KW_li, maxL) & "'", True: rUs(k).KW_li = Left$(rUs(k).KW_li, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_re: '" & rUs(k).Vibr_IK_re & "' -> '" & Left$(rUs(k).Vibr_IK_re, maxL) & "'", True: rUs(k).Vibr_IK_re = Left$(rUs(k).Vibr_IK_re, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_li: '" & rUs(k).Vibr_IK_li & "' -> '" & Left$(rUs(k).Vibr_IK_li, maxL) & "'", True: rUs(k).Vibr_IK_li = Left$(rUs(k).Vibr_IK_li, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_re: '" & rUs(k).Vibr_GZ_re & "' -> '" & Left$(rUs(k).Vibr_GZ_re, maxL) & "'", True: rUs(k).Vibr_GZ_re = Left$(rUs(k).Vibr_GZ_re, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_li: '" & rUs(k).Vibr_GZ_li & "' -> '" & Left$(rUs(k).Vibr_GZ_li, maxL) & "'", True: rUs(k).Vibr_GZ_li = Left$(rUs(k).Vibr_GZ_li, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_re: '" & rUs(k).PulsL_re & "' -> '" & Left$(rUs(k).PulsL_re, maxL) & "'", True: rUs(k).PulsL_re = Left$(rUs(k).PulsL_re, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_li: '" & rUs(k).PulsL_li & "' -> '" & Left$(rUs(k).PulsL_li, maxL) & "'", True: rUs(k).PulsL_li = Left$(rUs(k).PulsL_li, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_re: '" & rUs(k).PulsKK_re & "' -> '" & Left$(rUs(k).PulsKK_re, maxL) & "'", True: rUs(k).PulsKK_re = Left$(rUs(k).PulsKK_re, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_li: '" & rUs(k).PulsKK_li & "' -> '" & Left$(rUs(k).PulsKK_li, maxL) & "'", True: rUs(k).PulsKK_li = Left$(rUs(k).PulsKK_li, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_re: '" & rUs(k).PulsAtp_re & "' -> '" & Left$(rUs(k).PulsAtp_re, maxL) & "'", True: rUs(k).PulsAtp_re = Left$(rUs(k).PulsAtp_re, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_li: '" & rUs(k).PulsAtp_li & "' -> '" & Left$(rUs(k).PulsAtp_li, maxL) & "'", True: rUs(k).PulsAtp_li = Left$(rUs(k).PulsAtp_li, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_re: '" & rUs(k).PulsAdp_re & "' -> '" & Left$(rUs(k).PulsAdp_re, maxL) & "'", True: rUs(k).PulsAdp_re = Left$(rUs(k).PulsAdp_re, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_li: '" & rUs(k).PulsAdp_li & "' -> '" & Left$(rUs(k).PulsAdp_li, maxL) & "'", True: rUs(k).PulsAdp_li = Left$(rUs(k).PulsAdp_li, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rUs.Mitarbeiter: '" & rUs(k).Mitarbeiter & "' -> '" & Left$(rUs(k).Mitarbeiter, maxL) & "'", True: rUs(k).Mitarbeiter = Left$(rUs(k).Mitarbeiter, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rUs.QS: '" & rUs(k).QS & "' -> '" & Left$(rUs(k).QS, maxL) & "'", True: rUs(k).QS = Left$(rUs(k).QS, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rUs.QT: '" & rUs(k).QT & "' -> '" & Left$(rUs(k).QT, maxL) & "'", True: rUs(k).QT = Left$(rUs(k).QT, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' usdmSpeichern

Public Function roFuZuw(i&, j&)
 roFu(i).FID = rFu(j).FID
 roFu(i).Pat_id = rFu(j).Pat_id
 roFu(i).Zeitpunkt = rFu(j).Zeitpunkt
 roFu(i).art = rFu(j).art
 roFu(i).Fußdeform = rFu(j).Fußdeform
 roFu(i).Hyper_mEin = rFu(j).Hyper_mEin
 roFu(i).Weiteres = rFu(j).Weiteres
 roFu(i).Zn_Ulcus = rFu(j).Zn_Ulcus
 roFu(i).Zn_Amput = rFu(j).Zn_Amput
 roFu(i).Fuß_ang = rFu(j).Fuß_ang
 roFu(i).Ulcera = rFu(j).Ulcera
 roFu(i).Wundinfektion = rFu(j).Wundinfektion
 roFu(i).nae_US = rFu(j).nae_US
 roFu(i).Mitarbeiter = rFu(j).Mitarbeiter
 roFu(i).absPos = rFu(j).absPos
 roFu(i).AktZeit = rFu(j).AktZeit
 roFu(i).QS = rFu(j).QS
 roFu(i).QT = rFu(j).QT
 roFu(i).StByte = rFu(j).StByte
 roFu(i).id = rFu(j).id
End Function ' roFuZuw

Public Function FuZUnt%(i&, j&)
 If roFu(i).FID <> rFu(j).FID Then GoSub unter
 If roFu(i).Pat_id <> rFu(j).Pat_id Then GoSub unter
 If roFu(i).Zeitpunkt <> rFu(j).Zeitpunkt Then GoSub unter
 If roFu(i).art <> rFu(j).art Then GoSub unter
 If roFu(i).Fußdeform <> rFu(j).Fußdeform Then GoSub unter
 If roFu(i).Hyper_mEin <> rFu(j).Hyper_mEin Then GoSub unter
 If roFu(i).Weiteres <> rFu(j).Weiteres Then GoSub unter
 If roFu(i).Zn_Ulcus <> rFu(j).Zn_Ulcus Then GoSub unter
 If roFu(i).Zn_Amput <> rFu(j).Zn_Amput Then GoSub unter
 If roFu(i).Fuß_ang <> rFu(j).Fuß_ang Then GoSub unter
 If roFu(i).Ulcera <> rFu(j).Ulcera Then GoSub unter
 If roFu(i).Wundinfektion <> rFu(j).Wundinfektion Then GoSub unter
 If roFu(i).nae_US <> rFu(j).nae_US Then GoSub unter
 If roFu(i).Mitarbeiter <> rFu(j).Mitarbeiter Then GoSub unter
 If roFu(i).absPos <> rFu(j).absPos Then GoSub unter
 If roFu(i).AktZeit <> rFu(j).AktZeit Then GoSub unter
 If roFu(i).QS <> rFu(j).QS Then GoSub unter
 If roFu(i).QT <> rFu(j).QT Then GoSub unter
 If roFu(i).StByte <> rFu(j).StByte Then GoSub unter
 If roFu(i).id <> rFu(j).id Then GoSub unter
 Exit Function
unter:
 FuZUnt = FuZUnt + 1
 Return
End Function ' FuZUnt

Public Function fussLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Fußdeform,'') Fußdeform,COALESCE(Hyper_mEin,'') Hyper_mEin,COALESCE(Weiteres,'') Weiteres,COALESCE(Zn_Ulcus,'') Zn_Ulcus" & _
",COALESCE(Zn_Amput,'') Zn_Amput,COALESCE(Fuß_ang,'') Fuß_ang,COALESCE(Ulcera,'') Ulcera,COALESCE(Wundinfektion,'') Wundinfektion" & _
",COALESCE(nae_US,'') nae_US,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(id,0) id" & _
" FROM `fuss` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roFu(0)
 Else ' rs.EOF Then
  ReDim roFu(1)
  Do While Not rs.EOF
   akt = UBound(roFu)
   roFu(akt).FID = rs!FID
   roFu(akt).Pat_id = rs!Pat_id
   roFu(akt).Zeitpunkt = rs!Zeitpunkt
   roFu(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roFu(akt).Fußdeform = doUmwfSQL(rs!Fußdeform, lies.obMySQL, False)
   roFu(akt).Hyper_mEin = doUmwfSQL(rs!Hyper_mEin, lies.obMySQL, False)
   roFu(akt).Weiteres = doUmwfSQL(rs!Weiteres, lies.obMySQL, False)
   roFu(akt).Zn_Ulcus = doUmwfSQL(rs!Zn_Ulcus, lies.obMySQL, False)
   roFu(akt).Zn_Amput = doUmwfSQL(rs!Zn_Amput, lies.obMySQL, False)
   roFu(akt).Fuß_ang = doUmwfSQL(rs!Fuß_ang, lies.obMySQL, False)
   roFu(akt).Ulcera = doUmwfSQL(rs!Ulcera, lies.obMySQL, False)
   roFu(akt).Wundinfektion = doUmwfSQL(rs!Wundinfektion, lies.obMySQL, False)
   roFu(akt).nae_US = doUmwfSQL(rs!nae_US, lies.obMySQL, False)
   roFu(akt).Mitarbeiter = doUmwfSQL(rs!Mitarbeiter, lies.obMySQL, False)
   roFu(akt).absPos = rs!absPos
   roFu(akt).AktZeit = rs!AktZeit
   roFu(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roFu(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roFu(akt).StByte = rs!StByte
   roFu(akt).id = rs!id
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roFu(UBound(roFu) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' fussLaden

Function fussEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rFu) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rFu)
   If rFu(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roFu)
    If roFu(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roFu(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFu(robeg + UBound(rFu) - rbeg)
   For ri = rbeg To UBound(rFu)
    Call roFuZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rFu = roFu
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' fussEinf

Public Function fussSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFu) + 0 & " Sätze in `fuss`"
' sql0 = " INSERT " & sqlignore &  "INTO `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte)           VALUES"))
 For i = 1 To UBound(rFu)
'  rFu(i).AktZeit = now()
  rFu(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rFu(i).FID, ",", rFu(i).Pat_id, ",", DatFor_k(rFu(i).Zeitpunkt), ",'", rFu(i).art, "','", rFu(i).Fußdeform, "','", rFu(i).Hyper_mEin, "','", rFu(i).Weiteres, "','", rFu(i).Zn_Ulcus, "','", _
   rFu(i).Zn_Amput, "','", rFu(i).Fuß_ang, "','", rFu(i).Ulcera, "','", rFu(i).Wundinfektion, "','", rFu(i).nae_US, "','", rFu(i).Mitarbeiter, "',", rFu(i).absPos, ",", DatFor_k(rFu(i).AktZeit), ",'", _
   rFu(i).QS, "','", rFu(i).QT, "',", rFu(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rFu) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFu) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `fuss` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rFu)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""fussSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(12)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFu), i)
  If Len(rFu(k).art) > maxi(0) Then maxi(0) = Len(rFu(k).art)
  If Len(rFu(k).Fußdeform) > maxi(1) Then maxi(1) = Len(rFu(k).Fußdeform)
  If Len(rFu(k).Hyper_mEin) > maxi(2) Then maxi(2) = Len(rFu(k).Hyper_mEin)
  If Len(rFu(k).Weiteres) > maxi(3) Then maxi(3) = Len(rFu(k).Weiteres)
  If Len(rFu(k).Zn_Ulcus) > maxi(4) Then maxi(4) = Len(rFu(k).Zn_Ulcus)
  If Len(rFu(k).Zn_Amput) > maxi(5) Then maxi(5) = Len(rFu(k).Zn_Amput)
  If Len(rFu(k).Fuß_ang) > maxi(6) Then maxi(6) = Len(rFu(k).Fuß_ang)
  If Len(rFu(k).Ulcera) > maxi(7) Then maxi(7) = Len(rFu(k).Ulcera)
  If Len(rFu(k).Wundinfektion) > maxi(8) Then maxi(8) = Len(rFu(k).Wundinfektion)
  If Len(rFu(k).nae_US) > maxi(9) Then maxi(9) = Len(rFu(k).nae_US)
  If Len(rFu(k).Mitarbeiter) > maxi(10) Then maxi(10) = Len(rFu(k).Mitarbeiter)
  If Len(rFu(k).QS) > maxi(11) Then maxi(11) = Len(rFu(k).QS)
  If Len(rFu(k).QT) > maxi(12) Then maxi(12) = Len(rFu(k).QT)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "fuss", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "fuss", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFu), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFu.Art: '" & rFu(k).art & "' -> '" & Left$(rFu(k).art, maxL) & "'", True: rFu(k).art = Left$(rFu(k).art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fußdeform: '" & rFu(k).Fußdeform & "' -> '" & Left$(rFu(k).Fußdeform, maxL) & "'", True: rFu(k).Fußdeform = Left$(rFu(k).Fußdeform, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFu.Hyper_mEin: '" & rFu(k).Hyper_mEin & "' -> '" & Left$(rFu(k).Hyper_mEin, maxL) & "'", True: rFu(k).Hyper_mEin = Left$(rFu(k).Hyper_mEin, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFu.Weiteres: '" & rFu(k).Weiteres & "' -> '" & Left$(rFu(k).Weiteres, maxL) & "'", True: rFu(k).Weiteres = Left$(rFu(k).Weiteres, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Ulcus: '" & rFu(k).Zn_Ulcus & "' -> '" & Left$(rFu(k).Zn_Ulcus, maxL) & "'", True: rFu(k).Zn_Ulcus = Left$(rFu(k).Zn_Ulcus, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Amput: '" & rFu(k).Zn_Amput & "' -> '" & Left$(rFu(k).Zn_Amput, maxL) & "'", True: rFu(k).Zn_Amput = Left$(rFu(k).Zn_Amput, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fuß_ang: '" & rFu(k).Fuß_ang & "' -> '" & Left$(rFu(k).Fuß_ang, maxL) & "'", True: rFu(k).Fuß_ang = Left$(rFu(k).Fuß_ang, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFu.Ulcera: '" & rFu(k).Ulcera & "' -> '" & Left$(rFu(k).Ulcera, maxL) & "'", True: rFu(k).Ulcera = Left$(rFu(k).Ulcera, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFu.Wundinfektion: '" & rFu(k).Wundinfektion & "' -> '" & Left$(rFu(k).Wundinfektion, maxL) & "'", True: rFu(k).Wundinfektion = Left$(rFu(k).Wundinfektion, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFu.nae_US: '" & rFu(k).nae_US & "' -> '" & Left$(rFu(k).nae_US, maxL) & "'", True: rFu(k).nae_US = Left$(rFu(k).nae_US, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFu.Mitarbeiter: '" & rFu(k).Mitarbeiter & "' -> '" & Left$(rFu(k).Mitarbeiter, maxL) & "'", True: rFu(k).Mitarbeiter = Left$(rFu(k).Mitarbeiter, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFu.QS: '" & rFu(k).QS & "' -> '" & Left$(rFu(k).QS, maxL) & "'", True: rFu(k).QS = Left$(rFu(k).QS, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFu.QT: '" & rFu(k).QT & "' -> '" & Left$(rFu(k).QT, maxL) & "'", True: rFu(k).QT = Left$(rFu(k).QT, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' fussSpeichern

Public Function roUlZuw(i&, j&)
 roUl(i).FID = rUl(j).FID
 roUl(i).Pat_id = rUl(j).Pat_id
 roUl(i).Zeitpunkt = rUl(j).Zeitpunkt
 roUl(i).Lokalisation = rUl(j).Lokalisation
 roUl(i).Seite = rUl(j).Seite
 roUl(i).Größe = rUl(j).Größe
 roUl(i).Beläge = rUl(j).Beläge
 roUl(i).Exsudat = rUl(j).Exsudat
 roUl(i).Geruch = rUl(j).Geruch
 roUl(i).Wundrand = rUl(j).Wundrand
 roUl(i).Wundumgebung = rUl(j).Wundumgebung
 roUl(i).Temperatur = rUl(j).Temperatur
 roUl(i).Fotodoku = rUl(j).Fotodoku
 roUl(i).Wundversorgung = rUl(j).Wundversorgung
 roUl(i).Mitarbeiter = rUl(j).Mitarbeiter
 roUl(i).absPos = rUl(j).absPos
 roUl(i).AktZeit = rUl(j).AktZeit
 roUl(i).StByte = rUl(j).StByte
End Function ' roUlZuw

Public Function UlZUnt%(i&, j&)
 If roUl(i).FID <> rUl(j).FID Then GoSub unter
 If roUl(i).Pat_id <> rUl(j).Pat_id Then GoSub unter
 If roUl(i).Zeitpunkt <> rUl(j).Zeitpunkt Then GoSub unter
 If roUl(i).Lokalisation <> rUl(j).Lokalisation Then GoSub unter
 If roUl(i).Seite <> rUl(j).Seite Then GoSub unter
 If roUl(i).Größe <> rUl(j).Größe Then GoSub unter
 If roUl(i).Beläge <> rUl(j).Beläge Then GoSub unter
 If roUl(i).Exsudat <> rUl(j).Exsudat Then GoSub unter
 If roUl(i).Geruch <> rUl(j).Geruch Then GoSub unter
 If roUl(i).Wundrand <> rUl(j).Wundrand Then GoSub unter
 If roUl(i).Wundumgebung <> rUl(j).Wundumgebung Then GoSub unter
 If roUl(i).Temperatur <> rUl(j).Temperatur Then GoSub unter
 If roUl(i).Fotodoku <> rUl(j).Fotodoku Then GoSub unter
 If roUl(i).Wundversorgung <> rUl(j).Wundversorgung Then GoSub unter
 If roUl(i).Mitarbeiter <> rUl(j).Mitarbeiter Then GoSub unter
 If roUl(i).absPos <> rUl(j).absPos Then GoSub unter
 If roUl(i).AktZeit <> rUl(j).AktZeit Then GoSub unter
 If roUl(i).StByte <> rUl(j).StByte Then GoSub unter
 Exit Function
unter:
 UlZUnt = UlZUnt + 1
 Return
End Function ' UlZUnt

Public Function ulcusLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Lokalisation,'') Lokalisation" & _
",COALESCE(Seite,'') Seite,COALESCE(Größe,'') Größe,COALESCE(Beläge,'') Beläge,COALESCE(Exsudat,'') Exsudat" & _
",COALESCE(Geruch,'') Geruch,COALESCE(Wundrand,'') Wundrand,COALESCE(Wundumgebung,'') Wundumgebung,COALESCE(Temperatur,'') Temperatur" & _
",COALESCE(Fotodoku,'') Fotodoku,COALESCE(Wundversorgung,'') Wundversorgung,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `ulcus` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roUl(0)
 Else ' rs.EOF Then
  ReDim roUl(1)
  Do While Not rs.EOF
   akt = UBound(roUl)
   roUl(akt).FID = rs!FID
   roUl(akt).Pat_id = rs!Pat_id
   roUl(akt).Zeitpunkt = rs!Zeitpunkt
   roUl(akt).Lokalisation = doUmwfSQL(rs!Lokalisation, lies.obMySQL, False)
   roUl(akt).Seite = doUmwfSQL(rs!Seite, lies.obMySQL, False)
   roUl(akt).Größe = doUmwfSQL(rs!Größe, lies.obMySQL, False)
   roUl(akt).Beläge = doUmwfSQL(rs!Beläge, lies.obMySQL, False)
   roUl(akt).Exsudat = doUmwfSQL(rs!Exsudat, lies.obMySQL, False)
   roUl(akt).Geruch = doUmwfSQL(rs!Geruch, lies.obMySQL, False)
   roUl(akt).Wundrand = doUmwfSQL(rs!Wundrand, lies.obMySQL, False)
   roUl(akt).Wundumgebung = doUmwfSQL(rs!Wundumgebung, lies.obMySQL, False)
   roUl(akt).Temperatur = doUmwfSQL(rs!Temperatur, lies.obMySQL, False)
   roUl(akt).Fotodoku = doUmwfSQL(rs!Fotodoku, lies.obMySQL, False)
   roUl(akt).Wundversorgung = doUmwfSQL(rs!Wundversorgung, lies.obMySQL, False)
   roUl(akt).Mitarbeiter = doUmwfSQL(rs!Mitarbeiter, lies.obMySQL, False)
   roUl(akt).absPos = rs!absPos
   roUl(akt).AktZeit = rs!AktZeit
   roUl(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roUl(UBound(roUl) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' ulcusLaden

Function ulcusEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rUl) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rUl)
   If rUl(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roUl)
    If roUl(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roUl(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roUl(robeg + UBound(rUl) - rbeg)
   For ri = rbeg To UBound(rUl)
    Call roUlZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rUl = roUl
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' ulcusEinf

Public Function ulcusSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rUl) + 0 & " Sätze in `ulcus`"
' sql0 = " INSERT " & sqlignore &  "INTO `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte)  VALUES"))
 For i = 1 To UBound(rUl)
'  rUl(i).AktZeit = now()
  rUl(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rUl(i).FID, ",", rUl(i).Pat_id, ",", DatFor_k(rUl(i).Zeitpunkt), ",'", rUl(i).Lokalisation, "','", rUl(i).Seite, "','", rUl(i).Größe, "','", rUl(i).Beläge, "','", rUl(i).Exsudat, "','", _
   rUl(i).Geruch, "','", rUl(i).Wundrand, "','", rUl(i).Wundumgebung, "','", rUl(i).Temperatur, "','", rUl(i).Fotodoku, "','", rUl(i).Wundversorgung, "','", rUl(i).Mitarbeiter, "',", _
   rUl(i).absPos, ",", DatFor_k(rUl(i).AktZeit), ",", rUl(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rUl) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUl) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `ulcus` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rUl)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""ulcusSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(11)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUl), i)
  If Len(rUl(k).Lokalisation) > maxi(0) Then maxi(0) = Len(rUl(k).Lokalisation)
  If Len(rUl(k).Seite) > maxi(1) Then maxi(1) = Len(rUl(k).Seite)
  If Len(rUl(k).Größe) > maxi(2) Then maxi(2) = Len(rUl(k).Größe)
  If Len(rUl(k).Beläge) > maxi(3) Then maxi(3) = Len(rUl(k).Beläge)
  If Len(rUl(k).Exsudat) > maxi(4) Then maxi(4) = Len(rUl(k).Exsudat)
  If Len(rUl(k).Geruch) > maxi(5) Then maxi(5) = Len(rUl(k).Geruch)
  If Len(rUl(k).Wundrand) > maxi(6) Then maxi(6) = Len(rUl(k).Wundrand)
  If Len(rUl(k).Wundumgebung) > maxi(7) Then maxi(7) = Len(rUl(k).Wundumgebung)
  If Len(rUl(k).Temperatur) > maxi(8) Then maxi(8) = Len(rUl(k).Temperatur)
  If Len(rUl(k).Fotodoku) > maxi(9) Then maxi(9) = Len(rUl(k).Fotodoku)
  If Len(rUl(k).Wundversorgung) > maxi(10) Then maxi(10) = Len(rUl(k).Wundversorgung)
  If Len(rUl(k).Mitarbeiter) > maxi(11) Then maxi(11) = Len(rUl(k).Mitarbeiter)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "ulcus", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "ulcus", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUl), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUl.Lokalisation: '" & rUl(k).Lokalisation & "' -> '" & Left$(rUl(k).Lokalisation, maxL) & "'", True: rUl(k).Lokalisation = Left$(rUl(k).Lokalisation, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUl.Seite: '" & rUl(k).Seite & "' -> '" & Left$(rUl(k).Seite, maxL) & "'", True: rUl(k).Seite = Left$(rUl(k).Seite, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUl.Größe: '" & rUl(k).Größe & "' -> '" & Left$(rUl(k).Größe, maxL) & "'", True: rUl(k).Größe = Left$(rUl(k).Größe, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUl.Beläge: '" & rUl(k).Beläge & "' -> '" & Left$(rUl(k).Beläge, maxL) & "'", True: rUl(k).Beläge = Left$(rUl(k).Beläge, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUl.Exsudat: '" & rUl(k).Exsudat & "' -> '" & Left$(rUl(k).Exsudat, maxL) & "'", True: rUl(k).Exsudat = Left$(rUl(k).Exsudat, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUl.Geruch: '" & rUl(k).Geruch & "' -> '" & Left$(rUl(k).Geruch, maxL) & "'", True: rUl(k).Geruch = Left$(rUl(k).Geruch, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundrand: '" & rUl(k).Wundrand & "' -> '" & Left$(rUl(k).Wundrand, maxL) & "'", True: rUl(k).Wundrand = Left$(rUl(k).Wundrand, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundumgebung: '" & rUl(k).Wundumgebung & "' -> '" & Left$(rUl(k).Wundumgebung, maxL) & "'", True: rUl(k).Wundumgebung = Left$(rUl(k).Wundumgebung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUl.Temperatur: '" & rUl(k).Temperatur & "' -> '" & Left$(rUl(k).Temperatur, maxL) & "'", True: rUl(k).Temperatur = Left$(rUl(k).Temperatur, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUl.Fotodoku: '" & rUl(k).Fotodoku & "' -> '" & Left$(rUl(k).Fotodoku, maxL) & "'", True: rUl(k).Fotodoku = Left$(rUl(k).Fotodoku, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundversorgung: '" & rUl(k).Wundversorgung & "' -> '" & Left$(rUl(k).Wundversorgung, maxL) & "'", True: rUl(k).Wundversorgung = Left$(rUl(k).Wundversorgung, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUl.Mitarbeiter: '" & rUl(k).Mitarbeiter & "' -> '" & Left$(rUl(k).Mitarbeiter, maxL) & "'", True: rUl(k).Mitarbeiter = Left$(rUl(k).Mitarbeiter, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' ulcusSpeichern

Public Function roVkZuw(i&, j&)
 roVk(i).FID = rVk(j).FID
 roVk(i).Pat_id = rVk(j).Pat_id
 roVk(i).Zeitpunkt = rVk(j).Zeitpunkt
 roVk(i).Wohlempfinden = rVk(j).Wohlempfinden
 roVk(i).Saettigung = rVk(j).Saettigung
 roVk(i).Zielwerterreichung = rVk(j).Zielwerterreichung
 roVk(i).Ketonkörper = rVk(j).Ketonkörper
 roVk(i).Gynaekologenbefund = rVk(j).Gynaekologenbefund
 roVk(i).Gewichtsentwicklung = rVk(j).Gewichtsentwicklung
 roVk(i).HbA1c = rVk(j).HbA1c
 roVk(i).Bewegung = rVk(j).Bewegung
 roVk(i).Minuten = rVk(j).Minuten
 roVk(i).Blutdruck = rVk(j).Blutdruck
 roVk(i).Puls = rVk(j).Puls
 roVk(i).Mitarbeiter = rVk(j).Mitarbeiter
 roVk(i).absPos = rVk(j).absPos
 roVk(i).AktZeit = rVk(j).AktZeit
 roVk(i).StByte = rVk(j).StByte
End Function ' roVkZuw

Public Function VkZUnt%(i&, j&)
 If roVk(i).FID <> rVk(j).FID Then GoSub unter
 If roVk(i).Pat_id <> rVk(j).Pat_id Then GoSub unter
 If roVk(i).Zeitpunkt <> rVk(j).Zeitpunkt Then GoSub unter
 If roVk(i).Wohlempfinden <> rVk(j).Wohlempfinden Then GoSub unter
 If roVk(i).Saettigung <> rVk(j).Saettigung Then GoSub unter
 If roVk(i).Zielwerterreichung <> rVk(j).Zielwerterreichung Then GoSub unter
 If roVk(i).Ketonkörper <> rVk(j).Ketonkörper Then GoSub unter
 If roVk(i).Gynaekologenbefund <> rVk(j).Gynaekologenbefund Then GoSub unter
 If roVk(i).Gewichtsentwicklung <> rVk(j).Gewichtsentwicklung Then GoSub unter
 If roVk(i).HbA1c <> rVk(j).HbA1c Then GoSub unter
 If roVk(i).Bewegung <> rVk(j).Bewegung Then GoSub unter
 If roVk(i).Minuten <> rVk(j).Minuten Then GoSub unter
 If roVk(i).Blutdruck <> rVk(j).Blutdruck Then GoSub unter
 If roVk(i).Puls <> rVk(j).Puls Then GoSub unter
 If roVk(i).Mitarbeiter <> rVk(j).Mitarbeiter Then GoSub unter
 If roVk(i).absPos <> rVk(j).absPos Then GoSub unter
 If roVk(i).AktZeit <> rVk(j).AktZeit Then GoSub unter
 If roVk(i).StByte <> rVk(j).StByte Then GoSub unter
 Exit Function
unter:
 VkZUnt = VkZUnt + 1
 Return
End Function ' VkZUnt

Public Function vkgdLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Wohlempfinden,'') Wohlempfinden" & _
",COALESCE(Saettigung,'') Saettigung,COALESCE(Zielwerterreichung,'') Zielwerterreichung,COALESCE(Ketonkörper,'') Ketonkörper,COALESCE(Gynaekologenbefund,'') Gynaekologenbefund" & _
",COALESCE(Gewichtsentwicklung,'') Gewichtsentwicklung,COALESCE(HbA1c,'') HbA1c,COALESCE(Bewegung,'') Bewegung,COALESCE(Minuten,'') Minuten" & _
",COALESCE(Blutdruck,'') Blutdruck,COALESCE(Puls,'') Puls,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `vkgd` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roVk(0)
 Else ' rs.EOF Then
  ReDim roVk(1)
  Do While Not rs.EOF
   akt = UBound(roVk)
   roVk(akt).FID = rs!FID
   roVk(akt).Pat_id = rs!Pat_id
   roVk(akt).Zeitpunkt = rs!Zeitpunkt
   roVk(akt).Wohlempfinden = doUmwfSQL(rs!Wohlempfinden, lies.obMySQL, False)
   roVk(akt).Saettigung = doUmwfSQL(rs!Saettigung, lies.obMySQL, False)
   roVk(akt).Zielwerterreichung = doUmwfSQL(rs!Zielwerterreichung, lies.obMySQL, False)
   roVk(akt).Ketonkörper = doUmwfSQL(rs!Ketonkörper, lies.obMySQL, False)
   roVk(akt).Gynaekologenbefund = doUmwfSQL(rs!Gynaekologenbefund, lies.obMySQL, False)
   roVk(akt).Gewichtsentwicklung = doUmwfSQL(rs!Gewichtsentwicklung, lies.obMySQL, False)
   roVk(akt).HbA1c = doUmwfSQL(rs!HbA1c, lies.obMySQL, False)
   roVk(akt).Bewegung = doUmwfSQL(rs!Bewegung, lies.obMySQL, False)
   roVk(akt).Minuten = doUmwfSQL(rs!Minuten, lies.obMySQL, False)
   roVk(akt).Blutdruck = doUmwfSQL(rs!Blutdruck, lies.obMySQL, False)
   roVk(akt).Puls = doUmwfSQL(rs!Puls, lies.obMySQL, False)
   roVk(akt).Mitarbeiter = doUmwfSQL(rs!Mitarbeiter, lies.obMySQL, False)
   roVk(akt).absPos = rs!absPos
   roVk(akt).AktZeit = rs!AktZeit
   roVk(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roVk(UBound(roVk) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' vkgdLaden

Function vkgdEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rVk) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rVk)
   If rVk(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roVk)
    If roVk(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roVk(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roVk(robeg + UBound(rVk) - rbeg)
   For ri = rbeg To UBound(rVk)
    Call roVkZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rVk = roVk
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' vkgdEinf

Public Function vkgdSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rVk) + 0 & " Sätze in `vkgd`"
' sql0 = " INSERT " & sqlignore &  "INTO `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte)            VALUES"))
 For i = 1 To UBound(rVk)
'  rVk(i).AktZeit = now()
  rVk(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rVk(i).FID, ",", rVk(i).Pat_id, ",", DatFor_k(rVk(i).Zeitpunkt), ",'", rVk(i).Wohlempfinden, "','", rVk(i).Saettigung, "','", rVk(i).Zielwerterreichung, "','", rVk(i).Ketonkörper, "','", _
   rVk(i).Gynaekologenbefund, "','", rVk(i).Gewichtsentwicklung, "','", rVk(i).HbA1c, "','", rVk(i).Bewegung, "','", rVk(i).Minuten, "','", rVk(i).Blutdruck, "','", rVk(i).Puls, "','", _
   rVk(i).Mitarbeiter, "',", rVk(i).absPos, ",", DatFor_k(rVk(i).AktZeit), ",", rVk(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rVk) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rVk) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `vkgd` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rVk)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""vkgdSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(11)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rVk), i)
  If Len(rVk(k).Wohlempfinden) > maxi(0) Then maxi(0) = Len(rVk(k).Wohlempfinden)
  If Len(rVk(k).Saettigung) > maxi(1) Then maxi(1) = Len(rVk(k).Saettigung)
  If Len(rVk(k).Zielwerterreichung) > maxi(2) Then maxi(2) = Len(rVk(k).Zielwerterreichung)
  If Len(rVk(k).Ketonkörper) > maxi(3) Then maxi(3) = Len(rVk(k).Ketonkörper)
  If Len(rVk(k).Gynaekologenbefund) > maxi(4) Then maxi(4) = Len(rVk(k).Gynaekologenbefund)
  If Len(rVk(k).Gewichtsentwicklung) > maxi(5) Then maxi(5) = Len(rVk(k).Gewichtsentwicklung)
  If Len(rVk(k).HbA1c) > maxi(6) Then maxi(6) = Len(rVk(k).HbA1c)
  If Len(rVk(k).Bewegung) > maxi(7) Then maxi(7) = Len(rVk(k).Bewegung)
  If Len(rVk(k).Minuten) > maxi(8) Then maxi(8) = Len(rVk(k).Minuten)
  If Len(rVk(k).Blutdruck) > maxi(9) Then maxi(9) = Len(rVk(k).Blutdruck)
  If Len(rVk(k).Puls) > maxi(10) Then maxi(10) = Len(rVk(k).Puls)
  If Len(rVk(k).Mitarbeiter) > maxi(11) Then maxi(11) = Len(rVk(k).Mitarbeiter)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "vkgd", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "vkgd", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rVk), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rVk.Wohlempfinden: '" & rVk(k).Wohlempfinden & "' -> '" & Left$(rVk(k).Wohlempfinden, maxL) & "'", True: rVk(k).Wohlempfinden = Left$(rVk(k).Wohlempfinden, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rVk.Saettigung: '" & rVk(k).Saettigung & "' -> '" & Left$(rVk(k).Saettigung, maxL) & "'", True: rVk(k).Saettigung = Left$(rVk(k).Saettigung, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rVk.Zielwerterreichung: '" & rVk(k).Zielwerterreichung & "' -> '" & Left$(rVk(k).Zielwerterreichung, maxL) & "'", True: rVk(k).Zielwerterreichung = Left$(rVk(k).Zielwerterreichung, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rVk.Ketonkörper: '" & rVk(k).Ketonkörper & "' -> '" & Left$(rVk(k).Ketonkörper, maxL) & "'", True: rVk(k).Ketonkörper = Left$(rVk(k).Ketonkörper, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gynaekologenbefund: '" & rVk(k).Gynaekologenbefund & "' -> '" & Left$(rVk(k).Gynaekologenbefund, maxL) & "'", True: rVk(k).Gynaekologenbefund = Left$(rVk(k).Gynaekologenbefund, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gewichtsentwicklung: '" & rVk(k).Gewichtsentwicklung & "' -> '" & Left$(rVk(k).Gewichtsentwicklung, maxL) & "'", True: rVk(k).Gewichtsentwicklung = Left$(rVk(k).Gewichtsentwicklung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rVk.HbA1c: '" & rVk(k).HbA1c & "' -> '" & Left$(rVk(k).HbA1c, maxL) & "'", True: rVk(k).HbA1c = Left$(rVk(k).HbA1c, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rVk.Bewegung: '" & rVk(k).Bewegung & "' -> '" & Left$(rVk(k).Bewegung, maxL) & "'", True: rVk(k).Bewegung = Left$(rVk(k).Bewegung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rVk.Minuten: '" & rVk(k).Minuten & "' -> '" & Left$(rVk(k).Minuten, maxL) & "'", True: rVk(k).Minuten = Left$(rVk(k).Minuten, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rVk.Blutdruck: '" & rVk(k).Blutdruck & "' -> '" & Left$(rVk(k).Blutdruck, maxL) & "'", True: rVk(k).Blutdruck = Left$(rVk(k).Blutdruck, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rVk.Puls: '" & rVk(k).Puls & "' -> '" & Left$(rVk(k).Puls, maxL) & "'", True: rVk(k).Puls = Left$(rVk(k).Puls, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rVk.Mitarbeiter: '" & rVk(k).Mitarbeiter & "' -> '" & Left$(rVk(k).Mitarbeiter, maxL) & "'", True: rVk(k).Mitarbeiter = Left$(rVk(k).Mitarbeiter, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' vkgdSpeichern

Public Function roSwZuw(i&, j&)
 roSw(i).FID = rSw(j).FID
 roSw(i).Pat_id = rSw(j).Pat_id
 roSw(i).Zeitpunkt = rSw(j).Zeitpunkt
 roSw(i).lR = rSw(j).lR
 roSw(i).vorET = rSw(j).vorET
 roSw(i).ET = rSw(j).ET
 roSw(i).efLR = rSw(j).efLR
 roSw(i).erLR = rSw(j).erLR
 roSw(i).kGT = rSw(j).kGT
 roSw(i).MB = rSw(j).MB
 roSw(i).EndeArt = rSw(j).EndeArt
 roSw(i).ED = rSw(j).ED
 roSw(i).absPos = rSw(j).absPos
 roSw(i).AktZeit = rSw(j).AktZeit
 roSw(i).StByte = rSw(j).StByte
End Function ' roSwZuw

Public Function SwZUnt%(i&, j&)
 If roSw(i).FID <> rSw(j).FID Then GoSub unter
 If roSw(i).Pat_id <> rSw(j).Pat_id Then GoSub unter
 If roSw(i).Zeitpunkt <> rSw(j).Zeitpunkt Then GoSub unter
 If roSw(i).lR <> rSw(j).lR Then GoSub unter
 If roSw(i).vorET <> rSw(j).vorET Then GoSub unter
 If roSw(i).ET <> rSw(j).ET Then GoSub unter
 If roSw(i).efLR <> rSw(j).efLR Then GoSub unter
 If roSw(i).erLR <> rSw(j).erLR Then GoSub unter
 If roSw(i).kGT <> rSw(j).kGT Then GoSub unter
 If roSw(i).MB <> rSw(j).MB Then GoSub unter
 If roSw(i).EndeArt <> rSw(j).EndeArt Then GoSub unter
 If roSw(i).ED <> rSw(j).ED Then GoSub unter
 If roSw(i).absPos <> rSw(j).absPos Then GoSub unter
 If roSw(i).AktZeit <> rSw(j).AktZeit Then GoSub unter
 If roSw(i).StByte <> rSw(j).StByte Then GoSub unter
 Exit Function
unter:
 SwZUnt = SwZUnt + 1
 Return
End Function ' SwZUnt

Public Function swsLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_id
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(LR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) LR" & _
",COALESCE(vorET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) vorET,COALESCE(ET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ET,COALESCE(efLR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) efLR,COALESCE(erLR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) erLR" & _
",COALESCE(kGT - INTERVAL 0 DAY,CONVERT('18991230',DATE)) kGT,COALESCE(MB - INTERVAL 0 DAY,CONVERT('18991230',DATE)) MB,COALESCE(EndeArt,'') EndeArt,COALESCE(ED - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ED" & _
",COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `sws` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roSw(0)
 Else ' rs.EOF Then
  ReDim roSw(1)
  Do While Not rs.EOF
   akt = UBound(roSw)
   roSw(akt).FID = rs!FID
   roSw(akt).Pat_id = rs!Pat_id
   roSw(akt).Zeitpunkt = rs!Zeitpunkt
   roSw(akt).lR = rs!lR
   roSw(akt).vorET = rs!vorET
   roSw(akt).ET = rs!ET
   roSw(akt).efLR = rs!efLR
   roSw(akt).erLR = rs!erLR
   roSw(akt).kGT = rs!kGT
   roSw(akt).MB = rs!MB
   roSw(akt).EndeArt = doUmwfSQL(rs!EndeArt, lies.obMySQL, False)
   roSw(akt).ED = rs!ED
   roSw(akt).absPos = rs!absPos
   roSw(akt).AktZeit = rs!AktZeit
   roSw(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roSw(UBound(roSw) + 1)
  Loop ' While Not rs.EOF
 End If ' If rs.EOF
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' swsLaden

Function swsEinf()
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 If UBound(rSw) > 0 Then
  rbeg = 0
  For ri = 1 To UBound(rSw)
   If rSw(ri).Zeitpunkt >= qbeg Then
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For robeg = 0 To UBound(roSw)
    If roSw(robeg).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next robeg
   If robeg <= UBound(roFa) Then
    ReDim Preserve roSw(robeg - 1)
   End If ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roSw(robeg + UBound(rSw) - rbeg)
   For ri = rbeg To UBound(rSw)
    Call roSwZuw(robeg + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rSw = roSw
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' swsEinf

Public Function swsSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rSw) + 0 & " Sätze in `sws`"
' sql0 = " INSERT " & sqlignore &  "INTO `sws` (FID,Pat_ID,ZeitPunkt," & _
     "LR,vorET,ET,efLR,erLR,kGT,MB,EndeArt,ED,absPos," & _
     "AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `sws` (FID,Pat_ID,ZeitPunkt," & _
     "LR,vorET,ET,efLR,erLR,kGT,MB,EndeArt,ED,absPos," & _
     "AktZeit,StByte)        VALUES"))
 For i = 1 To UBound(rSw)
'  rSw(i).AktZeit = now()
  rSw(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rSw(i).FID, ",", rSw(i).Pat_id, ",", DatFor_k(rSw(i).Zeitpunkt), ",", DatFor_k(rSw(i).lR), ",", DatFor_k(rSw(i).vorET), ",", DatFor_k(rSw(i).ET), ",", DatFor_k(rSw(i).efLR), ",", DatFor_k( _
   rSw(i).erLR), ",", DatFor_k(rSw(i).kGT), ",", DatFor_k(rSw(i).MB), ",'", rSw(i).EndeArt, "',", DatFor_k(rSw(i).ED), ",", rSw(i).absPos, ",", DatFor_k(rSw(i).AktZeit), ",", _
   rSw(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rSw) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rSw) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   If Not AllePat Then
     sql = "DELETE FROM `sws` WHERE Pat_ID = " & CStr(rNa(0).Pat_id)
     Call myEFrag(sql)
   End If ' not AllePat
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rSw)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""swsSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(0)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rSw), i)
  If Len(rSw(k).EndeArt) > maxi(0) Then maxi(0) = Len(rSw(k).EndeArt)
 Next k
 If obTrans <> 0 Then If DBCn.Execute("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()").Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "sws", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "sws", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rSw), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rSw.EndeArt: '" & rSw(k).EndeArt & "' -> '" & Left$(rSw(k).EndeArt, maxL) & "'", True: rSw(k).EndeArt = Left$(rSw(k).EndeArt, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' swsSpeichern

Public Function laborxsaetzeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLs) + 0 & " Sätze in `laborxsaetze`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             VALUES"))
 For i = 0 To UBound(rLs)
'  rLs(i).AktZeit = now()
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLs(i).DatID, ",'", rLs(i).Satzart, "','", rLs(i).Satzlänge, "','", rLs(i).SatzlängeSchluss, "','", rLs(i).VersionSatzb, "','", rLs(i).Arztnr, "','", rLs(i).Arztname, "','", _
   rLs(i).StraßePraxis, "','", rLs(i).Arzt, "','", rLs(i).Lanr, "','", rLs(i).PLZPraxis, "','", rLs(i).OrtPraxis, "','", rLs(i).Labor, "','", rLs(i).StraßeLabor, "','", rLs(i).PLZLabor, "','", _
   rLs(i).OrtLabor, "','", rLs(i).KBVPrüfnr, "','", rLs(i).Zeichensatz, "','", rLs(i).Kundenarztnr, "','", rLs(i).Erstellungsdatum, "','", rLs(i).Gesamtlänge, "')")
  If SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLs) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLs)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxsaetzeSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(19)
 For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rLs), i)
  If Len(rLs(k).Satzart) > maxi(0) Then maxi(0) = Len(rLs(k).Satzart)
  If Len(rLs(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rLs(k).Satzlänge)
  If Len(rLs(k).SatzlängeSchluss) > maxi(2) Then maxi(2) = Len(rLs(k).SatzlängeSchluss)
  If Len(rLs(k).VersionSatzb) > maxi(3) Then maxi(3) = Len(rLs(k).VersionSatzb)
  If Len(rLs(k).Arztnr) > maxi(4) Then maxi(4) = Len(rLs(k).Arztnr)
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
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rLs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzart: '" & rLs(k).Satzart & "' -> '" & Left$(rLs(k).Satzart, maxL) & "'", True: rLs(k).Satzart = Left$(rLs(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzlänge: '" & rLs(k).Satzlänge & "' -> '" & Left$(rLs(k).Satzlänge, maxL) & "'", True: rLs(k).Satzlänge = Left$(rLs(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLs.SatzlängeSchluss: '" & rLs(k).SatzlängeSchluss & "' -> '" & Left$(rLs(k).SatzlängeSchluss, maxL) & "'", True: rLs(k).SatzlängeSchluss = Left$(rLs(k).SatzlängeSchluss, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLs.VersionSatzb: '" & rLs(k).VersionSatzb & "' -> '" & Left$(rLs(k).VersionSatzb, maxL) & "'", True: rLs(k).VersionSatzb = Left$(rLs(k).VersionSatzb, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztnr: '" & rLs(k).Arztnr & "' -> '" & Left$(rLs(k).Arztnr, maxL) & "'", True: rLs(k).Arztnr = Left$(rLs(k).Arztnr, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function laborxeingelSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLg) + 0 & " Sätze in `laborxeingel`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  VALUES"))
 For i = 1 To UBound(rLg)
'  rLg(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rLg(i).Pfad, "','", rLg(i).name, "',", DatFor_k(rLg(i).Zp), ",", CStr(-(rLg(i).Fertig <> 0)), ")")
  If SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLg) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLg)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxeingelSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLg), i)
  If Len(rLg(k).Pfad) > maxi(0) Then maxi(0) = Len(rLg(k).Pfad)
  If Len(rLg(k).name) > maxi(1) Then maxi(1) = Len(rLg(k).name)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLg), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL) & "'", True: rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).name & "' -> '" & Left$(rLg(k).name, maxL) & "'", True: rLg(k).name = Left$(rLg(k).name, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function laborxusSpeichern(SammelInsert%, BezfSp%, j&)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLu) + 0 & " Sätze in `laborxus`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN)     VALUES"))
 For i = j To j
'  rLu(i).AktZeit = now()
  If SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLu(i).DatID, ",", rLu(i).SatzID, ",'", rLu(i).Satzart, "','", rLu(i).Satzlänge, "','", rLu(i).Auftragsnummer, "','", rLu(i).Auftragsschlüssel, "',", DatFor_k(rLu(i).Eingang), ",'", _
   rLu(i).Berichtsdatum, "',", rLu(i).Pat_id, ",'", rLu(i).Nachname, "','", rLu(i).Vorname, "','", rLu(i).GebDat, "','", rLu(i).Titel, "','", rLu(i).NVorsatz, "','", rLu(i).BefArt, "','", _
   rLu(i).Abrechnungstyp, "','", rLu(i).GebüOrd, "','", rLu(i).Auftraggeber, "','", rLu(i).Patienteninformation, "','", rLu(i).geschlecht, "','", rLu(i).AuftrHinw, "','", rLu(i).Pat_idUrsp, "','", _
   rLu(i).Pat_idErwVNG, "','", rLu(i).Pat_idErwVN, "','", rLu(i).Pat_idErwG, "','", rLu(i).Pat_idErwGB, "','", rLu(i).Pat_idErwGL, "','", rLu(i).Pat_idLaborNeu, "',", DatFor_k(rLu(i).ZeitpunktLaborneu), ",", _
   rLu(i).ZdüP, ",", rLu(i).ZdiP, ",'", rLu(i).LWerte, "',", DatFor_k(rLu(i).verglichen), ",", rLu(i).AfN, ")")
  If SammelInsert <> 0 And i < j Then csql.Append ","
  If SammelInsert = 0 Or i = j Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = j
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxusSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(24)
 For k = IIf(SammelInsert <> 0, j, i) To IIf(SammelInsert <> 0, j, i)
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
  If Len(rLu(k).geschlecht) > maxi(15) Then maxi(15) = Len(rLu(k).geschlecht)
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
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, j, i) To IIf(SammelInsert <> 0, j, i)
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
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLu.Geschlecht: '" & rLu(k).geschlecht & "' -> '" & Left$(rLu(k).geschlecht, maxL) & "'", True: rLu(k).geschlecht = Left$(rLu(k).geschlecht, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function laborxbaktSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLo) + 0 & " Sätze in `laborxbakt`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  VALUES"))
 For i = 1 To UBound(rLo)
'  rLo(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLo(i).RefNr, ",'", rLo(i).Verf, "','", rLo(i).KuQu, "','", rLo(i).Quelle, "','", rLo(i).QSpez, "',", DatFor_k(rLo(i).AbnDat), ",'", rLo(i).Kommentar, "','", rLo(i).Erklärung, "','", _
   rLo(i).Keimzahl, "')")
  If SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLo) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLo)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxbaktSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(6)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLo), i)
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
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLo), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function laborxwertSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLw) + 0 & " Sätze in `laborxwert`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    VALUES"))
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
     myFrag rsdop, "SELECT 0 FROM `laborxwert` WHERE `RefNr` = " & rLw(i).RefNr & " AND `Abkü` = '" & rLw(i).Abkü & "' AND `Langname` = '" & rLw(i).Langname & "' AND `Quelle` = '" & _
   rLw(i).Quelle & "' AND `QSpez` = '" & rLw(i).QSpez & "' AND `AbnDat` = " & DatFor_k(rLw(i).AbnDat) & " AND `Wert` = '" & _
   rLw(i).Wert & "' AND `Einheit` = '" & rLw(i).Einheit & "' AND `Grenzwerti` = '" & rLw(i).Grenzwerti & "' AND `Kommentar` = '" & rLw(i).Kommentar & "' AND `Teststatus` = '" & _
   rLw(i).Teststatus & "' AND `Erklärung` = '" & rLw(i).Erklärung & "' AND `AuftrHinw` = '" & rLw(i).AuftrHinw & "' AND `nbid` = " & _
   rLw(i).nbid & ""
     If Not rsdop.EOF Then GoTo nexti
    Next j
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLw(i).RefNr, ",'", rLw(i).Abkü, "','", rLw(i).Langname, "','", rLw(i).Quelle, "','", rLw(i).QSpez, "',", DatFor_k(rLw(i).AbnDat), ",'", rLw(i).Wert, "','", rLw(i).Einheit, "','", _
   rLw(i).Grenzwerti, "','", rLw(i).Kommentar, "','", rLw(i).Teststatus, "','", rLw(i).Erklärung, "','", rLw(i).AuftrHinw, "',", rLw(i).nbid, ")")
  If SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLw) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLw)
nexti:
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxwertSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(10)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLw), i)
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
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLw), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function laborxleistSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLL) + 0 & " Sätze in `laborxleist`"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   VALUES"))
 For i = 1 To UBound(rLL)
'  rLL(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(", rLL(i).RefNr, ",'", rLL(i).Abkü, "','", rLL(i).Verf, "','", rLL(i).EBM, "','", rLL(i).goä, "','", rLL(i).Anzahl, "','", rLL(i).abrd, "')")
  If SammelInsert <> 0 And i < UBound(rLL) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLL) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLL)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxleistSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLL), i)
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
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLL), i)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function liuezSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAf&, pid$, m%, sfnr% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 T1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_id
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLi) + 0 & " Sätze in `liuez`"
' sql0 = " INSERT " & sqlignore &  "INTO `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel,zusatz,ursp," & _
     "aktzeit) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, " INTO `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel,zusatz,ursp," & _
     "aktzeit)               VALUES"))
 For i = 1 To UBound(rLi)
'  rLi(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('", rLi(i).name, "','", rLi(i).Vorname, "','", rLi(i).titelt, "','", rLi(i).fachgruppe, "','", rLi(i).strasse, "','", rLi(i).plz, "','", rLi(i).ort, "','", rLi(i).telefon, "','", _
   rLi(i).fax, "','", rLi(i).KVNr, "',", DatFor_k(rLi(i).aktdat), ",'", rLi(i).Überschrift, "','", rLi(i).DBNr, "','", rLi(i).BStelle, "','", rLi(i).anrede, "','", rLi(i).tel1, "','", _
   rLi(i).tel2, "','", rLi(i).tel3, "','", rLi(i).tel4, "','", rLi(i).fax1, "','", rLi(i).fax2, "','", rLi(i).fax3, "','", rLi(i).email, "','", rLi(i).zulg, "','", rLi(i).arzttyp, "','", _
   rLi(i).gemmit, "','", rLi(i).beme, "',", rLi(i).dmpt2, ",", rLi(i).dmpt1, ",'", rLi(i).geschlecht, "','", rLi(i).Titel, "','", rLi(i).zusatz, "','", rLi(i).ursp, "',", DatFor_k(rLi(i).AktZeit), ")")
  If SammelInsert <> 0 And i < UBound(rLi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLi) Then
'   Call myEFrag(csql.Value,rAf)', , adAsyncExecute) ' wegen unzureichender Fehlerverarbeitung wieder ausrangiert 19.8.23
sql:
   Call DBCn.Execute(csql.Value, rAf) ', , adAsyncExecute)
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then csql.m_Len = 0
   If Lese.obMySQL And obMitAlterTab Then
    Set rs = myEFrag("SHOW WARNINGS")
    If Not rs.BOF() Then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If ' rs!code = 1265 THEN
    End If ' lese.obmysql AND obmitAlterTab THEN
   End If ' lese.obmysql AND obmitAlterTab
  End If ' SammelInsert = 0 OR i = ubound(rLi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescr = Err.Description
ErrNumber = Err.Number
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""liuezSpeichern()"" bei Pat. " & rNa(0).Pat_id & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If
If ErrNumber = -2147217900 And InStrB(ErrDescr, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(29)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLi), i)
  If Len(rLi(k).name) > maxi(0) Then maxi(0) = Len(rLi(k).name)
  If Len(rLi(k).Vorname) > maxi(1) Then maxi(1) = Len(rLi(k).Vorname)
  If Len(rLi(k).titelt) > maxi(2) Then maxi(2) = Len(rLi(k).titelt)
  If Len(rLi(k).fachgruppe) > maxi(3) Then maxi(3) = Len(rLi(k).fachgruppe)
  If Len(rLi(k).strasse) > maxi(4) Then maxi(4) = Len(rLi(k).strasse)
  If Len(rLi(k).plz) > maxi(5) Then maxi(5) = Len(rLi(k).plz)
  If Len(rLi(k).ort) > maxi(6) Then maxi(6) = Len(rLi(k).ort)
  If Len(rLi(k).telefon) > maxi(7) Then maxi(7) = Len(rLi(k).telefon)
  If Len(rLi(k).fax) > maxi(8) Then maxi(8) = Len(rLi(k).fax)
  If Len(rLi(k).KVNr) > maxi(9) Then maxi(9) = Len(rLi(k).KVNr)
  If Len(rLi(k).Überschrift) > maxi(10) Then maxi(10) = Len(rLi(k).Überschrift)
  If Len(rLi(k).DBNr) > maxi(11) Then maxi(11) = Len(rLi(k).DBNr)
  If Len(rLi(k).BStelle) > maxi(12) Then maxi(12) = Len(rLi(k).BStelle)
  If Len(rLi(k).anrede) > maxi(13) Then maxi(13) = Len(rLi(k).anrede)
  If Len(rLi(k).tel1) > maxi(14) Then maxi(14) = Len(rLi(k).tel1)
  If Len(rLi(k).tel2) > maxi(15) Then maxi(15) = Len(rLi(k).tel2)
  If Len(rLi(k).tel3) > maxi(16) Then maxi(16) = Len(rLi(k).tel3)
  If Len(rLi(k).tel4) > maxi(17) Then maxi(17) = Len(rLi(k).tel4)
  If Len(rLi(k).fax1) > maxi(18) Then maxi(18) = Len(rLi(k).fax1)
  If Len(rLi(k).fax2) > maxi(19) Then maxi(19) = Len(rLi(k).fax2)
  If Len(rLi(k).fax3) > maxi(20) Then maxi(20) = Len(rLi(k).fax3)
  If Len(rLi(k).email) > maxi(21) Then maxi(21) = Len(rLi(k).email)
  If Len(rLi(k).zulg) > maxi(22) Then maxi(22) = Len(rLi(k).zulg)
  If Len(rLi(k).arzttyp) > maxi(23) Then maxi(23) = Len(rLi(k).arzttyp)
  If Len(rLi(k).gemmit) > maxi(24) Then maxi(24) = Len(rLi(k).gemmit)
  If Len(rLi(k).beme) > maxi(25) Then maxi(25) = Len(rLi(k).beme)
  If Len(rLi(k).geschlecht) > maxi(26) Then maxi(26) = Len(rLi(k).geschlecht)
  If Len(rLi(k).Titel) > maxi(27) Then maxi(27) = Len(rLi(k).Titel)
  If Len(rLi(k).zusatz) > maxi(28) Then maxi(28) = Len(rLi(k).zusatz)
  If Len(rLi(k).ursp) > maxi(29) Then maxi(29) = Len(rLi(k).ursp)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "liuez", Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "liuez", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLi.name: '" & rLi(k).name & "' -> '" & Left$(rLi(k).name, maxL) & "'", True: rLi(k).name = Left$(rLi(k).name, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLi.vorname: '" & rLi(k).Vorname & "' -> '" & Left$(rLi(k).Vorname, maxL) & "'", True: rLi(k).Vorname = Left$(rLi(k).Vorname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLi.titelt: '" & rLi(k).titelt & "' -> '" & Left$(rLi(k).titelt, maxL) & "'", True: rLi(k).titelt = Left$(rLi(k).titelt, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLi.fachgruppe: '" & rLi(k).fachgruppe & "' -> '" & Left$(rLi(k).fachgruppe, maxL) & "'", True: rLi(k).fachgruppe = Left$(rLi(k).fachgruppe, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLi.strasse: '" & rLi(k).strasse & "' -> '" & Left$(rLi(k).strasse, maxL) & "'", True: rLi(k).strasse = Left$(rLi(k).strasse, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLi.plz: '" & rLi(k).plz & "' -> '" & Left$(rLi(k).plz, maxL) & "'", True: rLi(k).plz = Left$(rLi(k).plz, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLi.ort: '" & rLi(k).ort & "' -> '" & Left$(rLi(k).ort, maxL) & "'", True: rLi(k).ort = Left$(rLi(k).ort, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLi.telefon: '" & rLi(k).telefon & "' -> '" & Left$(rLi(k).telefon, maxL) & "'", True: rLi(k).telefon = Left$(rLi(k).telefon, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax: '" & rLi(k).fax & "' -> '" & Left$(rLi(k).fax, maxL) & "'", True: rLi(k).fax = Left$(rLi(k).fax, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLi.kvnr: '" & rLi(k).KVNr & "' -> '" & Left$(rLi(k).KVNr, maxL) & "'", True: rLi(k).KVNr = Left$(rLi(k).KVNr, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLi.überschrift: '" & rLi(k).Überschrift & "' -> '" & Left$(rLi(k).Überschrift, maxL) & "'", True: rLi(k).Überschrift = Left$(rLi(k).Überschrift, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLi.dbnr: '" & rLi(k).DBNr & "' -> '" & Left$(rLi(k).DBNr, maxL) & "'", True: rLi(k).DBNr = Left$(rLi(k).DBNr, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLi.bstelle: '" & rLi(k).BStelle & "' -> '" & Left$(rLi(k).BStelle, maxL) & "'", True: rLi(k).BStelle = Left$(rLi(k).BStelle, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLi.anrede: '" & rLi(k).anrede & "' -> '" & Left$(rLi(k).anrede, maxL) & "'", True: rLi(k).anrede = Left$(rLi(k).anrede, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel1: '" & rLi(k).tel1 & "' -> '" & Left$(rLi(k).tel1, maxL) & "'", True: rLi(k).tel1 = Left$(rLi(k).tel1, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel2: '" & rLi(k).tel2 & "' -> '" & Left$(rLi(k).tel2, maxL) & "'", True: rLi(k).tel2 = Left$(rLi(k).tel2, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel3: '" & rLi(k).tel3 & "' -> '" & Left$(rLi(k).tel3, maxL) & "'", True: rLi(k).tel3 = Left$(rLi(k).tel3, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel4: '" & rLi(k).tel4 & "' -> '" & Left$(rLi(k).tel4, maxL) & "'", True: rLi(k).tel4 = Left$(rLi(k).tel4, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax1: '" & rLi(k).fax1 & "' -> '" & Left$(rLi(k).fax1, maxL) & "'", True: rLi(k).fax1 = Left$(rLi(k).fax1, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax2: '" & rLi(k).fax2 & "' -> '" & Left$(rLi(k).fax2, maxL) & "'", True: rLi(k).fax2 = Left$(rLi(k).fax2, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax3: '" & rLi(k).fax3 & "' -> '" & Left$(rLi(k).fax3, maxL) & "'", True: rLi(k).fax3 = Left$(rLi(k).fax3, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLi.email: '" & rLi(k).email & "' -> '" & Left$(rLi(k).email, maxL) & "'", True: rLi(k).email = Left$(rLi(k).email, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLi.zulg: '" & rLi(k).zulg & "' -> '" & Left$(rLi(k).zulg, maxL) & "'", True: rLi(k).zulg = Left$(rLi(k).zulg, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLi.arzttyp: '" & rLi(k).arzttyp & "' -> '" & Left$(rLi(k).arzttyp, maxL) & "'", True: rLi(k).arzttyp = Left$(rLi(k).arzttyp, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLi.gemmit: '" & rLi(k).gemmit & "' -> '" & Left$(rLi(k).gemmit, maxL) & "'", True: rLi(k).gemmit = Left$(rLi(k).gemmit, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rLi.beme: '" & rLi(k).beme & "' -> '" & Left$(rLi(k).beme, maxL) & "'", True: rLi(k).beme = Left$(rLi(k).beme, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rLi.geschlecht: '" & rLi(k).geschlecht & "' -> '" & Left$(rLi(k).geschlecht, maxL) & "'", True: rLi(k).geschlecht = Left$(rLi(k).geschlecht, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rLi.titel: '" & rLi(k).Titel & "' -> '" & Left$(rLi(k).Titel, maxL) & "'", True: rLi(k).Titel = Left$(rLi(k).Titel, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rLi.zusatz: '" & rLi(k).zusatz & "' -> '" & Left$(rLi(k).zusatz, maxL) & "'", True: rLi(k).zusatz = Left$(rLi(k).zusatz, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rLi.ursp: '" & rLi(k).ursp & "' -> '" & Left$(rLi(k).ursp, maxL) & "'", True: rLi(k).ursp = Left$(rLi(k).ursp, maxL)
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
 Resume
End If ' ErrNumber = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in liuezSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' liuezSpeichern

Public Function tuLaden()
 Call namenLaden
 Call faelleLaden
'  IF not lese.obmysql THEN
'   IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
'    Call DBCn.BeginTrans: obTrans = 1
'  END IF ' not lese.obmysql
   wechsTrans
 Call auLaden
 Call briefeLaden
 Call diagnosenLaden
 Call dokumenteLaden
 Call eintraegeLaden
'  IF not lese.obmysql THEN
'   IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
'    Call DBCn.BeginTrans: obTrans = 1
'  END IF ' not lese.obmysql
   wechsTrans
 Call forminhkopfLaden
 Call kheinweisLaden
 Call lbanforderungenLaden
 Call laborneuLaden
 Call leistungenLaden
 Call medplanLaden
 Call rezepteintraegeLaden
 Call rrLaden
 Call dmpreiheLaden
 Call desktopLaden
 Call usdmLaden
 Call fussLaden
 Call ulcusLaden
 Call vkgdLaden
 Call swsLaden
End Function ' tuLaden

Public Function tuSpeichern(frm As Lese, SI%, BfS%) ' frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern
 Dim rAf&, altsi$, altsam%
 altsi = sqlIGNORE
 sqlIGNORE = ""
 altsam = SI
 sqlIGNORE = ""
 On Error GoTo fehler
     DBCn.Execute ("delete FROM faelle WHERE pat_id=67789")
     Call DBCn.Execute("INSERT  INTO `faelle` (Pat_ID,Quartal,Nachname,Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn,FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK,KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb,PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf,dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)" & _
      "VALUES(67789,'32023','Will','Harald',5,'','64414','0','1','2006-07-01 00:00:00','Audi BKK',0,0,0,'2023-07-01 00:00:00','2023-09-30 00:00:00','1899-12-30 00:00:00','','1899-12-30 00:00:00','00','1','2023-07-18 00:00:00','8534160','','','5000100','AUDI BKK                    71',0,'1','07','','','','',0,0,0,0,'','','365389403','641954100','0000209','','0000209','365389403','FA Innere und Allgemeinmedizin (Hausarzt)','0000209','Wolff','Dr. med.','Ina','',6658,'Teilnahme HzV- Vertrag ','Diabetes mellitus (E14.90/G)','','','',0,24,'',0,'','0','','','','1899-12-30 00:00:00','','','','','','','','','','','','1899-12-30 00:00:00','','','2023-08-26 12:00:41','2023-07-18 00:00:00','','2023-07-01 00:00:00','2023-09-30 00:00:00','','',16557,0,1,'','',0,'1899-12-30 00:00:00','','','','','','','','','','','','','','','','')")
     If DBCn.Execute("select last_insert_id()").Fields(0) = 0 Then
      Stop
     End If
 Call namenSpeichern(SI, BfS)
 SI = 0
 DBCn.Execute "START TRANSACTION"
     DBCn.Execute ("delete FROM faelle WHERE pat_id=67789")
     Call DBCn.Execute("INSERT  INTO `faelle` (Pat_ID,Quartal,Nachname,Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn,FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK,KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb,PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf,dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)" & _
      "VALUES(67789,'32023','Will','Harald',5,'','64414','0','1','2006-07-01 00:00:00','Audi BKK',0,0,0,'2023-07-01 00:00:00','2023-09-30 00:00:00','1899-12-30 00:00:00','','1899-12-30 00:00:00','00','1','2023-07-18 00:00:00','8534160','','','5000100','AUDI BKK                    71',0,'1','07','','','','',0,0,0,0,'','','365389403','641954100','0000209','','0000209','365389403','FA Innere und Allgemeinmedizin (Hausarzt)','0000209','Wolff','Dr. med.','Ina','',6658,'Teilnahme HzV- Vertrag ','Diabetes mellitus (E14.90/G)','','','',0,24,'',0,'','0','','','','1899-12-30 00:00:00','','','','','','','','','','','','1899-12-30 00:00:00','','','2023-08-26 12:00:41','2023-07-18 00:00:00','','2023-07-01 00:00:00','2023-09-30 00:00:00','','',16557,0,1,'','',0,'1899-12-30 00:00:00','','','','','','','','','','','','','','','','')")
     If DBCn.Execute("select last_insert_id()").Fields(0) = 0 Then
      Stop
     End If
 Call faelleSpeichern(SI, BfS)
     DBCn.Execute ("delete FROM faelle WHERE pat_id=67789")
     Call DBCn.Execute("INSERT  INTO `faelle` (Pat_ID,Quartal,Nachname,Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn,FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK,KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb,PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf,dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)" & _
      "VALUES(67789,'32023','Will','Harald',5,'','64414','0','1','2006-07-01 00:00:00','Audi BKK',0,0,0,'2023-07-01 00:00:00','2023-09-30 00:00:00','1899-12-30 00:00:00','','1899-12-30 00:00:00','00','1','2023-07-18 00:00:00','8534160','','','5000100','AUDI BKK                    71',0,'1','07','','','','',0,0,0,0,'','','365389403','641954100','0000209','','0000209','365389403','FA Innere und Allgemeinmedizin (Hausarzt)','0000209','Wolff','Dr. med.','Ina','',6658,'Teilnahme HzV- Vertrag ','Diabetes mellitus (E14.90/G)','','','',0,24,'',0,'','0','','','','1899-12-30 00:00:00','','','','','','','','','','','','1899-12-30 00:00:00','','','2023-08-26 12:00:41','2023-07-18 00:00:00','','2023-07-01 00:00:00','2023-09-30 00:00:00','','',16557,0,1,'','',0,'1899-12-30 00:00:00','','','','','','','','','','','','','','','','')")
     If DBCn.Execute("select last_insert_id()").Fields(0) = 0 Then
      Stop
     End If
 SI = altsam
' DBCn.Execute "COMMIT"
 ComTrans
'   IF not lese.obmysql THEN
'    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
'    Call DBCn.BeginTrans: obTrans = 1
'   END IF ' not lese.obmysql
    wechsTrans
 Call auSpeichern(SI, BfS)
 Call briefeSpeichern(SI, BfS)
 Call diagnosenSpeichern(SI, BfS)
 Call dokumenteSpeichern(SI, BfS)
 Call eintraegeSpeichern(SI, BfS)
 Call forminhaltform_abkSpeichern(SI, BfS)
'   IF not lese.obmysql THEN
'    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
'    Call DBCn.BeginTrans: obTrans = 1
'   END IF ' not lese.obmysql
    wechsTrans
 sqlIGNORE = " IGNORE "
 Call formulareSpeichern(SI, BfS)
 sqlIGNORE = ""
 Call forminhkopfSpeichern(SI, BfS)
 Call forminhfeldSpeichern(SI, BfS)
 Call kheinweisSpeichern(SI, BfS)
 Call lbanforderungenSpeichern(SI, BfS)
 Call laborneuSpeichern(SI, BfS)
 Call leistungenSpeichern(SI, BfS)
 Call medplanSpeichern(SI, BfS)
 Call rezepteintraegeSpeichern(SI, BfS)
 Call rrSpeichern(SI, BfS)
 Call kvnrueSpeichern(SI, BfS)
 Call unbekannte_kennungenSpeichern(SI, BfS)
 Call dmpreiheSpeichern(SI, BfS)
 Call desktopSpeichern(SI, BfS)
 Call usdmSpeichern(SI, BfS)
 Call fussSpeichern(SI, BfS)
 Call ulcusSpeichern(SI, BfS)
 Call vkgdSpeichern(SI, BfS)
 Call swsSpeichern(SI, BfS)
 Call myEFrag("UPDATE `namen` SET aktZeit = " & DatFor_k(rNa(0).AktZeit) & " WHERE pat_id = " & rNa(0).Pat_id, rAf)
 If rAf <> 1 Then
  frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_id & " " & rNa(0).Nachname & " " & rNa(0).Vorname, True
 End If ' rAf <> 0
 sqlIGNORE = altsi
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 ErrDescr = Err.Description
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
  Resume
 End If
 Select Case MsgBox("FNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in tuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' tuSpeichern

