Option Explicit
Public obFor%
Dim sql$, T1!, T2!, rs As Adodb.Recordset, maxL%

Public type namen
 Pat_ID AS long 'Pat_ID int '3000
 lfdnr AS long 'lfdnr int 'laufende Patientennummer
 NVorsatz AS string 'NVorsatz varchar '3100
 Nachname AS string 'Nachname varchar '3101
 Vorname AS string 'Vorname varchar '3102
 GebDat AS date 'GebDat datetime '3103
 f3004 AS string 'f3004 varchar '3004 unbek.Feld, bis jetzt 0 oder 2
 f3006 AS string 'f3006 varchar '3006 unbek.Feld, bis jetzt 5.1.0 oder 5.2.0
 Straße AS string 'Straße varchar '3107
 KVKStatus AS string 'KVKStatus varchar '3108
 Hausnr AS string 'Hausnr varchar '3109
 Geschlecht AS string 'Geschlecht varchar '3110
 Plz AS string 'Plz varchar '3112
 Ort AS string 'Ort varchar '3113
 Lkz AS string 'Lkz varchar '3114 Länderkennzeichen
 Anschrzus AS string 'Anschrzus varchar '3115 Anschriftenzusatz
 NVors AS string 'NVors varchar '3120 Namensvorsatz
 PFPlz AS string 'PFPlz varchar '3121 Postfach-Postleitzahl
 PFOrt AS string 'PFOrt varchar '3122 Postfach-Ort
 PFNr AS string 'PFNr varchar '3123 Postfach-Nr.
 f3124 AS string 'f3124 varchar '3124 unbek.Feld, bisher immer leer
 AnschrZus_2 AS string 'AnschrZus_2 varchar '3215 Anschriftenzusatz, aufgeteiltes Feld
 Postfach_2 AS string 'Postfach_2 varchar '3216 Postfach, aufgeteiltes Feld
 LK_2 AS string 'LK_2 varchar '3217 Ländercode vermutlich Herkunftsland, aufgeteiltes Feld
 Postfach AS string 'Postfach varchar '3216
 Beruf AS string 'Beruf varchar '3620 Beruf
 Weggeldzone AS string 'Weggeldzone varchar '3631 (1) Weggeldzone mit Z
 WeggzZahl AS double 'WeggzZahl decimal '3631 (2) Weggeldzone, Zahl in Feld 2
 AufnDat AS date 'AufnDat datetime '3610
 kAufDat AS date 'kAufDat date 'korrigiertes Aufnahmedatum: if(min(fanf)>min(bhfb),(min(fanf),min(bhfb))
 LANR AS string 'LANR varchar '3635, LANR, interne Zuordnung Arzt bei GP, zuvor IntZoGP
 BStNr AS string 'BStNr varchar '3536 Betriebsstättennummer
 Titel AS string 'Titel varchar '3104
 Versichertennummer AS string 'Versichertennummer varchar '3105
 PrivatTel AS string 'PrivatTel varchar '3629
 KVNr AS string 'KVNr varchar '3630 Hausarzt
 KVNr2 AS string 'KVNr2 varchar '3630 Hausarzt (2.Eintrag)
 KVNr3 AS string 'KVNr3 varchar '3630 Hausarzt (3.Eintrag)
 KVNr4 AS string 'KVNr4 varchar '3630 Hausarzt (4.Eintrag)
 PrivatTel_2 AS string 'PrivatTel_2 varchar '3629
 PrivatFax AS string 'PrivatFax varchar '3629
 DienstTel AS string 'DienstTel varchar '3629
 PrivatMobil AS string 'PrivatMobil varchar '3629
 Email AS string 'Email varchar 'Email
 Arbeitgeber AS string 'Arbeitgeber varchar '3625
 AnAllgda AS integer 'AnAllgda bit 'Anamnese allgemein da
 An1da AS integer 'An1da bit 'Anamnese S.1 da
 An2da AS integer 'An2da bit 'Anamnese S.2 da
 Checkda AS integer 'Checkda bit 'Checkliste da
 DMTypaD AS string 'DMTypaD varchar 'aus Diagnosen
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 StByteA AS long 'StByteA int 'Ordnungsnummer der Datenübertragung, Beginn der Übertragung
 Cave AS string 'Cave varchar '3654
 Notiz AS string 'Notiz varchar '3634 DMP-Infos: DMP hier <datum>, DMP HA <datum>, DMP nein <datum>, DMP ausgeschrieben <datum>
 obChk AS string 'obChk varchar 'obChroniker (Feld 3800)
 dmpklass AS long 'dmpklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier, 4 = DMP ausgeschrieben
 dmpbeg AS date 'dmpbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass AS long 'dmpkhkklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg AS date 'dmpkhkbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass AS long 'dmpcopdklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg AS date 'dmpcopdbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpabklass AS long 'dmpabklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpabbeg AS date 'dmpabbeg date 'Datum der aktuellen DMP-Klassifikation
 dakab AS date 'dakab date 'DAK-Einverständnis-Datum
 HzV AS long 'HzV int '1 = HzV-Eintrag im Notizfeld
 HzVbeg AS date 'HzVbeg date 'HzV ab
 DS AS long 'DS int '1 = Datenschutzerklärung laut Notizfeld da
 DSbeg AS date 'DSbeg date 'DS ab
 getHA0 AS long 'getHA0 int 'KVNr aus getHausarzt -> Üw(12,0)
 fnHA0 AS string 'fnHA0 varchar 'Funktion aus getHausarzt -> Üw(10,0)
 getHA1 AS long 'getHA1 int 'KVNr aus getHausarzt -> Üw(12,1)
 fnHA1 AS string 'fnHA1 varchar 'Funktion aus getHausarzt -> Üw(10,1)
 getHA2 AS long 'getHA2 int 'KVNr aus getHausarzt -> Üw(12,2)
 fnHA2 AS string 'fnHA2 varchar 'Funktion aus getHausarzt -> Üw(10,2)
 zubenach AS string 'zubenach varchar '3633
 Verwandt AS string 'Verwandt varchar '3632
 Sprache AS string 'Sprache varchar '3628
 lAktTM AS date 'lAktTM datetime 'letzte Aktualisierung in Turbomed
 Mitarbeiter AS long 'Mitarbeiter int 'ob Pat. Mitarbeiter ist
end type

Public type faelle
 FID AS long 'FID int '
 Pat_ID AS long 'Pat_ID int '3000 Pat_id
 Quartal AS string 'Quartal varchar '4101 Quartal der Ausstellung
 Nachname AS string 'Nachname varchar '3101 Nachname
 Vorname AS string 'Vorname varchar '3102 Vorname
 lfdnr AS long 'lfdnr int 'laufende Fallnummer
 TMFNr AS string 'TMFNr varchar '4144 Fallnummer in Turbomed
 VKNr AS string 'VKNr varchar '4104 VK-Nr.
 f4131 AS string 'f4131 varchar '4131 unbekanntes  Feld, bis jetzt immer 0
 f4132 AS string 'f4132 varchar '4132 unbekanntes  Feld, 0-6
 VschBeg AS date 'VschBeg datetime '4133 Versichternschutzbeginn
 KKasse_2 AS string 'KKasse_2 varchar '4134 Krankenkasse aus Fall
 FaktPers AS single 'FaktPers float '4136 Faktor persönlich
 FaktTechn AS single 'FaktTechn float '4137 Faktor technisch
 FaktLabor AS single 'FaktLabor float '4138 Faktor Labor
 BhFB AS date 'BhFB datetime '4150 Behandlungsfall: Beginn
 BhFE1 AS date 'BhFE1 datetime '4151 Behandlungsfall: Ende (Musterwoman) / Wohl Ende des Behandlungsfallbeginnquartals
 BhFE2 AS date 'BhFE2 datetime '4152 Behandlungsfall: Ende (Musterwoman),bei offenem Fall 00000000, sonst z.B. 30092006 für 3/06
 f4202 AS string 'f4202 varchar '4202 Unfall, Unfallfolgen nach 4152
 ausgst AS date 'ausgst datetime '4102 ('ausgestellt am')
 KtrAbrB AS string 'KtrAbrB varchar '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr AS string 'AbrAr varchar '4107, Abrechnungsart (1 = Primärkassen)
 lVorl AS date 'lVorl datetime '4109, letzte Vorlage
 IK AS string 'IK varchar '4111 Krankenkassennummer (IK)
 KVKs AS string 'KVKs varchar '4112 Versichertenstatus VK
 KVKserg AS string 'KVKserg varchar '4113 Ost/West-Status VK
 Status AS string 'Status varchar 'Feld Status in verschiedenen Formularen
 Kasse AS string 'Kasse varchar '6299 Kasse (aus Formularen)
 KID AS long 'KID int 'Bezug auf kassenliste.id
 GebOr AS string 'GebOr varchar '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb AS string 'AbrGb varchar '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis AS string 'PersKreis varchar '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz AS string 'SKtZusatz varchar '4124 SKT-Zusatzangaben
 letzteRegel AS string 'letzteRegel varchar '4206, letzter Tag der Regel
 ÜwText AS string 'ÜwText longtext '4209: Auftrags- / erläuternder Text zur Überweisung
 f4210 AS byte 'f4210 tinyint '4210, Ankreuzfeld LSR
 AkfHAH AS byte 'AkfHAH tinyint '4211 Ankreuzfeld HAH
 AkfAB0 AS byte 'AkfAB0 tinyint '4212 Ankreuzfeld AB0.RH
 AkfAK AS byte 'AkfAK tinyint '4213 Ankreuzfeld AK
 statNuller AS string 'statNuller varchar '4216, nu bei Musterfrau 16 Nuller
 ÜbwV AS string 'ÜbwV varchar '4218, überwiesen von Arztnummer
 ÜbWVLANR AS string 'ÜbWVLANR varchar '4218(1) überwiesen von LANR
 ÜbWVBSNR AS string 'ÜbWVBSNR varchar '4218(2) überwiesen von BSNR
 ÜbWVKVNR AS string 'ÜbWVKVNR varchar '4218(3) überwiesen von KVNR
 AndÜw AS string 'AndÜw varchar '4219, anderer Überweiser
 Übwr AS string 'Übwr varchar 'resultierender Überweiser (BSNR): 4218 oder 4219, je nachdem, was befüllt
 ÜbwLANR AS string 'ÜbwLANR varchar '4242 LANR des Überweisers
 ÜWZiel AS string 'ÜWZiel varchar '4220 Überweisung an
 ÜWNNr AS string 'ÜWNNr varchar '4231(4): KV-Nummer des Überweisers
 ÜWNaN AS string 'ÜWNaN varchar '4231(3): Nachname des Überweisers
 ÜWTit AS string 'ÜWTit varchar '4231(3): Titel des Überweisers
 ÜWVor AS string 'ÜWVor varchar '4231(2): Vorname des Überweisers
 ÜWVsw AS string 'ÜWVsw varchar '4231(2b): Vorsatzwort des Überweisers
 üwvid AS long 'üwvid int '4247 Bezug auf ueberwvon
 Auftrag AS string 'Auftrag varchar '4205 Auftrag bei Überweisung
 Verdacht AS string 'Verdacht varchar '4207 Verdacht bei Überweisung

 Befund AS string 'Befund varchar '4208 Befund bei Überweisung
 statKlasse AS string 'statKlasse varchar '4236 Klasse bei Behandlung
 f4237 AS string 'f4237 varchar '4237 Krankenhausname
 statBehTage AS long 'statBehTage int '4238 Krankenhausaufenthalt
 SchGr AS double 'SchGr decimal '4239, Schein(unter)gruppe
 Weiterbeh AS string 'Weiterbeh varchar '4243, Weiterbehandelnder
 f4266 AS single 'f4266 float '4266 unbekanntes Datumsfeld, bisher nur bei Musterwoman
 PGeb AS string 'PGeb varchar '4401, Praxisgebühr
 PGebErg AS string 'PGebErg varchar '4402, Array
 Mahnfrist AS string 'Mahnfrist varchar '4403, Mahnfrist bis
 Unfallort AS string 'Unfallort varchar '4505 Unfallort
 BeschAls AS string 'BeschAls varchar '4506 Beschäftigt als
 BeschSeit AS date 'BeschSeit datetime '4506 Beschäftigt seit
 Unfallbetrieb AS string 'Unfallbetrieb varchar '4509 Unfallbetrieb
 f4570 AS string 'f4570 varchar 'f4570 unbekanntes Feld, nur einamal 0
 GOÄKatNr AS string 'GOÄKatNr varchar '4580 (1): Katalog-Nummer
 GOÄKatName AS string 'GOÄKatName varchar '4580 (2): Privat-Abrechnungskatalog
 abrArzt AS string 'abrArzt varchar '4585 abrechnender Arzt
 privVers AS string 'privVers varchar '4586 private Versicherung
 AdNam AS string 'AdNam varchar '4602(1) Name Rechnungsanschrift
 AdStr AS string 'AdStr varchar '4602(2) Straße Rechnungsanschrift
 AdPlz AS string 'AdPlz varchar '4602(3) PLZ Rechnungsanschrift
 AdOrt AS string 'AdOrt varchar '4602(4) Ort Rechnungsanschrift
 ÜwBG AS string 'ÜwBG varchar 'f4603 Überweiser BG
 BhFE AS date 'BhFE datetime '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 AS string 's8000 varchar '8000, Satzidentifikation
 s8100 AS string 's8100 varchar '8100 Satzlänge
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 Fanf AS date 'Fanf datetime 'Fallanfang
 altQuart AS string 'altQuart varchar '
 QAnf AS date 'QAnf datetime 'Quartalsanfang
 QEnd AS date 'QEnd datetime 'Quartalsende
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 LANRid AS long 'LANRid int 'Bezug auf lanrpraxis.id
 f4108 AS string 'f4108 varchar '4108
 BGFallNr AS string 'BGFallNr varchar '3603 BG-Fall-Nummer	
 lGewicht AS double 'lGewicht decimal 'letztes Gewicht in kg
 vorET AS date 'vorET date 'voraussichtlicher Entbindungstermin
 dmpVertret AS string 'dmpVertret varchar 'DMP: Vertretung
 dmpArztw AS string 'dmpArztw varchar 'DMP: Arztwechsel
 dmpHypos AS string 'dmpHypos varchar 'DMP: Zahl der schweren Hypos im letzten Quartal
 dmpKhsA AS string 'dmpKhsA varchar 'DMP: Zahl der Khs-Aufenthalte wg.Diabetes im letzten Quartal
 dmpDMSchulEmpf AS string 'dmpDMSchulEmpf varchar 'DMP: Schulung D.m. empfohlen
 dmpDMSchulWahrg AS string 'dmpDMSchulWahrg varchar 'DMP: Schulung D.m. wahrgenommen (bei letzter Doku)
 dmpHypertSchulEmpf AS string 'dmpHypertSchulEmpf varchar 'DMP-Schulung zur Hypertonie empfohlen
 dmpHypertSchulWahrg AS string 'dmpHypertSchulWahrg varchar 'DMP: Schulung Hypertonie wahrgenommen (bei letzter Doku)
 dmpKKTabakEmpf AS string 'dmpKKTabakEmpf varchar 'DMP-Schulung zur Tabakentwöhnung empfohlen
 dmpKKErnEmpf AS string 'dmpKKErnEmpf varchar 'DMP-Schulung zur Ernährung empfohlen
 dmpKKkTrainEmpf AS string 'dmpKKkTrainEmpf varchar 'DMP-Schulung zum körperlichen Training empfohlen
 dmpHbA1cZiel AS string 'dmpHbA1cZiel varchar 'DMP: HbA1c-Ziel erreicht
 dmpUewFuss AS string 'dmpUewFuss varchar 'DMP: Überweisung Fußeinrichtung veranlasst
 dmpEinwDM AS string 'dmpEinwDM varchar 'DMP: Einweisung wegen D.m. veranlasst
 dmphalbj AS string 'dmphalbj varchar 'j=DMP nur halbjährlich dokumentieren
 dmpMA AS string 'dmpMA varchar 'DMP: Mitarbeiter, der Makro eingegeben hat
end type

Public type au
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 Beginn AS string 'Beginn varchar '6285 1. Hälfte
 Ende AS string 'Ende varchar '6285 2. Hälfte
 ICDs AS string 'ICDs varchar '6286
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type briefe
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '
 ZeitPunkt AS date 'ZeitPunkt datetime '
 Pfad AS string 'Pfad varchar '
 Art AS string 'Art varchar '
 Name AS string 'Name varchar '
 autor AS string 'autor varchar 'Autor
 Quelldatum AS date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 Typ AS string 'Typ varchar '
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe AS long 'DokGroe int 'Größe der Datei
 DokAenD AS date 'DokAenD datetime 'Dokument-letzte Änderung
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 ID AS long 'ID int '
end type

Public type diagnosen
 ID1 AS long 'ID1 int '
 FID AS long 'FID int 'Fall-Bezug
 Pat_id AS long 'Pat_id int 'Bezug auf Anamneseblattt
 DiagDatum AS date 'DiagDatum datetime '
 DiagSicherheit AS string 'DiagSicherheit varchar '6003
 DiagText AS string 'DiagText longtext '
 DiagSeite AS string 'DiagSeite varchar '6004
 DiagAttr AS string 'DiagAttr varchar '6006 Diagnosenattribut (optionale Erläuterung)
 ICD AS string 'ICD varchar '
 obDauer AS byte 'obDauer tinyint 'ob Dauerdiagnose
 intBemerk AS string 'intBemerk varchar '6009 interne Bemerkung
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 AusnBegr AS string 'AusnBegr varchar '6008 Ausnahmebegründung
 f6010 AS byte 'f6010 tinyint 'Falsch=dd-Eintrag, Wahr=bdd-Eintragurspr.: 6010 Diagnose gelöscht
 obKasse AS byte 'obKasse tinyint 'ob nach Kodierrichtlinien an Kasse zu übermitteln
 lKasse AS date 'lKasse datetime 'wann zuletzt Kassenübermittlung nach Kodierrichtlinien eingetragen (bd/bdd, f6010 Wahr)
 f6011 AS string 'f6011 varchar '6011 8.12.10: bisher nur """"""""TM#?""""""""
end type

Public type dokumente
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '
 ZeitPunkt AS date 'ZeitPunkt datetime '
 DokPfad AS string 'DokPfad varchar '
 DokArt AS string 'DokArt varchar '
 DokName AS string 'DokName varchar '
 Quelldatum AS date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe AS long 'DokGroe int 'Dokument-Größe
 DokAenD AS date 'DokAenD datetime 'Dokument-letzte Änderung
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type eintraege
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '
 Art AS string 'Art varchar '6330
 Inhalt AS string 'Inhalt varchar '8480
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnr. der Datenübertragung
 id AS long 'id int '
 inhNum AS double 'inhNum double 'Inhalt numerisch
end type

Public type forminhaltform_abk
 Form_AbkVW AS long 'Form_AbkVW int '
 Form_Abk AS string 'Form_Abk varchar '
end type

Public type formulare
 FormID AS long 'FormID int '
 Form_Abk AS string 'Form_Abk varchar '
 FormBez AS string 'FormBez longtext '
 FormVorl AS string 'FormVorl varchar '
 AktZeit AS date 'AktZeit datetime 'Zeitpunkt der Aktualisierung
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type forminhkopf
 FoID AS long 'FoID int '
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '
 Form_ID AS long 'Form_ID int '
 ZeitPunkt AS date 'ZeitPunkt datetime '
 AbsPos AS long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 Satzart AS string 'Satzart varchar '8000
 Satzlänge AS string 'Satzlänge varchar '8100
 LANRid AS long 'LANRid int 'Bezug auf lanrpraxis.id
end type

Public type forminhfeld
 FoID AS long 'FoID int '
 Nr AS integer 'Nr smallint '
 FeldNr AS integer 'FeldNr smallint '
 FeldVW AS long 'FeldVW int '
 Feld AS string
 FeldInhVW AS long 'FeldInhVW int '
 FeldInh AS string
end type

Public type kheinweis
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '
 Ziel AS string 'Ziel varchar '6291
 Diagnose AS string 'Diagnose longtext '6230
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type lbanforderungen
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 AnfText AS string 'AnfText longtext '6280
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Statusbyte
end type

Public type laborneu
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '
 FertigStGrad AS string 'FertigStGrad varchar '8401
 Abkü AS string 'Abkü varchar '8410
 LangtextVW AS long 'LangtextVW int '8411
 Langtext AS string
 Wert AS string 'Wert varchar '8420
 Einheit AS string 'Einheit varchar '8421
 KommentarVW AS long 'KommentarVW int '8480
 Kommentar AS string
 AbsPos AS long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 Refnr AS long 'Refnr int 'Bezug auf LaborXUS
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 ID AS long 'ID int '
end type

Public type leistungen
 id AS long 'id int 'eindeutige ID, hinzugefügt 26.3.11
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '5000 + 6201
 Leistung AS string 'Leistung varchar '5001 Leistungsziffer
 f5002 AS string 'f5002 varchar '5002 Art der Untersuchung
 f5005 AS string 'f5005 varchar '5005 Anzahl
 f5006 AS string 'f5006 varchar '5006 um Uhrzeit
 f5009 AS string 'f5009 varchar '5009 freier Begründungstext
 Med AS string 'Med varchar '5010 Medikament
 f5015 AS string 'f5015 varchar '5015 Organ
 f5016 AS string 'f5016 varchar '5016 Name des Arztes (Briefempfänger)
 f5021 AS string 'f5021 varchar '5021 Datum letzte Krebsvorsorge
 f5026 AS string 'f5026 varchar '5026 Entlassungsdatum
 Faktor AS string 'Faktor varchar '5062 Multiplikator für GOÄ-Rechnung
 f5098 AS string 'f5098 varchar '5098 0000000000
 Charge AS string 'Charge varchar '9999 Charge
 LANR AS string 'LANR varchar '5099 LANR
 letzVorg AS date 'letzVorg datetime '5101 letzter Vorgang
 Ausn AS string 'Ausn varchar '3677 Ausnahme/Begründung für abweichendes Geschlecht
 Beme AS string 'Beme varchar '         Bemerkung
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 LANRid AS long 'LANRid int 'Bezug auf lanrpraxis.id
 Sachkbez AS string 'Sachkbez varchar '5011 Sachkostenbezeichnung
 Sachkct AS long 'Sachkct int '5012 Sach-/Materialkosten in ct
 Zone AS string 'Zone varchar '5018 Zone bei Besuchen
end type

Public type medplan
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 MPNr AS long 'MPNr int 'Ordnungsziffer für Medikamentenplan
 ZeitPunkt AS date 'ZeitPunkt datetime 'Zeitpunkt, der Speicherung im Turbomed
 Datum AS date 'Datum datetime 'Zeitpunkt aus dem Kopf des Medikamentenplans
 Medikament AS string 'Medikament varchar '
 MedAnfang AS string 'MedAnfang varchar 'Referenz auf medarten.Medikament
 Wirkstoff AS string 'Wirkstoff varchar 'Stärke aus BMP
 PZN AS long 'PZN int 'Pharmazentralnummer
 FeldNr AS integer 'FeldNr smallint '
 mo AS string 'mo varchar '
 mi AS string 'mi varchar '
 nm AS string 'nm varchar '
 ab AS string 'ab varchar '
 zn AS string 'zn varchar '
 bBed AS integer 'bBed bit '
 Bemerkung AS string 'Bemerkung longtext '
 Grund AS string 'Grund varchar 'Grund aus BMP
 Stärke AS string 'Stärke varchar 'Stärke aus BMP
 Einheit AS string 'Einheit varchar 'Stärke aus BMP
 Form AS string 'Form varchar 'Stärke aus BMP
 AbsPos AS long 'AbsPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 ergaenzt AS integer 'ergaenzt bit 'PZN ergaenzt von pznbdt (bdtnachw)
end type

Public type rezepteintraege
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 Rezept AS string 'Rezept varchar '6210, 3652(1), 6218(1)
 RKlnm AS string 'RKlnm varchar 'Anfang des Rezeptklassennamens bei Langrezepten (HeilHilfsmittel, LangRezeptEintrag)
 Rezeptklasse AS string 'Rezeptklasse varchar '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Rezklkurz AS string 'Rezklkurz varchar 'letztes Split-Feld, z.B. 'rp' oder 'lar'
 Rezkllang AS string 'Rezkllang varchar 'erstes Split-Feld, z.B. 'Sprechstundenbedarf': Langrezept, Sprechstundenbedarf, Heilmittel
 kbez AS string 'kbez varchar 'vorletztes Split-Feld, z.B. 'DTronAnthra'
 Medikament AS string 'Medikament varchar '3652(2), 6218(4)
 auti AS long 'auti int 'aut-idem-Eintrag im Rz. (1=kein Ausschluß, 0=Ausschluß)
 anzl AS long 'anzl int 'Anzahl
 PZN AS string 'PZN varchar '6210(2), 6218(3)
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Statusbyte
 LANRid AS long 'LANRid int 'Bezug auf lanrpraxis.id
 id AS long 'id int '
end type

Public type rr
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 RR AS string 'RR longtext '6230
 Puls AS long 'Puls int 'Puls
 RRsyst AS integer 'RRsyst smallint '
 RRdiast AS integer 'RRdiast smallint '
 RRzahl AS integer 'RRzahl smallint '
 Quelle AS string 'Quelle varchar 'Informationsquelle
 Bemerkung AS string 'Bemerkung varchar 'Bemerkung
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type kvnrue
 lfdnr AS long 'lfdnr int '
 Pat_ID AS long 'Pat_ID int '
 KVNr AS string 'KVNr varchar '
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type unbekannte_kennungen
 Kennung AS string 'Kennung varchar '
 absPos AS long 'absPos int '
 StByte AS long 'StByte int '
 Pat_id AS long 'Pat_id int 'zugehöriger Patient für spätere Ermittlungen
 Inhalt AS string 'Inhalt varchar 'Inhalt Zeile zum Wiederauffinden
end type

Public type dmpreihe
 Abk AS string 'Abk varchar 'Abkürzung der DMP-Art
 Art AS string 'Art varchar 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum AS date 'KarteiDatum date 'Datum des Karteikarteneintrags der Dokumentation
 exportiert AS date 'exportiert datetime 'Datum des Exports
 DokuDatum AS date 'DokuDatum datetime 'Datum der Dokumentation
 obvoll AS integer 'obvoll bit 'ob vollständig
 ok AS integer 'ok bit 'ob """"""""ok""""""""
 ausgedruckt AS integer 'ausgedruckt bit 'ob """"""""ausgedruckt""""""""
 NachName AS string 'NachName varchar '
 VorName AS string 'VorName varchar '
 GebDat AS date 'GebDat date '
 Pat_id AS long 'Pat_id int '
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
 AktZeit AS date 'AktZeit datetime 'Aktualisierungzeit
 lanrid AS long 'lanrid int 'Bezug auf lanrpraxis.id
 Zusatzdaten AS string 'Zusatzdaten varchar 'Zusatzdaten
end type

Public type desktop
 id AS long 'id int 'Primärschlüssel
 IDS AS string 'IDS varchar 'id=
 Pat_ID AS long 'Pat_ID int '3000
 erstZP AS date 'erstZP datetime 'erstellungsZeitpunkt
 exoL AS string 'exoL varchar 'executeonLoad
 hideT AS byte 'hideT tinyint 'hideTitel
 iconPath AS string 'iconPath varchar 'iconPath
 noteBkColor AS long 'noteBkColor int 'noteBkColor
 noteFgColor AS long 'noteFgColor int 'noteFgColor
 positionBottom AS long 'positionBottom int 'positionBottom
 positionLeft AS long 'positionLeft int 'positionLeft
 positionRight AS long 'positionRight int 'positionRight
 positionTop AS long 'positionTop int 'positionTop
 showAsNote AS byte 'showAsNote tinyint 'showAsNote
 syncInfoList AS string 'syncInfoList varchar 'syncInfoList
 titel AS string 'titel varchar 'titel
 toolTipText AS string 'toolTipText varchar 'toolTipText
 verankert AS byte 'verankert tinyint 'verankert
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type usdm
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '
 Art AS string 'Art varchar '6330
 Spritzst AS string 'Spritzst text 'Spritzstellen/Liphyper./~atr.
 Fußbef_re AS string 'Fußbef_re text 'Fußbefund rechts
 Fußbef_li AS string 'Fußbef_li text 'Fußbefund links
 Hyperk_re AS string 'Hyperk_re text 'Hyperkeratosen rechts
 Hyperk_li AS string 'Hyperk_li text 'Hyperkeratosen links
 Ulcera_re AS string 'Ulcera_re text 'Ulcera rechts
 Ulcera_li AS string 'Ulcera_li text 'Ulcera links
 Kraft_Zh_re AS string 'Kraft_Zh_re text 'Kraft Zehenheber re
 Kraft_Zh_li AS string 'Kraft_Zh_li text 'Kraft Zehenheber li
 Kraft_Zb_re AS string 'Kraft_Zb_re text 'Kraft Zehenbeuger re
 Kraft_Zb_li AS string 'Kraft_Zb_li text 'Kraft Zehenbeuger li
 Kraft_Knie_re AS string 'Kraft_Knie_re text 'Kraft Knie re
 Kraft_Knie_li AS string 'Kraft_Knie_li text 'Kraft Knie li
 ASR_re AS string 'ASR_re text 'Achillessehnenreflex rechts
 ASR_li AS string 'ASR_li text 'Achillessehnenreflex links
 PSR_re AS string 'PSR_re text 'Patellarsehnenreflex rechts
 PSR_li AS string 'PSR_li text 'Patellarsehnenreflex links
 Oberfl_re AS string 'Oberfl_re text 'Oberflächensensibilität rechts
 Oberfl_li AS string 'Oberfl_li text 'Oberflächensensibilität links
 MF_re AS string 'MF_re text 'Monofilament rechts
 MF_li AS string 'MF_li text 'Monofilament links
 KW_re AS string 'KW_re text 'Kalt-warm rechts
 KW_li AS string 'KW_li text 'Kalt-warm links
 Vibr_IK_re AS string 'Vibr_IK_re text 'Vibration Innenknöchel rechts
 Vibr_IK_li AS string 'Vibr_IK_li text 'Vibration Innenknöchel links
 Vibr_GZ_re AS string 'Vibr_GZ_re text 'Vibration Großzehe rechts
 Vibr_GZ_li AS string 'Vibr_GZ_li text 'Vibration Großzehe links
 PulsL_re AS string 'PulsL_re text 'Puls Leiste rechts
 PulsL_li AS string 'PulsL_li text 'Puls Leiste links
 PulsKK_re AS string 'PulsKK_re text 'Puls Kniekehle rechts
 PulsKK_li AS string 'PulsKK_li text 'Puls Kniekehle links
 PulsAtp_re AS string 'PulsAtp_re text 'Puls Arteria tibialis posterior rechts
 PulsAtp_li AS string 'PulsAtp_li text 'Puls Arteria tibialis posterior links
 PulsAdp_re AS string 'PulsAdp_re text 'Puls Arteria dorsalis pedis rechts
 PulsAdp_li AS string 'PulsAdp_li text 'Puls Arteria dorsalis pedis links
 Mitarbeiter AS string 'Mitarbeiter text 'Mitarbeiter
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnr. der Datenübertragung
 id AS long 'id int '
end type

Public type fuss
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '
 Art AS string 'Art varchar '6330
 Fußdeform AS string 'Fußdeform varchar 'Fußdeformität re/li
 Hyper_mEin AS string 'Hyper_mEin varchar 'Hyperkeratose mit Einblutung re/li
 Weiteres AS string 'Weiteres varchar 'Hyperkeratose ohne Einblutung, Stauungszeichen, Ekzem, Nagelmykose usw.
 Zn_Ulcus AS string 'Zn_Ulcus varchar 'Z.n. Ulcus re/li
 Zn_Amput AS string 'Zn_Amput varchar 'Z.n. Amputation re/li
 Fuß_ang AS string 'Fuß_ang varchar 'Füße genau angeschaut
 Ulcera AS string 'Ulcera varchar 'Ulcera re/li
 Wundinfektion AS string 'Wundinfektion varchar 'Wundinfektion re/li
 nae_US AS string 'nae_US varchar 'nächste Untersuchung
 Mitarbeiter AS string 'Mitarbeiter varchar 'Mitarbeiter
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte AS long 'StByte int 'Ordnungsnr. der Datenübertragung
 id AS long 'id int '
end type

Public type ulcus
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 Lokalisation AS string 'Lokalisation varchar 'Lokalisation des Ulcus am Fuß, ohne Seite
 Seite AS string 'Seite varchar 're oder li
 Größe AS string 'Größe varchar 'Größe in mm
 Beläge AS string 'Beläge varchar 'Beläge auf Ulcus
 Exsudat AS string 'Exsudat varchar 'Exsudat auf Ulcus
 Geruch AS string 'Geruch varchar 'Geruch aus 3 cm Entfernung
 Wundrand AS string 'Wundrand varchar 'Wundrand
 Wundumgebung AS string 'Wundumgebung varchar 'Wundumgebung
 Temperatur AS string 'Temperatur varchar 'Temperatur im Vergleich zu gesunden Stellen
 Fotodoku AS string 'Fotodoku varchar 'Fotodoku durchgeführt (ja/nein)
 Wundversorgung AS string 'Wundversorgung varchar 'Wundversorgung
 Mitarbeiter AS string 'Mitarbeiter varchar 'Mitarbeiterkürzel
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type vkgd
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 Wohlempfinden AS string 'Wohlempfinden varchar 'Wohlempfinden
 Saettigung AS string 'Saettigung varchar 'Sättigung
 Zielwerterreichung AS string 'Zielwerterreichung varchar 'Zielwerterreichung
 Ketonkörper AS string 'Ketonkörper varchar 'Ketonkörper
 Gynaekologenbefund AS string 'Gynaekologenbefund varchar 'aktueller Gynäkologenbefund
 Gewichtsentwicklung AS string 'Gewichtsentwicklung varchar 'Gewichtsentwicklung der Schwangeren
 HbA1c AS string 'HbA1c varchar 'HbA1c (monatlich)+ ggf. TSH (dreimonatlich) abgenommen
 Bewegung AS string 'Bewegung varchar 'Bewegung: Art
 Minuten AS string 'Minuten varchar 'Minuten pro Woche
 Blutdruck AS string 'Blutdruck varchar 'Blutdruck
 Puls AS string 'Puls varchar 'Puls
 Mitarbeiter AS string 'Mitarbeiter varchar 'Mitarbeiterkürzel
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type sws
 FID AS long 'FID int 'Fall-Bezug
 Pat_ID AS long 'Pat_ID int '3000
 ZeitPunkt AS date 'ZeitPunkt datetime '6200 + 6201
 LR AS date 'LR date 'anzeigeText:>LR: 
 vorET AS date 'vorET date 'voraussichtlicher ET: 
 ET AS date 'ET date 'Entbindung ... Datum: 
 efLR AS date 'efLR date 'EffektiveLetzteRegel
 erLR AS date 'erLR date 'ErfassteLetzteRegel
 kGT AS date 'kGT date 'KorrigierterGeburtstermin
 MB AS date 'MB date 'MutterschutzBeginn
 EndeArt AS string 'EndeArt varchar 'EndeArt
 ED AS date 'ED date 'EndeDatum
 absPos AS long 'absPos int 'Zeile in BDT-Datei
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
end type

Public type laborxsaetze
 SatzID AS long 'SatzID int 'zum Bezug für LaborUS
 DatID AS long 'DatID int 'Bezug zu LaborEingelesen
 Satzart AS string 'Satzart varchar '8000 Satzart (Turbomed)
 Satzlänge AS string 'Satzlänge varchar '8100 Satzlänge (Turbomed)
 SatzlängeSchluss AS string 'SatzlängeSchluss varchar '8100 Satzlänge (Turbomed), nach 8221 in Feld 8000
 VersionSatzb AS string 'VersionSatzb varchar '9212 Version der Satzbeschreibung (Turbomed)
 Arztnr AS string 'Arztnr varchar '201 Arztnummer (Turbomed)
 Arztname AS string 'Arztname varchar '203 Arztname (Turbomed)
 StraßePraxis AS string 'StraßePraxis varchar '205 Straße der Praxis (Turbomed)
 Arzt AS string 'Arzt varchar ' 211 Ausführender Arzt
 LANR AS string 'LANR varchar ' 212 LANR
 PLZPraxis AS string 'PLZPraxis varchar '215 PLZ der Praxis (Turbomed)
 OrtPraxis AS string 'OrtPraxis varchar '216 Ort der Praxis (Turbomed)
 Labor AS string 'Labor varchar '8320 Labor
 StraßeLabor AS string 'StraßeLabor varchar '8321 Straße der Laboradresse (Turbomed)
 PLZLabor AS string 'PLZLabor varchar '8322 PLZ der Laboradresse (Turbomed)
 OrtLabor AS string 'OrtLabor varchar '8323 Ort der Laboradresse (Turbomed)
 KBVPrüfnr AS string 'KBVPrüfnr varchar '101 KBV-Prüfnummer (Turbomed)
 Zeichensatz AS string 'Zeichensatz varchar '9106 verwendeter Zeichensatz (Turbomed)
 Kundenarztnr AS string 'Kundenarztnr varchar '8312 Kundenarztnummer (Turbomed)
 Erstellungsdatum AS string 'Erstellungsdatum varchar '9103 Erstellungsdatum (Turbomed)
 Gesamtlänge AS string 'Gesamtlänge varchar '9202 Gesamtlänge des Datenpaketes (Turbomed)
end type

Public type laborxeingel
 DatID AS long 'DatID int 'Bezug auf LaborEingelesen
 Pfad AS string 'Pfad varchar 'Pfadname
 Name AS string 'Name varchar 'Name der eingelesenen Labordatei ohne Endung
 Zp AS date 'Zp datetime 'Einlesezeitpunkt
 fertig AS integer 'fertig bit 'ob Einlesen fertig
end type

Public type laborxus
 RefNr AS long 'RefNr int 'Bezug auf LaborWert
 DatID AS long 'DatID int 'Bezug auf LaborEingelesen
 SatzID AS long 'SatzID int 'Bezug auf LaborXSätze
 Satzart AS string 'Satzart varchar '8000 Satzart (Turbomed)
 Satzlänge AS string 'Satzlänge varchar '8100 Satzlänge (Turbomed)
 Auftragsnummer AS string 'Auftragsnummer varchar '8310 Anforderungsident (Turbomed)
 Auftragsschlüssel AS string 'Auftragsschlüssel varchar '8311 Anforderungsnr d Labors (Turbomed)
 Eingang AS date 'Eingang datetime '8301 Eingangsdatum in Datumsform
 Berichtsdatum AS string 'Berichtsdatum varchar '8302 Berichtsdatum
 Pat_id AS long 'Pat_id int '
 Nachname AS string 'Nachname varchar '3101
 Vorname AS string 'Vorname varchar '3102
 GebDat AS string 'GebDat varchar '3103
 Titel AS string 'Titel varchar '3104
 NVorsatz AS string 'NVorsatz varchar '3100
 BefArt AS string 'BefArt varchar '8401 Befundart (Turbomed) / Fertigstellungsgrad (""""""""E""""""""=Endbefund, """"""""T"""""""" = Teilbefund)
 Abrechnungstyp AS string 'Abrechnungstyp varchar '8609 Abrechnungstyp (Kasse/Privat/X=anderer RE/Einsender) (Turbomed)
 GebüOrd AS string 'GebüOrd varchar '8403 Gebührenordnung (Turbomed)
 Auftraggeber AS string 'Auftraggeber varchar '8615 Auftraggeber (LANR)
 Patienteninformation AS string 'Patienteninformation varchar '8405 Patienteninformation (Turbomed)
 Geschlecht AS string 'Geschlecht varchar '8407 Geschlecht (Turbomed)
 AuftrHinw AS string 'AuftrHinw varchar '8490 Auftragsbezogene Hinweise (Turbomed)
 Pat_idUrsp AS string 'Pat_idUrsp varchar 'Ursprung der Pat_id: E = erwogene Pat_id s.u., L = vergleich mit über Turbomed eingelesenem Labor
 Pat_idErwVNG AS string 'Pat_idErwVNG varchar 'erwogene Pat_id mit gleichem Vornamen, Nachnamen und Geburtstag
 Pat_idErwVN AS string 'Pat_idErwVN varchar 'erwogene Pat_id mit gleichem Vornamen und Nachnamen
 Pat_idErwG AS string 'Pat_idErwG varchar 'erwogene Pat_id mit gleichem Geburtstag
 Pat_idErwGB AS string 'Pat_idErwGB varchar 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passender Behandlung
 Pat_idErwGL AS string 'Pat_idErwGL varchar 'erwogene Pat_id mit gleichem Geburtstag und zeitlich passendem Labor
 Pat_idLaborNeu AS string 'Pat_idLaborNeu varchar 'Pat_ids von in Laborneu zuordnbaren Patienten
 ZeitpunktLaborneu AS date 'ZeitpunktLaborneu datetime 'Zeitpunkt der Untersuchung, die in Laborneu zugeordnet wurde
 ZdüP AS integer 'ZdüP smallint 'Zahl der verglichenen Parameter
 ZdiP AS long 'ZdiP int 'Zahl der infragekommenden Patienten
 LWerte AS string 'LWerte longtext 'Laborwerte, die zur Zuordnung geführt haben
 verglichen AS date 'verglichen datetime 'Datum, zu dem Datensatz zuletzt verglichen wurde
 AfN AS integer 'AfN smallint 'Affected Number: Zahl der zugehörigen Datensätze in Laborneu
end type

Public type laborxbakt
 RefNr AS long 'RefNr int '
 Verf AS string 'Verf varchar '
 KuQu AS string 'KuQu varchar '8428 Probenmaterial-Ident (Turbomed)
 Quelle AS string 'Quelle varchar '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez AS string 'QSpez varchar '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat AS date 'AbnDat datetime '8432 Abnahmedatum (Turbomed)
 Kommentar AS string 'Kommentar longtext '8480 Ergebnistest (Turbomed)
 Erklärung AS string 'Erklärung longtext '
 Keimzahl AS string 'Keimzahl varchar '
end type

Public type laborxwert
 RefNr AS long 'RefNr int 'Bezug auf LaborUS
 Abkü AS string 'Abkü varchar '8410 Test-Ident  (Turbomed)
 Langname AS string 'Langname varchar '8411 Testbezeichnung (Turbomed)
 Quelle AS string 'Quelle varchar '8430 Probenmaterial-Bezeichnung (Turbomed)
 QSpez AS string 'QSpez varchar '8431 Probenmaterial-Spezifikation (Turbomed)
 AbnDat AS date 'AbnDat datetime '8432 Abnahmedatum (Turbomed)
 Wert AS string 'Wert varchar '8420 Ergebniswert (Turbomed)
 Einheit AS string 'Einheit varchar '8421 Einheit (Turbomed)
 Grenzwerti AS string 'Grenzwerti varchar '8422 Grenzwertindikator (Turbomed)
 Kommentar AS string 'Kommentar varchar '8480 Ergebnistext (Turbomed)
 Teststatus AS string 'Teststatus varchar '8418 Teststatus (Turbomed)
 Erklärung AS string 'Erklärung varchar '8470 Testbezogene Hinweise (Turbomed)
 AuftrHinw AS string 'AuftrHinw varchar '8490 Auftragsbezogene Hinweise (Turbomed)
 nbid AS long 'nbid int 'Bezug zu laborxplab.id
end type

Public type laborxleist
 RefNr AS long 'RefNr int 'Bezug auf LaborUS
 Abkü AS string 'Abkü varchar '8410 Test-Ident (Turbomed)
 Verf AS string 'Verf varchar '8434
 EBM AS string 'EBM varchar '5001 GNR (Turbomed)
 goä AS string 'goä varchar '8406
 Anzahl AS string 'Anzahl varchar '5005
 abrd AS string 'abrd varchar '8614 Abrechnung durch: 1 = Labor, 2 = Einweiser
end type

Public type liuez
 name AS string 'name varchar '
 vorname AS string 'vorname varchar '
 titelt AS string 'titelt varchar '
 fachgruppe AS string 'fachgruppe varchar '
 strasse AS string 'strasse varchar '
 plz AS string 'plz varchar '
 ort AS string 'ort varchar '
 telefon AS string 'telefon varchar '
 fax AS string 'fax varchar '
 kvnr AS string 'kvnr varchar '
 aktdat AS date 'aktdat date '
 id AS long 'id int '
 überschrift AS string 'überschrift varchar '
 dbnr AS string 'dbnr varchar '
 bstelle AS string 'bstelle varchar '
 anrede AS string 'anrede varchar '
 tel1 AS string 'tel1 varchar '
 tel2 AS string 'tel2 varchar '
 tel3 AS string 'tel3 varchar '
 tel4 AS string 'tel4 varchar '
 fax1 AS string 'fax1 varchar '
 fax2 AS string 'fax2 varchar '
 fax3 AS string 'fax3 varchar '
 email AS string 'email varchar '
 zulg AS string 'zulg varchar '
 arzttyp AS string 'arzttyp varchar '
 gemmit AS string 'gemmit longtext '
 beme AS string 'beme longtext '
 dmpt2 AS byte 'dmpt2 tinyint '
 dmpt1 AS byte 'dmpt1 tinyint '
 geschlecht AS string 'geschlecht varchar '
 titel AS string 'titel varchar '
 zusatz AS string 'zusatz varchar '
 ursp AS string 'ursp varchar '
 aktzeit AS date 'aktzeit datetime '
end type

Public type anamnesebogen
 Prim AS long 'Prim int 'Primärschlüssel
 Pat_id AS long 'Pat_id int '
 Nachname AS string 'Nachname varchar '-
 Vorname AS string 'Vorname varchar '
 NVorsatz AS string 'NVorsatz varchar '
 Titel AS string 'Titel varchar '
 Anrede AS string 'Anrede varchar '
 GebDat AS date 'GebDat datetime ', geb.
 Tkz AS byte 'Tkz tinyint 'Tod-Kennzeichen
 Versicherungsart AS string 'Versicherungsart varchar '
 Diabetestyp AS string 'Diabetestyp varchar '^Diabetes Typ
 Diabetes_seit AS string 'Diabetes_seit text '<seit
 Tabletten_seit AS string 'Tabletten_seit text ', Tabletten seit
 Insulin_seit AS string 'Insulin_seit text ', Insulin seit
 Grund_für_Vorstellung AS string 'Grund_für_Vorstellung text '^:
 Familienanamnese AS string 'Familienanamnese text '^:
 Größe AS double 'Größe double '^:
 Gewicht AS double 'Gewicht double ',:
 bmi AS double 'bmi decimal '
 Tendenz AS string 'Tendenz varchar '<, Tendenz
 DiabetesMedikament_1 AS string 'DiabetesMedikament_1 text '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge AS string 'DiabetesMedikament_1_Menge text '<
 DiabetesMedikament_2 AS string 'DiabetesMedikament_2 text '<,
 DiabetesMedikament_2_Menge AS string 'DiabetesMedikament_2_Menge text '<
 DiabetesMedikament_3 AS string 'DiabetesMedikament_3 text '<,
 DiabetesMedikament_3_Menge AS string 'DiabetesMedikament_3_Menge text '<
 DiabetesMedikament_4 AS string 'DiabetesMedikament_4 text '<,
 DiabetesMedikament_4_Menge AS string 'DiabetesMedikament_4_Menge text '<,
 Insulinpumpe AS byte 'Insulinpumpe tinyint '^:
 Insulinpumpe_seit AS string 'Insulinpumpe_seit text '<seit
 Insulinpumpe_Marke AS string 'Insulinpumpe_Marke text '<, Marke:
 Broteinheiten_gesamt AS string 'Broteinheiten_gesamt text '^Broteinheiten:gesamt
 Broteinheiten_früh AS string 'Broteinheiten_früh text '<, früh
 Broteinheiten_ZM_früh AS string 'Broteinheiten_ZM_früh text '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags AS string 'Broteinheiten_mittags text '<, mittags
 Broteinheiten_nachmittags AS string 'Broteinheiten_nachmittags text '<, nachmittags
 Broteinheiten_abends AS string 'Broteinheiten_abends text '<, abends
 Broteinheiten_nachts AS string 'Broteinheiten_nachts text '<, nachts
 Essenszeit_früh AS string 'Essenszeit_früh text '^Essenszeiten:früh
 Essenszeit_vormittags AS string 'Essenszeit_vormittags text '<, vormittags
 Essenszeit_mittags AS string 'Essenszeit_mittags text '<, mittags
 Essenszeit_nachmittags AS string 'Essenszeit_nachmittags text '<, nachmittags
 Essenszeit_abends AS string 'Essenszeit_abends text '<, abends
 Essenszeit_spät AS string 'Essenszeit_spät text '<, spät
 Spritz_Eß_Abstand_früh AS string 'Spritz_Eß_Abstand_früh text '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags AS string 'Spritz_Eß_Abstand_mittags text '<, mittags
 Spritz_Eß_Abstand_abends AS string 'Spritz_Eß_Abstand_abends text '<, abends
 Spritzstelle_früh AS string 'Spritzstelle_früh text '^Spritzstellen:früh
 Spritzstelle_mittags AS string 'Spritzstelle_mittags text '<, mittags
 Spritzstelle_abends AS string 'Spritzstelle_abends text '<, abends
 Spritzstelle_nachts AS string 'Spritzstelle_nachts text '<, nachts
 Ernährung AS string 'Ernährung text '^:
 Jahr_letzte_Diabetesschulung AS string 'Jahr_letzte_Diabetesschulung varchar '^Letzte Diabetesschulung:
 Ort_Schulung AS string 'Ort_Schulung text '<in
 letztes_HbA1c AS string 'letztes_HbA1c text '^Letztes HbA1c:
 gemessen_am AS string 'gemessen_am text '<, gemessen
 vorherige_Werte AS string 'vorherige_Werte text '<, vorher:
 subcutane_Zuckermessung AS string 'subcutane_Zuckermessung text '^:
 CGM_seit AS string 'CGM_seit text '<
 BZMessungen_selbst AS string 'BZMessungen_selbst text '^Blutzuckermessung:Selbstmessung?
 Gerät AS string 'Gerät text '<:
 BZMessungen_pW AS string 'BZMessungen_pW text '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE AS string 'BZMessungen_pW_ndE text '<, davon nach dem Essen:
 BZMessungen_p_W_nachts AS string 'BZMessungen_p_W_nachts text '<, nachts:
 Aufschreiben AS string 'Aufschreiben text '<, Dokumentation:
 BZWerte_v_d_Essen AS string 'BZWerte_v_d_Essen text '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen AS string 'BZWerte_n_d_Essen text '<, nach dem Essen:
 UZ_Tageszeit AS string 'UZ_Tageszeit text '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM AS string 'Unterzucker_pM text '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig AS string 'UZ_rechtzeitig text '<, rechtzeitig bemerkt:
 Schwere_Uzu AS string 'Schwere_Uzu text 'Schwere Unterzucker
 Fremde_Hilfe_pa AS string 'Fremde_Hilfe_pa text '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa AS string 'Bewußtlos_pa text '<, bewußtlos deshalb:
 Keto AS string 'Keto text '^Bisher Ketoazidosen mit Krankenhauseinweisung:
 Keto_pa AS string 'Keto_pa text '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM AS string 'BZgr300_pM text ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck AS string 'Bluthochdruck text '^Bluthochdruck:
 BHD_seit AS string 'BHD_seit text '<seit:
 BHD_beh_mit AS string 'BHD_beh_mit text '<, behandelt mit:
 Blutdruckwerte AS string 'Blutdruckwerte text '^Blutdruckwerte:
 BDselbst AS string 'BDselbst text '^Blutdruckselbstmessung:
 Schwanger AS string 'Schwanger text '^Aktuelle Schwangerschaft:
 Schwanger_seit AS string 'Schwanger_seit text '<, seit:
 Augensp_zuletzt AS string 'Augensp_zuletzt text '^Letzte Augenspiegelung:
 Augensp_Befund AS string 'Augensp_Befund text '<, Befund:
 Netzhaut_gelasert AS string 'Netzhaut_gelasert text ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar AS string 'Sehminderung_unbehebbar text ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden AS string 'Diabet_Nierenschaden text '^Diabetischer Nierenschaden:
 Albumin_zuletzt AS string 'Albumin_zuletzt text ', letztes Albumin:
 erhöht AS string 'erhöht text '<, Befund:
 Kreatinin AS string 'Kreatinin text ',:
 Dialyse AS byte 'Dialyse tinyint ',:
 Dialyse_seit AS string 'Dialyse_seit text '<seit
 andere_Nierenerkrankung AS string 'andere_Nierenerkrankung text ', andere Nierenerkrankung:
 Herzkrankheit AS string 'Herzkrankheit text '^Herzkrankheit:
 Angina_pectoris AS string 'Angina_pectoris text ',:
 Herzinfarkt AS string 'Herzinfarkt text ',:
 Herzinfarkt_wann AS string 'Herzinfarkt_wann text '<, wann:
 PTCA_oder_Stent AS string 'PTCA_oder_Stent text ',:
 Bypass_kardial AS byte 'Bypass_kardial tinyint ',:
 Bypass_wann AS string 'Bypass_wann text '<, wann:
 Herzschwäche AS string 'Herzschwäche text ',:
 Herzkrankheit_Beschreibung AS string 'Herzkrankheit_Beschreibung text ', Beschreibung:
 Hirndurchblutungsstörung AS string 'Hirndurchblutungsstörung text '^:
 Schlaganfall AS string 'Schlaganfall text ',:
 Beindurchblutungsstörung AS string 'Beindurchblutungsstörung text '^:
 Schaufensterkrankheit AS string 'Schaufensterkrankheit text ',:
 Bypaß_peripher AS byte 'Bypaß_peripher tinyint ',:
 Geschwür AS string 'Geschwür text ',:
 Amputation AS string 'Amputation text ',:
 pAVK_Beschreibung AS string 'pAVK_Beschreibung text ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen AS string 'Ameisenlaufen text '^:
 Ameisen_Ausmaß AS string 'Ameisen_Ausmaß text '<, Ausmaß:
 Druckstellen AS string 'Druckstellen text ',:
 Verformungen AS string 'Verformungen text ',:
 Verformungen_Beschreibung AS string 'Verformungen_Beschreibung text '<Beschreibung:
 Fußpflege AS string 'Fußpflege text '^:
 Podologie AS string 'Podologie text ',:
 Einlagen AS string 'Einlagen text ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen AS string 'Neue_Fußkomplikationen text '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen AS string 'Entleerungsstörungen_Magen text '^:
 Entleerungsstörungen_Harnblase AS string 'Entleerungsstörungen_Harnblase text ',:
 Schwindel_Aufstehen AS string 'Schwindel_Aufstehen text ',:
 Folgeerkrankungen_Haut AS string 'Folgeerkrankungen_Haut text '^:
 Bewegungseinschränkungen AS string 'Bewegungseinschränkungen text ',:
 Sexualstörung AS string 'Sexualstörung text '^:
 Sexualstörung_seit AS string 'Sexualstörung_seit text '<seit
 Weitere_Anamnese AS string 'Weitere_Anamnese text '^:
 Tabak AS string 'Tabak text '^Tabak:
 tabakex AS string 'tabakex text '<, früher:
 tabakbis AS string 'tabakbis text '<, bis:
 tabakakt AS string 'tabakakt text '<, aktuell:
 tabakmenge AS string 'tabakmenge text '<, Menge:
 Alkohol AS string 'Alkohol text '^Alkohol pro Woche:
 Mitarbeiter AS string 'Mitarbeiter text '<, Mitarbeiter:
 Weitere_Medikation AS string 'Weitere_Medikation text '^:
 Liphypertrophien_Abdomen AS string 'Liphypertrophien_Abdomen text '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine AS string 'Liphypertrophien_Beine text '<, Beine:
 Liphypertrophien_Arme AS string 'Liphypertrophien_Arme varchar '<, Arme:
 Beinbefund AS string 'Beinbefund text '^:
 Hyperkeratosen AS string 'Hyperkeratosen text ',:
 Ulcera AS string 'Ulcera text ',:
 Kraft_Zehenheber AS string 'Kraft_Zehenheber text '^Kraft:Zehenheber
 Kraft_Zehenbeuger AS string 'Kraft_Zehenbeuger text '<, Zehenbeuger:
 Kraft_Knie AS string 'Kraft_Knie text '<, Knie:
 ASR AS string 'ASR text ',:
 PSR AS string 'PSR text ',:
 Oberflächensensibilität AS string 'Oberflächensensibilität text '^:
 Monofilamenttest AS string 'Monofilamenttest text ',:
 Kalt_Warm AS string 'Kalt_Warm text ', Kalt-Warm-Diskrimination:
 Vibration_IK AS string 'Vibration_IK text ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe AS string 'Vibration_Großzehe text '<, Großzehe:
 Puls_Leiste AS string 'Puls_Leiste text '^Pulse:Leiste
 Puls_Kniekehle AS string 'Puls_Kniekehle text '<,Kniekehle:
 Puls_Atp AS string 'Puls_Atp text '<,Innenknöchel:
 Puls_Adp AS string 'Puls_Adp text '<,Fußrücken:
 RR AS string 'RR text '^Blutdruck:
 RRTurboMed AS string 'RRTurboMed text '
 Herz AS string 'Herz text '^:
 Lunge AS string 'Lunge text ',:
 Bauch AS string 'Bauch text ', Abdomen:
 WS AS string 'WS text ', Wirbelsäule:
 NL AS string 'NL text ', Nierenlager:
 SD AS string 'SD text ', Schilddrüse:
 Carotiden AS string 'Carotiden text ', Halsschlagadern:
 NNH AS string 'NNH text ', Nasennebenhöhlen:
 Zähne AS string 'Zähne text ',:
 Parodontitis AS string 'Parodontitis text ',:
 Mundhöhle AS string 'Mundhöhle text ',:
 LK AS string 'LK text ', Lymphknoten:
 BeinödVen AS string 'BeinödVen text ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst AS string 'Neuro_sonst text '^Sonstige neurologische Befunde:
 Weitere_Befunde AS string 'Weitere_Befunde text ', weitere Befunde:
 Schulung AS string 'Schulung varchar 'ob Schulungsbedarf
 DMP AS string 'DMP text 'ob Pat. bei HA im DMP
 DMSchulz AS integer 'DMSchulz smallint 'Zahl der DMP-Schulungen hier
 DMSchL AS integer 'DMSchL smallint 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz AS integer 'RRSchulz smallint 'Zahl der Hypertonie-Schulungen hier
 DMPhier AS date 'DMPhier datetime 'ob Pat hier im DMP
 HANr AS string 'HANr varchar 'mit """"""""/""""""""
 HANr2 AS string 'HANr2 varchar 'mit """"""""/""""""""
 letzte_Änderung AS date 'letzte_Änderung datetime 'Datum der letzten Änderung
 Diagnosen AS string 'Diagnosen text '
 Vorgestellt AS date 'Vorgestellt datetime 'Erstvorstellung
 Versicherung AS string 'Versicherung varchar '
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 Ther1 AS string 'Ther1 varchar 'Diät, OAD, CT, Komb, ICT, CSII
 TherAkt AS string 'TherAkt varchar 'Diät, OAD, CT, Komb, ICT, CSII
 obAn1eing AS byte 'obAn1eing tinyint 'ob Anamneseblatt S. 1 eingegeben wurde
 obAn2eing AS byte 'obAn2eing tinyint 'ob Anamneseblatt S. 2 eingegeben wurde
 obAnAeing AS byte 'obAnAeing tinyint 'ob Anamneseblatt allgemein eingegeben wurde
 obCheck AS byte 'obCheck tinyint 'ob Checkliste vorliegt
 obBZausgew AS byte 'obBZausgew tinyint 'ob Blutzuckergerät ausgewechselt
 obOSaufgek AS byte 'obOSaufgek tinyint 'ob über orthopäd Schuhmacher aufgeklärt
 obPodAufgek AS byte 'obPodAufgek tinyint 'ob über Podologie aufgeklärt
 obMBlAusgeh AS byte 'obMBlAusgeh tinyint 'ob Merkblatt Fußsyndrom ausgehändigt
 obSchulaufgek AS string 'obSchulaufgek varchar 'ob über Podologie aufgeklärt
 obDMPaufgekl AS string 'obDMPaufgekl varchar 'ob Merkblatt Fußsyndrom ausgehändigt
 obMedNetz AS byte 'obMedNetz tinyint 'ob von Med. Netz geschickt
 Hausarzt AS string 'Hausarzt text 'Hausarzt laut Anamnesebogen
 ob AS byte 'ob tinyint 'für verschiedene Aktionen
 QS AS string 'QS varchar 'Quartal sortiert von vorgestellt
 QT AS string 'QT varchar 'Quartal sortiert von vorgestellt
end type

Public rNa() AS namen
Public rFa() AS faelle
Public rAu() AS au
Public rBr() AS briefe
Public rDi() AS diagnosen
Public rDo() AS dokumente
Public rEi() AS eintraege
Public rFi() AS forminhaltform_abk
Public rFo() AS formulare
Public rFr() AS forminhkopf
Public rFm() AS forminhfeld
Public rKh() AS kheinweis
Public rLb() AS lbanforderungen
Public rLa() AS laborneu
Public rLe() AS leistungen
Public rMe() AS medplan
Public rRe() AS rezepteintraege
Public rRr() AS rr
Public rKv() AS kvnrue
Public rUn() AS unbekannte_kennungen
Public rDm() AS dmpreihe
Public rDe() AS desktop
Public rUs() AS usdm
Public rFu() AS fuss
Public rUl() AS ulcus
Public rVk() AS vkgd
Public rSw() AS sws
Public rLs() AS laborxsaetze
Public rLg() AS laborxeingel
Public rLu() AS laborxus
Public rLo() AS laborxbakt
Public rLw() AS laborxwert
Public rLL() AS laborxleist
Public rLi() AS liuez
Public rAna() AS anamnesebogen

Public roNa() AS namen
Public roFa() AS faelle
Public roAu() AS au
Public roBr() AS briefe
Public roDi() AS diagnosen
Public roDo() AS dokumente
Public roEi() AS eintraege
Public roFi() AS forminhaltform_abk
Public roFo() AS formulare
Public roFr() AS forminhkopf
Public roFm() AS forminhfeld
Public roKh() AS kheinweis
Public roLb() AS lbanforderungen
Public roLa() AS laborneu
Public roLe() AS leistungen
Public roMe() AS medplan
Public roRe() AS rezepteintraege
Public roRr() AS rr
Public roKv() AS kvnrue
Public roUn() AS unbekannte_kennungen
Public roDm() AS dmpreihe
Public roDe() AS desktop
Public roUs() AS usdm
Public roFu() AS fuss
Public roUl() AS ulcus
Public roVk() AS vkgd
Public roSw() AS sws
Public roLs() AS laborxsaetze
Public roLg() AS laborxeingel
Public roLu() AS laborxus
Public roLo() AS laborxbakt
Public roLw() AS laborxwert
Public roLL() AS laborxleist
Public roLi() AS liuez
Public roAna() AS anamnesebogen

Public FUNCTION Tinit()
 static wdh%
 ReDim rAna(0)
 ReDim rNa(0)
 ReDim rFa(0)
 ReDim rAu(0)
 ReDim rBr(0)
 ReDim rDi(0)
 ReDim rDo(0)
 ReDim rEi(0)
 IF wdh = 0 THEN ReDim rFi(0)
 IF wdh = 0 THEN ReDim rFo(0)
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
End FUNCTION ' Tinit

Public FUNCTION LabInit()
 static wdh%
 ReDim rLs(0)
 ReDim rLg(0)
 ReDim rLu(0)
 ReDim rLo(0)
 ReDim rLw(0)
 ReDim rLL(0)
 ReDim rLi(0)
 wdh = -1
End FUNCTION ' Tinit

' in AllesLösch, LabLösch
Public FUNCTION doEntleer(frm AS lese, Tbl$)
 Dim rs As ADODB.Recordset
' SET rs = DBCn.Execute("SELECT COUNT(0) ct FROM `" & Tbl & "`")
 myFrag rs, "SELECT COUNT(0) ct FROM `" & Tbl & "`"
 frm.Ausgeb "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)", True
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call DBCn.Execute(sql) ' ,,adAsyncExecute
 DoEvents
End FUNCTION ' doEntleer

Public FUNCTION AllesLösch(frm AS lese)
 Dim ct&, rs As New ADODB.recordset
 ON Error GoTo fehler
 call ForeignNo0
 call ForeignNo1
 call doEntleer(frm, "sws")
 call doEntleer(frm, "vkgd")
 call doEntleer(frm, "ulcus")
 call doEntleer(frm, "fuss")
 call doEntleer(frm, "usdm")
 call doEntleer(frm, "desktop")
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
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' AllesLösch

Public FUNCTION LabLösch(frm AS lese)
 Dim ct&, rs As New ADODB.recordset
 ON Error GoTo fehler
 call ForeignNo0
 call ForeignNo1
 call doEntleer(frm, "liuez")
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
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' LabLösch

Function doBezFeh(csqlVal$, obSpei%, ErrDes$)
 Call ForeignNo0
 Call ForeignNo1
 obfor = True
 IF obSpei <> 0 THEN
  Open BezFeh For Append AS #299
  Print #299, vbCrLf & vbCrLf & Now() & ": " & csqlVal
  Print #299, vbCrLf & ErrDes
  Close #299
 END IF
End FUNCTION 'doBezFeh

Public FUNCTION roNaZuw(i&, j&)
 roNa(i).Pat_ID = rNa(j).Pat_ID
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
 roNa(i).Geschlecht = rNa(j).Geschlecht
 roNa(i).Plz = rNa(j).Plz
 roNa(i).Ort = rNa(j).Ort
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
 roNa(i).LANR = rNa(j).LANR
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
 roNa(i).Email = rNa(j).Email
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
End FUNCTION ' roNaZuw

Public FUNCTION NaZUnt%(i&, j&)
 IF roNa(i).Pat_ID <> rNa(j).Pat_ID THEN gosub unter
 IF roNa(i).lfdnr <> rNa(j).lfdnr THEN gosub unter
 IF roNa(i).NVorsatz <> rNa(j).NVorsatz THEN gosub unter
 IF roNa(i).Nachname <> rNa(j).Nachname THEN gosub unter
 IF roNa(i).Vorname <> rNa(j).Vorname THEN gosub unter
 IF roNa(i).GebDat <> rNa(j).GebDat THEN gosub unter
 IF roNa(i).f3004 <> rNa(j).f3004 THEN gosub unter
 IF roNa(i).f3006 <> rNa(j).f3006 THEN gosub unter
 IF roNa(i).Straße <> rNa(j).Straße THEN gosub unter
 IF roNa(i).KVKStatus <> rNa(j).KVKStatus THEN gosub unter
 IF roNa(i).Hausnr <> rNa(j).Hausnr THEN gosub unter
 IF roNa(i).Geschlecht <> rNa(j).Geschlecht THEN gosub unter
 IF roNa(i).Plz <> rNa(j).Plz THEN gosub unter
 IF roNa(i).Ort <> rNa(j).Ort THEN gosub unter
 IF roNa(i).Lkz <> rNa(j).Lkz THEN gosub unter
 IF roNa(i).Anschrzus <> rNa(j).Anschrzus THEN gosub unter
 IF roNa(i).NVors <> rNa(j).NVors THEN gosub unter
 IF roNa(i).PFPlz <> rNa(j).PFPlz THEN gosub unter
 IF roNa(i).PFOrt <> rNa(j).PFOrt THEN gosub unter
 IF roNa(i).PFNr <> rNa(j).PFNr THEN gosub unter
 IF roNa(i).f3124 <> rNa(j).f3124 THEN gosub unter
 IF roNa(i).AnschrZus_2 <> rNa(j).AnschrZus_2 THEN gosub unter
 IF roNa(i).Postfach_2 <> rNa(j).Postfach_2 THEN gosub unter
 IF roNa(i).LK_2 <> rNa(j).LK_2 THEN gosub unter
 IF roNa(i).Postfach <> rNa(j).Postfach THEN gosub unter
 IF roNa(i).Beruf <> rNa(j).Beruf THEN gosub unter
 IF roNa(i).Weggeldzone <> rNa(j).Weggeldzone THEN gosub unter
 IF roNa(i).WeggzZahl <> rNa(j).WeggzZahl THEN gosub unter
 IF roNa(i).AufnDat <> rNa(j).AufnDat THEN gosub unter
 IF roNa(i).kAufDat <> rNa(j).kAufDat THEN gosub unter
 IF roNa(i).LANR <> rNa(j).LANR THEN gosub unter
 IF roNa(i).BStNr <> rNa(j).BStNr THEN gosub unter
 IF roNa(i).Titel <> rNa(j).Titel THEN gosub unter
 IF roNa(i).Versichertennummer <> rNa(j).Versichertennummer THEN gosub unter
 IF roNa(i).PrivatTel <> rNa(j).PrivatTel THEN gosub unter
 IF roNa(i).KVNr <> rNa(j).KVNr THEN gosub unter
 IF roNa(i).KVNr2 <> rNa(j).KVNr2 THEN gosub unter
 IF roNa(i).KVNr3 <> rNa(j).KVNr3 THEN gosub unter
 IF roNa(i).KVNr4 <> rNa(j).KVNr4 THEN gosub unter
 IF roNa(i).PrivatTel_2 <> rNa(j).PrivatTel_2 THEN gosub unter
 IF roNa(i).PrivatFax <> rNa(j).PrivatFax THEN gosub unter
 IF roNa(i).DienstTel <> rNa(j).DienstTel THEN gosub unter
 IF roNa(i).PrivatMobil <> rNa(j).PrivatMobil THEN gosub unter
 IF roNa(i).Email <> rNa(j).Email THEN gosub unter
 IF roNa(i).Arbeitgeber <> rNa(j).Arbeitgeber THEN gosub unter
 IF roNa(i).AnAllgda <> rNa(j).AnAllgda THEN gosub unter
 IF roNa(i).An1da <> rNa(j).An1da THEN gosub unter
 IF roNa(i).An2da <> rNa(j).An2da THEN gosub unter
 IF roNa(i).Checkda <> rNa(j).Checkda THEN gosub unter
 IF roNa(i).DMTypaD <> rNa(j).DMTypaD THEN gosub unter
 IF roNa(i).AktZeit <> rNa(j).AktZeit THEN gosub unter
 IF roNa(i).absPos <> rNa(j).absPos THEN gosub unter
 IF roNa(i).StByte <> rNa(j).StByte THEN gosub unter
 IF roNa(i).StByteA <> rNa(j).StByteA THEN gosub unter
 IF roNa(i).Cave <> rNa(j).Cave THEN gosub unter
 IF roNa(i).Notiz <> rNa(j).Notiz THEN gosub unter
 IF roNa(i).obChk <> rNa(j).obChk THEN gosub unter
 IF roNa(i).dmpklass <> rNa(j).dmpklass THEN gosub unter
 IF roNa(i).dmpbeg <> rNa(j).dmpbeg THEN gosub unter
 IF roNa(i).dmpkhkklass <> rNa(j).dmpkhkklass THEN gosub unter
 IF roNa(i).dmpkhkbeg <> rNa(j).dmpkhkbeg THEN gosub unter
 IF roNa(i).dmpcopdklass <> rNa(j).dmpcopdklass THEN gosub unter
 IF roNa(i).dmpcopdbeg <> rNa(j).dmpcopdbeg THEN gosub unter
 IF roNa(i).dmpabklass <> rNa(j).dmpabklass THEN gosub unter
 IF roNa(i).dmpabbeg <> rNa(j).dmpabbeg THEN gosub unter
 IF roNa(i).dakab <> rNa(j).dakab THEN gosub unter
 IF roNa(i).HzV <> rNa(j).HzV THEN gosub unter
 IF roNa(i).HzVbeg <> rNa(j).HzVbeg THEN gosub unter
 IF roNa(i).DS <> rNa(j).DS THEN gosub unter
 IF roNa(i).DSbeg <> rNa(j).DSbeg THEN gosub unter
 IF roNa(i).getHA0 <> rNa(j).getHA0 THEN gosub unter
 IF roNa(i).fnHA0 <> rNa(j).fnHA0 THEN gosub unter
 IF roNa(i).getHA1 <> rNa(j).getHA1 THEN gosub unter
 IF roNa(i).fnHA1 <> rNa(j).fnHA1 THEN gosub unter
 IF roNa(i).getHA2 <> rNa(j).getHA2 THEN gosub unter
 IF roNa(i).fnHA2 <> rNa(j).fnHA2 THEN gosub unter
 IF roNa(i).zubenach <> rNa(j).zubenach THEN gosub unter
 IF roNa(i).Verwandt <> rNa(j).Verwandt THEN gosub unter
 IF roNa(i).Sprache <> rNa(j).Sprache THEN gosub unter
 IF roNa(i).lAktTM <> rNa(j).lAktTM THEN gosub unter
 IF roNa(i).Mitarbeiter <> rNa(j).Mitarbeiter THEN gosub unter
 Exit Function
unter:
 NaZUnt = NaZUnt + 1
 Return
End FUNCTION ' NaZUnt

Public FUNCTION namenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roNa(1)
 sql = "SELECT COALESCE(Pat_ID,0) Pat_ID,COALESCE(lfdnr,0) lfdnr,COALESCE(NVorsatz,'') NVorsatz,COALESCE(Nachname,'') Nachname" & _
",COALESCE(Vorname,'') Vorname,IF(GebDat IS NULL OR GebDat=0,CONVERT('18991230',DATE),GebDat) GebDat,COALESCE(f3004,'') f3004,COALESCE(f3006,'') f3006" & _
",COALESCE(Straße,'') Straße,COALESCE(KVKStatus,'') KVKStatus,COALESCE(Hausnr,'') Hausnr,COALESCE(Geschlecht,'') Geschlecht" & _
",COALESCE(Plz,'') Plz,COALESCE(Ort,'') Ort,COALESCE(Lkz,'') Lkz,COALESCE(Anschrzus,'') Anschrzus" & _
",COALESCE(NVors,'') NVors,COALESCE(PFPlz,'') PFPlz,COALESCE(PFOrt,'') PFOrt,COALESCE(PFNr,'') PFNr" & _
",COALESCE(f3124,'') f3124,COALESCE(AnschrZus_2,'') AnschrZus_2,COALESCE(Postfach_2,'') Postfach_2,COALESCE(LK_2,'') LK_2" & _
",COALESCE(Postfach,'') Postfach,COALESCE(Beruf,'') Beruf,COALESCE(Weggeldzone,'') Weggeldzone,COALESCE(WeggzZahl,0) WeggzZahl" & _
",IF(AufnDat IS NULL OR AufnDat=0,CONVERT('18991230',DATE),AufnDat) AufnDat,IF(kAufDat IS NULL OR kAufDat=0,CONVERT('18991230',DATE),kAufDat) kAufDat,COALESCE(LANR,'') LANR,COALESCE(BStNr,'') BStNr" & _
",COALESCE(Titel,'') Titel,COALESCE(Versichertennummer,'') Versichertennummer,COALESCE(PrivatTel,'') PrivatTel,COALESCE(KVNr,'') KVNr" & _
",COALESCE(KVNr2,'') KVNr2,COALESCE(KVNr3,'') KVNr3,COALESCE(KVNr4,'') KVNr4,COALESCE(PrivatTel_2,'') PrivatTel_2" & _
",COALESCE(PrivatFax,'') PrivatFax,COALESCE(DienstTel,'') DienstTel,COALESCE(PrivatMobil,'') PrivatMobil,COALESCE(Email,'') Email" & _
",COALESCE(Arbeitgeber,'') Arbeitgeber,COALESCE(AnAllgda,0) AnAllgda,COALESCE(An1da,0) An1da,COALESCE(An2da,0) An2da" & _
",COALESCE(Checkda,0) Checkda,COALESCE(DMTypaD,'') DMTypaD,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(absPos,0) absPos" & _
",COALESCE(StByte,0) StByte,COALESCE(StByteA,0) StByteA,COALESCE(Cave,'') Cave,COALESCE(Notiz,'') Notiz" & _
",COALESCE(obChk,'') obChk,COALESCE(dmpklass,0) dmpklass,IF(dmpbeg IS NULL OR dmpbeg=0,CONVERT('18991230',DATE),dmpbeg) dmpbeg,COALESCE(dmpkhkklass,0) dmpkhkklass" & _
",IF(dmpkhkbeg IS NULL OR dmpkhkbeg=0,CONVERT('18991230',DATE),dmpkhkbeg) dmpkhkbeg,COALESCE(dmpcopdklass,0) dmpcopdklass,IF(dmpcopdbeg IS NULL OR dmpcopdbeg=0,CONVERT('18991230',DATE),dmpcopdbeg) dmpcopdbeg,COALESCE(dmpabklass,0) dmpabklass" & _
",IF(dmpabbeg IS NULL OR dmpabbeg=0,CONVERT('18991230',DATE),dmpabbeg) dmpabbeg,IF(dakab IS NULL OR dakab=0,CONVERT('18991230',DATE),dakab) dakab,COALESCE(HzV,0) HzV,IF(HzVbeg IS NULL OR HzVbeg=0,CONVERT('18991230',DATE),HzVbeg) HzVbeg" & _
",COALESCE(DS,0) DS,IF(DSbeg IS NULL OR DSbeg=0,CONVERT('18991230',DATE),DSbeg) DSbeg,COALESCE(getHA0,0) getHA0,COALESCE(fnHA0,'') fnHA0" & _
",COALESCE(getHA1,0) getHA1,COALESCE(fnHA1,'') fnHA1,COALESCE(getHA2,0) getHA2,COALESCE(fnHA2,'') fnHA2" & _
",COALESCE(zubenach,'') zubenach,COALESCE(Verwandt,'') Verwandt,COALESCE(Sprache,'') Sprache,IF(lAktTM IS NULL OR lAktTM=0,CONVERT('18991230',DATE),lAktTM) lAktTM" & _
",COALESCE(Mitarbeiter,0) Mitarbeiter FROM `namen` WHERE Pat_ID=" & pid & " ORDER BY `kAufDat`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roNa)
  roNa(akt).Pat_ID = rs!Pat_ID
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
  roNa(akt).Geschlecht = doUmwfSQL(rs!Geschlecht, lies.obMySQL, False)
  roNa(akt).Plz = doUmwfSQL(rs!Plz, lies.obMySQL, False)
  roNa(akt).Ort = doUmwfSQL(rs!Ort, lies.obMySQL, False)
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
  roNa(akt).LANR = doUmwfSQL(rs!LANR, lies.obMySQL, False)
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
  roNa(akt).Email = doUmwfSQL(rs!Email, lies.obMySQL, False)
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
  IF Not rs.EOF THEN ReDim Preserve roNa(UBound(roNa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' namenLaden

Function namenEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rNa) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rNa)
   IF rNa(ri).kAufDat >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roNa)
    IF roNa(robeg).kAufDat >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roNa(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roNa(robeg + UBound(rNa) - rbeg)
   For ri = rbeg To UBound(rNa)
    Call roNaZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rNa = roNa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' namenEinf

Public FUNCTION namenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere namen"
 IF not Allepat THEN
   sql = "DELETE FROM `namen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,obChk,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,HzV,HzVbeg,DS,DSbeg,getHA0,fnHA0,getHA1," & _
     "fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM,Mitarbeiter) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,obChk,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,HzV,HzVbeg,DS,DSbeg,getHA0,fnHA0,getHA1," & _
     "fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM,Mitarbeiter)               VALUES"))
 For i = 0 to ubound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 0 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rNa(i).Pat_ID, "," , rNa(i).lfdnr, ",'" , rNa(i).NVorsatz, "','" , rNa(i).Nachname, "','" , rNa(i).Vorname, "'," , DatFor_k(rNa(i).GebDat), ",'" , rNa(i).f3004, "','" , rNa(i).f3006, "','" ,  _
   rNa(i).Straße, "','" , rNa(i).KVKStatus, "','" , rNa(i).Hausnr, "','" , rNa(i).Geschlecht, "','" , rNa(i).Plz, "','" , rNa(i).Ort, "','" , rNa(i).Lkz, "','" , rNa(i).Anschrzus, "','" ,  _
   rNa(i).NVors, "','" , rNa(i).PFPlz, "','" , rNa(i).PFOrt, "','" , rNa(i).PFNr, "','" , rNa(i).f3124, "','" , rNa(i).AnschrZus_2, "','" , rNa(i).Postfach_2, "','" , rNa(i).LK_2, "','" , rNa(i).Postfach, "','" ,  _
   rNa(i).Beruf, "','" , rNa(i).Weggeldzone, "'," , rNa(i).WeggzZahl, "," , DatFor_k(rNa(i).AufnDat), "," , DatFor_k(rNa(i).kAufDat), ",'" , rNa(i).LANR, "','" , rNa(i).BStNr, "','" ,  _
   rNa(i).Titel, "','" , rNa(i).Versichertennummer, "','" , rNa(i).PrivatTel, "','" , rNa(i).KVNr, "','" , rNa(i).KVNr2, "','" , rNa(i).KVNr3, "','" , rNa(i).KVNr4, "','" , rNa(i).PrivatTel_2, "','" ,  _
   rNa(i).PrivatFax, "','" , rNa(i).DienstTel, "','" , rNa(i).PrivatMobil, "','" , rNa(i).Email, "','" , rNa(i).Arbeitgeber, "'," , cstr(-(rNa(i).AnAllgda<>0)) , "," , cstr(-(rNa(i).An1da<>0)) , "," , cstr(-( _
   rNa(i).An2da<>0)) , "," , cstr(-(rNa(i).Checkda<>0)) , ",'" , rNa(i).DMTypaD, "'," , DatFor_k( 0 ), "," , rNa(i).absPos, "," , rNa(i).StByte, "," , rNa(i).StByteA, ",'" ,  _
   rNa(i).Cave, "','" , rNa(i).Notiz, "','" , rNa(i).obChk, "'," , rNa(i).dmpklass, "," , DatFor_k(rNa(i).dmpbeg), "," , rNa(i).dmpkhkklass, "," , DatFor_k(rNa(i).dmpkhkbeg), "," , rNa(i).dmpcopdklass, "," , DatFor_k( _
   rNa(i).dmpcopdbeg), "," , rNa(i).dmpabklass, "," , DatFor_k(rNa(i).dmpabbeg), "," , DatFor_k(rNa(i).dakab), "," , rNa(i).HzV, "," , DatFor_k(rNa(i).HzVbeg), "," , rNa(i).DS, "," , DatFor_k( _
   rNa(i).DSbeg), "," , rNa(i).getHA0, ",'" , rNa(i).fnHA0, "'," , rNa(i).getHA1, ",'" , rNa(i).fnHA1, "'," , rNa(i).getHA2, ",'" , rNa(i).fnHA2, "','" , rNa(i).zubenach, "','" ,  _
   rNa(i).Verwandt, "','" , rNa(i).Sprache, "'," , DatFor_k(rNa(i).lAktTM), "," , rNa(i).Mitarbeiter, ")")
  IF SammelInsert <> 0 AND i < ubound(rNa) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rNa) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(48)
 for k = iif(SammelInsert<>0,0,i) to iif(SammelInsert<>0,ubound(rNa),i)
  IF Len(rNa(k).NVorsatz) > maxi(0) THEN maxi(0) = Len(rNa(k).NVorsatz)
  IF Len(rNa(k).Nachname) > maxi(1) THEN maxi(1) = Len(rNa(k).Nachname)
  IF Len(rNa(k).Vorname) > maxi(2) THEN maxi(2) = Len(rNa(k).Vorname)
  IF Len(rNa(k).f3004) > maxi(3) THEN maxi(3) = Len(rNa(k).f3004)
  IF Len(rNa(k).f3006) > maxi(4) THEN maxi(4) = Len(rNa(k).f3006)
  IF Len(rNa(k).Straße) > maxi(5) THEN maxi(5) = Len(rNa(k).Straße)
  IF Len(rNa(k).KVKStatus) > maxi(6) THEN maxi(6) = Len(rNa(k).KVKStatus)
  IF Len(rNa(k).Hausnr) > maxi(7) THEN maxi(7) = Len(rNa(k).Hausnr)
  IF Len(rNa(k).Geschlecht) > maxi(8) THEN maxi(8) = Len(rNa(k).Geschlecht)
  IF Len(rNa(k).Plz) > maxi(9) THEN maxi(9) = Len(rNa(k).Plz)
  IF Len(rNa(k).Ort) > maxi(10) THEN maxi(10) = Len(rNa(k).Ort)
  IF Len(rNa(k).Lkz) > maxi(11) THEN maxi(11) = Len(rNa(k).Lkz)
  IF Len(rNa(k).Anschrzus) > maxi(12) THEN maxi(12) = Len(rNa(k).Anschrzus)
  IF Len(rNa(k).NVors) > maxi(13) THEN maxi(13) = Len(rNa(k).NVors)
  IF Len(rNa(k).PFPlz) > maxi(14) THEN maxi(14) = Len(rNa(k).PFPlz)
  IF Len(rNa(k).PFOrt) > maxi(15) THEN maxi(15) = Len(rNa(k).PFOrt)
  IF Len(rNa(k).PFNr) > maxi(16) THEN maxi(16) = Len(rNa(k).PFNr)
  IF Len(rNa(k).f3124) > maxi(17) THEN maxi(17) = Len(rNa(k).f3124)
  IF Len(rNa(k).AnschrZus_2) > maxi(18) THEN maxi(18) = Len(rNa(k).AnschrZus_2)
  IF Len(rNa(k).Postfach_2) > maxi(19) THEN maxi(19) = Len(rNa(k).Postfach_2)
  IF Len(rNa(k).LK_2) > maxi(20) THEN maxi(20) = Len(rNa(k).LK_2)
  IF Len(rNa(k).Postfach) > maxi(21) THEN maxi(21) = Len(rNa(k).Postfach)
  IF Len(rNa(k).Beruf) > maxi(22) THEN maxi(22) = Len(rNa(k).Beruf)
  IF Len(rNa(k).Weggeldzone) > maxi(23) THEN maxi(23) = Len(rNa(k).Weggeldzone)
  IF Len(rNa(k).LANR) > maxi(24) THEN maxi(24) = Len(rNa(k).LANR)
  IF Len(rNa(k).BStNr) > maxi(25) THEN maxi(25) = Len(rNa(k).BStNr)
  IF Len(rNa(k).Titel) > maxi(26) THEN maxi(26) = Len(rNa(k).Titel)
  IF Len(rNa(k).Versichertennummer) > maxi(27) THEN maxi(27) = Len(rNa(k).Versichertennummer)
  IF Len(rNa(k).PrivatTel) > maxi(28) THEN maxi(28) = Len(rNa(k).PrivatTel)
  IF Len(rNa(k).KVNr) > maxi(29) THEN maxi(29) = Len(rNa(k).KVNr)
  IF Len(rNa(k).KVNr2) > maxi(30) THEN maxi(30) = Len(rNa(k).KVNr2)
  IF Len(rNa(k).KVNr3) > maxi(31) THEN maxi(31) = Len(rNa(k).KVNr3)
  IF Len(rNa(k).KVNr4) > maxi(32) THEN maxi(32) = Len(rNa(k).KVNr4)
  IF Len(rNa(k).PrivatTel_2) > maxi(33) THEN maxi(33) = Len(rNa(k).PrivatTel_2)
  IF Len(rNa(k).PrivatFax) > maxi(34) THEN maxi(34) = Len(rNa(k).PrivatFax)
  IF Len(rNa(k).DienstTel) > maxi(35) THEN maxi(35) = Len(rNa(k).DienstTel)
  IF Len(rNa(k).PrivatMobil) > maxi(36) THEN maxi(36) = Len(rNa(k).PrivatMobil)
  IF Len(rNa(k).Email) > maxi(37) THEN maxi(37) = Len(rNa(k).Email)
  IF Len(rNa(k).Arbeitgeber) > maxi(38) THEN maxi(38) = Len(rNa(k).Arbeitgeber)
  IF Len(rNa(k).DMTypaD) > maxi(39) THEN maxi(39) = Len(rNa(k).DMTypaD)
  IF Len(rNa(k).Cave) > maxi(40) THEN maxi(40) = Len(rNa(k).Cave)
  IF Len(rNa(k).Notiz) > maxi(41) THEN maxi(41) = Len(rNa(k).Notiz)
  IF Len(rNa(k).obChk) > maxi(42) THEN maxi(42) = Len(rNa(k).obChk)
  IF Len(rNa(k).fnHA0) > maxi(43) THEN maxi(43) = Len(rNa(k).fnHA0)
  IF Len(rNa(k).fnHA1) > maxi(44) THEN maxi(44) = Len(rNa(k).fnHA1)
  IF Len(rNa(k).fnHA2) > maxi(45) THEN maxi(45) = Len(rNa(k).fnHA2)
  IF Len(rNa(k).zubenach) > maxi(46) THEN maxi(46) = Len(rNa(k).zubenach)
  IF Len(rNa(k).Verwandt) > maxi(47) THEN maxi(47) = Len(rNa(k).Verwandt)
  IF Len(rNa(k).Sprache) > maxi(48) THEN maxi(48) = Len(rNa(k).Sprache)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,0, i) To IIf(SammelInsert <> 0,ubound(rNa), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & Left$(rNa(k).NVorsatz, maxL)  & "'",true : rNa(k).NVorsatz = Left$(rNa(k).NVorsatz, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & Left$(rNa(k).Nachname, maxL)  & "'",true : rNa(k).Nachname = Left$(rNa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & Left$(rNa(k).Vorname, maxL)  & "'",true : rNa(k).Vorname = Left$(rNa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3004: '" & rNa(k).f3004 & "' -> '" & Left$(rNa(k).f3004, maxL)  & "'",true : rNa(k).f3004 = Left$(rNa(k).f3004, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3006: '" & rNa(k).f3006 & "' -> '" & Left$(rNa(k).f3006, maxL)  & "'",true : rNa(k).f3006 = Left$(rNa(k).f3006, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & Left$(rNa(k).Straße, maxL)  & "'",true : rNa(k).Straße = Left$(rNa(k).Straße, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & Left$(rNa(k).KVKStatus, maxL)  & "'",true : rNa(k).KVKStatus = Left$(rNa(k).KVKStatus, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Hausnr: '" & rNa(k).Hausnr & "' -> '" & Left$(rNa(k).Hausnr, maxL)  & "'",true : rNa(k).Hausnr = Left$(rNa(k).Hausnr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).Geschlecht & "' -> '" & Left$(rNa(k).Geschlecht, maxL)  & "'",true : rNa(k).Geschlecht = Left$(rNa(k).Geschlecht, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).Plz & "' -> '" & Left$(rNa(k).Plz, maxL)  & "'",true : rNa(k).Plz = Left$(rNa(k).Plz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).Ort & "' -> '" & Left$(rNa(k).Ort, maxL)  & "'",true : rNa(k).Ort = Left$(rNa(k).Ort, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rNa.Lkz: '" & rNa(k).Lkz & "' -> '" & Left$(rNa(k).Lkz, maxL)  & "'",true : rNa(k).Lkz = Left$(rNa(k).Lkz, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rNa.Anschrzus: '" & rNa(k).Anschrzus & "' -> '" & Left$(rNa(k).Anschrzus, maxL)  & "'",true : rNa(k).Anschrzus = Left$(rNa(k).Anschrzus, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVors: '" & rNa(k).NVors & "' -> '" & Left$(rNa(k).NVors, maxL)  & "'",true : rNa(k).NVors = Left$(rNa(k).NVors, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFPlz: '" & rNa(k).PFPlz & "' -> '" & Left$(rNa(k).PFPlz, maxL)  & "'",true : rNa(k).PFPlz = Left$(rNa(k).PFPlz, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFOrt: '" & rNa(k).PFOrt & "' -> '" & Left$(rNa(k).PFOrt, maxL)  & "'",true : rNa(k).PFOrt = Left$(rNa(k).PFOrt, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFNr: '" & rNa(k).PFNr & "' -> '" & Left$(rNa(k).PFNr, maxL)  & "'",true : rNa(k).PFNr = Left$(rNa(k).PFNr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3124: '" & rNa(k).f3124 & "' -> '" & Left$(rNa(k).f3124, maxL)  & "'",true : rNa(k).f3124 = Left$(rNa(k).f3124, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rNa.AnschrZus_2: '" & rNa(k).AnschrZus_2 & "' -> '" & Left$(rNa(k).AnschrZus_2, maxL)  & "'",true : rNa(k).AnschrZus_2 = Left$(rNa(k).AnschrZus_2, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach_2: '" & rNa(k).Postfach_2 & "' -> '" & Left$(rNa(k).Postfach_2, maxL)  & "'",true : rNa(k).Postfach_2 = Left$(rNa(k).Postfach_2, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rNa.LK_2: '" & rNa(k).LK_2 & "' -> '" & Left$(rNa(k).LK_2, maxL)  & "'",true : rNa(k).LK_2 = Left$(rNa(k).LK_2, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach: '" & rNa(k).Postfach & "' -> '" & Left$(rNa(k).Postfach, maxL)  & "'",true : rNa(k).Postfach = Left$(rNa(k).Postfach, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rNa.Beruf: '" & rNa(k).Beruf & "' -> '" & Left$(rNa(k).Beruf, maxL)  & "'",true : rNa(k).Beruf = Left$(rNa(k).Beruf, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & Left$(rNa(k).Weggeldzone, maxL)  & "'",true : rNa(k).Weggeldzone = Left$(rNa(k).Weggeldzone, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).LANR & "' -> '" & Left$(rNa(k).LANR, maxL)  & "'",true : rNa(k).LANR = Left$(rNa(k).LANR, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.BStNr: '" & rNa(k).BStNr & "' -> '" & Left$(rNa(k).BStNr, maxL)  & "'",true : rNa(k).BStNr = Left$(rNa(k).BStNr, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.Titel: '" & rNa(k).Titel & "' -> '" & Left$(rNa(k).Titel, maxL)  & "'",true : rNa(k).Titel = Left$(rNa(k).Titel, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.Versichertennummer: '" & rNa(k).Versichertennummer & "' -> '" & Left$(rNa(k).Versichertennummer, maxL)  & "'",true : rNa(k).Versichertennummer = Left$(rNa(k).Versichertennummer, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel: '" & rNa(k).PrivatTel & "' -> '" & Left$(rNa(k).PrivatTel, maxL)  & "'",true : rNa(k).PrivatTel = Left$(rNa(k).PrivatTel, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr: '" & rNa(k).KVNr & "' -> '" & Left$(rNa(k).KVNr, maxL)  & "'",true : rNa(k).KVNr = Left$(rNa(k).KVNr, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr2: '" & rNa(k).KVNr2 & "' -> '" & Left$(rNa(k).KVNr2, maxL)  & "'",true : rNa(k).KVNr2 = Left$(rNa(k).KVNr2, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr3: '" & rNa(k).KVNr3 & "' -> '" & Left$(rNa(k).KVNr3, maxL)  & "'",true : rNa(k).KVNr3 = Left$(rNa(k).KVNr3, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr4: '" & rNa(k).KVNr4 & "' -> '" & Left$(rNa(k).KVNr4, maxL)  & "'",true : rNa(k).KVNr4 = Left$(rNa(k).KVNr4, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel_2: '" & rNa(k).PrivatTel_2 & "' -> '" & Left$(rNa(k).PrivatTel_2, maxL)  & "'",true : rNa(k).PrivatTel_2 = Left$(rNa(k).PrivatTel_2, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatFax: '" & rNa(k).PrivatFax & "' -> '" & Left$(rNa(k).PrivatFax, maxL)  & "'",true : rNa(k).PrivatFax = Left$(rNa(k).PrivatFax, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rNa.DienstTel: '" & rNa(k).DienstTel & "' -> '" & Left$(rNa(k).DienstTel, maxL)  & "'",true : rNa(k).DienstTel = Left$(rNa(k).DienstTel, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatMobil: '" & rNa(k).PrivatMobil & "' -> '" & Left$(rNa(k).PrivatMobil, maxL)  & "'",true : rNa(k).PrivatMobil = Left$(rNa(k).PrivatMobil, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rNa.Email: '" & rNa(k).Email & "' -> '" & Left$(rNa(k).Email, maxL)  & "'",true : rNa(k).Email = Left$(rNa(k).Email, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rNa.Arbeitgeber: '" & rNa(k).Arbeitgeber & "' -> '" & Left$(rNa(k).Arbeitgeber, maxL)  & "'",true : rNa(k).Arbeitgeber = Left$(rNa(k).Arbeitgeber, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rNa.DMTypaD: '" & rNa(k).DMTypaD & "' -> '" & Left$(rNa(k).DMTypaD, maxL)  & "'",true : rNa(k).DMTypaD = Left$(rNa(k).DMTypaD, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rNa.Cave: '" & rNa(k).Cave & "' -> '" & Left$(rNa(k).Cave, maxL)  & "'",true : rNa(k).Cave = Left$(rNa(k).Cave, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rNa.Notiz: '" & rNa(k).Notiz & "' -> '" & Left$(rNa(k).Notiz, maxL)  & "'",true : rNa(k).Notiz = Left$(rNa(k).Notiz, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rNa.obChk: '" & rNa(k).obChk & "' -> '" & Left$(rNa(k).obChk, maxL)  & "'",true : rNa(k).obChk = Left$(rNa(k).obChk, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & Left$(rNa(k).fnHA0, maxL)  & "'",true : rNa(k).fnHA0 = Left$(rNa(k).fnHA0, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & Left$(rNa(k).fnHA1, maxL)  & "'",true : rNa(k).fnHA1 = Left$(rNa(k).fnHA1, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & Left$(rNa(k).fnHA2, maxL)  & "'",true : rNa(k).fnHA2 = Left$(rNa(k).fnHA2, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL)  & "'",true : rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL)  & "'",true : rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL)  & "'",true : rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
ElseIf Err.Number = -2147217871 OR Err.Number = -2147217859 OR err.Number = -2147467259 THEN
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' namenSpeichern

Public FUNCTION roFaZuw(i&, j&)
 roFa(i).FID = rFa(j).FID
 roFa(i).Pat_ID = rFa(j).Pat_ID
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
 roFa(i).LANRid = rFa(j).LANRid
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
End FUNCTION ' roFaZuw

Public FUNCTION FaZUnt%(i&, j&)
 IF roFa(i).FID <> rFa(j).FID THEN gosub unter
 IF roFa(i).Pat_ID <> rFa(j).Pat_ID THEN gosub unter
 IF roFa(i).Quartal <> rFa(j).Quartal THEN gosub unter
 IF roFa(i).Nachname <> rFa(j).Nachname THEN gosub unter
 IF roFa(i).Vorname <> rFa(j).Vorname THEN gosub unter
 IF roFa(i).lfdnr <> rFa(j).lfdnr THEN gosub unter
 IF roFa(i).TMFNr <> rFa(j).TMFNr THEN gosub unter
 IF roFa(i).VKNr <> rFa(j).VKNr THEN gosub unter
 IF roFa(i).f4131 <> rFa(j).f4131 THEN gosub unter
 IF roFa(i).f4132 <> rFa(j).f4132 THEN gosub unter
 IF roFa(i).VschBeg <> rFa(j).VschBeg THEN gosub unter
 IF roFa(i).KKasse_2 <> rFa(j).KKasse_2 THEN gosub unter
 IF roFa(i).FaktPers <> rFa(j).FaktPers THEN gosub unter
 IF roFa(i).FaktTechn <> rFa(j).FaktTechn THEN gosub unter
 IF roFa(i).FaktLabor <> rFa(j).FaktLabor THEN gosub unter
 IF roFa(i).BhFB <> rFa(j).BhFB THEN gosub unter
 IF roFa(i).BhFE1 <> rFa(j).BhFE1 THEN gosub unter
 IF roFa(i).BhFE2 <> rFa(j).BhFE2 THEN gosub unter
 IF roFa(i).f4202 <> rFa(j).f4202 THEN gosub unter
 IF roFa(i).ausgst <> rFa(j).ausgst THEN gosub unter
 IF roFa(i).KtrAbrB <> rFa(j).KtrAbrB THEN gosub unter
 IF roFa(i).AbrAr <> rFa(j).AbrAr THEN gosub unter
 IF roFa(i).lVorl <> rFa(j).lVorl THEN gosub unter
 IF roFa(i).IK <> rFa(j).IK THEN gosub unter
 IF roFa(i).KVKs <> rFa(j).KVKs THEN gosub unter
 IF roFa(i).KVKserg <> rFa(j).KVKserg THEN gosub unter
 IF roFa(i).Status <> rFa(j).Status THEN gosub unter
 IF roFa(i).Kasse <> rFa(j).Kasse THEN gosub unter
 IF roFa(i).KID <> rFa(j).KID THEN gosub unter
 IF roFa(i).GebOr <> rFa(j).GebOr THEN gosub unter
 IF roFa(i).AbrGb <> rFa(j).AbrGb THEN gosub unter
 IF roFa(i).PersKreis <> rFa(j).PersKreis THEN gosub unter
 IF roFa(i).SKtZusatz <> rFa(j).SKtZusatz THEN gosub unter
 IF roFa(i).letzteRegel <> rFa(j).letzteRegel THEN gosub unter
 IF roFa(i).ÜwText <> rFa(j).ÜwText THEN gosub unter
 IF roFa(i).f4210 <> rFa(j).f4210 THEN gosub unter
 IF roFa(i).AkfHAH <> rFa(j).AkfHAH THEN gosub unter
 IF roFa(i).AkfAB0 <> rFa(j).AkfAB0 THEN gosub unter
 IF roFa(i).AkfAK <> rFa(j).AkfAK THEN gosub unter
 IF roFa(i).statNuller <> rFa(j).statNuller THEN gosub unter
 IF roFa(i).ÜbwV <> rFa(j).ÜbwV THEN gosub unter
 IF roFa(i).ÜbWVLANR <> rFa(j).ÜbWVLANR THEN gosub unter
 IF roFa(i).ÜbWVBSNR <> rFa(j).ÜbWVBSNR THEN gosub unter
 IF roFa(i).ÜbWVKVNR <> rFa(j).ÜbWVKVNR THEN gosub unter
 IF roFa(i).AndÜw <> rFa(j).AndÜw THEN gosub unter
 IF roFa(i).Übwr <> rFa(j).Übwr THEN gosub unter
 IF roFa(i).ÜbwLANR <> rFa(j).ÜbwLANR THEN gosub unter
 IF roFa(i).ÜWZiel <> rFa(j).ÜWZiel THEN gosub unter
 IF roFa(i).ÜWNNr <> rFa(j).ÜWNNr THEN gosub unter
 IF roFa(i).ÜWNaN <> rFa(j).ÜWNaN THEN gosub unter
 IF roFa(i).ÜWTit <> rFa(j).ÜWTit THEN gosub unter
 IF roFa(i).ÜWVor <> rFa(j).ÜWVor THEN gosub unter
 IF roFa(i).ÜWVsw <> rFa(j).ÜWVsw THEN gosub unter
 IF roFa(i).üwvid <> rFa(j).üwvid THEN gosub unter
 IF roFa(i).Auftrag <> rFa(j).Auftrag THEN gosub unter
 IF roFa(i).Verdacht <> rFa(j).Verdacht THEN gosub unter
 IF roFa(i).Befund <> rFa(j).Befund THEN gosub unter
 IF roFa(i).statKlasse <> rFa(j).statKlasse THEN gosub unter
 IF roFa(i).f4237 <> rFa(j).f4237 THEN gosub unter
 IF roFa(i).statBehTage <> rFa(j).statBehTage THEN gosub unter
 IF roFa(i).SchGr <> rFa(j).SchGr THEN gosub unter
 IF roFa(i).Weiterbeh <> rFa(j).Weiterbeh THEN gosub unter
 IF roFa(i).f4266 <> rFa(j).f4266 THEN gosub unter
 IF roFa(i).PGeb <> rFa(j).PGeb THEN gosub unter
 IF roFa(i).PGebErg <> rFa(j).PGebErg THEN gosub unter
 IF roFa(i).Mahnfrist <> rFa(j).Mahnfrist THEN gosub unter
 IF roFa(i).Unfallort <> rFa(j).Unfallort THEN gosub unter
 IF roFa(i).BeschAls <> rFa(j).BeschAls THEN gosub unter
 IF roFa(i).BeschSeit <> rFa(j).BeschSeit THEN gosub unter
 IF roFa(i).Unfallbetrieb <> rFa(j).Unfallbetrieb THEN gosub unter
 IF roFa(i).f4570 <> rFa(j).f4570 THEN gosub unter
 IF roFa(i).GOÄKatNr <> rFa(j).GOÄKatNr THEN gosub unter
 IF roFa(i).GOÄKatName <> rFa(j).GOÄKatName THEN gosub unter
 IF roFa(i).abrArzt <> rFa(j).abrArzt THEN gosub unter
 IF roFa(i).privVers <> rFa(j).privVers THEN gosub unter
 IF roFa(i).AdNam <> rFa(j).AdNam THEN gosub unter
 IF roFa(i).AdStr <> rFa(j).AdStr THEN gosub unter
 IF roFa(i).AdPlz <> rFa(j).AdPlz THEN gosub unter
 IF roFa(i).AdOrt <> rFa(j).AdOrt THEN gosub unter
 IF roFa(i).ÜwBG <> rFa(j).ÜwBG THEN gosub unter
 IF roFa(i).BhFE <> rFa(j).BhFE THEN gosub unter
 IF roFa(i).s8000 <> rFa(j).s8000 THEN gosub unter
 IF roFa(i).s8100 <> rFa(j).s8100 THEN gosub unter
 IF roFa(i).AktZeit <> rFa(j).AktZeit THEN gosub unter
 IF roFa(i).Fanf <> rFa(j).Fanf THEN gosub unter
 IF roFa(i).altQuart <> rFa(j).altQuart THEN gosub unter
 IF roFa(i).QAnf <> rFa(j).QAnf THEN gosub unter
 IF roFa(i).QEnd <> rFa(j).QEnd THEN gosub unter
 IF roFa(i).QS <> rFa(j).QS THEN gosub unter
 IF roFa(i).QT <> rFa(j).QT THEN gosub unter
 IF roFa(i).StByte <> rFa(j).StByte THEN gosub unter
 IF roFa(i).absPos <> rFa(j).absPos THEN gosub unter
 IF roFa(i).LANRid <> rFa(j).LANRid THEN gosub unter
 IF roFa(i).f4108 <> rFa(j).f4108 THEN gosub unter
 IF roFa(i).BGFallNr <> rFa(j).BGFallNr THEN gosub unter
 IF roFa(i).lGewicht <> rFa(j).lGewicht THEN gosub unter
 IF roFa(i).vorET <> rFa(j).vorET THEN gosub unter
 IF roFa(i).dmpVertret <> rFa(j).dmpVertret THEN gosub unter
 IF roFa(i).dmpArztw <> rFa(j).dmpArztw THEN gosub unter
 IF roFa(i).dmpHypos <> rFa(j).dmpHypos THEN gosub unter
 IF roFa(i).dmpKhsA <> rFa(j).dmpKhsA THEN gosub unter
 IF roFa(i).dmpDMSchulEmpf <> rFa(j).dmpDMSchulEmpf THEN gosub unter
 IF roFa(i).dmpDMSchulWahrg <> rFa(j).dmpDMSchulWahrg THEN gosub unter
 IF roFa(i).dmpHypertSchulEmpf <> rFa(j).dmpHypertSchulEmpf THEN gosub unter
 IF roFa(i).dmpHypertSchulWahrg <> rFa(j).dmpHypertSchulWahrg THEN gosub unter
 IF roFa(i).dmpKKTabakEmpf <> rFa(j).dmpKKTabakEmpf THEN gosub unter
 IF roFa(i).dmpKKErnEmpf <> rFa(j).dmpKKErnEmpf THEN gosub unter
 IF roFa(i).dmpKKkTrainEmpf <> rFa(j).dmpKKkTrainEmpf THEN gosub unter
 IF roFa(i).dmpHbA1cZiel <> rFa(j).dmpHbA1cZiel THEN gosub unter
 IF roFa(i).dmpUewFuss <> rFa(j).dmpUewFuss THEN gosub unter
 IF roFa(i).dmpEinwDM <> rFa(j).dmpEinwDM THEN gosub unter
 IF roFa(i).dmphalbj <> rFa(j).dmphalbj THEN gosub unter
 IF roFa(i).dmpMA <> rFa(j).dmpMA THEN gosub unter
 Exit Function
unter:
 FaZUnt = FaZUnt + 1
 Return
End FUNCTION ' FaZUnt

Public FUNCTION faelleLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFa(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(Quartal,'') Quartal,COALESCE(Nachname,'') Nachname" & _
",COALESCE(Vorname,'') Vorname,COALESCE(lfdnr,0) lfdnr,COALESCE(TMFNr,'') TMFNr,COALESCE(VKNr,'') VKNr" & _
",COALESCE(f4131,'') f4131,COALESCE(f4132,'') f4132,IF(VschBeg IS NULL OR VschBeg=0,CONVERT('18991230',DATE),VschBeg) VschBeg,COALESCE(KKasse_2,'') KKasse_2" & _
",COALESCE(FaktPers,0) FaktPers,COALESCE(FaktTechn,0) FaktTechn,COALESCE(FaktLabor,0) FaktLabor,IF(BhFB IS NULL OR BhFB=0,CONVERT('18991230',DATE),BhFB) BhFB" & _
",IF(BhFE1 IS NULL OR BhFE1=0,CONVERT('18991230',DATE),BhFE1) BhFE1,IF(BhFE2 IS NULL OR BhFE2=0,CONVERT('18991230',DATE),BhFE2) BhFE2,COALESCE(f4202,'') f4202,IF(ausgst IS NULL OR ausgst=0,CONVERT('18991230',DATE),ausgst) ausgst" & _
",COALESCE(KtrAbrB,'') KtrAbrB,COALESCE(AbrAr,'') AbrAr,IF(lVorl IS NULL OR lVorl=0,CONVERT('18991230',DATE),lVorl) lVorl,COALESCE(IK,'') IK" & _
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
",IF(BeschSeit IS NULL OR BeschSeit=0,CONVERT('18991230',DATE),BeschSeit) BeschSeit,COALESCE(Unfallbetrieb,'') Unfallbetrieb,COALESCE(f4570,'') f4570,COALESCE(GOÄKatNr,'') GOÄKatNr" & _
",COALESCE(GOÄKatName,'') GOÄKatName,COALESCE(abrArzt,'') abrArzt,COALESCE(privVers,'') privVers,COALESCE(AdNam,'') AdNam" & _
",COALESCE(AdStr,'') AdStr,COALESCE(AdPlz,'') AdPlz,COALESCE(AdOrt,'') AdOrt,COALESCE(ÜwBG,'') ÜwBG" & _
",IF(BhFE IS NULL OR BhFE=0,CONVERT('18991230',DATE),BhFE) BhFE,COALESCE(s8000,'') s8000,COALESCE(s8100,'') s8100,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit" & _
",IF(Fanf IS NULL OR Fanf=0,CONVERT('18991230',DATE),Fanf) Fanf,COALESCE(altQuart,'') altQuart,IF(QAnf IS NULL OR QAnf=0,CONVERT('18991230',DATE),QAnf) QAnf,IF(QEnd IS NULL OR QEnd=0,CONVERT('18991230',DATE),QEnd) QEnd" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(absPos,0) absPos" & _
",COALESCE(LANRid,0) LANRid,COALESCE(f4108,'') f4108,COALESCE(BGFallNr,'') BGFallNr,COALESCE(lGewicht,0) lGewicht" & _
",IF(vorET IS NULL OR vorET=0,CONVERT('18991230',DATE),vorET) vorET,COALESCE(dmpVertret,'') dmpVertret,COALESCE(dmpArztw,'') dmpArztw,COALESCE(dmpHypos,'') dmpHypos"
sql = sql & ",COALESCE(dmpKhsA,'') dmpKhsA,COALESCE(dmpDMSchulEmpf,'') dmpDMSchulEmpf,COALESCE(dmpDMSchulWahrg,'') dmpDMSchulWahrg,COALESCE(dmpHypertSchulEmpf,'') dmpHypertSchulEmpf" & _
",COALESCE(dmpHypertSchulWahrg,'') dmpHypertSchulWahrg,COALESCE(dmpKKTabakEmpf,'') dmpKKTabakEmpf,COALESCE(dmpKKErnEmpf,'') dmpKKErnEmpf,COALESCE(dmpKKkTrainEmpf,'') dmpKKkTrainEmpf" & _
",COALESCE(dmpHbA1cZiel,'') dmpHbA1cZiel,COALESCE(dmpUewFuss,'') dmpUewFuss,COALESCE(dmpEinwDM,'') dmpEinwDM,COALESCE(dmphalbj,'') dmphalbj" & _
",COALESCE(dmpMA,'') dmpMA FROM `faelle` WHERE Pat_ID=" & pid & " ORDER BY `fanf`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roFa)
  roFa(akt).FID = rs!FID
  roFa(akt).Pat_ID = rs!Pat_ID
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
  roFa(akt).LANRid = rs!LANRid
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
  IF Not rs.EOF THEN ReDim Preserve roFa(UBound(roFa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' faelleLaden

Function faelleEinf
 Dim rbeg&, robeg&, ri&, roi&, jj&, fazu%()
 On Error GoTo fehler
 IF UBound(rFa) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rFa)
   IF rFa(ri).fanf >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roFa)
    IF roFa(robeg).fanf >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim fazu(robeg To UBound(roFa))
    For roi = robeg To UBound(roFa)
     For ri = rbeg To UBound(rFa) - 1
      IF roFa(roi).fanf >= rFa(ri).fanf AND roFa(roi).fanf < rFa(ri).BhFE THEN
       fazu(roi) = ri
       GoTo fertig
      END IF
     Next ri
     fazu(roi) = UBound(rFa)
fertig:
    Next roi
    For roi = robeg To UBound(roFa)
     For jj = 1 To UBound(roAu)
      IF roAu(jj).FID = roFa(roi).FID THEN roAu(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roBr)
      IF roBr(jj).FID = roFa(roi).FID THEN roBr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roDi)
      IF roDi(jj).FID = roFa(roi).FID THEN roDi(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roDo)
      IF roDo(jj).FID = roFa(roi).FID THEN roDo(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roEi)
      IF roEi(jj).FID = roFa(roi).FID THEN roEi(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roFr)
      IF roFr(jj).FID = roFa(roi).FID THEN roFr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roKh)
      IF roKh(jj).FID = roFa(roi).FID THEN roKh(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLb)
      IF roLb(jj).FID = roFa(roi).FID THEN roLb(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLa)
      IF roLa(jj).FID = roFa(roi).FID THEN roLa(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roLe)
      IF roLe(jj).FID = roFa(roi).FID THEN roLe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roMe)
      IF roMe(jj).FID = roFa(roi).FID THEN roMe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roRe)
      IF roRe(jj).FID = roFa(roi).FID THEN roRe(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roRr)
      IF roRr(jj).FID = roFa(roi).FID THEN roRr(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roUs)
      IF roUs(jj).FID = roFa(roi).FID THEN roUs(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roFu)
      IF roFu(jj).FID = roFa(roi).FID THEN roFu(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roUl)
      IF roUl(jj).FID = roFa(roi).FID THEN roUl(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roVk)
      IF roVk(jj).FID = roFa(roi).FID THEN roVk(jj).FID = rFa(fazu(roi)).FID
     Next jj
     For jj = 1 To UBound(roSw)
      IF roSw(jj).FID = roFa(roi).FID THEN roSw(jj).FID = rFa(fazu(roi)).FID
     Next jj
    Next roi
    ReDim Preserve roFa(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFa(robeg + UBound(rFa) - rbeg)
   For ri = rbeg To UBound(rFa)
    Call roFaZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rFa = roFa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' faelleEinf

Public FUNCTION faelleSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere faelle"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `faelle` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `faelle` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
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
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `faelle` (Pat_ID,Quartal,Nachname," & _
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
 For i = 1 to ubound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rFa(i).Pat_ID, ",'" , rFa(i).Quartal, "','" , rFa(i).Nachname, "','" , rFa(i).Vorname, "'," , rFa(i).lfdnr, ",'" , rFa(i).TMFNr, "','" , rFa(i).VKNr, "','" , rFa(i).f4131, "','" ,  _
   rFa(i).f4132, "'," , DatFor_k(rFa(i).VschBeg), ",'" , rFa(i).KKasse_2, "'," , replace$(rFa(i).FaktPers,",","."), "," , replace$(rFa(i).FaktTechn,",","."), "," , replace$(rFa(i).FaktLabor,",","."), "," , DatFor_k( _
   rFa(i).BhFB), "," , DatFor_k(rFa(i).BhFE1), "," , DatFor_k(rFa(i).BhFE2), ",'" , rFa(i).f4202, "'," , DatFor_k(rFa(i).ausgst), ",'" , rFa(i).KtrAbrB, "','" , rFa(i).AbrAr, "'," , DatFor_k( _
   rFa(i).lVorl), ",'" , rFa(i).IK, "','" , rFa(i).KVKs, "','" , rFa(i).KVKserg, "','" , rFa(i).Status, "','" , rFa(i).Kasse, "'," , rFa(i).KID, ",'" , rFa(i).GebOr, "','" , rFa(i).AbrGb, "','" ,  _
   rFa(i).PersKreis, "','" , rFa(i).SKtZusatz, "','" , rFa(i).letzteRegel, "','" , rFa(i).ÜwText, "'," , rFa(i).f4210, "," , rFa(i).AkfHAH, "," , rFa(i).AkfAB0, "," , rFa(i).AkfAK, ",'" ,  _
   rFa(i).statNuller, "','" , rFa(i).ÜbwV, "','" , rFa(i).ÜbWVLANR, "','" , rFa(i).ÜbWVBSNR, "','" , rFa(i).ÜbWVKVNR, "','" , rFa(i).AndÜw, "','" , rFa(i).Übwr, "','" , rFa(i).ÜbwLANR, "','" ,  _
   rFa(i).ÜWZiel, "','" , rFa(i).ÜWNNr, "','" , rFa(i).ÜWNaN, "','" , rFa(i).ÜWTit, "','" , rFa(i).ÜWVor, "','" , rFa(i).ÜWVsw, "'," , rFa(i).üwvid, ",'" , rFa(i).Auftrag, "','" , rFa(i).Verdacht, "','" ,  _
   rFa(i).Befund, "','" , rFa(i).statKlasse, "','" , rFa(i).f4237, "'," , rFa(i).statBehTage, "," , rFa(i).SchGr, ",'" , rFa(i).Weiterbeh, "'," , replace$(rFa(i).f4266,",","."), ",'" ,  _
   rFa(i).PGeb, "','" , rFa(i).PGebErg, "','" , rFa(i).Mahnfrist, "','" , rFa(i).Unfallort, "','" , rFa(i).BeschAls, "'," , DatFor_k(rFa(i).BeschSeit), ",'" , rFa(i).Unfallbetrieb, "','" , rFa(i).f4570, "','" ,  _
   rFa(i).GOÄKatNr, "','" , rFa(i).GOÄKatName, "','" , rFa(i).abrArzt, "','" , rFa(i).privVers, "','" , rFa(i).AdNam, "','" , rFa(i).AdStr, "','" , rFa(i).AdPlz, "','" , rFa(i).AdOrt, "','" ,  _
   rFa(i).ÜwBG, "'," , DatFor_k(rFa(i).BhFE), ",'" , rFa(i).s8000, "','" , rFa(i).s8100, "'," , DatFor_k(rFa(i).AktZeit), "," , DatFor_k(rFa(i).Fanf), ",'" , rFa(i).altQuart, "'," , DatFor_k(rFa(i).QAnf), "," , DatFor_k( _
   rFa(i).QEnd), ",'" , rFa(i).QS, "','" , rFa(i).QT, "'," , rFa(i).StByte, "," , rFa(i).absPos, "," , rFa(i).LANRid, ",'" , rFa(i).f4108, "','" , rFa(i).BGFallNr, "'," ,  _
   rFa(i).lGewicht, "," , DatFor_k(rFa(i).vorET), ",'" , rFa(i).dmpVertret, "','" , rFa(i).dmpArztw, "','" , rFa(i).dmpHypos, "','" , rFa(i).dmpKhsA, "','" , rFa(i).dmpDMSchulEmpf, "','" ,  _
   rFa(i).dmpDMSchulWahrg, "','" , rFa(i).dmpHypertSchulEmpf, "','" , rFa(i).dmpHypertSchulWahrg, "','" , rFa(i).dmpKKTabakEmpf, "','" , rFa(i).dmpKKErnEmpf, "','" , rFa(i).dmpKKkTrainEmpf, "','" ,  _
   rFa(i).dmpHbA1cZiel, "','" , rFa(i).dmpUewFuss, "','" , rFa(i).dmpEinwDM, "','" , rFa(i).dmphalbj, "','" , rFa(i).dmpMA, "')")
  IF SammelInsert <> 0 AND i < ubound(rFa) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFa) THEN
   IF Not obFor THEN ForeignNo0
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF Not obFor THEN ForeignYes0
   IF rAF = 0 THEN
    Err.Raise 998, , "Fehler in faelleSpeichern b.Pat. " & rFa(i).Pat_id & "Err.Number " & Err.Number & ", err.description: " & Err.Description
   END IF
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
  IF SammelInsert = 0 THEN
  'Hier gibts mit Sammelins noch ein Problem ...
   ' SET rs = DBCn.Execute("SELECT FID FROM `faelle` WHERE pat_id = " & rFa(i).Pat_ID & " AND quartal = '" & rFa(i).Quartal & "' AND bhfb = " & DatFor_k(rFa(i).BhFB) & " AND bhfe1 = " & DatFor_k(rFa(i).BhFE1) & " AND ausgst = " & DatFor_k(rFa(i).ausgst))
   SET rs = DBCn.Execute("SELECT LAST_INSERT_ID() FID")
   IF rs.BOF THEN
    Err.Raise 999, , "Fehler bei der Fallaktualisierung b.Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
   Else
    IF rs!FID <> rFa(i).FID THEN
     Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & rs!FID & " in zu speichernden Tabellen mit fallid", True 
     Dim jjj&
      For jjj = 1 To UBound(rAu)
       IF rAu(jjj).FID = rFa(i).FID THEN
        rAu(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rBr)
       IF rBr(jjj).FID = rFa(i).FID THEN
        rBr(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rDi)
       IF rDi(jjj).FID = rFa(i).FID THEN
        rDi(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rDo)
       IF rDo(jjj).FID = rFa(i).FID THEN
        rDo(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rEi)
       IF rEi(jjj).FID = rFa(i).FID THEN
        rEi(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rFr)
       IF rFr(jjj).FID = rFa(i).FID THEN
        rFr(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rKh)
       IF rKh(jjj).FID = rFa(i).FID THEN
        rKh(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rLb)
       IF rLb(jjj).FID = rFa(i).FID THEN
        rLb(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rLa)
       IF rLa(jjj).FID = rFa(i).FID THEN
        rLa(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rLe)
       IF rLe(jjj).FID = rFa(i).FID THEN
        rLe(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rMe)
       IF rMe(jjj).FID = rFa(i).FID THEN
        rMe(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rRe)
       IF rRe(jjj).FID = rFa(i).FID THEN
        rRe(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rRr)
       IF rRr(jjj).FID = rFa(i).FID THEN
        rRr(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rUs)
       IF rUs(jjj).FID = rFa(i).FID THEN
        rUs(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rFu)
       IF rFu(jjj).FID = rFa(i).FID THEN
        rFu(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rUl)
       IF rUl(jjj).FID = rFa(i).FID THEN
        rUl(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rVk)
       IF rVk(jjj).FID = rFa(i).FID THEN
        rVk(jjj).FID = rs!FID
       END IF
      Next jjj
      For jjj = 1 To UBound(rSw)
       IF rSw(jjj).FID = rFa(i).FID THEN
        rSw(jjj).FID = rs!FID
       END IF
      Next jjj
     END IF
    END IF
  END IF ' IF Sammelinsert = 0 
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(80)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFa),i)
  IF Len(rFa(k).Quartal) > maxi(0) THEN maxi(0) = Len(rFa(k).Quartal)
  IF Len(rFa(k).Nachname) > maxi(1) THEN maxi(1) = Len(rFa(k).Nachname)
  IF Len(rFa(k).Vorname) > maxi(2) THEN maxi(2) = Len(rFa(k).Vorname)
  IF Len(rFa(k).TMFNr) > maxi(3) THEN maxi(3) = Len(rFa(k).TMFNr)
  IF Len(rFa(k).VKNr) > maxi(4) THEN maxi(4) = Len(rFa(k).VKNr)
  IF Len(rFa(k).f4131) > maxi(5) THEN maxi(5) = Len(rFa(k).f4131)
  IF Len(rFa(k).f4132) > maxi(6) THEN maxi(6) = Len(rFa(k).f4132)
  IF Len(rFa(k).KKasse_2) > maxi(7) THEN maxi(7) = Len(rFa(k).KKasse_2)
  IF Len(rFa(k).f4202) > maxi(8) THEN maxi(8) = Len(rFa(k).f4202)
  IF Len(rFa(k).KtrAbrB) > maxi(9) THEN maxi(9) = Len(rFa(k).KtrAbrB)
  IF Len(rFa(k).AbrAr) > maxi(10) THEN maxi(10) = Len(rFa(k).AbrAr)
  IF Len(rFa(k).IK) > maxi(11) THEN maxi(11) = Len(rFa(k).IK)
  IF Len(rFa(k).KVKs) > maxi(12) THEN maxi(12) = Len(rFa(k).KVKs)
  IF Len(rFa(k).KVKserg) > maxi(13) THEN maxi(13) = Len(rFa(k).KVKserg)
  IF Len(rFa(k).Status) > maxi(14) THEN maxi(14) = Len(rFa(k).Status)
  IF Len(rFa(k).Kasse) > maxi(15) THEN maxi(15) = Len(rFa(k).Kasse)
  IF Len(rFa(k).GebOr) > maxi(16) THEN maxi(16) = Len(rFa(k).GebOr)
  IF Len(rFa(k).AbrGb) > maxi(17) THEN maxi(17) = Len(rFa(k).AbrGb)
  IF Len(rFa(k).PersKreis) > maxi(18) THEN maxi(18) = Len(rFa(k).PersKreis)
  IF Len(rFa(k).SKtZusatz) > maxi(19) THEN maxi(19) = Len(rFa(k).SKtZusatz)
  IF Len(rFa(k).letzteRegel) > maxi(20) THEN maxi(20) = Len(rFa(k).letzteRegel)
  IF Len(rFa(k).ÜwText) > maxi(21) THEN maxi(21) = Len(rFa(k).ÜwText)
  IF Len(rFa(k).statNuller) > maxi(22) THEN maxi(22) = Len(rFa(k).statNuller)
  IF Len(rFa(k).ÜbwV) > maxi(23) THEN maxi(23) = Len(rFa(k).ÜbwV)
  IF Len(rFa(k).ÜbWVLANR) > maxi(24) THEN maxi(24) = Len(rFa(k).ÜbWVLANR)
  IF Len(rFa(k).ÜbWVBSNR) > maxi(25) THEN maxi(25) = Len(rFa(k).ÜbWVBSNR)
  IF Len(rFa(k).ÜbWVKVNR) > maxi(26) THEN maxi(26) = Len(rFa(k).ÜbWVKVNR)
  IF Len(rFa(k).AndÜw) > maxi(27) THEN maxi(27) = Len(rFa(k).AndÜw)
  IF Len(rFa(k).Übwr) > maxi(28) THEN maxi(28) = Len(rFa(k).Übwr)
  IF Len(rFa(k).ÜbwLANR) > maxi(29) THEN maxi(29) = Len(rFa(k).ÜbwLANR)
  IF Len(rFa(k).ÜWZiel) > maxi(30) THEN maxi(30) = Len(rFa(k).ÜWZiel)
  IF Len(rFa(k).ÜWNNr) > maxi(31) THEN maxi(31) = Len(rFa(k).ÜWNNr)
  IF Len(rFa(k).ÜWNaN) > maxi(32) THEN maxi(32) = Len(rFa(k).ÜWNaN)
  IF Len(rFa(k).ÜWTit) > maxi(33) THEN maxi(33) = Len(rFa(k).ÜWTit)
  IF Len(rFa(k).ÜWVor) > maxi(34) THEN maxi(34) = Len(rFa(k).ÜWVor)
  IF Len(rFa(k).ÜWVsw) > maxi(35) THEN maxi(35) = Len(rFa(k).ÜWVsw)
  IF Len(rFa(k).Auftrag) > maxi(36) THEN maxi(36) = Len(rFa(k).Auftrag)
  IF Len(rFa(k).Verdacht) > maxi(37) THEN maxi(37) = Len(rFa(k).Verdacht)
  IF Len(rFa(k).Befund) > maxi(38) THEN maxi(38) = Len(rFa(k).Befund)
  IF Len(rFa(k).statKlasse) > maxi(39) THEN maxi(39) = Len(rFa(k).statKlasse)
  IF Len(rFa(k).f4237) > maxi(40) THEN maxi(40) = Len(rFa(k).f4237)
  IF Len(rFa(k).Weiterbeh) > maxi(41) THEN maxi(41) = Len(rFa(k).Weiterbeh)
  IF Len(rFa(k).PGeb) > maxi(42) THEN maxi(42) = Len(rFa(k).PGeb)
  IF Len(rFa(k).PGebErg) > maxi(43) THEN maxi(43) = Len(rFa(k).PGebErg)
  IF Len(rFa(k).Mahnfrist) > maxi(44) THEN maxi(44) = Len(rFa(k).Mahnfrist)
  IF Len(rFa(k).Unfallort) > maxi(45) THEN maxi(45) = Len(rFa(k).Unfallort)
  IF Len(rFa(k).BeschAls) > maxi(46) THEN maxi(46) = Len(rFa(k).BeschAls)
  IF Len(rFa(k).Unfallbetrieb) > maxi(47) THEN maxi(47) = Len(rFa(k).Unfallbetrieb)
  IF Len(rFa(k).f4570) > maxi(48) THEN maxi(48) = Len(rFa(k).f4570)
  IF Len(rFa(k).GOÄKatNr) > maxi(49) THEN maxi(49) = Len(rFa(k).GOÄKatNr)
  IF Len(rFa(k).GOÄKatName) > maxi(50) THEN maxi(50) = Len(rFa(k).GOÄKatName)
  IF Len(rFa(k).abrArzt) > maxi(51) THEN maxi(51) = Len(rFa(k).abrArzt)
  IF Len(rFa(k).privVers) > maxi(52) THEN maxi(52) = Len(rFa(k).privVers)
  IF Len(rFa(k).AdNam) > maxi(53) THEN maxi(53) = Len(rFa(k).AdNam)
  IF Len(rFa(k).AdStr) > maxi(54) THEN maxi(54) = Len(rFa(k).AdStr)
  IF Len(rFa(k).AdPlz) > maxi(55) THEN maxi(55) = Len(rFa(k).AdPlz)
  IF Len(rFa(k).AdOrt) > maxi(56) THEN maxi(56) = Len(rFa(k).AdOrt)
  IF Len(rFa(k).ÜwBG) > maxi(57) THEN maxi(57) = Len(rFa(k).ÜwBG)
  IF Len(rFa(k).s8000) > maxi(58) THEN maxi(58) = Len(rFa(k).s8000)
  IF Len(rFa(k).s8100) > maxi(59) THEN maxi(59) = Len(rFa(k).s8100)
  IF Len(rFa(k).altQuart) > maxi(60) THEN maxi(60) = Len(rFa(k).altQuart)
  IF Len(rFa(k).QS) > maxi(61) THEN maxi(61) = Len(rFa(k).QS)
  IF Len(rFa(k).QT) > maxi(62) THEN maxi(62) = Len(rFa(k).QT)
  IF Len(rFa(k).f4108) > maxi(63) THEN maxi(63) = Len(rFa(k).f4108)
  IF Len(rFa(k).BGFallNr) > maxi(64) THEN maxi(64) = Len(rFa(k).BGFallNr)
  IF Len(rFa(k).dmpVertret) > maxi(65) THEN maxi(65) = Len(rFa(k).dmpVertret)
  IF Len(rFa(k).dmpArztw) > maxi(66) THEN maxi(66) = Len(rFa(k).dmpArztw)
  IF Len(rFa(k).dmpHypos) > maxi(67) THEN maxi(67) = Len(rFa(k).dmpHypos)
  IF Len(rFa(k).dmpKhsA) > maxi(68) THEN maxi(68) = Len(rFa(k).dmpKhsA)
  IF Len(rFa(k).dmpDMSchulEmpf) > maxi(69) THEN maxi(69) = Len(rFa(k).dmpDMSchulEmpf)
  IF Len(rFa(k).dmpDMSchulWahrg) > maxi(70) THEN maxi(70) = Len(rFa(k).dmpDMSchulWahrg)
  IF Len(rFa(k).dmpHypertSchulEmpf) > maxi(71) THEN maxi(71) = Len(rFa(k).dmpHypertSchulEmpf)
  IF Len(rFa(k).dmpHypertSchulWahrg) > maxi(72) THEN maxi(72) = Len(rFa(k).dmpHypertSchulWahrg)
  IF Len(rFa(k).dmpKKTabakEmpf) > maxi(73) THEN maxi(73) = Len(rFa(k).dmpKKTabakEmpf)
  IF Len(rFa(k).dmpKKErnEmpf) > maxi(74) THEN maxi(74) = Len(rFa(k).dmpKKErnEmpf)
  IF Len(rFa(k).dmpKKkTrainEmpf) > maxi(75) THEN maxi(75) = Len(rFa(k).dmpKKkTrainEmpf)
  IF Len(rFa(k).dmpHbA1cZiel) > maxi(76) THEN maxi(76) = Len(rFa(k).dmpHbA1cZiel)
  IF Len(rFa(k).dmpUewFuss) > maxi(77) THEN maxi(77) = Len(rFa(k).dmpUewFuss)
  IF Len(rFa(k).dmpEinwDM) > maxi(78) THEN maxi(78) = Len(rFa(k).dmpEinwDM)
  IF Len(rFa(k).dmphalbj) > maxi(79) THEN maxi(79) = Len(rFa(k).dmphalbj)
  IF Len(rFa(k).dmpMA) > maxi(80) THEN maxi(80) = Len(rFa(k).dmpMA)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFa), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & Left$(rFa(k).Quartal, maxL)  & "'",true : rFa(k).Quartal = Left$(rFa(k).Quartal, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & Left$(rFa(k).Nachname, maxL)  & "'",true : rFa(k).Nachname = Left$(rFa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & Left$(rFa(k).Vorname, maxL)  & "'",true : rFa(k).Vorname = Left$(rFa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & Left$(rFa(k).TMFNr, maxL)  & "'",true : rFa(k).TMFNr = Left$(rFa(k).TMFNr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & Left$(rFa(k).VKNr, maxL)  & "'",true : rFa(k).VKNr = Left$(rFa(k).VKNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4131: '" & rFa(k).f4131 & "' -> '" & Left$(rFa(k).f4131, maxL)  & "'",true : rFa(k).f4131 = Left$(rFa(k).f4131, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4132: '" & rFa(k).f4132 & "' -> '" & Left$(rFa(k).f4132, maxL)  & "'",true : rFa(k).f4132 = Left$(rFa(k).f4132, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.KKasse_2: '" & rFa(k).KKasse_2 & "' -> '" & Left$(rFa(k).KKasse_2, maxL)  & "'",true : rFa(k).KKasse_2 = Left$(rFa(k).KKasse_2, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4202: '" & rFa(k).f4202 & "' -> '" & Left$(rFa(k).f4202, maxL)  & "'",true : rFa(k).f4202 = Left$(rFa(k).f4202, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & Left$(rFa(k).KtrAbrB, maxL)  & "'",true : rFa(k).KtrAbrB = Left$(rFa(k).KtrAbrB, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & Left$(rFa(k).AbrAr, maxL)  & "'",true : rFa(k).AbrAr = Left$(rFa(k).AbrAr, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & Left$(rFa(k).IK, maxL)  & "'",true : rFa(k).IK = Left$(rFa(k).IK, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & Left$(rFa(k).KVKs, maxL)  & "'",true : rFa(k).KVKs = Left$(rFa(k).KVKs, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & Left$(rFa(k).KVKserg, maxL)  & "'",true : rFa(k).KVKserg = Left$(rFa(k).KVKserg, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.Status: '" & rFa(k).Status & "' -> '" & Left$(rFa(k).Status, maxL)  & "'",true : rFa(k).Status = Left$(rFa(k).Status, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.Kasse: '" & rFa(k).Kasse & "' -> '" & Left$(rFa(k).Kasse, maxL)  & "'",true : rFa(k).Kasse = Left$(rFa(k).Kasse, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & Left$(rFa(k).GebOr, maxL)  & "'",true : rFa(k).GebOr = Left$(rFa(k).GebOr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & Left$(rFa(k).AbrGb, maxL)  & "'",true : rFa(k).AbrGb = Left$(rFa(k).AbrGb, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & Left$(rFa(k).PersKreis, maxL)  & "'",true : rFa(k).PersKreis = Left$(rFa(k).PersKreis, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & Left$(rFa(k).SKtZusatz, maxL)  & "'",true : rFa(k).SKtZusatz = Left$(rFa(k).SKtZusatz, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.letzteRegel: '" & rFa(k).letzteRegel & "' -> '" & Left$(rFa(k).letzteRegel, maxL)  & "'",true : rFa(k).letzteRegel = Left$(rFa(k).letzteRegel, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & Left$(rFa(k).ÜwText, maxL)  & "'",true : rFa(k).ÜwText = Left$(rFa(k).ÜwText, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & Left$(rFa(k).statNuller, maxL)  & "'",true : rFa(k).statNuller = Left$(rFa(k).statNuller, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & Left$(rFa(k).ÜbwV, maxL)  & "'",true : rFa(k).ÜbwV = Left$(rFa(k).ÜbwV, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVLANR: '" & rFa(k).ÜbWVLANR & "' -> '" & Left$(rFa(k).ÜbWVLANR, maxL)  & "'",true : rFa(k).ÜbWVLANR = Left$(rFa(k).ÜbWVLANR, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVBSNR: '" & rFa(k).ÜbWVBSNR & "' -> '" & Left$(rFa(k).ÜbWVBSNR, maxL)  & "'",true : rFa(k).ÜbWVBSNR = Left$(rFa(k).ÜbWVBSNR, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVKVNR: '" & rFa(k).ÜbWVKVNR & "' -> '" & Left$(rFa(k).ÜbWVKVNR, maxL)  & "'",true : rFa(k).ÜbWVKVNR = Left$(rFa(k).ÜbWVKVNR, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & Left$(rFa(k).AndÜw, maxL)  & "'",true : rFa(k).AndÜw = Left$(rFa(k).AndÜw, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übwr: '" & rFa(k).Übwr & "' -> '" & Left$(rFa(k).Übwr, maxL)  & "'",true : rFa(k).Übwr = Left$(rFa(k).Übwr, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & Left$(rFa(k).ÜbwLANR, maxL)  & "'",true : rFa(k).ÜbwLANR = Left$(rFa(k).ÜbwLANR, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & Left$(rFa(k).ÜWZiel, maxL)  & "'",true : rFa(k).ÜWZiel = Left$(rFa(k).ÜWZiel, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & Left$(rFa(k).ÜWNNr, maxL)  & "'",true : rFa(k).ÜWNNr = Left$(rFa(k).ÜWNNr, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & Left$(rFa(k).ÜWNaN, maxL)  & "'",true : rFa(k).ÜWNaN = Left$(rFa(k).ÜWNaN, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & Left$(rFa(k).ÜWTit, maxL)  & "'",true : rFa(k).ÜWTit = Left$(rFa(k).ÜWTit, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & Left$(rFa(k).ÜWVor, maxL)  & "'",true : rFa(k).ÜWVor = Left$(rFa(k).ÜWVor, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & Left$(rFa(k).ÜWVsw, maxL)  & "'",true : rFa(k).ÜWVsw = Left$(rFa(k).ÜWVsw, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.Auftrag: '" & rFa(k).Auftrag & "' -> '" & Left$(rFa(k).Auftrag, maxL)  & "'",true : rFa(k).Auftrag = Left$(rFa(k).Auftrag, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.Verdacht: '" & rFa(k).Verdacht & "' -> '" & Left$(rFa(k).Verdacht, maxL)  & "'",true : rFa(k).Verdacht = Left$(rFa(k).Verdacht, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.Befund: '" & rFa(k).Befund & "' -> '" & Left$(rFa(k).Befund, maxL)  & "'",true : rFa(k).Befund = Left$(rFa(k).Befund, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & Left$(rFa(k).statKlasse, maxL)  & "'",true : rFa(k).statKlasse = Left$(rFa(k).statKlasse, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4237: '" & rFa(k).f4237 & "' -> '" & Left$(rFa(k).f4237, maxL)  & "'",true : rFa(k).f4237 = Left$(rFa(k).f4237, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & Left$(rFa(k).Weiterbeh, maxL)  & "'",true : rFa(k).Weiterbeh = Left$(rFa(k).Weiterbeh, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & Left$(rFa(k).PGeb, maxL)  & "'",true : rFa(k).PGeb = Left$(rFa(k).PGeb, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & Left$(rFa(k).PGebErg, maxL)  & "'",true : rFa(k).PGebErg = Left$(rFa(k).PGebErg, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & Left$(rFa(k).Mahnfrist, maxL)  & "'",true : rFa(k).Mahnfrist = Left$(rFa(k).Mahnfrist, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallort: '" & rFa(k).Unfallort & "' -> '" & Left$(rFa(k).Unfallort, maxL)  & "'",true : rFa(k).Unfallort = Left$(rFa(k).Unfallort, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.BeschAls: '" & rFa(k).BeschAls & "' -> '" & Left$(rFa(k).BeschAls, maxL)  & "'",true : rFa(k).BeschAls = Left$(rFa(k).BeschAls, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallbetrieb: '" & rFa(k).Unfallbetrieb & "' -> '" & Left$(rFa(k).Unfallbetrieb, maxL)  & "'",true : rFa(k).Unfallbetrieb = Left$(rFa(k).Unfallbetrieb, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4570: '" & rFa(k).f4570 & "' -> '" & Left$(rFa(k).f4570, maxL)  & "'",true : rFa(k).f4570 = Left$(rFa(k).f4570, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & Left$(rFa(k).GOÄKatNr, maxL)  & "'",true : rFa(k).GOÄKatNr = Left$(rFa(k).GOÄKatNr, maxL)
       Case 50: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & Left$(rFa(k).GOÄKatName, maxL)  & "'",true : rFa(k).GOÄKatName = Left$(rFa(k).GOÄKatName, maxL)
       Case 51: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & Left$(rFa(k).abrArzt, maxL)  & "'",true : rFa(k).abrArzt = Left$(rFa(k).abrArzt, maxL)
       Case 52: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & Left$(rFa(k).privVers, maxL)  & "'",true : rFa(k).privVers = Left$(rFa(k).privVers, maxL)
       Case 53: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & Left$(rFa(k).AdNam, maxL)  & "'",true : rFa(k).AdNam = Left$(rFa(k).AdNam, maxL)
       Case 54: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & Left$(rFa(k).AdStr, maxL)  & "'",true : rFa(k).AdStr = Left$(rFa(k).AdStr, maxL)
       Case 55: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & Left$(rFa(k).AdPlz, maxL)  & "'",true : rFa(k).AdPlz = Left$(rFa(k).AdPlz, maxL)
       Case 56: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & Left$(rFa(k).AdOrt, maxL)  & "'",true : rFa(k).AdOrt = Left$(rFa(k).AdOrt, maxL)
       Case 57: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwBG: '" & rFa(k).ÜwBG & "' -> '" & Left$(rFa(k).ÜwBG, maxL)  & "'",true : rFa(k).ÜwBG = Left$(rFa(k).ÜwBG, maxL)
       Case 58: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & Left$(rFa(k).s8000, maxL)  & "'",true : rFa(k).s8000 = Left$(rFa(k).s8000, maxL)
       Case 59: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & Left$(rFa(k).s8100, maxL)  & "'",true : rFa(k).s8100 = Left$(rFa(k).s8100, maxL)
       Case 60: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & Left$(rFa(k).altQuart, maxL)  & "'",true : rFa(k).altQuart = Left$(rFa(k).altQuart, maxL)
       Case 61: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & Left$(rFa(k).QS, maxL)  & "'",true : rFa(k).QS = Left$(rFa(k).QS, maxL)
       Case 62: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & Left$(rFa(k).QT, maxL)  & "'",true : rFa(k).QT = Left$(rFa(k).QT, maxL)
       Case 63: Lese.Ausgeb "   Verkürze Inhalt von rFa.f4108: '" & rFa(k).f4108 & "' -> '" & Left$(rFa(k).f4108, maxL)  & "'",true : rFa(k).f4108 = Left$(rFa(k).f4108, maxL)
       Case 64: Lese.Ausgeb "   Verkürze Inhalt von rFa.BGFallNr: '" & rFa(k).BGFallNr & "' -> '" & Left$(rFa(k).BGFallNr, maxL)  & "'",true : rFa(k).BGFallNr = Left$(rFa(k).BGFallNr, maxL)
       Case 65: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpVertret: '" & rFa(k).dmpVertret & "' -> '" & Left$(rFa(k).dmpVertret, maxL)  & "'",true : rFa(k).dmpVertret = Left$(rFa(k).dmpVertret, maxL)
       Case 66: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpArztw: '" & rFa(k).dmpArztw & "' -> '" & Left$(rFa(k).dmpArztw, maxL)  & "'",true : rFa(k).dmpArztw = Left$(rFa(k).dmpArztw, maxL)
       Case 67: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypos: '" & rFa(k).dmpHypos & "' -> '" & Left$(rFa(k).dmpHypos, maxL)  & "'",true : rFa(k).dmpHypos = Left$(rFa(k).dmpHypos, maxL)
       Case 68: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKhsA: '" & rFa(k).dmpKhsA & "' -> '" & Left$(rFa(k).dmpKhsA, maxL)  & "'",true : rFa(k).dmpKhsA = Left$(rFa(k).dmpKhsA, maxL)
       Case 69: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulEmpf: '" & rFa(k).dmpDMSchulEmpf & "' -> '" & Left$(rFa(k).dmpDMSchulEmpf, maxL)  & "'",true : rFa(k).dmpDMSchulEmpf = Left$(rFa(k).dmpDMSchulEmpf, maxL)
       Case 70: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulWahrg: '" & rFa(k).dmpDMSchulWahrg & "' -> '" & Left$(rFa(k).dmpDMSchulWahrg, maxL)  & "'",true : rFa(k).dmpDMSchulWahrg = Left$(rFa(k).dmpDMSchulWahrg, maxL)
       Case 71: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulEmpf: '" & rFa(k).dmpHypertSchulEmpf & "' -> '" & Left$(rFa(k).dmpHypertSchulEmpf, maxL)  & "'",true : rFa(k).dmpHypertSchulEmpf = Left$(rFa(k).dmpHypertSchulEmpf, maxL)
       Case 72: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulWahrg: '" & rFa(k).dmpHypertSchulWahrg & "' -> '" & Left$(rFa(k).dmpHypertSchulWahrg, maxL)  & "'",true : rFa(k).dmpHypertSchulWahrg = Left$(rFa(k).dmpHypertSchulWahrg, maxL)
       Case 73: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKTabakEmpf: '" & rFa(k).dmpKKTabakEmpf & "' -> '" & Left$(rFa(k).dmpKKTabakEmpf, maxL)  & "'",true : rFa(k).dmpKKTabakEmpf = Left$(rFa(k).dmpKKTabakEmpf, maxL)
       Case 74: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKErnEmpf: '" & rFa(k).dmpKKErnEmpf & "' -> '" & Left$(rFa(k).dmpKKErnEmpf, maxL)  & "'",true : rFa(k).dmpKKErnEmpf = Left$(rFa(k).dmpKKErnEmpf, maxL)
       Case 75: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKkTrainEmpf: '" & rFa(k).dmpKKkTrainEmpf & "' -> '" & Left$(rFa(k).dmpKKkTrainEmpf, maxL)  & "'",true : rFa(k).dmpKKkTrainEmpf = Left$(rFa(k).dmpKKkTrainEmpf, maxL)
       Case 76: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHbA1cZiel: '" & rFa(k).dmpHbA1cZiel & "' -> '" & Left$(rFa(k).dmpHbA1cZiel, maxL)  & "'",true : rFa(k).dmpHbA1cZiel = Left$(rFa(k).dmpHbA1cZiel, maxL)
       Case 77: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpUewFuss: '" & rFa(k).dmpUewFuss & "' -> '" & Left$(rFa(k).dmpUewFuss, maxL)  & "'",true : rFa(k).dmpUewFuss = Left$(rFa(k).dmpUewFuss, maxL)
       Case 78: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpEinwDM: '" & rFa(k).dmpEinwDM & "' -> '" & Left$(rFa(k).dmpEinwDM, maxL)  & "'",true : rFa(k).dmpEinwDM = Left$(rFa(k).dmpEinwDM, maxL)
       Case 79: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmphalbj: '" & rFa(k).dmphalbj & "' -> '" & Left$(rFa(k).dmphalbj, maxL)  & "'",true : rFa(k).dmphalbj = Left$(rFa(k).dmphalbj, maxL)
       Case 80: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpMA: '" & rFa(k).dmpMA & "' -> '" & Left$(rFa(k).dmpMA, maxL)  & "'",true : rFa(k).dmpMA = Left$(rFa(k).dmpMA, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
IF Err.Number = -2147467259 THEN
 Dim sqlquer$
 sqlquer = "INSERT INTO `kassenliste`(name,kurzname,`GO`,`VKNR`,`IK`,`eingef`,pid) values (" & "'" & rFa(I).kasse & "', '" & rFa(I).kkasse_2 & "', '" & rFa(I).GOÄKatName & "', '" & rFa(I).VKNr & "', '" & rFa(I).IK & "'," & Format(NOW(), "yyyymmddHHMMSS") & "," & rFa(I).Pat_id & ")"
 InsKorr DBCn, DBCnS, sqlquer, rAF
 Resume
END IF ' Err.Number = -2147467259 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' faelleSpeichern

Public FUNCTION roAuZuw(i&, j&)
 roAu(i).FID = rAu(j).FID
 roAu(i).Pat_ID = rAu(j).Pat_ID
 roAu(i).ZeitPunkt = rAu(j).ZeitPunkt
 roAu(i).Beginn = rAu(j).Beginn
 roAu(i).Ende = rAu(j).Ende
 roAu(i).ICDs = rAu(j).ICDs
 roAu(i).absPos = rAu(j).absPos
 roAu(i).AktZeit = rAu(j).AktZeit
 roAu(i).StByte = rAu(j).StByte
End FUNCTION ' roAuZuw

Public FUNCTION AuZUnt%(i&, j&)
 IF roAu(i).FID <> rAu(j).FID THEN gosub unter
 IF roAu(i).Pat_ID <> rAu(j).Pat_ID THEN gosub unter
 IF roAu(i).ZeitPunkt <> rAu(j).ZeitPunkt THEN gosub unter
 IF roAu(i).Beginn <> rAu(j).Beginn THEN gosub unter
 IF roAu(i).Ende <> rAu(j).Ende THEN gosub unter
 IF roAu(i).ICDs <> rAu(j).ICDs THEN gosub unter
 IF roAu(i).absPos <> rAu(j).absPos THEN gosub unter
 IF roAu(i).AktZeit <> rAu(j).AktZeit THEN gosub unter
 IF roAu(i).StByte <> rAu(j).StByte THEN gosub unter
 Exit Function
unter:
 AuZUnt = AuZUnt + 1
 Return
End FUNCTION ' AuZUnt

Public FUNCTION auLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roAu(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Beginn,'') Beginn" & _
",COALESCE(Ende,'') Ende,COALESCE(ICDs,'') ICDs,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `au` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roAu)
  roAu(akt).FID = rs!FID
  roAu(akt).Pat_ID = rs!Pat_ID
  roAu(akt).ZeitPunkt = rs!ZeitPunkt
  roAu(akt).Beginn = doUmwfSQL(rs!Beginn, lies.obMySQL, False)
  roAu(akt).Ende = doUmwfSQL(rs!Ende, lies.obMySQL, False)
  roAu(akt).ICDs = doUmwfSQL(rs!ICDs, lies.obMySQL, False)
  roAu(akt).absPos = rs!absPos
  roAu(akt).AktZeit = rs!AktZeit
  roAu(akt).StByte = rs!StByte
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roAu(UBound(roAu) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' auLaden

Function auEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rAu) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rAu)
   IF rAu(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roAu)
    IF roAu(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roAu(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roAu(robeg + UBound(rAu) - rbeg)
   For ri = rbeg To UBound(rAu)
    Call roAuZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rAu = roAu
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' auEinf

Public FUNCTION auSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere au"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `au` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `au` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `au` (FID,Pat_ID,ZeitPunkt," & _
     "Beginn,Ende,ICDs,absPos,AktZeit,StByte)            VALUES"))
 For i = 1 to ubound(rAu)
'  rAu(i).AktZeit = now()
  rAu(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rAu(i).FID, "," , rAu(i).Pat_ID, "," , DatFor_k(rAu(i).ZeitPunkt), ",'" , rAu(i).Beginn, "','" , rAu(i).Ende, "','" , rAu(i).ICDs, "'," , rAu(i).absPos, "," , DatFor_k(rAu(i).AktZeit), "," ,  _
   rAu(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rAu) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rAu) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rAu),i)
  IF Len(rAu(k).Beginn) > maxi(0) THEN maxi(0) = Len(rAu(k).Beginn)
  IF Len(rAu(k).Ende) > maxi(1) THEN maxi(1) = Len(rAu(k).Ende)
  IF Len(rAu(k).ICDs) > maxi(2) THEN maxi(2) = Len(rAu(k).ICDs)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rAu), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & Left$(rAu(k).Beginn, maxL)  & "'",true : rAu(k).Beginn = Left$(rAu(k).Beginn, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & Left$(rAu(k).Ende, maxL)  & "'",true : rAu(k).Ende = Left$(rAu(k).Ende, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & Left$(rAu(k).ICDs, maxL)  & "'",true : rAu(k).ICDs = Left$(rAu(k).ICDs, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' auSpeichern

Public FUNCTION roBrZuw(i&, j&)
 roBr(i).FID = rBr(j).FID
 roBr(i).Pat_ID = rBr(j).Pat_ID
 roBr(i).ZeitPunkt = rBr(j).ZeitPunkt
 roBr(i).Pfad = rBr(j).Pfad
 roBr(i).Art = rBr(j).Art
 roBr(i).Name = rBr(j).Name
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
 roBr(i).ID = rBr(j).ID
End FUNCTION ' roBrZuw

Public FUNCTION BrZUnt%(i&, j&)
 IF roBr(i).FID <> rBr(j).FID THEN gosub unter
 IF roBr(i).Pat_ID <> rBr(j).Pat_ID THEN gosub unter
 IF roBr(i).ZeitPunkt <> rBr(j).ZeitPunkt THEN gosub unter
 IF roBr(i).Pfad <> rBr(j).Pfad THEN gosub unter
 IF roBr(i).Art <> rBr(j).Art THEN gosub unter
 IF roBr(i).Name <> rBr(j).Name THEN gosub unter
 IF roBr(i).autor <> rBr(j).autor THEN gosub unter
 IF roBr(i).Quelldatum <> rBr(j).Quelldatum THEN gosub unter
 IF roBr(i).Typ <> rBr(j).Typ THEN gosub unter
 IF roBr(i).AktZeit <> rBr(j).AktZeit THEN gosub unter
 IF roBr(i).DokGroe <> rBr(j).DokGroe THEN gosub unter
 IF roBr(i).DokAenD <> rBr(j).DokAenD THEN gosub unter
 IF roBr(i).QS <> rBr(j).QS THEN gosub unter
 IF roBr(i).QT <> rBr(j).QT THEN gosub unter
 IF roBr(i).absPos <> rBr(j).absPos THEN gosub unter
 IF roBr(i).StByte <> rBr(j).StByte THEN gosub unter
 IF roBr(i).ID <> rBr(j).ID THEN gosub unter
 Exit Function
unter:
 BrZUnt = BrZUnt + 1
 Return
End FUNCTION ' BrZUnt

Public FUNCTION briefeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roBr(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Pfad,'') Pfad" & _
",COALESCE(Art,'') Art,COALESCE(Name,'') Name,COALESCE(autor,'') autor,IF(Quelldatum IS NULL OR Quelldatum=0,CONVERT('18991230',DATE),Quelldatum) Quelldatum" & _
",COALESCE(Typ,'') Typ,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(DokGroe,0) DokGroe,IF(DokAenD IS NULL OR DokAenD=0,CONVERT('18991230',DATE),DokAenD) DokAenD" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(absPos,0) absPos,COALESCE(StByte,0) StByte" & _
",COALESCE(ID,0) ID FROM `briefe` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roBr)
  roBr(akt).FID = rs!FID
  roBr(akt).Pat_ID = rs!Pat_ID
  roBr(akt).ZeitPunkt = rs!ZeitPunkt
  roBr(akt).Pfad = doUmwfSQL(rs!Pfad, lies.obMySQL, False)
  roBr(akt).Art = doUmwfSQL(rs!Art, lies.obMySQL, False)
  roBr(akt).Name = doUmwfSQL(rs!Name, lies.obMySQL, False)
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
  roBr(akt).ID = rs!ID
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roBr(UBound(roBr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' briefeLaden

Function briefeEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rBr) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rBr)
   IF rBr(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roBr)
    IF roBr(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roBr(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roBr(robeg + UBound(rBr) - rbeg)
   For ri = rbeg To UBound(rBr)
    Call roBrZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rBr = roBr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' briefeEinf

Public FUNCTION briefeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere briefe"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `briefe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `briefe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,autor,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS," & _
     "QT,absPos,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,autor,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS," & _
     "QT,absPos,StByte)      VALUES"))
 For i = 1 to ubound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rBr(i).FID, "," , rBr(i).Pat_ID, "," , DatFor_k(rBr(i).ZeitPunkt), ",'" , rBr(i).Pfad, "','" , rBr(i).Art, "','" , rBr(i).Name, "','" , rBr(i).autor, "'," , DatFor_k(rBr(i).Quelldatum), ",'" ,  _
   rBr(i).Typ, "'," , DatFor_k(rBr(i).AktZeit), "," , rBr(i).DokGroe, "," , DatFor_k(rBr(i).DokAenD), ",'" , rBr(i).QS, "','" , rBr(i).QT, "'," , rBr(i).absPos, "," , rBr(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rBr) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rBr) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(6)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rBr),i)
  IF Len(rBr(k).Pfad) > maxi(0) THEN maxi(0) = Len(rBr(k).Pfad)
  IF Len(rBr(k).Art) > maxi(1) THEN maxi(1) = Len(rBr(k).Art)
  IF Len(rBr(k).Name) > maxi(2) THEN maxi(2) = Len(rBr(k).Name)
  IF Len(rBr(k).autor) > maxi(3) THEN maxi(3) = Len(rBr(k).autor)
  IF Len(rBr(k).Typ) > maxi(4) THEN maxi(4) = Len(rBr(k).Typ)
  IF Len(rBr(k).QS) > maxi(5) THEN maxi(5) = Len(rBr(k).QS)
  IF Len(rBr(k).QT) > maxi(6) THEN maxi(6) = Len(rBr(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rBr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL)  & "'",true : rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).Art & "' -> '" & Left$(rBr(k).Art, maxL)  & "'",true : rBr(k).Art = Left$(rBr(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).Name & "' -> '" & Left$(rBr(k).Name, maxL)  & "'",true : rBr(k).Name = Left$(rBr(k).Name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.autor: '" & rBr(k).autor & "' -> '" & Left$(rBr(k).autor, maxL)  & "'",true : rBr(k).autor = Left$(rBr(k).autor, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & Left$(rBr(k).Typ, maxL)  & "'",true : rBr(k).Typ = Left$(rBr(k).Typ, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & Left$(rBr(k).QS, maxL)  & "'",true : rBr(k).QS = Left$(rBr(k).QS, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & Left$(rBr(k).QT, maxL)  & "'",true : rBr(k).QT = Left$(rBr(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' briefeSpeichern

Public FUNCTION roDiZuw(i&, j&)
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
End FUNCTION ' roDiZuw

Public FUNCTION DiZUnt%(i&, j&)
 IF roDi(i).ID1 <> rDi(j).ID1 THEN gosub unter
 IF roDi(i).FID <> rDi(j).FID THEN gosub unter
 IF roDi(i).Pat_id <> rDi(j).Pat_id THEN gosub unter
 IF roDi(i).DiagDatum <> rDi(j).DiagDatum THEN gosub unter
 IF roDi(i).DiagSicherheit <> rDi(j).DiagSicherheit THEN gosub unter
 IF roDi(i).DiagText <> rDi(j).DiagText THEN gosub unter
 IF roDi(i).DiagSeite <> rDi(j).DiagSeite THEN gosub unter
 IF roDi(i).DiagAttr <> rDi(j).DiagAttr THEN gosub unter
 IF roDi(i).ICD <> rDi(j).ICD THEN gosub unter
 IF roDi(i).obDauer <> rDi(j).obDauer THEN gosub unter
 IF roDi(i).intBemerk <> rDi(j).intBemerk THEN gosub unter
 IF roDi(i).absPos <> rDi(j).absPos THEN gosub unter
 IF roDi(i).AktZeit <> rDi(j).AktZeit THEN gosub unter
 IF roDi(i).StByte <> rDi(j).StByte THEN gosub unter
 IF roDi(i).AusnBegr <> rDi(j).AusnBegr THEN gosub unter
 IF roDi(i).f6010 <> rDi(j).f6010 THEN gosub unter
 IF roDi(i).obKasse <> rDi(j).obKasse THEN gosub unter
 IF roDi(i).lKasse <> rDi(j).lKasse THEN gosub unter
 IF roDi(i).f6011 <> rDi(j).f6011 THEN gosub unter
 Exit Function
unter:
 DiZUnt = DiZUnt + 1
 Return
End FUNCTION ' DiZUnt

Public FUNCTION diagnosenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDi(1)
 sql = "SELECT COALESCE(ID1,0) ID1,COALESCE(FID,0) FID,COALESCE(Pat_id,0) Pat_id,IF(DiagDatum IS NULL OR DiagDatum=0,CONVERT('18991230',DATE),DiagDatum) DiagDatum" & _
",COALESCE(DiagSicherheit,'') DiagSicherheit,COALESCE(DiagText,'') DiagText,COALESCE(DiagSeite,'') DiagSeite,COALESCE(DiagAttr,'') DiagAttr" & _
",COALESCE(ICD,'') ICD,COALESCE(obDauer,0) obDauer,COALESCE(intBemerk,'') intBemerk,COALESCE(absPos,0) absPos" & _
",IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte,COALESCE(AusnBegr,'') AusnBegr,COALESCE(f6010,0) f6010" & _
",COALESCE(obKasse,0) obKasse,IF(lKasse IS NULL OR lKasse=0,CONVERT('18991230',DATE),lKasse) lKasse,COALESCE(f6011,'') f6011 FROM `diagnosen` WHERE Pat_ID=" & pid & " ORDER BY `DiagDatum`
 myfrag rs, sql
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
  IF Not rs.EOF THEN ReDim Preserve roDi(UBound(roDi) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' diagnosenLaden

Function diagnosenEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rDi) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rDi)
   IF rDi(ri).DiagDatum >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roDi)
    IF roDi(robeg).DiagDatum >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roDi(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDi(robeg + UBound(rDi) - rbeg)
   For ri = rbeg To UBound(rDi)
    Call roDiZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rDi = roDi
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' diagnosenEinf

Public FUNCTION diagnosenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere diagnosen"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `diagnosen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `diagnosen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,obKasse,lKasse,f6011) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,obKasse,lKasse,f6011)               VALUES"))
 For i = 1 to ubound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rDi(i).FID, "," , rDi(i).Pat_id, "," , DatFor_k(rDi(i).DiagDatum), ",'" , rDi(i).DiagSicherheit, "','" , rDi(i).DiagText, "','" , rDi(i).DiagSeite, "','" , rDi(i).DiagAttr, "','" ,  _
   rDi(i).ICD, "'," , rDi(i).obDauer, ",'" , rDi(i).intBemerk, "'," , rDi(i).absPos, "," , DatFor_k(rDi(i).AktZeit), "," , rDi(i).StByte, ",'" , rDi(i).AusnBegr, "'," , rDi(i).f6010, "," , rDi(i).obKasse, "," , DatFor_k( _
   rDi(i).lKasse), ",'" , rDi(i).f6011, "')")
  IF SammelInsert <> 0 AND i < ubound(rDi) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rDi) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(7)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDi),i)
  IF Len(rDi(k).DiagSicherheit) > maxi(0) THEN maxi(0) = Len(rDi(k).DiagSicherheit)
  IF Len(rDi(k).DiagText) > maxi(1) THEN maxi(1) = Len(rDi(k).DiagText)
  IF Len(rDi(k).DiagSeite) > maxi(2) THEN maxi(2) = Len(rDi(k).DiagSeite)
  IF Len(rDi(k).DiagAttr) > maxi(3) THEN maxi(3) = Len(rDi(k).DiagAttr)
  IF Len(rDi(k).ICD) > maxi(4) THEN maxi(4) = Len(rDi(k).ICD)
  IF Len(rDi(k).intBemerk) > maxi(5) THEN maxi(5) = Len(rDi(k).intBemerk)
  IF Len(rDi(k).AusnBegr) > maxi(6) THEN maxi(6) = Len(rDi(k).AusnBegr)
  IF Len(rDi(k).f6011) > maxi(7) THEN maxi(7) = Len(rDi(k).f6011)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & Left$(rDi(k).DiagSicherheit, maxL)  & "'",true : rDi(k).DiagSicherheit = Left$(rDi(k).DiagSicherheit, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & Left$(rDi(k).DiagText, maxL)  & "'",true : rDi(k).DiagText = Left$(rDi(k).DiagText, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & Left$(rDi(k).DiagSeite, maxL)  & "'",true : rDi(k).DiagSeite = Left$(rDi(k).DiagSeite, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & Left$(rDi(k).DiagAttr, maxL)  & "'",true : rDi(k).DiagAttr = Left$(rDi(k).DiagAttr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & Left$(rDi(k).ICD, maxL)  & "'",true : rDi(k).ICD = Left$(rDi(k).ICD, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & Left$(rDi(k).intBemerk, maxL)  & "'",true : rDi(k).intBemerk = Left$(rDi(k).intBemerk, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.AusnBegr: '" & rDi(k).AusnBegr & "' -> '" & Left$(rDi(k).AusnBegr, maxL)  & "'",true : rDi(k).AusnBegr = Left$(rDi(k).AusnBegr, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.f6011: '" & rDi(k).f6011 & "' -> '" & Left$(rDi(k).f6011, maxL)  & "'",true : rDi(k).f6011 = Left$(rDi(k).f6011, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' diagnosenSpeichern

Public FUNCTION roDoZuw(i&, j&)
 roDo(i).FID = rDo(j).FID
 roDo(i).Pat_ID = rDo(j).Pat_ID
 roDo(i).ZeitPunkt = rDo(j).ZeitPunkt
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
End FUNCTION ' roDoZuw

Public FUNCTION DoZUnt%(i&, j&)
 IF roDo(i).FID <> rDo(j).FID THEN gosub unter
 IF roDo(i).Pat_ID <> rDo(j).Pat_ID THEN gosub unter
 IF roDo(i).ZeitPunkt <> rDo(j).ZeitPunkt THEN gosub unter
 IF roDo(i).DokPfad <> rDo(j).DokPfad THEN gosub unter
 IF roDo(i).DokArt <> rDo(j).DokArt THEN gosub unter
 IF roDo(i).DokName <> rDo(j).DokName THEN gosub unter
 IF roDo(i).Quelldatum <> rDo(j).Quelldatum THEN gosub unter
 IF roDo(i).absPos <> rDo(j).absPos THEN gosub unter
 IF roDo(i).AktZeit <> rDo(j).AktZeit THEN gosub unter
 IF roDo(i).DokGroe <> rDo(j).DokGroe THEN gosub unter
 IF roDo(i).DokAenD <> rDo(j).DokAenD THEN gosub unter
 IF roDo(i).QS <> rDo(j).QS THEN gosub unter
 IF roDo(i).QT <> rDo(j).QT THEN gosub unter
 IF roDo(i).StByte <> rDo(j).StByte THEN gosub unter
 Exit Function
unter:
 DoZUnt = DoZUnt + 1
 Return
End FUNCTION ' DoZUnt

Public FUNCTION dokumenteLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDo(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(DokPfad,'') DokPfad" & _
",COALESCE(DokArt,'') DokArt,COALESCE(DokName,'') DokName,IF(Quelldatum IS NULL OR Quelldatum=0,CONVERT('18991230',DATE),Quelldatum) Quelldatum,COALESCE(absPos,0) absPos" & _
",IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(DokGroe,0) DokGroe,IF(DokAenD IS NULL OR DokAenD=0,CONVERT('18991230',DATE),DokAenD) DokAenD,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte FROM `dokumente` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roDo)
  roDo(akt).FID = rs!FID
  roDo(akt).Pat_ID = rs!Pat_ID
  roDo(akt).ZeitPunkt = rs!ZeitPunkt
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
  IF Not rs.EOF THEN ReDim Preserve roDo(UBound(roDo) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' dokumenteLaden

Function dokumenteEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rDo) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rDo)
   IF rDo(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roDo)
    IF roDo(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roDo(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDo(robeg + UBound(rDo) - rbeg)
   For ri = rbeg To UBound(rDo)
    Call roDoZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rDo = roDo
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' dokumenteEinf

Public FUNCTION dokumenteSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere dokumente"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `dokumente` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `dokumente` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `dokumente` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte)  VALUES"))
 For i = 1 to ubound(rDo)
'  rDo(i).AktZeit = now()
  rDo(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rDo(i).FID, "," , rDo(i).Pat_ID, "," , DatFor_k(rDo(i).ZeitPunkt), ",'" , rDo(i).DokPfad, "','" , rDo(i).DokArt, "','" , rDo(i).DokName, "'," , DatFor_k(rDo(i).Quelldatum), "," ,  _
   rDo(i).absPos, "," , DatFor_k(rDo(i).AktZeit), "," , rDo(i).DokGroe, "," , DatFor_k(rDo(i).DokAenD), ",'" , rDo(i).QS, "','" , rDo(i).QT, "'," , rDo(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rDo) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rDo) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(4)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDo),i)
  IF Len(rDo(k).DokPfad) > maxi(0) THEN maxi(0) = Len(rDo(k).DokPfad)
  IF Len(rDo(k).DokArt) > maxi(1) THEN maxi(1) = Len(rDo(k).DokArt)
  IF Len(rDo(k).DokName) > maxi(2) THEN maxi(2) = Len(rDo(k).DokName)
  IF Len(rDo(k).QS) > maxi(3) THEN maxi(3) = Len(rDo(k).QS)
  IF Len(rDo(k).QT) > maxi(4) THEN maxi(4) = Len(rDo(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & Left$(rDo(k).DokPfad, maxL)  & "'",true : rDo(k).DokPfad = Left$(rDo(k).DokPfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & Left$(rDo(k).DokArt, maxL)  & "'",true : rDo(k).DokArt = Left$(rDo(k).DokArt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & Left$(rDo(k).DokName, maxL)  & "'",true : rDo(k).DokName = Left$(rDo(k).DokName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & Left$(rDo(k).QS, maxL)  & "'",true : rDo(k).QS = Left$(rDo(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & Left$(rDo(k).QT, maxL)  & "'",true : rDo(k).QT = Left$(rDo(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' dokumenteSpeichern

Public FUNCTION roEiZuw(i&, j&)
 roEi(i).FID = rEi(j).FID
 roEi(i).Pat_ID = rEi(j).Pat_ID
 roEi(i).ZeitPunkt = rEi(j).ZeitPunkt
 roEi(i).Art = rEi(j).Art
 roEi(i).Inhalt = rEi(j).Inhalt
 roEi(i).absPos = rEi(j).absPos
 roEi(i).AktZeit = rEi(j).AktZeit
 roEi(i).QS = rEi(j).QS
 roEi(i).QT = rEi(j).QT
 roEi(i).StByte = rEi(j).StByte
 roEi(i).id = rEi(j).id
 roEi(i).inhNum = rEi(j).inhNum
End FUNCTION ' roEiZuw

Public FUNCTION EiZUnt%(i&, j&)
 IF roEi(i).FID <> rEi(j).FID THEN gosub unter
 IF roEi(i).Pat_ID <> rEi(j).Pat_ID THEN gosub unter
 IF roEi(i).ZeitPunkt <> rEi(j).ZeitPunkt THEN gosub unter
 IF roEi(i).Art <> rEi(j).Art THEN gosub unter
 IF roEi(i).Inhalt <> rEi(j).Inhalt THEN gosub unter
 IF roEi(i).absPos <> rEi(j).absPos THEN gosub unter
 IF roEi(i).AktZeit <> rEi(j).AktZeit THEN gosub unter
 IF roEi(i).QS <> rEi(j).QS THEN gosub unter
 IF roEi(i).QT <> rEi(j).QT THEN gosub unter
 IF roEi(i).StByte <> rEi(j).StByte THEN gosub unter
 IF roEi(i).id <> rEi(j).id THEN gosub unter
 IF roEi(i).inhNum <> rEi(j).inhNum THEN gosub unter
 Exit Function
unter:
 EiZUnt = EiZUnt + 1
 Return
End FUNCTION ' EiZUnt

Public FUNCTION eintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roEi(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Inhalt,'') Inhalt,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(id,0) id,COALESCE(inhNum,0) inhNum" & _
" FROM `eintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roEi)
  roEi(akt).FID = rs!FID
  roEi(akt).Pat_ID = rs!Pat_ID
  roEi(akt).ZeitPunkt = rs!ZeitPunkt
  roEi(akt).Art = doUmwfSQL(rs!Art, lies.obMySQL, False)
  roEi(akt).Inhalt = doUmwfSQL(rs!Inhalt, lies.obMySQL, False)
  roEi(akt).absPos = rs!absPos
  roEi(akt).AktZeit = rs!AktZeit
  roEi(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
  roEi(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
  roEi(akt).StByte = rs!StByte
  roEi(akt).id = rs!id
  roEi(akt).inhNum = rs!inhNum
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roEi(UBound(roEi) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' eintraegeLaden

Function eintraegeEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rEi) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rEi)
   IF rEi(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roEi)
    IF roEi(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roEi(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roEi(robeg + UBound(rEi) - rbeg)
   For ri = rbeg To UBound(rEi)
    Call roEiZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rEi = roEi
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' eintraegeEinf

Public FUNCTION eintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere eintraege"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `eintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `eintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `eintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum)     VALUES"))
 For i = 1 to ubound(rEi)
'  rEi(i).AktZeit = now()
  rEi(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rEi(i).FID, "," , rEi(i).Pat_ID, "," , DatFor_k(rEi(i).ZeitPunkt), ",'" , rEi(i).Art, "','" , rEi(i).Inhalt, "'," , rEi(i).absPos, "," , DatFor_k(rEi(i).AktZeit), ",'" , rEi(i).QS, "','" ,  _
   rEi(i).QT, "'," , rEi(i).StByte, "," , replace$(rEi(i).inhNum,",","."), ")")
  IF SammelInsert <> 0 AND i < ubound(rEi) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rEi) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(3)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rEi),i)
  IF Len(rEi(k).Art) > maxi(0) THEN maxi(0) = Len(rEi(k).Art)
  IF Len(rEi(k).Inhalt) > maxi(1) THEN maxi(1) = Len(rEi(k).Inhalt)
  IF Len(rEi(k).QS) > maxi(2) THEN maxi(2) = Len(rEi(k).QS)
  IF Len(rEi(k).QT) > maxi(3) THEN maxi(3) = Len(rEi(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rEi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).Art & "' -> '" & Left$(rEi(k).Art, maxL)  & "'",true : rEi(k).Art = Left$(rEi(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & Left$(rEi(k).Inhalt, maxL)  & "'",true : rEi(k).Inhalt = Left$(rEi(k).Inhalt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & Left$(rEi(k).QS, maxL)  & "'",true : rEi(k).QS = Left$(rEi(k).QS, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & Left$(rEi(k).QT, maxL)  & "'",true : rEi(k).QT = Left$(rEi(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' eintraegeSpeichern

Public FUNCTION forminhaltform_abkSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhaltform_abk"
 IF not Allepat THEN
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `forminhaltform_abk` (Form_Abk) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `forminhaltform_abk` (Form_Abk)             VALUES"))
 For i = rFi1 + 1 to ubound(rFi)
'  rFi(i).AktZeit = now()
  IF SammelInsert = 0 OR i = rFi1 + 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rFi(i).Form_Abk, "')")
  IF SammelInsert <> 0 AND i < ubound(rFi) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFi) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 rFi1 = ubound(rFi)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,rFi1 + 1,i) to iif(SammelInsert<>0,ubound(rFi),i)
  IF Len(rFi(k).Form_Abk) > maxi(0) THEN maxi(0) = Len(rFi(k).Form_Abk)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,rFi1 + 1, i) To IIf(SammelInsert <> 0,ubound(rFi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFi.Form_Abk: '" & rFi(k).Form_Abk & "' -> '" & Left$(rFi(k).Form_Abk, maxL)  & "'",true : rFi(k).Form_Abk = Left$(rFi(k).Form_Abk, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' forminhaltform_abkSpeichern

Public FUNCTION formulareSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere formulare"
 IF not Allepat THEN
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `formulare` (FormID,Form_Abk,FormBez," & _
     "FormVorl,AktZeit,absPos,StByte)      VALUES"))
 For i = rFo1 + 1 to ubound(rFo)
'  rFo(i).AktZeit = now()
  rFo(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = rFo1 + 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rFo(i).FormID, ",'" , rFo(i).Form_Abk, "','" , rFo(i).FormBez, "','" , rFo(i).FormVorl, "'," , DatFor_k(rFo(i).AktZeit), "," , rFo(i).absPos, "," , rFo(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rFo) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFo) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 rFo1 = ubound(rFo)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,rFo1 + 1,i) to iif(SammelInsert<>0,ubound(rFo),i)
  IF Len(rFo(k).Form_Abk) > maxi(0) THEN maxi(0) = Len(rFo(k).Form_Abk)
  IF Len(rFo(k).FormBez) > maxi(1) THEN maxi(1) = Len(rFo(k).FormBez)
  IF Len(rFo(k).FormVorl) > maxi(2) THEN maxi(2) = Len(rFo(k).FormVorl)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,rFo1 + 1, i) To IIf(SammelInsert <> 0,ubound(rFo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & Left$(rFo(k).Form_Abk, maxL)  & "'",true : rFo(k).Form_Abk = Left$(rFo(k).Form_Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & Left$(rFo(k).FormBez, maxL)  & "'",true : rFo(k).FormBez = Left$(rFo(k).FormBez, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & Left$(rFo(k).FormVorl, maxL)  & "'",true : rFo(k).FormVorl = Left$(rFo(k).FormVorl, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' formulareSpeichern

Public FUNCTION roFrZuw(i&, j&)
 roFr(i).FoID = rFr(j).FoID
 roFr(i).FID = rFr(j).FID
 roFr(i).Pat_ID = rFr(j).Pat_ID
 roFr(i).Form_ID = rFr(j).Form_ID
 roFr(i).ZeitPunkt = rFr(j).ZeitPunkt
 roFr(i).AbsPos = rFr(j).AbsPos
 roFr(i).AktZeit = rFr(j).AktZeit
 roFr(i).StByte = rFr(j).StByte
 roFr(i).Satzart = rFr(j).Satzart
 roFr(i).Satzlänge = rFr(j).Satzlänge
 roFr(i).LANRid = rFr(j).LANRid
End FUNCTION ' roFrZuw

Public FUNCTION FrZUnt%(i&, j&)
 IF roFr(i).FoID <> rFr(j).FoID THEN gosub unter
 IF roFr(i).FID <> rFr(j).FID THEN gosub unter
 IF roFr(i).Pat_ID <> rFr(j).Pat_ID THEN gosub unter
 IF roFr(i).Form_ID <> rFr(j).Form_ID THEN gosub unter
 IF roFr(i).ZeitPunkt <> rFr(j).ZeitPunkt THEN gosub unter
 IF roFr(i).AbsPos <> rFr(j).AbsPos THEN gosub unter
 IF roFr(i).AktZeit <> rFr(j).AktZeit THEN gosub unter
 IF roFr(i).StByte <> rFr(j).StByte THEN gosub unter
 IF roFr(i).Satzart <> rFr(j).Satzart THEN gosub unter
 IF roFr(i).Satzlänge <> rFr(j).Satzlänge THEN gosub unter
 IF roFr(i).LANRid <> rFr(j).LANRid THEN gosub unter
 Exit Function
unter:
 FrZUnt = FrZUnt + 1
 Return
End FUNCTION ' FrZUnt

Public FUNCTION forminhkopfLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFr(1)
 sql = "SELECT COALESCE(FoID,0) FoID,COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(Form_ID,0) Form_ID" & _
",IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(AbsPos,0) AbsPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte" & _
",COALESCE(Satzart,'') Satzart,COALESCE(Satzlänge,'') Satzlänge,COALESCE(LANRid,0) LANRid FROM `forminhkopf` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roFr)
  roFr(akt).FoID = rs!FoID
  roFr(akt).FID = rs!FID
  roFr(akt).Pat_ID = rs!Pat_ID
  roFr(akt).Form_ID = rs!Form_ID
  roFr(akt).ZeitPunkt = rs!ZeitPunkt
  roFr(akt).AbsPos = rs!AbsPos
  roFr(akt).AktZeit = rs!AktZeit
  roFr(akt).StByte = rs!StByte
  roFr(akt).Satzart = doUmwfSQL(rs!Satzart, lies.obMySQL, False)
  roFr(akt).Satzlänge = doUmwfSQL(rs!Satzlänge, lies.obMySQL, False)
  roFr(akt).LANRid = rs!LANRid
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roFr(UBound(roFr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' forminhkopfLaden

Function forminhkopfEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rFr) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rFr)
   IF rFr(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roFr)
    IF roFr(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roFr(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFr(robeg + UBound(rFr) - rbeg)
   For ri = rbeg To UBound(rFr)
    Call roFrZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rFr = roFr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' forminhkopfEinf

Public FUNCTION forminhkopfSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhkopf"
 IF not Allepat THEN
'  sql = "DELETE FROM `forminhfeld` WHERE foid IN (SELECT foid FROM `forminhkopf` WHERE pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "SELECT pat_id FROM `forminhkopf` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `forminhkopf` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `forminhkopf` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid)              VALUES"))
 For i = 1 to ubound(rFr)
'  rFr(i).AktZeit = now()
  rFr(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rFr(i).FoID, "," , rFr(i).FID, "," , rFr(i).Pat_ID, "," , rFr(i).Form_ID, "," , DatFor_k(rFr(i).ZeitPunkt), "," , rFr(i).AbsPos, "," , DatFor_k(rFr(i).AktZeit), "," , rFr(i).StByte, ",'" ,  _
   rFr(i).Satzart, "','" , rFr(i).Satzlänge, "'," , rFr(i).LANRid, ")")
  IF SammelInsert <> 0 AND i < ubound(rFr) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFr) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFr),i)
  IF Len(rFr(k).Satzart) > maxi(0) THEN maxi(0) = Len(rFr(k).Satzart)
  IF Len(rFr(k).Satzlänge) > maxi(1) THEN maxi(1) = Len(rFr(k).Satzlänge)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & Left$(rFr(k).Satzart, maxL)  & "'",true : rFr(k).Satzart = Left$(rFr(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & Left$(rFr(k).Satzlänge, maxL)  & "'",true : rFr(k).Satzlänge = Left$(rFr(k).Satzlänge, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' forminhkopfSpeichern

Public FUNCTION forminhfeldSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhfeld"
 IF not Allepat THEN
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      VALUES"))
 For i = 1 to ubound(rFm)
'  rFm(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rFm(i).FoID, "," , rFm(i).Nr, "," , rFm(i).FeldNr, "," , rFm(i).FeldVW, "," , rFm(i).FeldInhVW, ")")
  IF SammelInsert <> 0 AND i < ubound(rFm) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFm) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(-1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFm),i)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhfeld", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFm), i)
      SELECT CASE m
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' forminhfeldSpeichern

Public FUNCTION roKhZuw(i&, j&)
 roKh(i).FID = rKh(j).FID
 roKh(i).Pat_ID = rKh(j).Pat_ID
 roKh(i).ZeitPunkt = rKh(j).ZeitPunkt
 roKh(i).Ziel = rKh(j).Ziel
 roKh(i).Diagnose = rKh(j).Diagnose
 roKh(i).absPos = rKh(j).absPos
 roKh(i).AktZeit = rKh(j).AktZeit
 roKh(i).StByte = rKh(j).StByte
End FUNCTION ' roKhZuw

Public FUNCTION KhZUnt%(i&, j&)
 IF roKh(i).FID <> rKh(j).FID THEN gosub unter
 IF roKh(i).Pat_ID <> rKh(j).Pat_ID THEN gosub unter
 IF roKh(i).ZeitPunkt <> rKh(j).ZeitPunkt THEN gosub unter
 IF roKh(i).Ziel <> rKh(j).Ziel THEN gosub unter
 IF roKh(i).Diagnose <> rKh(j).Diagnose THEN gosub unter
 IF roKh(i).absPos <> rKh(j).absPos THEN gosub unter
 IF roKh(i).AktZeit <> rKh(j).AktZeit THEN gosub unter
 IF roKh(i).StByte <> rKh(j).StByte THEN gosub unter
 Exit Function
unter:
 KhZUnt = KhZUnt + 1
 Return
End FUNCTION ' KhZUnt

Public FUNCTION kheinweisLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roKh(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Ziel,'') Ziel" & _
",COALESCE(Diagnose,'') Diagnose,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte" & _
" FROM `kheinweis` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roKh)
  roKh(akt).FID = rs!FID
  roKh(akt).Pat_ID = rs!Pat_ID
  roKh(akt).ZeitPunkt = rs!ZeitPunkt
  roKh(akt).Ziel = doUmwfSQL(rs!Ziel, lies.obMySQL, False)
  roKh(akt).Diagnose = doUmwfSQL(rs!Diagnose, lies.obMySQL, False)
  roKh(akt).absPos = rs!absPos
  roKh(akt).AktZeit = rs!AktZeit
  roKh(akt).StByte = rs!StByte
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roKh(UBound(roKh) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' kheinweisLaden

Function kheinweisEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rKh) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rKh)
   IF rKh(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roKh)
    IF roKh(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roKh(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roKh(robeg + UBound(rKh) - rbeg)
   For ri = rbeg To UBound(rKh)
    Call roKhZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rKh = roKh
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' kheinweisEinf

Public FUNCTION kheinweisSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere kheinweis"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `kheinweis` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `kheinweis` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `kheinweis` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,Diagnose,absPos,AktZeit,StByte)               VALUES"))
 For i = 1 to ubound(rKh)
'  rKh(i).AktZeit = now()
  rKh(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rKh(i).FID, "," , rKh(i).Pat_ID, "," , DatFor_k(rKh(i).ZeitPunkt), ",'" , rKh(i).Ziel, "','" , rKh(i).Diagnose, "'," , rKh(i).absPos, "," , DatFor_k(rKh(i).AktZeit), "," , rKh(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rKh) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rKh) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rKh),i)
  IF Len(rKh(k).Ziel) > maxi(0) THEN maxi(0) = Len(rKh(k).Ziel)
  IF Len(rKh(k).Diagnose) > maxi(1) THEN maxi(1) = Len(rKh(k).Diagnose)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rKh), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & Left$(rKh(k).Ziel, maxL)  & "'",true : rKh(k).Ziel = Left$(rKh(k).Ziel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & Left$(rKh(k).Diagnose, maxL)  & "'",true : rKh(k).Diagnose = Left$(rKh(k).Diagnose, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' kheinweisSpeichern

Public FUNCTION roLbZuw(i&, j&)
 roLb(i).FID = rLb(j).FID
 roLb(i).Pat_ID = rLb(j).Pat_ID
 roLb(i).ZeitPunkt = rLb(j).ZeitPunkt
 roLb(i).AnfText = rLb(j).AnfText
 roLb(i).absPos = rLb(j).absPos
 roLb(i).AktZeit = rLb(j).AktZeit
 roLb(i).StByte = rLb(j).StByte
End FUNCTION ' roLbZuw

Public FUNCTION LbZUnt%(i&, j&)
 IF roLb(i).FID <> rLb(j).FID THEN gosub unter
 IF roLb(i).Pat_ID <> rLb(j).Pat_ID THEN gosub unter
 IF roLb(i).ZeitPunkt <> rLb(j).ZeitPunkt THEN gosub unter
 IF roLb(i).AnfText <> rLb(j).AnfText THEN gosub unter
 IF roLb(i).absPos <> rLb(j).absPos THEN gosub unter
 IF roLb(i).AktZeit <> rLb(j).AktZeit THEN gosub unter
 IF roLb(i).StByte <> rLb(j).StByte THEN gosub unter
 Exit Function
unter:
 LbZUnt = LbZUnt + 1
 Return
End FUNCTION ' LbZUnt

Public FUNCTION lbanforderungenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLb(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(AnfText,'') AnfText" & _
",COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte FROM `lbanforderungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roLb)
  roLb(akt).FID = rs!FID
  roLb(akt).Pat_ID = rs!Pat_ID
  roLb(akt).ZeitPunkt = rs!ZeitPunkt
  roLb(akt).AnfText = doUmwfSQL(rs!AnfText, lies.obMySQL, False)
  roLb(akt).absPos = rs!absPos
  roLb(akt).AktZeit = rs!AktZeit
  roLb(akt).StByte = rs!StByte
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roLb(UBound(roLb) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' lbanforderungenLaden

Function lbanforderungenEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rLb) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rLb)
   IF rLb(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roLb)
    IF roLb(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roLb(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLb(robeg + UBound(rLb) - rbeg)
   For ri = rbeg To UBound(rLb)
    Call roLbZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rLb = roLb
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' lbanforderungenEinf

Public FUNCTION lbanforderungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere lbanforderungen"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `lbanforderungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `lbanforderungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `lbanforderungen` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       VALUES"))
 For i = 1 to ubound(rLb)
'  rLb(i).AktZeit = now()
  rLb(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLb(i).FID, "," , rLb(i).Pat_ID, "," , DatFor_k(rLb(i).ZeitPunkt), ",'" , rLb(i).AnfText, "'," , rLb(i).absPos, "," , DatFor_k(rLb(i).AktZeit), "," , rLb(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rLb) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLb) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLb),i)
  IF Len(rLb(k).AnfText) > maxi(0) THEN maxi(0) = Len(rLb(k).AnfText)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLb), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & Left$(rLb(k).AnfText, maxL)  & "'",true : rLb(k).AnfText = Left$(rLb(k).AnfText, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' lbanforderungenSpeichern

Public FUNCTION roLaZuw(i&, j&)
 roLa(i).FID = rLa(j).FID
 roLa(i).Pat_ID = rLa(j).Pat_ID
 roLa(i).ZeitPunkt = rLa(j).ZeitPunkt
 roLa(i).FertigStGrad = rLa(j).FertigStGrad
 roLa(i).Abkü = rLa(j).Abkü
 roLa(i).LangtextVW = rLa(j).LangtextVW
 roLa(i).Wert = rLa(j).Wert
 roLa(i).Einheit = rLa(j).Einheit
 roLa(i).KommentarVW = rLa(j).KommentarVW
 roLa(i).AbsPos = rLa(j).AbsPos
 roLa(i).AktZeit = rLa(j).AktZeit
 roLa(i).Refnr = rLa(j).Refnr
 roLa(i).StByte = rLa(j).StByte
 roLa(i).ID = rLa(j).ID
End FUNCTION ' roLaZuw

Public FUNCTION LaZUnt%(i&, j&)
 IF roLa(i).FID <> rLa(j).FID THEN gosub unter
 IF roLa(i).Pat_ID <> rLa(j).Pat_ID THEN gosub unter
 IF roLa(i).ZeitPunkt <> rLa(j).ZeitPunkt THEN gosub unter
 IF roLa(i).FertigStGrad <> rLa(j).FertigStGrad THEN gosub unter
 IF roLa(i).Abkü <> rLa(j).Abkü THEN gosub unter
 IF roLa(i).LangtextVW <> rLa(j).LangtextVW THEN gosub unter
 IF roLa(i).Wert <> rLa(j).Wert THEN gosub unter
 IF roLa(i).Einheit <> rLa(j).Einheit THEN gosub unter
 IF roLa(i).KommentarVW <> rLa(j).KommentarVW THEN gosub unter
 IF roLa(i).AbsPos <> rLa(j).AbsPos THEN gosub unter
 IF roLa(i).AktZeit <> rLa(j).AktZeit THEN gosub unter
 IF roLa(i).Refnr <> rLa(j).Refnr THEN gosub unter
 IF roLa(i).StByte <> rLa(j).StByte THEN gosub unter
 IF roLa(i).ID <> rLa(j).ID THEN gosub unter
 Exit Function
unter:
 LaZUnt = LaZUnt + 1
 Return
End FUNCTION ' LaZUnt

Public FUNCTION laborneuLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLa(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(FertigStGrad,'') FertigStGrad" & _
",COALESCE(Abkü,'') Abkü,COALESCE(LangtextVW,0) LangtextVW,COALESCE(Wert,'') Wert,COALESCE(Einheit,'') Einheit" & _
",COALESCE(KommentarVW,0) KommentarVW,COALESCE(AbsPos,0) AbsPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(Refnr,0) Refnr" & _
",COALESCE(StByte,0) StByte,COALESCE(ID,0) ID FROM `laborneu` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roLa)
  roLa(akt).FID = rs!FID
  roLa(akt).Pat_ID = rs!Pat_ID
  roLa(akt).ZeitPunkt = rs!ZeitPunkt
  roLa(akt).FertigStGrad = doUmwfSQL(rs!FertigStGrad, lies.obMySQL, False)
  roLa(akt).Abkü = doUmwfSQL(rs!Abkü, lies.obMySQL, False)
  roLa(akt).LangtextVW = rs!LangtextVW
  roLa(akt).Wert = doUmwfSQL(rs!Wert, lies.obMySQL, False)
  roLa(akt).Einheit = doUmwfSQL(rs!Einheit, lies.obMySQL, False)
  roLa(akt).KommentarVW = rs!KommentarVW
  roLa(akt).AbsPos = rs!AbsPos
  roLa(akt).AktZeit = rs!AktZeit
  roLa(akt).Refnr = rs!Refnr
  roLa(akt).StByte = rs!StByte
  roLa(akt).ID = rs!ID
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roLa(UBound(roLa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborneuLaden

Function laborneuEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rLa) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rLa)
   IF rLa(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roLa)
    IF roLa(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roLa(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLa(robeg + UBound(rLa) - rbeg)
   For ri = rbeg To UBound(rLa)
    Call roLaZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rLa = roLa
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' laborneuEinf

Public FUNCTION laborneuSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborneu"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `laborneu` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `laborneu` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborneu` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte)           VALUES"))
 For i = 1 to ubound(rLa)
'  rLa(i).AktZeit = now()
  rLa(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLa(i).FID, "," , rLa(i).Pat_ID, "," , DatFor_k(rLa(i).ZeitPunkt), ",'" , rLa(i).FertigStGrad, "','" , rLa(i).Abkü, "'," , rLa(i).LangtextVW, ",'" , rLa(i).Wert, "','" , rLa(i).Einheit, "'," ,  _
   rLa(i).KommentarVW, "," , rLa(i).AbsPos, "," , DatFor_k(rLa(i).AktZeit), "," , rLa(i).Refnr, "," , rLa(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rLa) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLa) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(3)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLa),i)
  IF Len(rLa(k).FertigStGrad) > maxi(0) THEN maxi(0) = Len(rLa(k).FertigStGrad)
  IF Len(rLa(k).Abkü) > maxi(1) THEN maxi(1) = Len(rLa(k).Abkü)
  IF Len(rLa(k).Wert) > maxi(2) THEN maxi(2) = Len(rLa(k).Wert)
  IF Len(rLa(k).Einheit) > maxi(3) THEN maxi(3) = Len(rLa(k).Einheit)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLa), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & Left$(rLa(k).FertigStGrad, maxL)  & "'",true : rLa(k).FertigStGrad = Left$(rLa(k).FertigStGrad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & Left$(rLa(k).Abkü, maxL)  & "'",true : rLa(k).Abkü = Left$(rLa(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & Left$(rLa(k).Wert, maxL)  & "'",true : rLa(k).Wert = Left$(rLa(k).Wert, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & Left$(rLa(k).Einheit, maxL)  & "'",true : rLa(k).Einheit = Left$(rLa(k).Einheit, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborneuSpeichern

Public FUNCTION roLeZuw(i&, j&)
 roLe(i).id = rLe(j).id
 roLe(i).FID = rLe(j).FID
 roLe(i).Pat_ID = rLe(j).Pat_ID
 roLe(i).ZeitPunkt = rLe(j).ZeitPunkt
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
 roLe(i).LANR = rLe(j).LANR
 roLe(i).letzVorg = rLe(j).letzVorg
 roLe(i).Ausn = rLe(j).Ausn
 roLe(i).Beme = rLe(j).Beme
 roLe(i).absPos = rLe(j).absPos
 roLe(i).AktZeit = rLe(j).AktZeit
 roLe(i).QS = rLe(j).QS
 roLe(i).QT = rLe(j).QT
 roLe(i).StByte = rLe(j).StByte
 roLe(i).LANRid = rLe(j).LANRid
 roLe(i).Sachkbez = rLe(j).Sachkbez
 roLe(i).Sachkct = rLe(j).Sachkct
 roLe(i).Zone = rLe(j).Zone
End FUNCTION ' roLeZuw

Public FUNCTION LeZUnt%(i&, j&)
 IF roLe(i).id <> rLe(j).id THEN gosub unter
 IF roLe(i).FID <> rLe(j).FID THEN gosub unter
 IF roLe(i).Pat_ID <> rLe(j).Pat_ID THEN gosub unter
 IF roLe(i).ZeitPunkt <> rLe(j).ZeitPunkt THEN gosub unter
 IF roLe(i).Leistung <> rLe(j).Leistung THEN gosub unter
 IF roLe(i).f5002 <> rLe(j).f5002 THEN gosub unter
 IF roLe(i).f5005 <> rLe(j).f5005 THEN gosub unter
 IF roLe(i).f5006 <> rLe(j).f5006 THEN gosub unter
 IF roLe(i).f5009 <> rLe(j).f5009 THEN gosub unter
 IF roLe(i).Med <> rLe(j).Med THEN gosub unter
 IF roLe(i).f5015 <> rLe(j).f5015 THEN gosub unter
 IF roLe(i).f5016 <> rLe(j).f5016 THEN gosub unter
 IF roLe(i).f5021 <> rLe(j).f5021 THEN gosub unter
 IF roLe(i).f5026 <> rLe(j).f5026 THEN gosub unter
 IF roLe(i).Faktor <> rLe(j).Faktor THEN gosub unter
 IF roLe(i).f5098 <> rLe(j).f5098 THEN gosub unter
 IF roLe(i).Charge <> rLe(j).Charge THEN gosub unter
 IF roLe(i).LANR <> rLe(j).LANR THEN gosub unter
 IF roLe(i).letzVorg <> rLe(j).letzVorg THEN gosub unter
 IF roLe(i).Ausn <> rLe(j).Ausn THEN gosub unter
 IF roLe(i).Beme <> rLe(j).Beme THEN gosub unter
 IF roLe(i).absPos <> rLe(j).absPos THEN gosub unter
 IF roLe(i).AktZeit <> rLe(j).AktZeit THEN gosub unter
 IF roLe(i).QS <> rLe(j).QS THEN gosub unter
 IF roLe(i).QT <> rLe(j).QT THEN gosub unter
 IF roLe(i).StByte <> rLe(j).StByte THEN gosub unter
 IF roLe(i).LANRid <> rLe(j).LANRid THEN gosub unter
 IF roLe(i).Sachkbez <> rLe(j).Sachkbez THEN gosub unter
 IF roLe(i).Sachkct <> rLe(j).Sachkct THEN gosub unter
 IF roLe(i).Zone <> rLe(j).Zone THEN gosub unter
 Exit Function
unter:
 LeZUnt = LeZUnt + 1
 Return
End FUNCTION ' LeZUnt

Public FUNCTION leistungenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLe(1)
 sql = "SELECT COALESCE(id,0) id,COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt" & _
",COALESCE(Leistung,'') Leistung,COALESCE(f5002,'') f5002,COALESCE(f5005,'') f5005,COALESCE(f5006,'') f5006" & _
",COALESCE(f5009,'') f5009,COALESCE(Med,'') Med,COALESCE(f5015,'') f5015,COALESCE(f5016,'') f5016" & _
",COALESCE(f5021,'') f5021,COALESCE(f5026,'') f5026,COALESCE(Faktor,'') Faktor,COALESCE(f5098,'') f5098" & _
",COALESCE(Charge,'') Charge,COALESCE(LANR,'') LANR,IF(letzVorg IS NULL OR letzVorg=0,CONVERT('18991230',DATE),letzVorg) letzVorg,COALESCE(Ausn,'') Ausn" & _
",COALESCE(Beme,'') Beme,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid,COALESCE(Sachkbez,'') Sachkbez" & _
",COALESCE(Sachkct,0) Sachkct,COALESCE(Zone,'') Zone FROM `leistungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roLe)
  roLe(akt).id = rs!id
  roLe(akt).FID = rs!FID
  roLe(akt).Pat_ID = rs!Pat_ID
  roLe(akt).ZeitPunkt = rs!ZeitPunkt
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
  roLe(akt).LANR = doUmwfSQL(rs!LANR, lies.obMySQL, False)
  roLe(akt).letzVorg = rs!letzVorg
  roLe(akt).Ausn = doUmwfSQL(rs!Ausn, lies.obMySQL, False)
  roLe(akt).Beme = doUmwfSQL(rs!Beme, lies.obMySQL, False)
  roLe(akt).absPos = rs!absPos
  roLe(akt).AktZeit = rs!AktZeit
  roLe(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
  roLe(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
  roLe(akt).StByte = rs!StByte
  roLe(akt).LANRid = rs!LANRid
  roLe(akt).Sachkbez = doUmwfSQL(rs!Sachkbez, lies.obMySQL, False)
  roLe(akt).Sachkct = rs!Sachkct
  roLe(akt).Zone = doUmwfSQL(rs!Zone, lies.obMySQL, False)
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roLe(UBound(roLe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' leistungenLaden

Function leistungenEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rLe) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rLe)
   IF rLe(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roLe)
    IF roLe(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roLe(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roLe(robeg + UBound(rLe) - rbeg)
   For ri = rbeg To UBound(rLe)
    Call roLeZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rLe = roLe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' leistungenEinf

Public FUNCTION leistungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere leistungen"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `leistungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `leistungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,Charge,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS," & _
     "QT,StByte,LANRid,Sachkbez,Sachkct,Zone) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `leistungen` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026," & _
     "Faktor,f5098,Charge,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS," & _
     "QT,StByte,LANRid,Sachkbez,Sachkct,Zone)            VALUES"))
 For i = 1 to ubound(rLe)
'  rLe(i).AktZeit = now()
  rLe(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLe(i).FID, "," , rLe(i).Pat_ID, "," , DatFor_k(rLe(i).ZeitPunkt), ",'" , rLe(i).Leistung, "','" , rLe(i).f5002, "','" , rLe(i).f5005, "','" , rLe(i).f5006, "','" , rLe(i).f5009, "','" ,  _
   rLe(i).Med, "','" , rLe(i).f5015, "','" , rLe(i).f5016, "','" , rLe(i).f5021, "','" , rLe(i).f5026, "','" , rLe(i).Faktor, "','" , rLe(i).f5098, "','" , rLe(i).Charge, "','" , rLe(i).LANR, "'," , DatFor_k( _
   rLe(i).letzVorg), ",'" , rLe(i).Ausn, "','" , rLe(i).Beme, "'," , rLe(i).absPos, "," , DatFor_k(rLe(i).AktZeit), ",'" , rLe(i).QS, "','" , rLe(i).QT, "'," , rLe(i).StByte, "," ,  _
   rLe(i).LANRid, ",'" , rLe(i).Sachkbez, "'," , rLe(i).Sachkct, ",'" , rLe(i).Zone, "')")
  IF SammelInsert <> 0 AND i < ubound(rLe) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLe) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(19)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLe),i)
  IF Len(rLe(k).Leistung) > maxi(0) THEN maxi(0) = Len(rLe(k).Leistung)
  IF Len(rLe(k).f5002) > maxi(1) THEN maxi(1) = Len(rLe(k).f5002)
  IF Len(rLe(k).f5005) > maxi(2) THEN maxi(2) = Len(rLe(k).f5005)
  IF Len(rLe(k).f5006) > maxi(3) THEN maxi(3) = Len(rLe(k).f5006)
  IF Len(rLe(k).f5009) > maxi(4) THEN maxi(4) = Len(rLe(k).f5009)
  IF Len(rLe(k).Med) > maxi(5) THEN maxi(5) = Len(rLe(k).Med)
  IF Len(rLe(k).f5015) > maxi(6) THEN maxi(6) = Len(rLe(k).f5015)
  IF Len(rLe(k).f5016) > maxi(7) THEN maxi(7) = Len(rLe(k).f5016)
  IF Len(rLe(k).f5021) > maxi(8) THEN maxi(8) = Len(rLe(k).f5021)
  IF Len(rLe(k).f5026) > maxi(9) THEN maxi(9) = Len(rLe(k).f5026)
  IF Len(rLe(k).Faktor) > maxi(10) THEN maxi(10) = Len(rLe(k).Faktor)
  IF Len(rLe(k).f5098) > maxi(11) THEN maxi(11) = Len(rLe(k).f5098)
  IF Len(rLe(k).Charge) > maxi(12) THEN maxi(12) = Len(rLe(k).Charge)
  IF Len(rLe(k).LANR) > maxi(13) THEN maxi(13) = Len(rLe(k).LANR)
  IF Len(rLe(k).Ausn) > maxi(14) THEN maxi(14) = Len(rLe(k).Ausn)
  IF Len(rLe(k).Beme) > maxi(15) THEN maxi(15) = Len(rLe(k).Beme)
  IF Len(rLe(k).QS) > maxi(16) THEN maxi(16) = Len(rLe(k).QS)
  IF Len(rLe(k).QT) > maxi(17) THEN maxi(17) = Len(rLe(k).QT)
  IF Len(rLe(k).Sachkbez) > maxi(18) THEN maxi(18) = Len(rLe(k).Sachkbez)
  IF Len(rLe(k).Zone) > maxi(19) THEN maxi(19) = Len(rLe(k).Zone)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLe), i)
      SELECT CASE m
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
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.Charge: '" & rLe(k).Charge & "' -> '" & Left$(rLe(k).Charge, maxL)  & "'",true : rLe(k).Charge = Left$(rLe(k).Charge, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).LANR & "' -> '" & Left$(rLe(k).LANR, maxL)  & "'",true : rLe(k).LANR = Left$(rLe(k).LANR, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL)  & "'",true : rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).Beme & "' -> '" & Left$(rLe(k).Beme, maxL)  & "'",true : rLe(k).Beme = Left$(rLe(k).Beme, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL)  & "'",true : rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL)  & "'",true : rLe(k).QT = Left$(rLe(k).QT, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & Left$(rLe(k).Sachkbez, maxL)  & "'",true : rLe(k).Sachkbez = Left$(rLe(k).Sachkbez, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & Left$(rLe(k).Zone, maxL)  & "'",true : rLe(k).Zone = Left$(rLe(k).Zone, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' leistungenSpeichern

Public FUNCTION roMeZuw(i&, j&)
 roMe(i).FID = rMe(j).FID
 roMe(i).Pat_ID = rMe(j).Pat_ID
 roMe(i).MPNr = rMe(j).MPNr
 roMe(i).ZeitPunkt = rMe(j).ZeitPunkt
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
 roMe(i).zn = rMe(j).zn
 roMe(i).bBed = rMe(j).bBed
 roMe(i).Bemerkung = rMe(j).Bemerkung
 roMe(i).Grund = rMe(j).Grund
 roMe(i).Stärke = rMe(j).Stärke
 roMe(i).Einheit = rMe(j).Einheit
 roMe(i).Form = rMe(j).Form
 roMe(i).AbsPos = rMe(j).AbsPos
 roMe(i).AktZeit = rMe(j).AktZeit
 roMe(i).StByte = rMe(j).StByte
 roMe(i).ergaenzt = rMe(j).ergaenzt
End FUNCTION ' roMeZuw

Public FUNCTION MeZUnt%(i&, j&)
 IF roMe(i).FID <> rMe(j).FID THEN gosub unter
 IF roMe(i).Pat_ID <> rMe(j).Pat_ID THEN gosub unter
 IF roMe(i).MPNr <> rMe(j).MPNr THEN gosub unter
 IF roMe(i).ZeitPunkt <> rMe(j).ZeitPunkt THEN gosub unter
 IF roMe(i).Datum <> rMe(j).Datum THEN gosub unter
 IF roMe(i).Medikament <> rMe(j).Medikament THEN gosub unter
 IF roMe(i).MedAnfang <> rMe(j).MedAnfang THEN gosub unter
 IF roMe(i).Wirkstoff <> rMe(j).Wirkstoff THEN gosub unter
 IF roMe(i).PZN <> rMe(j).PZN THEN gosub unter
 IF roMe(i).FeldNr <> rMe(j).FeldNr THEN gosub unter
 IF roMe(i).mo <> rMe(j).mo THEN gosub unter
 IF roMe(i).mi <> rMe(j).mi THEN gosub unter
 IF roMe(i).nm <> rMe(j).nm THEN gosub unter
 IF roMe(i).ab <> rMe(j).ab THEN gosub unter
 IF roMe(i).zn <> rMe(j).zn THEN gosub unter
 IF roMe(i).bBed <> rMe(j).bBed THEN gosub unter
 IF roMe(i).Bemerkung <> rMe(j).Bemerkung THEN gosub unter
 IF roMe(i).Grund <> rMe(j).Grund THEN gosub unter
 IF roMe(i).Stärke <> rMe(j).Stärke THEN gosub unter
 IF roMe(i).Einheit <> rMe(j).Einheit THEN gosub unter
 IF roMe(i).Form <> rMe(j).Form THEN gosub unter
 IF roMe(i).AbsPos <> rMe(j).AbsPos THEN gosub unter
 IF roMe(i).AktZeit <> rMe(j).AktZeit THEN gosub unter
 IF roMe(i).StByte <> rMe(j).StByte THEN gosub unter
 IF roMe(i).ergaenzt <> rMe(j).ergaenzt THEN gosub unter
 Exit Function
unter:
 MeZUnt = MeZUnt + 1
 Return
End FUNCTION ' MeZUnt

Public FUNCTION medplanLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roMe(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(MPNr,0) MPNr,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt" & _
",IF(Datum IS NULL OR Datum=0,CONVERT('18991230',DATE),Datum) Datum,COALESCE(Medikament,'') Medikament,COALESCE(MedAnfang,'') MedAnfang,COALESCE(Wirkstoff,'') Wirkstoff" & _
",COALESCE(PZN,0) PZN,COALESCE(FeldNr,0) FeldNr,COALESCE(mo,'') mo,COALESCE(mi,'') mi" & _
",COALESCE(nm,'') nm,COALESCE(ab,'') ab,COALESCE(zn,'') zn,COALESCE(bBed,0) bBed" & _
",COALESCE(Bemerkung,'') Bemerkung,COALESCE(Grund,'') Grund,COALESCE(Stärke,'') Stärke,COALESCE(Einheit,'') Einheit" & _
",COALESCE(Form,'') Form,COALESCE(AbsPos,0) AbsPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte" & _
",COALESCE(ergaenzt,0) ergaenzt FROM `medplan` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roMe)
  roMe(akt).FID = rs!FID
  roMe(akt).Pat_ID = rs!Pat_ID
  roMe(akt).MPNr = rs!MPNr
  roMe(akt).ZeitPunkt = rs!ZeitPunkt
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
  roMe(akt).zn = doUmwfSQL(rs!zn, lies.obMySQL, False)
  roMe(akt).bBed = rs!bBed
  roMe(akt).Bemerkung = doUmwfSQL(rs!Bemerkung, lies.obMySQL, False)
  roMe(akt).Grund = doUmwfSQL(rs!Grund, lies.obMySQL, False)
  roMe(akt).Stärke = doUmwfSQL(rs!Stärke, lies.obMySQL, False)
  roMe(akt).Einheit = doUmwfSQL(rs!Einheit, lies.obMySQL, False)
  roMe(akt).Form = doUmwfSQL(rs!Form, lies.obMySQL, False)
  roMe(akt).AbsPos = rs!AbsPos
  roMe(akt).AktZeit = rs!AktZeit
  roMe(akt).StByte = rs!StByte
  roMe(akt).ergaenzt = rs!ergaenzt
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roMe(UBound(roMe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' medplanLaden

Function medplanEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rMe) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rMe)
   IF rMe(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roMe)
    IF roMe(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roMe(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roMe(robeg + UBound(rMe) - rbeg)
   For ri = rbeg To UBound(rMe)
    Call roMeZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rMe = roMe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' medplanEinf

Public FUNCTION medplanSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere medplan"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `medplan` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `medplan` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte,ergaenzt) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte,ergaenzt)       VALUES"))
 For i = 1 to ubound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rMe(i).FID, "," , rMe(i).Pat_ID, "," , rMe(i).MPNr, "," , DatFor_k(rMe(i).ZeitPunkt), "," , DatFor_k(rMe(i).Datum), ",'" , rMe(i).Medikament, "','" , rMe(i).MedAnfang, "','" , rMe(i).Wirkstoff, "'," ,  _
   rMe(i).PZN, "," , rMe(i).FeldNr, ",'" , rMe(i).mo, "','" , rMe(i).mi, "','" , rMe(i).nm, "','" , rMe(i).ab, "','" , rMe(i).zn, "'," , cstr(-(rMe(i).bBed<>0)) , ",'" , rMe(i).Bemerkung, "','" ,  _
   rMe(i).Grund, "','" , rMe(i).Stärke, "','" , rMe(i).Einheit, "','" , rMe(i).Form, "'," , rMe(i).AbsPos, "," , DatFor_k(rMe(i).AktZeit), "," , rMe(i).StByte, "," , cstr(-(rMe(i).ergaenzt<>0)) , ")")
  IF SammelInsert <> 0 AND i < ubound(rMe) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rMe) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(12)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rMe),i)
  IF Len(rMe(k).Medikament) > maxi(0) THEN maxi(0) = Len(rMe(k).Medikament)
  IF Len(rMe(k).MedAnfang) > maxi(1) THEN maxi(1) = Len(rMe(k).MedAnfang)
  IF Len(rMe(k).Wirkstoff) > maxi(2) THEN maxi(2) = Len(rMe(k).Wirkstoff)
  IF Len(rMe(k).mo) > maxi(3) THEN maxi(3) = Len(rMe(k).mo)
  IF Len(rMe(k).mi) > maxi(4) THEN maxi(4) = Len(rMe(k).mi)
  IF Len(rMe(k).nm) > maxi(5) THEN maxi(5) = Len(rMe(k).nm)
  IF Len(rMe(k).ab) > maxi(6) THEN maxi(6) = Len(rMe(k).ab)
  IF Len(rMe(k).zn) > maxi(7) THEN maxi(7) = Len(rMe(k).zn)
  IF Len(rMe(k).Bemerkung) > maxi(8) THEN maxi(8) = Len(rMe(k).Bemerkung)
  IF Len(rMe(k).Grund) > maxi(9) THEN maxi(9) = Len(rMe(k).Grund)
  IF Len(rMe(k).Stärke) > maxi(10) THEN maxi(10) = Len(rMe(k).Stärke)
  IF Len(rMe(k).Einheit) > maxi(11) THEN maxi(11) = Len(rMe(k).Einheit)
  IF Len(rMe(k).Form) > maxi(12) THEN maxi(12) = Len(rMe(k).Form)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rMe), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & Left$(rMe(k).Medikament, maxL)  & "'",true : rMe(k).Medikament = Left$(rMe(k).Medikament, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & Left$(rMe(k).MedAnfang, maxL)  & "'",true : rMe(k).MedAnfang = Left$(rMe(k).MedAnfang, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.Wirkstoff: '" & rMe(k).Wirkstoff & "' -> '" & Left$(rMe(k).Wirkstoff, maxL)  & "'",true : rMe(k).Wirkstoff = Left$(rMe(k).Wirkstoff, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & Left$(rMe(k).mo, maxL)  & "'",true : rMe(k).mo = Left$(rMe(k).mo, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & Left$(rMe(k).mi, maxL)  & "'",true : rMe(k).mi = Left$(rMe(k).mi, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & Left$(rMe(k).nm, maxL)  & "'",true : rMe(k).nm = Left$(rMe(k).nm, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & Left$(rMe(k).ab, maxL)  & "'",true : rMe(k).ab = Left$(rMe(k).ab, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).zn & "' -> '" & Left$(rMe(k).zn, maxL)  & "'",true : rMe(k).zn = Left$(rMe(k).zn, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & Left$(rMe(k).Bemerkung, maxL)  & "'",true : rMe(k).Bemerkung = Left$(rMe(k).Bemerkung, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rMe.Grund: '" & rMe(k).Grund & "' -> '" & Left$(rMe(k).Grund, maxL)  & "'",true : rMe(k).Grund = Left$(rMe(k).Grund, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rMe.Stärke: '" & rMe(k).Stärke & "' -> '" & Left$(rMe(k).Stärke, maxL)  & "'",true : rMe(k).Stärke = Left$(rMe(k).Stärke, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rMe.Einheit: '" & rMe(k).Einheit & "' -> '" & Left$(rMe(k).Einheit, maxL)  & "'",true : rMe(k).Einheit = Left$(rMe(k).Einheit, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rMe.Form: '" & rMe(k).Form & "' -> '" & Left$(rMe(k).Form, maxL)  & "'",true : rMe(k).Form = Left$(rMe(k).Form, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' medplanSpeichern

Public FUNCTION roReZuw(i&, j&)
 roRe(i).FID = rRe(j).FID
 roRe(i).Pat_ID = rRe(j).Pat_ID
 roRe(i).ZeitPunkt = rRe(j).ZeitPunkt
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
 roRe(i).LANRid = rRe(j).LANRid
 roRe(i).id = rRe(j).id
End FUNCTION ' roReZuw

Public FUNCTION ReZUnt%(i&, j&)
 IF roRe(i).FID <> rRe(j).FID THEN gosub unter
 IF roRe(i).Pat_ID <> rRe(j).Pat_ID THEN gosub unter
 IF roRe(i).ZeitPunkt <> rRe(j).ZeitPunkt THEN gosub unter
 IF roRe(i).Rezept <> rRe(j).Rezept THEN gosub unter
 IF roRe(i).RKlnm <> rRe(j).RKlnm THEN gosub unter
 IF roRe(i).Rezeptklasse <> rRe(j).Rezeptklasse THEN gosub unter
 IF roRe(i).Rezklkurz <> rRe(j).Rezklkurz THEN gosub unter
 IF roRe(i).Rezkllang <> rRe(j).Rezkllang THEN gosub unter
 IF roRe(i).kbez <> rRe(j).kbez THEN gosub unter
 IF roRe(i).Medikament <> rRe(j).Medikament THEN gosub unter
 IF roRe(i).auti <> rRe(j).auti THEN gosub unter
 IF roRe(i).anzl <> rRe(j).anzl THEN gosub unter
 IF roRe(i).PZN <> rRe(j).PZN THEN gosub unter
 IF roRe(i).absPos <> rRe(j).absPos THEN gosub unter
 IF roRe(i).AktZeit <> rRe(j).AktZeit THEN gosub unter
 IF roRe(i).QS <> rRe(j).QS THEN gosub unter
 IF roRe(i).QT <> rRe(j).QT THEN gosub unter
 IF roRe(i).StByte <> rRe(j).StByte THEN gosub unter
 IF roRe(i).LANRid <> rRe(j).LANRid THEN gosub unter
 IF roRe(i).id <> rRe(j).id THEN gosub unter
 Exit Function
unter:
 ReZUnt = ReZUnt + 1
 Return
End FUNCTION ' ReZUnt

Public FUNCTION rezepteintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roRe(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Rezept,'') Rezept" & _
",COALESCE(RKlnm,'') RKlnm,COALESCE(Rezeptklasse,'') Rezeptklasse,COALESCE(Rezklkurz,'') Rezklkurz,COALESCE(Rezkllang,'') Rezkllang" & _
",COALESCE(kbez,'') kbez,COALESCE(Medikament,'') Medikament,COALESCE(auti,0) auti,COALESCE(anzl,0) anzl" & _
",COALESCE(PZN,'') PZN,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid,COALESCE(id,0) id" & _
" FROM `rezepteintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roRe)
  roRe(akt).FID = rs!FID
  roRe(akt).Pat_ID = rs!Pat_ID
  roRe(akt).ZeitPunkt = rs!ZeitPunkt
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
  roRe(akt).LANRid = rs!LANRid
  roRe(akt).id = rs!id
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roRe(UBound(roRe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' rezepteintraegeLaden

Function rezepteintraegeEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rRe) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rRe)
   IF rRe(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roRe)
    IF roRe(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roRe(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roRe(robeg + UBound(rRe) - rbeg)
   For ri = rbeg To UBound(rRe)
    Call roReZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rRe = roRe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' rezepteintraegeEinf

Public FUNCTION rezepteintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere rezepteintraege"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `rezepteintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `rezepteintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid)  VALUES"))
 For i = 1 to ubound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rRe(i).FID, "," , rRe(i).Pat_ID, "," , DatFor_k(rRe(i).ZeitPunkt), ",'" , rRe(i).Rezept, "','" , rRe(i).RKlnm, "','" , rRe(i).Rezeptklasse, "','" , rRe(i).Rezklkurz, "','" , rRe(i).Rezkllang, "','" ,  _
   rRe(i).kbez, "','" , rRe(i).Medikament, "'," , rRe(i).auti, "," , rRe(i).anzl, ",'" , rRe(i).PZN, "'," , rRe(i).absPos, "," , DatFor_k(rRe(i).AktZeit), ",'" , rRe(i).QS, "','" ,  _
   rRe(i).QT, "'," , rRe(i).StByte, "," , rRe(i).LANRid, ")")
  IF SammelInsert <> 0 AND i < ubound(rRe) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rRe) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(9)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rRe),i)
  IF Len(rRe(k).Rezept) > maxi(0) THEN maxi(0) = Len(rRe(k).Rezept)
  IF Len(rRe(k).RKlnm) > maxi(1) THEN maxi(1) = Len(rRe(k).RKlnm)
  IF Len(rRe(k).Rezeptklasse) > maxi(2) THEN maxi(2) = Len(rRe(k).Rezeptklasse)
  IF Len(rRe(k).Rezklkurz) > maxi(3) THEN maxi(3) = Len(rRe(k).Rezklkurz)
  IF Len(rRe(k).Rezkllang) > maxi(4) THEN maxi(4) = Len(rRe(k).Rezkllang)
  IF Len(rRe(k).kbez) > maxi(5) THEN maxi(5) = Len(rRe(k).kbez)
  IF Len(rRe(k).Medikament) > maxi(6) THEN maxi(6) = Len(rRe(k).Medikament)
  IF Len(rRe(k).PZN) > maxi(7) THEN maxi(7) = Len(rRe(k).PZN)
  IF Len(rRe(k).QS) > maxi(8) THEN maxi(8) = Len(rRe(k).QS)
  IF Len(rRe(k).QT) > maxi(9) THEN maxi(9) = Len(rRe(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rRe), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & Left$(rRe(k).Rezept, maxL)  & "'",true : rRe(k).Rezept = Left$(rRe(k).Rezept, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.RKlnm: '" & rRe(k).RKlnm & "' -> '" & Left$(rRe(k).RKlnm, maxL)  & "'",true : rRe(k).RKlnm = Left$(rRe(k).RKlnm, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & Left$(rRe(k).Rezeptklasse, maxL)  & "'",true : rRe(k).Rezeptklasse = Left$(rRe(k).Rezeptklasse, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezklkurz: '" & rRe(k).Rezklkurz & "' -> '" & Left$(rRe(k).Rezklkurz, maxL)  & "'",true : rRe(k).Rezklkurz = Left$(rRe(k).Rezklkurz, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezkllang: '" & rRe(k).Rezkllang & "' -> '" & Left$(rRe(k).Rezkllang, maxL)  & "'",true : rRe(k).Rezkllang = Left$(rRe(k).Rezkllang, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.kbez: '" & rRe(k).kbez & "' -> '" & Left$(rRe(k).kbez, maxL)  & "'",true : rRe(k).kbez = Left$(rRe(k).kbez, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & Left$(rRe(k).Medikament, maxL)  & "'",true : rRe(k).Medikament = Left$(rRe(k).Medikament, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & Left$(rRe(k).PZN, maxL)  & "'",true : rRe(k).PZN = Left$(rRe(k).PZN, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & Left$(rRe(k).QS, maxL)  & "'",true : rRe(k).QS = Left$(rRe(k).QS, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & Left$(rRe(k).QT, maxL)  & "'",true : rRe(k).QT = Left$(rRe(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' rezepteintraegeSpeichern

Public FUNCTION roRrZuw(i&, j&)
 roRr(i).FID = rRr(j).FID
 roRr(i).Pat_ID = rRr(j).Pat_ID
 roRr(i).ZeitPunkt = rRr(j).ZeitPunkt
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
End FUNCTION ' roRrZuw

Public FUNCTION RrZUnt%(i&, j&)
 IF roRr(i).FID <> rRr(j).FID THEN gosub unter
 IF roRr(i).Pat_ID <> rRr(j).Pat_ID THEN gosub unter
 IF roRr(i).ZeitPunkt <> rRr(j).ZeitPunkt THEN gosub unter
 IF roRr(i).RR <> rRr(j).RR THEN gosub unter
 IF roRr(i).Puls <> rRr(j).Puls THEN gosub unter
 IF roRr(i).RRsyst <> rRr(j).RRsyst THEN gosub unter
 IF roRr(i).RRdiast <> rRr(j).RRdiast THEN gosub unter
 IF roRr(i).RRzahl <> rRr(j).RRzahl THEN gosub unter
 IF roRr(i).Quelle <> rRr(j).Quelle THEN gosub unter
 IF roRr(i).Bemerkung <> rRr(j).Bemerkung THEN gosub unter
 IF roRr(i).absPos <> rRr(j).absPos THEN gosub unter
 IF roRr(i).AktZeit <> rRr(j).AktZeit THEN gosub unter
 IF roRr(i).StByte <> rRr(j).StByte THEN gosub unter
 Exit Function
unter:
 RrZUnt = RrZUnt + 1
 Return
End FUNCTION ' RrZUnt

Public FUNCTION rrLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roRr(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(RR,'') RR" & _
",COALESCE(Puls,0) Puls,COALESCE(RRsyst,0) RRsyst,COALESCE(RRdiast,0) RRdiast,COALESCE(RRzahl,0) RRzahl" & _
",COALESCE(Quelle,'') Quelle,COALESCE(Bemerkung,'') Bemerkung,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `rr` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roRr)
  roRr(akt).FID = rs!FID
  roRr(akt).Pat_ID = rs!Pat_ID
  roRr(akt).ZeitPunkt = rs!ZeitPunkt
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
  IF Not rs.EOF THEN ReDim Preserve roRr(UBound(roRr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' rrLaden

Function rrEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rRr) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rRr)
   IF rRr(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roRr)
    IF roRr(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roRr(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roRr(robeg + UBound(rRr) - rbeg)
   For ri = rbeg To UBound(rRr)
    Call roRrZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rRr = roRr
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' rrEinf

Public FUNCTION rrSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere rr"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `rr` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `rr` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte)          VALUES"))
 For i = 1 to ubound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rRr(i).FID, "," , rRr(i).Pat_ID, "," , DatFor_k(rRr(i).ZeitPunkt), ",'" , rRr(i).RR, "'," , rRr(i).Puls, "," , rRr(i).RRsyst, "," , rRr(i).RRdiast, "," , rRr(i).RRzahl, ",'" , rRr(i).Quelle, "','" ,  _
   rRr(i).Bemerkung, "'," , rRr(i).absPos, "," , DatFor_k(rRr(i).AktZeit), "," , rRr(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rRr) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rRr) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rRr),i)
  IF Len(rRr(k).RR) > maxi(0) THEN maxi(0) = Len(rRr(k).RR)
  IF Len(rRr(k).Quelle) > maxi(1) THEN maxi(1) = Len(rRr(k).Quelle)
  IF Len(rRr(k).Bemerkung) > maxi(2) THEN maxi(2) = Len(rRr(k).Bemerkung)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rRr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & Left$(rRr(k).RR, maxL)  & "'",true : rRr(k).RR = Left$(rRr(k).RR, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.Quelle: '" & rRr(k).Quelle & "' -> '" & Left$(rRr(k).Quelle, maxL)  & "'",true : rRr(k).Quelle = Left$(rRr(k).Quelle, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRr.Bemerkung: '" & rRr(k).Bemerkung & "' -> '" & Left$(rRr(k).Bemerkung, maxL)  & "'",true : rRr(k).Bemerkung = Left$(rRr(k).Bemerkung, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' rrSpeichern

Public FUNCTION kvnrueSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere kvnrue"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `kvnrue` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `kvnrue` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `kvnrue` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        VALUES"))
 For i = 1 to ubound(rKv)
'  rKv(i).AktZeit = now()
  rKv(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rKv(i).Pat_ID, ",'" , rKv(i).KVNr, "'," , rKv(i).absPos, "," , DatFor_k(rKv(i).AktZeit), "," , rKv(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rKv) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rKv) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rKv),i)
  IF Len(rKv(k).KVNr) > maxi(0) THEN maxi(0) = Len(rKv(k).KVNr)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rKv), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & Left$(rKv(k).KVNr, maxL)  & "'",true : rKv(k).KVNr = Left$(rKv(k).KVNr, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' kvnrueSpeichern

Public FUNCTION unbekannte_kennungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 syscmd 4, pid & ": Speichere unbekannte_kennungen"
 IF not Allepat THEN
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `unbekannte kennungen` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt)         VALUES"))
 For i = rUn1 + 1 to ubound(rUn)
'  rUn(i).AktZeit = now()
  rUn(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = rUn1 + 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rUn(i).Kennung, "'," , rUn(i).absPos, "," , rUn(i).StByte, "," , rUn(i).Pat_id, ",'" , rUn(i).Inhalt, "')")
  IF SammelInsert <> 0 AND i < ubound(rUn) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rUn) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 rUn1 = ubound(rUn)
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,rUn1 + 1,i) to iif(SammelInsert<>0,ubound(rUn),i)
  IF Len(rUn(k).Kennung) > maxi(0) THEN maxi(0) = Len(rUn(k).Kennung)
  IF Len(rUn(k).Inhalt) > maxi(1) THEN maxi(1) = Len(rUn(k).Inhalt)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,rUn1 + 1, i) To IIf(SammelInsert <> 0,ubound(rUn), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & Left$(rUn(k).Kennung, maxL)  & "'",true : rUn(k).Kennung = Left$(rUn(k).Kennung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & Left$(rUn(k).Inhalt, maxL)  & "'",true : rUn(k).Inhalt = Left$(rUn(k).Inhalt, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' unbekannte_kennungenSpeichern

Public FUNCTION roDmZuw(i&, j&)
 roDm(i).Abk = rDm(j).Abk
 roDm(i).Art = rDm(j).Art
 roDm(i).KarteiDatum = rDm(j).KarteiDatum
 roDm(i).exportiert = rDm(j).exportiert
 roDm(i).DokuDatum = rDm(j).DokuDatum
 roDm(i).obvoll = rDm(j).obvoll
 roDm(i).ok = rDm(j).ok
 roDm(i).ausgedruckt = rDm(j).ausgedruckt
 roDm(i).NachName = rDm(j).NachName
 roDm(i).VorName = rDm(j).VorName
 roDm(i).GebDat = rDm(j).GebDat
 roDm(i).Pat_id = rDm(j).Pat_id
 roDm(i).StByte = rDm(j).StByte
 roDm(i).AktZeit = rDm(j).AktZeit
 roDm(i).lanrid = rDm(j).lanrid
 roDm(i).Zusatzdaten = rDm(j).Zusatzdaten
End FUNCTION ' roDmZuw

Public FUNCTION DmZUnt%(i&, j&)
 IF roDm(i).Abk <> rDm(j).Abk THEN gosub unter
 IF roDm(i).Art <> rDm(j).Art THEN gosub unter
 IF roDm(i).KarteiDatum <> rDm(j).KarteiDatum THEN gosub unter
 IF roDm(i).exportiert <> rDm(j).exportiert THEN gosub unter
 IF roDm(i).DokuDatum <> rDm(j).DokuDatum THEN gosub unter
 IF roDm(i).obvoll <> rDm(j).obvoll THEN gosub unter
 IF roDm(i).ok <> rDm(j).ok THEN gosub unter
 IF roDm(i).ausgedruckt <> rDm(j).ausgedruckt THEN gosub unter
 IF roDm(i).NachName <> rDm(j).NachName THEN gosub unter
 IF roDm(i).VorName <> rDm(j).VorName THEN gosub unter
 IF roDm(i).GebDat <> rDm(j).GebDat THEN gosub unter
 IF roDm(i).Pat_id <> rDm(j).Pat_id THEN gosub unter
 IF roDm(i).StByte <> rDm(j).StByte THEN gosub unter
 IF roDm(i).AktZeit <> rDm(j).AktZeit THEN gosub unter
 IF roDm(i).lanrid <> rDm(j).lanrid THEN gosub unter
 IF roDm(i).Zusatzdaten <> rDm(j).Zusatzdaten THEN gosub unter
 Exit Function
unter:
 DmZUnt = DmZUnt + 1
 Return
End FUNCTION ' DmZUnt

Public FUNCTION dmpreiheLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDm(1)
 sql = "SELECT COALESCE(Abk,'') Abk,COALESCE(Art,'') Art,IF(KarteiDatum IS NULL OR KarteiDatum=0,CONVERT('18991230',DATE),KarteiDatum) KarteiDatum,IF(exportiert IS NULL OR exportiert=0,CONVERT('18991230',DATE),exportiert) exportiert" & _
",IF(DokuDatum IS NULL OR DokuDatum=0,CONVERT('18991230',DATE),DokuDatum) DokuDatum,COALESCE(obvoll,0) obvoll,COALESCE(ok,0) ok,COALESCE(ausgedruckt,0) ausgedruckt" & _
",COALESCE(NachName,'') NachName,COALESCE(VorName,'') VorName,IF(GebDat IS NULL OR GebDat=0,CONVERT('18991230',DATE),GebDat) GebDat,COALESCE(Pat_id,0) Pat_id" & _
",COALESCE(StByte,0) StByte,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(lanrid,0) lanrid,COALESCE(Zusatzdaten,'') Zusatzdaten" & _
" FROM `dmpreihe` WHERE Pat_ID=" & pid & " ORDER BY `Dokudatum`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roDm)
  roDm(akt).Abk = doUmwfSQL(rs!Abk, lies.obMySQL, False)
  roDm(akt).Art = doUmwfSQL(rs!Art, lies.obMySQL, False)
  roDm(akt).KarteiDatum = rs!KarteiDatum
  roDm(akt).exportiert = rs!exportiert
  roDm(akt).DokuDatum = rs!DokuDatum
  roDm(akt).obvoll = rs!obvoll
  roDm(akt).ok = rs!ok
  roDm(akt).ausgedruckt = rs!ausgedruckt
  roDm(akt).NachName = doUmwfSQL(rs!NachName, lies.obMySQL, False)
  roDm(akt).VorName = doUmwfSQL(rs!VorName, lies.obMySQL, False)
  roDm(akt).GebDat = rs!GebDat
  roDm(akt).Pat_id = rs!Pat_id
  roDm(akt).StByte = rs!StByte
  roDm(akt).AktZeit = rs!AktZeit
  roDm(akt).lanrid = rs!lanrid
  roDm(akt).Zusatzdaten = doUmwfSQL(rs!Zusatzdaten, lies.obMySQL, False)
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roDm(UBound(roDm) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' dmpreiheLaden

Function dmpreiheEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rDm) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rDm)
   IF rDm(ri).Dokudatum >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roDm)
    IF roDm(robeg).Dokudatum >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roDm(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDm(robeg + UBound(rDm) - rbeg)
   For ri = rbeg To UBound(rDm)
    Call roDmZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rDm = roDm
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' dmpreiheEinf

Public FUNCTION dmpreiheSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere dmpreihe"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `dmpreihe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `dmpreihe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten)          VALUES"))
 For i = 1 to ubound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rDm(i).Abk, "','" , rDm(i).Art, "'," , DatFor_k(rDm(i).KarteiDatum), "," , DatFor_k(rDm(i).exportiert), "," , DatFor_k(rDm(i).DokuDatum), "," , cstr(-(rDm(i).obvoll<>0)) , "," , cstr(-( _
   rDm(i).ok<>0)) , "," , cstr(-(rDm(i).ausgedruckt<>0)) , ",'" , rDm(i).NachName, "','" , rDm(i).VorName, "'," , DatFor_k(rDm(i).GebDat), "," , rDm(i).Pat_id, "," , rDm(i).StByte, "," , DatFor_k(rDm(i).AktZeit), "," ,  _
   rDm(i).lanrid, ",'" , rDm(i).Zusatzdaten, "')")
  IF SammelInsert <> 0 AND i < ubound(rDm) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rDm) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(4)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDm),i)
  IF Len(rDm(k).Abk) > maxi(0) THEN maxi(0) = Len(rDm(k).Abk)
  IF Len(rDm(k).Art) > maxi(1) THEN maxi(1) = Len(rDm(k).Art)
  IF Len(rDm(k).NachName) > maxi(2) THEN maxi(2) = Len(rDm(k).NachName)
  IF Len(rDm(k).VorName) > maxi(3) THEN maxi(3) = Len(rDm(k).VorName)
  IF Len(rDm(k).Zusatzdaten) > maxi(4) THEN maxi(4) = Len(rDm(k).Zusatzdaten)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDm), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL)  & "'",true : rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).Art & "' -> '" & Left$(rDm(k).Art, maxL)  & "'",true : rDm(k).Art = Left$(rDm(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).NachName & "' -> '" & Left$(rDm(k).NachName, maxL)  & "'",true : rDm(k).NachName = Left$(rDm(k).NachName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).VorName & "' -> '" & Left$(rDm(k).VorName, maxL)  & "'",true : rDm(k).VorName = Left$(rDm(k).VorName, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDm.Zusatzdaten: '" & rDm(k).Zusatzdaten & "' -> '" & Left$(rDm(k).Zusatzdaten, maxL)  & "'",true : rDm(k).Zusatzdaten = Left$(rDm(k).Zusatzdaten, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' dmpreiheSpeichern

Public FUNCTION roDeZuw(i&, j&)
 roDe(i).id = rDe(j).id
 roDe(i).IDS = rDe(j).IDS
 roDe(i).Pat_ID = rDe(j).Pat_ID
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
 roDe(i).titel = rDe(j).titel
 roDe(i).toolTipText = rDe(j).toolTipText
 roDe(i).verankert = rDe(j).verankert
 roDe(i).absPos = rDe(j).absPos
 roDe(i).AktZeit = rDe(j).AktZeit
 roDe(i).StByte = rDe(j).StByte
End FUNCTION ' roDeZuw

Public FUNCTION DeZUnt%(i&, j&)
 IF roDe(i).id <> rDe(j).id THEN gosub unter
 IF roDe(i).IDS <> rDe(j).IDS THEN gosub unter
 IF roDe(i).Pat_ID <> rDe(j).Pat_ID THEN gosub unter
 IF roDe(i).erstZP <> rDe(j).erstZP THEN gosub unter
 IF roDe(i).exoL <> rDe(j).exoL THEN gosub unter
 IF roDe(i).hideT <> rDe(j).hideT THEN gosub unter
 IF roDe(i).iconPath <> rDe(j).iconPath THEN gosub unter
 IF roDe(i).noteBkColor <> rDe(j).noteBkColor THEN gosub unter
 IF roDe(i).noteFgColor <> rDe(j).noteFgColor THEN gosub unter
 IF roDe(i).positionBottom <> rDe(j).positionBottom THEN gosub unter
 IF roDe(i).positionLeft <> rDe(j).positionLeft THEN gosub unter
 IF roDe(i).positionRight <> rDe(j).positionRight THEN gosub unter
 IF roDe(i).positionTop <> rDe(j).positionTop THEN gosub unter
 IF roDe(i).showAsNote <> rDe(j).showAsNote THEN gosub unter
 IF roDe(i).syncInfoList <> rDe(j).syncInfoList THEN gosub unter
 IF roDe(i).titel <> rDe(j).titel THEN gosub unter
 IF roDe(i).toolTipText <> rDe(j).toolTipText THEN gosub unter
 IF roDe(i).verankert <> rDe(j).verankert THEN gosub unter
 IF roDe(i).absPos <> rDe(j).absPos THEN gosub unter
 IF roDe(i).AktZeit <> rDe(j).AktZeit THEN gosub unter
 IF roDe(i).StByte <> rDe(j).StByte THEN gosub unter
 Exit Function
unter:
 DeZUnt = DeZUnt + 1
 Return
End FUNCTION ' DeZUnt

Public FUNCTION desktopLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDe(1)
 sql = "SELECT COALESCE(id,0) id,COALESCE(IDS,'') IDS,COALESCE(Pat_ID,0) Pat_ID,IF(erstZP IS NULL OR erstZP=0,CONVERT('18991230',DATE),erstZP) erstZP" & _
",COALESCE(exoL,'') exoL,COALESCE(hideT,0) hideT,COALESCE(iconPath,'') iconPath,COALESCE(noteBkColor,0) noteBkColor" & _
",COALESCE(noteFgColor,0) noteFgColor,COALESCE(positionBottom,0) positionBottom,COALESCE(positionLeft,0) positionLeft,COALESCE(positionRight,0) positionRight" & _
",COALESCE(positionTop,0) positionTop,COALESCE(showAsNote,0) showAsNote,COALESCE(syncInfoList,'') syncInfoList,COALESCE(titel,'') titel" & _
",COALESCE(toolTipText,'') toolTipText,COALESCE(verankert,0) verankert,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit" & _
",COALESCE(StByte,0) StByte FROM `desktop` WHERE Pat_ID=" & pid & " ORDER BY `erstZP`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roDe)
  roDe(akt).id = rs!id
  roDe(akt).IDS = doUmwfSQL(rs!IDS, lies.obMySQL, False)
  roDe(akt).Pat_ID = rs!Pat_ID
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
  roDe(akt).titel = doUmwfSQL(rs!titel, lies.obMySQL, False)
  roDe(akt).toolTipText = doUmwfSQL(rs!toolTipText, lies.obMySQL, False)
  roDe(akt).verankert = rs!verankert
  roDe(akt).absPos = rs!absPos
  roDe(akt).AktZeit = rs!AktZeit
  roDe(akt).StByte = rs!StByte
  rs.MoveNext
  IF Not rs.EOF THEN ReDim Preserve roDe(UBound(roDe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' desktopLaden

Function desktopEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rDe) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rDe)
   IF rDe(ri).erstZP >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roDe)
    IF roDe(robeg).erstZP >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roDe(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roDe(robeg + UBound(rDe) - rbeg)
   For ri = rbeg To UBound(rDe)
    Call roDeZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rDe = roDe
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' desktopEinf

Public FUNCTION desktopSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere desktop"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `desktop` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `desktop` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte)  VALUES"))
 For i = 1 to ubound(rDe)
'  rDe(i).AktZeit = now()
  rDe(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rDe(i).IDS, "'," , rDe(i).Pat_ID, "," , DatFor_k(rDe(i).erstZP), ",'" , rDe(i).exoL, "'," , rDe(i).hideT, ",'" , rDe(i).iconPath, "'," , rDe(i).noteBkColor, "," , rDe(i).noteFgColor, "," ,  _
   rDe(i).positionBottom, "," , rDe(i).positionLeft, "," , rDe(i).positionRight, "," , rDe(i).positionTop, "," , rDe(i).showAsNote, ",'" , rDe(i).syncInfoList, "','" , rDe(i).titel, "','" ,  _
   rDe(i).toolTipText, "'," , rDe(i).verankert, "," , rDe(i).absPos, "," , DatFor_k(rDe(i).AktZeit), "," , rDe(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rDe) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rDe) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(5)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDe),i)
  IF Len(rDe(k).IDS) > maxi(0) THEN maxi(0) = Len(rDe(k).IDS)
  IF Len(rDe(k).exoL) > maxi(1) THEN maxi(1) = Len(rDe(k).exoL)
  IF Len(rDe(k).iconPath) > maxi(2) THEN maxi(2) = Len(rDe(k).iconPath)
  IF Len(rDe(k).syncInfoList) > maxi(3) THEN maxi(3) = Len(rDe(k).syncInfoList)
  IF Len(rDe(k).titel) > maxi(4) THEN maxi(4) = Len(rDe(k).titel)
  IF Len(rDe(k).toolTipText) > maxi(5) THEN maxi(5) = Len(rDe(k).toolTipText)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "desktop", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "desktop", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDe), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDe.IDS: '" & rDe(k).IDS & "' -> '" & Left$(rDe(k).IDS, maxL)  & "'",true : rDe(k).IDS = Left$(rDe(k).IDS, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDe.exoL: '" & rDe(k).exoL & "' -> '" & Left$(rDe(k).exoL, maxL)  & "'",true : rDe(k).exoL = Left$(rDe(k).exoL, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDe.iconPath: '" & rDe(k).iconPath & "' -> '" & Left$(rDe(k).iconPath, maxL)  & "'",true : rDe(k).iconPath = Left$(rDe(k).iconPath, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDe.syncInfoList: '" & rDe(k).syncInfoList & "' -> '" & Left$(rDe(k).syncInfoList, maxL)  & "'",true : rDe(k).syncInfoList = Left$(rDe(k).syncInfoList, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDe.titel: '" & rDe(k).titel & "' -> '" & Left$(rDe(k).titel, maxL)  & "'",true : rDe(k).titel = Left$(rDe(k).titel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDe.toolTipText: '" & rDe(k).toolTipText & "' -> '" & Left$(rDe(k).toolTipText, maxL)  & "'",true : rDe(k).toolTipText = Left$(rDe(k).toolTipText, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' desktopSpeichern

Public FUNCTION roUsZuw(i&, j&)
 roUs(i).FID = rUs(j).FID
 roUs(i).Pat_ID = rUs(j).Pat_ID
 roUs(i).ZeitPunkt = rUs(j).ZeitPunkt
 roUs(i).Art = rUs(j).Art
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
End FUNCTION ' roUsZuw

Public FUNCTION UsZUnt%(i&, j&)
 IF roUs(i).FID <> rUs(j).FID THEN gosub unter
 IF roUs(i).Pat_ID <> rUs(j).Pat_ID THEN gosub unter
 IF roUs(i).ZeitPunkt <> rUs(j).ZeitPunkt THEN gosub unter
 IF roUs(i).Art <> rUs(j).Art THEN gosub unter
 IF roUs(i).Spritzst <> rUs(j).Spritzst THEN gosub unter
 IF roUs(i).Fußbef_re <> rUs(j).Fußbef_re THEN gosub unter
 IF roUs(i).Fußbef_li <> rUs(j).Fußbef_li THEN gosub unter
 IF roUs(i).Hyperk_re <> rUs(j).Hyperk_re THEN gosub unter
 IF roUs(i).Hyperk_li <> rUs(j).Hyperk_li THEN gosub unter
 IF roUs(i).Ulcera_re <> rUs(j).Ulcera_re THEN gosub unter
 IF roUs(i).Ulcera_li <> rUs(j).Ulcera_li THEN gosub unter
 IF roUs(i).Kraft_Zh_re <> rUs(j).Kraft_Zh_re THEN gosub unter
 IF roUs(i).Kraft_Zh_li <> rUs(j).Kraft_Zh_li THEN gosub unter
 IF roUs(i).Kraft_Zb_re <> rUs(j).Kraft_Zb_re THEN gosub unter
 IF roUs(i).Kraft_Zb_li <> rUs(j).Kraft_Zb_li THEN gosub unter
 IF roUs(i).Kraft_Knie_re <> rUs(j).Kraft_Knie_re THEN gosub unter
 IF roUs(i).Kraft_Knie_li <> rUs(j).Kraft_Knie_li THEN gosub unter
 IF roUs(i).ASR_re <> rUs(j).ASR_re THEN gosub unter
 IF roUs(i).ASR_li <> rUs(j).ASR_li THEN gosub unter
 IF roUs(i).PSR_re <> rUs(j).PSR_re THEN gosub unter
 IF roUs(i).PSR_li <> rUs(j).PSR_li THEN gosub unter
 IF roUs(i).Oberfl_re <> rUs(j).Oberfl_re THEN gosub unter
 IF roUs(i).Oberfl_li <> rUs(j).Oberfl_li THEN gosub unter
 IF roUs(i).MF_re <> rUs(j).MF_re THEN gosub unter
 IF roUs(i).MF_li <> rUs(j).MF_li THEN gosub unter
 IF roUs(i).KW_re <> rUs(j).KW_re THEN gosub unter
 IF roUs(i).KW_li <> rUs(j).KW_li THEN gosub unter
 IF roUs(i).Vibr_IK_re <> rUs(j).Vibr_IK_re THEN gosub unter
 IF roUs(i).Vibr_IK_li <> rUs(j).Vibr_IK_li THEN gosub unter
 IF roUs(i).Vibr_GZ_re <> rUs(j).Vibr_GZ_re THEN gosub unter
 IF roUs(i).Vibr_GZ_li <> rUs(j).Vibr_GZ_li THEN gosub unter
 IF roUs(i).PulsL_re <> rUs(j).PulsL_re THEN gosub unter
 IF roUs(i).PulsL_li <> rUs(j).PulsL_li THEN gosub unter
 IF roUs(i).PulsKK_re <> rUs(j).PulsKK_re THEN gosub unter
 IF roUs(i).PulsKK_li <> rUs(j).PulsKK_li THEN gosub unter
 IF roUs(i).PulsAtp_re <> rUs(j).PulsAtp_re THEN gosub unter
 IF roUs(i).PulsAtp_li <> rUs(j).PulsAtp_li THEN gosub unter
 IF roUs(i).PulsAdp_re <> rUs(j).PulsAdp_re THEN gosub unter
 IF roUs(i).PulsAdp_li <> rUs(j).PulsAdp_li THEN gosub unter
 IF roUs(i).Mitarbeiter <> rUs(j).Mitarbeiter THEN gosub unter
 IF roUs(i).absPos <> rUs(j).absPos THEN gosub unter
 IF roUs(i).AktZeit <> rUs(j).AktZeit THEN gosub unter
 IF roUs(i).QS <> rUs(j).QS THEN gosub unter
 IF roUs(i).QT <> rUs(j).QT THEN gosub unter
 IF roUs(i).StByte <> rUs(j).StByte THEN gosub unter
 IF roUs(i).id <> rUs(j).id THEN gosub unter
 Exit Function
unter:
 UsZUnt = UsZUnt + 1
 Return
End FUNCTION ' UsZUnt

Public FUNCTION usdmLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roUs(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Spritzst,0) Spritzst,COALESCE(Fußbef_re,0) Fußbef_re,COALESCE(Fußbef_li,0) Fußbef_li,COALESCE(Hyperk_re,0) Hyperk_re" & _
",COALESCE(Hyperk_li,0) Hyperk_li,COALESCE(Ulcera_re,0) Ulcera_re,COALESCE(Ulcera_li,0) Ulcera_li,COALESCE(Kraft_Zh_re,0) Kraft_Zh_re" & _
",COALESCE(Kraft_Zh_li,0) Kraft_Zh_li,COALESCE(Kraft_Zb_re,0) Kraft_Zb_re,COALESCE(Kraft_Zb_li,0) Kraft_Zb_li,COALESCE(Kraft_Knie_re,0) Kraft_Knie_re" & _
",COALESCE(Kraft_Knie_li,0) Kraft_Knie_li,COALESCE(ASR_re,0) ASR_re,COALESCE(ASR_li,0) ASR_li,COALESCE(PSR_re,0) PSR_re" & _
",COALESCE(PSR_li,0) PSR_li,COALESCE(Oberfl_re,0) Oberfl_re,COALESCE(Oberfl_li,0) Oberfl_li,COALESCE(MF_re,0) MF_re" & _
",COALESCE(MF_li,0) MF_li,COALESCE(KW_re,0) KW_re,COALESCE(KW_li,0) KW_li,COALESCE(Vibr_IK_re,0) Vibr_IK_re" & _
",COALESCE(Vibr_IK_li,0) Vibr_IK_li,COALESCE(Vibr_GZ_re,0) Vibr_GZ_re,COALESCE(Vibr_GZ_li,0) Vibr_GZ_li,COALESCE(PulsL_re,0) PulsL_re" & _
",COALESCE(PulsL_li,0) PulsL_li,COALESCE(PulsKK_re,0) PulsKK_re,COALESCE(PulsKK_li,0) PulsKK_li,COALESCE(PulsAtp_re,0) PulsAtp_re" & _
",COALESCE(PulsAtp_li,0) PulsAtp_li,COALESCE(PulsAdp_re,0) PulsAdp_re,COALESCE(PulsAdp_li,0) PulsAdp_li,COALESCE(Mitarbeiter,0) Mitarbeiter" & _
",COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(QS,'') QS,COALESCE(QT,'') QT" & _
",COALESCE(StByte,0) StByte,COALESCE(id,0) id FROM `usdm` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roUs)
  roUs(akt).FID = rs!FID
  roUs(akt).Pat_ID = rs!Pat_ID
  roUs(akt).ZeitPunkt = rs!ZeitPunkt
  roUs(akt).Art = doUmwfSQL(rs!Art, lies.obMySQL, False)
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
  IF Not rs.EOF THEN ReDim Preserve roUs(UBound(roUs) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' usdmLaden

Function usdmEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rUs) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rUs)
   IF rUs(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roUs)
    IF roUs(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roUs(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roUs(robeg + UBound(rUs) - rbeg)
   For ri = rbeg To UBound(rUs)
    Call roUsZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rUs = roUs
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' usdmEinf

Public FUNCTION usdmSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere usdm"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `usdm` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `usdm` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte)             VALUES"))
 For i = 1 to ubound(rUs)
'  rUs(i).AktZeit = now()
  rUs(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rUs(i).FID, "," , rUs(i).Pat_ID, "," , DatFor_k(rUs(i).ZeitPunkt), ",'" , rUs(i).Art, "','" , rUs(i).Spritzst, "','" , rUs(i).Fußbef_re, "','" , rUs(i).Fußbef_li, "','" , rUs(i).Hyperk_re, "','" ,  _
   rUs(i).Hyperk_li, "','" , rUs(i).Ulcera_re, "','" , rUs(i).Ulcera_li, "','" , rUs(i).Kraft_Zh_re, "','" , rUs(i).Kraft_Zh_li, "','" , rUs(i).Kraft_Zb_re, "','" , rUs(i).Kraft_Zb_li, "','" ,  _
   rUs(i).Kraft_Knie_re, "','" , rUs(i).Kraft_Knie_li, "','" , rUs(i).ASR_re, "','" , rUs(i).ASR_li, "','" , rUs(i).PSR_re, "','" , rUs(i).PSR_li, "','" , rUs(i).Oberfl_re, "','" , rUs(i).Oberfl_li, "','" ,  _
   rUs(i).MF_re, "','" , rUs(i).MF_li, "','" , rUs(i).KW_re, "','" , rUs(i).KW_li, "','" , rUs(i).Vibr_IK_re, "','" , rUs(i).Vibr_IK_li, "','" , rUs(i).Vibr_GZ_re, "','" ,  _
   rUs(i).Vibr_GZ_li, "','" , rUs(i).PulsL_re, "','" , rUs(i).PulsL_li, "','" , rUs(i).PulsKK_re, "','" , rUs(i).PulsKK_li, "','" , rUs(i).PulsAtp_re, "','" , rUs(i).PulsAtp_li, "','" , rUs(i).PulsAdp_re, "','" ,  _
   rUs(i).PulsAdp_li, "','" , rUs(i).Mitarbeiter, "'," , rUs(i).absPos, "," , DatFor_k(rUs(i).AktZeit), ",'" , rUs(i).QS, "','" , rUs(i).QT, "'," , rUs(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rUs) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rUs) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(38)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rUs),i)
  IF Len(rUs(k).Art) > maxi(0) THEN maxi(0) = Len(rUs(k).Art)
  IF Len(rUs(k).Spritzst) > maxi(1) THEN maxi(1) = Len(rUs(k).Spritzst)
  IF Len(rUs(k).Fußbef_re) > maxi(2) THEN maxi(2) = Len(rUs(k).Fußbef_re)
  IF Len(rUs(k).Fußbef_li) > maxi(3) THEN maxi(3) = Len(rUs(k).Fußbef_li)
  IF Len(rUs(k).Hyperk_re) > maxi(4) THEN maxi(4) = Len(rUs(k).Hyperk_re)
  IF Len(rUs(k).Hyperk_li) > maxi(5) THEN maxi(5) = Len(rUs(k).Hyperk_li)
  IF Len(rUs(k).Ulcera_re) > maxi(6) THEN maxi(6) = Len(rUs(k).Ulcera_re)
  IF Len(rUs(k).Ulcera_li) > maxi(7) THEN maxi(7) = Len(rUs(k).Ulcera_li)
  IF Len(rUs(k).Kraft_Zh_re) > maxi(8) THEN maxi(8) = Len(rUs(k).Kraft_Zh_re)
  IF Len(rUs(k).Kraft_Zh_li) > maxi(9) THEN maxi(9) = Len(rUs(k).Kraft_Zh_li)
  IF Len(rUs(k).Kraft_Zb_re) > maxi(10) THEN maxi(10) = Len(rUs(k).Kraft_Zb_re)
  IF Len(rUs(k).Kraft_Zb_li) > maxi(11) THEN maxi(11) = Len(rUs(k).Kraft_Zb_li)
  IF Len(rUs(k).Kraft_Knie_re) > maxi(12) THEN maxi(12) = Len(rUs(k).Kraft_Knie_re)
  IF Len(rUs(k).Kraft_Knie_li) > maxi(13) THEN maxi(13) = Len(rUs(k).Kraft_Knie_li)
  IF Len(rUs(k).ASR_re) > maxi(14) THEN maxi(14) = Len(rUs(k).ASR_re)
  IF Len(rUs(k).ASR_li) > maxi(15) THEN maxi(15) = Len(rUs(k).ASR_li)
  IF Len(rUs(k).PSR_re) > maxi(16) THEN maxi(16) = Len(rUs(k).PSR_re)
  IF Len(rUs(k).PSR_li) > maxi(17) THEN maxi(17) = Len(rUs(k).PSR_li)
  IF Len(rUs(k).Oberfl_re) > maxi(18) THEN maxi(18) = Len(rUs(k).Oberfl_re)
  IF Len(rUs(k).Oberfl_li) > maxi(19) THEN maxi(19) = Len(rUs(k).Oberfl_li)
  IF Len(rUs(k).MF_re) > maxi(20) THEN maxi(20) = Len(rUs(k).MF_re)
  IF Len(rUs(k).MF_li) > maxi(21) THEN maxi(21) = Len(rUs(k).MF_li)
  IF Len(rUs(k).KW_re) > maxi(22) THEN maxi(22) = Len(rUs(k).KW_re)
  IF Len(rUs(k).KW_li) > maxi(23) THEN maxi(23) = Len(rUs(k).KW_li)
  IF Len(rUs(k).Vibr_IK_re) > maxi(24) THEN maxi(24) = Len(rUs(k).Vibr_IK_re)
  IF Len(rUs(k).Vibr_IK_li) > maxi(25) THEN maxi(25) = Len(rUs(k).Vibr_IK_li)
  IF Len(rUs(k).Vibr_GZ_re) > maxi(26) THEN maxi(26) = Len(rUs(k).Vibr_GZ_re)
  IF Len(rUs(k).Vibr_GZ_li) > maxi(27) THEN maxi(27) = Len(rUs(k).Vibr_GZ_li)
  IF Len(rUs(k).PulsL_re) > maxi(28) THEN maxi(28) = Len(rUs(k).PulsL_re)
  IF Len(rUs(k).PulsL_li) > maxi(29) THEN maxi(29) = Len(rUs(k).PulsL_li)
  IF Len(rUs(k).PulsKK_re) > maxi(30) THEN maxi(30) = Len(rUs(k).PulsKK_re)
  IF Len(rUs(k).PulsKK_li) > maxi(31) THEN maxi(31) = Len(rUs(k).PulsKK_li)
  IF Len(rUs(k).PulsAtp_re) > maxi(32) THEN maxi(32) = Len(rUs(k).PulsAtp_re)
  IF Len(rUs(k).PulsAtp_li) > maxi(33) THEN maxi(33) = Len(rUs(k).PulsAtp_li)
  IF Len(rUs(k).PulsAdp_re) > maxi(34) THEN maxi(34) = Len(rUs(k).PulsAdp_re)
  IF Len(rUs(k).PulsAdp_li) > maxi(35) THEN maxi(35) = Len(rUs(k).PulsAdp_li)
  IF Len(rUs(k).Mitarbeiter) > maxi(36) THEN maxi(36) = Len(rUs(k).Mitarbeiter)
  IF Len(rUs(k).QS) > maxi(37) THEN maxi(37) = Len(rUs(k).QS)
  IF Len(rUs(k).QT) > maxi(38) THEN maxi(38) = Len(rUs(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "usdm", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "usdm", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rUs), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUs.Art: '" & rUs(k).Art & "' -> '" & Left$(rUs(k).Art, maxL)  & "'",true : rUs(k).Art = Left$(rUs(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUs.Spritzst: '" & rUs(k).Spritzst & "' -> '" & Left$(rUs(k).Spritzst, maxL)  & "'",true : rUs(k).Spritzst = Left$(rUs(k).Spritzst, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_re: '" & rUs(k).Fußbef_re & "' -> '" & Left$(rUs(k).Fußbef_re, maxL)  & "'",true : rUs(k).Fußbef_re = Left$(rUs(k).Fußbef_re, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_li: '" & rUs(k).Fußbef_li & "' -> '" & Left$(rUs(k).Fußbef_li, maxL)  & "'",true : rUs(k).Fußbef_li = Left$(rUs(k).Fußbef_li, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_re: '" & rUs(k).Hyperk_re & "' -> '" & Left$(rUs(k).Hyperk_re, maxL)  & "'",true : rUs(k).Hyperk_re = Left$(rUs(k).Hyperk_re, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_li: '" & rUs(k).Hyperk_li & "' -> '" & Left$(rUs(k).Hyperk_li, maxL)  & "'",true : rUs(k).Hyperk_li = Left$(rUs(k).Hyperk_li, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_re: '" & rUs(k).Ulcera_re & "' -> '" & Left$(rUs(k).Ulcera_re, maxL)  & "'",true : rUs(k).Ulcera_re = Left$(rUs(k).Ulcera_re, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_li: '" & rUs(k).Ulcera_li & "' -> '" & Left$(rUs(k).Ulcera_li, maxL)  & "'",true : rUs(k).Ulcera_li = Left$(rUs(k).Ulcera_li, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_re: '" & rUs(k).Kraft_Zh_re & "' -> '" & Left$(rUs(k).Kraft_Zh_re, maxL)  & "'",true : rUs(k).Kraft_Zh_re = Left$(rUs(k).Kraft_Zh_re, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_li: '" & rUs(k).Kraft_Zh_li & "' -> '" & Left$(rUs(k).Kraft_Zh_li, maxL)  & "'",true : rUs(k).Kraft_Zh_li = Left$(rUs(k).Kraft_Zh_li, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_re: '" & rUs(k).Kraft_Zb_re & "' -> '" & Left$(rUs(k).Kraft_Zb_re, maxL)  & "'",true : rUs(k).Kraft_Zb_re = Left$(rUs(k).Kraft_Zb_re, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_li: '" & rUs(k).Kraft_Zb_li & "' -> '" & Left$(rUs(k).Kraft_Zb_li, maxL)  & "'",true : rUs(k).Kraft_Zb_li = Left$(rUs(k).Kraft_Zb_li, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_re: '" & rUs(k).Kraft_Knie_re & "' -> '" & Left$(rUs(k).Kraft_Knie_re, maxL)  & "'",true : rUs(k).Kraft_Knie_re = Left$(rUs(k).Kraft_Knie_re, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_li: '" & rUs(k).Kraft_Knie_li & "' -> '" & Left$(rUs(k).Kraft_Knie_li, maxL)  & "'",true : rUs(k).Kraft_Knie_li = Left$(rUs(k).Kraft_Knie_li, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_re: '" & rUs(k).ASR_re & "' -> '" & Left$(rUs(k).ASR_re, maxL)  & "'",true : rUs(k).ASR_re = Left$(rUs(k).ASR_re, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_li: '" & rUs(k).ASR_li & "' -> '" & Left$(rUs(k).ASR_li, maxL)  & "'",true : rUs(k).ASR_li = Left$(rUs(k).ASR_li, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_re: '" & rUs(k).PSR_re & "' -> '" & Left$(rUs(k).PSR_re, maxL)  & "'",true : rUs(k).PSR_re = Left$(rUs(k).PSR_re, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_li: '" & rUs(k).PSR_li & "' -> '" & Left$(rUs(k).PSR_li, maxL)  & "'",true : rUs(k).PSR_li = Left$(rUs(k).PSR_li, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_re: '" & rUs(k).Oberfl_re & "' -> '" & Left$(rUs(k).Oberfl_re, maxL)  & "'",true : rUs(k).Oberfl_re = Left$(rUs(k).Oberfl_re, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_li: '" & rUs(k).Oberfl_li & "' -> '" & Left$(rUs(k).Oberfl_li, maxL)  & "'",true : rUs(k).Oberfl_li = Left$(rUs(k).Oberfl_li, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_re: '" & rUs(k).MF_re & "' -> '" & Left$(rUs(k).MF_re, maxL)  & "'",true : rUs(k).MF_re = Left$(rUs(k).MF_re, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_li: '" & rUs(k).MF_li & "' -> '" & Left$(rUs(k).MF_li, maxL)  & "'",true : rUs(k).MF_li = Left$(rUs(k).MF_li, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_re: '" & rUs(k).KW_re & "' -> '" & Left$(rUs(k).KW_re, maxL)  & "'",true : rUs(k).KW_re = Left$(rUs(k).KW_re, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_li: '" & rUs(k).KW_li & "' -> '" & Left$(rUs(k).KW_li, maxL)  & "'",true : rUs(k).KW_li = Left$(rUs(k).KW_li, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_re: '" & rUs(k).Vibr_IK_re & "' -> '" & Left$(rUs(k).Vibr_IK_re, maxL)  & "'",true : rUs(k).Vibr_IK_re = Left$(rUs(k).Vibr_IK_re, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_li: '" & rUs(k).Vibr_IK_li & "' -> '" & Left$(rUs(k).Vibr_IK_li, maxL)  & "'",true : rUs(k).Vibr_IK_li = Left$(rUs(k).Vibr_IK_li, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_re: '" & rUs(k).Vibr_GZ_re & "' -> '" & Left$(rUs(k).Vibr_GZ_re, maxL)  & "'",true : rUs(k).Vibr_GZ_re = Left$(rUs(k).Vibr_GZ_re, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_li: '" & rUs(k).Vibr_GZ_li & "' -> '" & Left$(rUs(k).Vibr_GZ_li, maxL)  & "'",true : rUs(k).Vibr_GZ_li = Left$(rUs(k).Vibr_GZ_li, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_re: '" & rUs(k).PulsL_re & "' -> '" & Left$(rUs(k).PulsL_re, maxL)  & "'",true : rUs(k).PulsL_re = Left$(rUs(k).PulsL_re, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_li: '" & rUs(k).PulsL_li & "' -> '" & Left$(rUs(k).PulsL_li, maxL)  & "'",true : rUs(k).PulsL_li = Left$(rUs(k).PulsL_li, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_re: '" & rUs(k).PulsKK_re & "' -> '" & Left$(rUs(k).PulsKK_re, maxL)  & "'",true : rUs(k).PulsKK_re = Left$(rUs(k).PulsKK_re, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_li: '" & rUs(k).PulsKK_li & "' -> '" & Left$(rUs(k).PulsKK_li, maxL)  & "'",true : rUs(k).PulsKK_li = Left$(rUs(k).PulsKK_li, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_re: '" & rUs(k).PulsAtp_re & "' -> '" & Left$(rUs(k).PulsAtp_re, maxL)  & "'",true : rUs(k).PulsAtp_re = Left$(rUs(k).PulsAtp_re, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_li: '" & rUs(k).PulsAtp_li & "' -> '" & Left$(rUs(k).PulsAtp_li, maxL)  & "'",true : rUs(k).PulsAtp_li = Left$(rUs(k).PulsAtp_li, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_re: '" & rUs(k).PulsAdp_re & "' -> '" & Left$(rUs(k).PulsAdp_re, maxL)  & "'",true : rUs(k).PulsAdp_re = Left$(rUs(k).PulsAdp_re, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_li: '" & rUs(k).PulsAdp_li & "' -> '" & Left$(rUs(k).PulsAdp_li, maxL)  & "'",true : rUs(k).PulsAdp_li = Left$(rUs(k).PulsAdp_li, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rUs.Mitarbeiter: '" & rUs(k).Mitarbeiter & "' -> '" & Left$(rUs(k).Mitarbeiter, maxL)  & "'",true : rUs(k).Mitarbeiter = Left$(rUs(k).Mitarbeiter, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rUs.QS: '" & rUs(k).QS & "' -> '" & Left$(rUs(k).QS, maxL)  & "'",true : rUs(k).QS = Left$(rUs(k).QS, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rUs.QT: '" & rUs(k).QT & "' -> '" & Left$(rUs(k).QT, maxL)  & "'",true : rUs(k).QT = Left$(rUs(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' usdmSpeichern

Public FUNCTION roFuZuw(i&, j&)
 roFu(i).FID = rFu(j).FID
 roFu(i).Pat_ID = rFu(j).Pat_ID
 roFu(i).ZeitPunkt = rFu(j).ZeitPunkt
 roFu(i).Art = rFu(j).Art
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
End FUNCTION ' roFuZuw

Public FUNCTION FuZUnt%(i&, j&)
 IF roFu(i).FID <> rFu(j).FID THEN gosub unter
 IF roFu(i).Pat_ID <> rFu(j).Pat_ID THEN gosub unter
 IF roFu(i).ZeitPunkt <> rFu(j).ZeitPunkt THEN gosub unter
 IF roFu(i).Art <> rFu(j).Art THEN gosub unter
 IF roFu(i).Fußdeform <> rFu(j).Fußdeform THEN gosub unter
 IF roFu(i).Hyper_mEin <> rFu(j).Hyper_mEin THEN gosub unter
 IF roFu(i).Weiteres <> rFu(j).Weiteres THEN gosub unter
 IF roFu(i).Zn_Ulcus <> rFu(j).Zn_Ulcus THEN gosub unter
 IF roFu(i).Zn_Amput <> rFu(j).Zn_Amput THEN gosub unter
 IF roFu(i).Fuß_ang <> rFu(j).Fuß_ang THEN gosub unter
 IF roFu(i).Ulcera <> rFu(j).Ulcera THEN gosub unter
 IF roFu(i).Wundinfektion <> rFu(j).Wundinfektion THEN gosub unter
 IF roFu(i).nae_US <> rFu(j).nae_US THEN gosub unter
 IF roFu(i).Mitarbeiter <> rFu(j).Mitarbeiter THEN gosub unter
 IF roFu(i).absPos <> rFu(j).absPos THEN gosub unter
 IF roFu(i).AktZeit <> rFu(j).AktZeit THEN gosub unter
 IF roFu(i).QS <> rFu(j).QS THEN gosub unter
 IF roFu(i).QT <> rFu(j).QT THEN gosub unter
 IF roFu(i).StByte <> rFu(j).StByte THEN gosub unter
 IF roFu(i).id <> rFu(j).id THEN gosub unter
 Exit Function
unter:
 FuZUnt = FuZUnt + 1
 Return
End FUNCTION ' FuZUnt

Public FUNCTION fussLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFu(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Fußdeform,'') Fußdeform,COALESCE(Hyper_mEin,'') Hyper_mEin,COALESCE(Weiteres,'') Weiteres,COALESCE(Zn_Ulcus,'') Zn_Ulcus" & _
",COALESCE(Zn_Amput,'') Zn_Amput,COALESCE(Fuß_ang,'') Fuß_ang,COALESCE(Ulcera,'') Ulcera,COALESCE(Wundinfektion,'') Wundinfektion" & _
",COALESCE(nae_US,'') nae_US,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(id,0) id" & _
" FROM `fuss` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roFu)
  roFu(akt).FID = rs!FID
  roFu(akt).Pat_ID = rs!Pat_ID
  roFu(akt).ZeitPunkt = rs!ZeitPunkt
  roFu(akt).Art = doUmwfSQL(rs!Art, lies.obMySQL, False)
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
  IF Not rs.EOF THEN ReDim Preserve roFu(UBound(roFu) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' fussLaden

Function fussEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rFu) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rFu)
   IF rFu(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roFu)
    IF roFu(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roFu(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roFu(robeg + UBound(rFu) - rbeg)
   For ri = rbeg To UBound(rFu)
    Call roFuZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rFu = roFu
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' fussEinf

Public FUNCTION fussSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere fuss"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `fuss` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `fuss` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte)           VALUES"))
 For i = 1 to ubound(rFu)
'  rFu(i).AktZeit = now()
  rFu(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rFu(i).FID, "," , rFu(i).Pat_ID, "," , DatFor_k(rFu(i).ZeitPunkt), ",'" , rFu(i).Art, "','" , rFu(i).Fußdeform, "','" , rFu(i).Hyper_mEin, "','" , rFu(i).Weiteres, "','" , rFu(i).Zn_Ulcus, "','" ,  _
   rFu(i).Zn_Amput, "','" , rFu(i).Fuß_ang, "','" , rFu(i).Ulcera, "','" , rFu(i).Wundinfektion, "','" , rFu(i).nae_US, "','" , rFu(i).Mitarbeiter, "'," , rFu(i).absPos, "," , DatFor_k(rFu(i).AktZeit), ",'" ,  _
   rFu(i).QS, "','" , rFu(i).QT, "'," , rFu(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rFu) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rFu) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(12)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFu),i)
  IF Len(rFu(k).Art) > maxi(0) THEN maxi(0) = Len(rFu(k).Art)
  IF Len(rFu(k).Fußdeform) > maxi(1) THEN maxi(1) = Len(rFu(k).Fußdeform)
  IF Len(rFu(k).Hyper_mEin) > maxi(2) THEN maxi(2) = Len(rFu(k).Hyper_mEin)
  IF Len(rFu(k).Weiteres) > maxi(3) THEN maxi(3) = Len(rFu(k).Weiteres)
  IF Len(rFu(k).Zn_Ulcus) > maxi(4) THEN maxi(4) = Len(rFu(k).Zn_Ulcus)
  IF Len(rFu(k).Zn_Amput) > maxi(5) THEN maxi(5) = Len(rFu(k).Zn_Amput)
  IF Len(rFu(k).Fuß_ang) > maxi(6) THEN maxi(6) = Len(rFu(k).Fuß_ang)
  IF Len(rFu(k).Ulcera) > maxi(7) THEN maxi(7) = Len(rFu(k).Ulcera)
  IF Len(rFu(k).Wundinfektion) > maxi(8) THEN maxi(8) = Len(rFu(k).Wundinfektion)
  IF Len(rFu(k).nae_US) > maxi(9) THEN maxi(9) = Len(rFu(k).nae_US)
  IF Len(rFu(k).Mitarbeiter) > maxi(10) THEN maxi(10) = Len(rFu(k).Mitarbeiter)
  IF Len(rFu(k).QS) > maxi(11) THEN maxi(11) = Len(rFu(k).QS)
  IF Len(rFu(k).QT) > maxi(12) THEN maxi(12) = Len(rFu(k).QT)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "fuss", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "fuss", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFu), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFu.Art: '" & rFu(k).Art & "' -> '" & Left$(rFu(k).Art, maxL)  & "'",true : rFu(k).Art = Left$(rFu(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fußdeform: '" & rFu(k).Fußdeform & "' -> '" & Left$(rFu(k).Fußdeform, maxL)  & "'",true : rFu(k).Fußdeform = Left$(rFu(k).Fußdeform, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFu.Hyper_mEin: '" & rFu(k).Hyper_mEin & "' -> '" & Left$(rFu(k).Hyper_mEin, maxL)  & "'",true : rFu(k).Hyper_mEin = Left$(rFu(k).Hyper_mEin, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFu.Weiteres: '" & rFu(k).Weiteres & "' -> '" & Left$(rFu(k).Weiteres, maxL)  & "'",true : rFu(k).Weiteres = Left$(rFu(k).Weiteres, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Ulcus: '" & rFu(k).Zn_Ulcus & "' -> '" & Left$(rFu(k).Zn_Ulcus, maxL)  & "'",true : rFu(k).Zn_Ulcus = Left$(rFu(k).Zn_Ulcus, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Amput: '" & rFu(k).Zn_Amput & "' -> '" & Left$(rFu(k).Zn_Amput, maxL)  & "'",true : rFu(k).Zn_Amput = Left$(rFu(k).Zn_Amput, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fuß_ang: '" & rFu(k).Fuß_ang & "' -> '" & Left$(rFu(k).Fuß_ang, maxL)  & "'",true : rFu(k).Fuß_ang = Left$(rFu(k).Fuß_ang, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFu.Ulcera: '" & rFu(k).Ulcera & "' -> '" & Left$(rFu(k).Ulcera, maxL)  & "'",true : rFu(k).Ulcera = Left$(rFu(k).Ulcera, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFu.Wundinfektion: '" & rFu(k).Wundinfektion & "' -> '" & Left$(rFu(k).Wundinfektion, maxL)  & "'",true : rFu(k).Wundinfektion = Left$(rFu(k).Wundinfektion, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFu.nae_US: '" & rFu(k).nae_US & "' -> '" & Left$(rFu(k).nae_US, maxL)  & "'",true : rFu(k).nae_US = Left$(rFu(k).nae_US, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFu.Mitarbeiter: '" & rFu(k).Mitarbeiter & "' -> '" & Left$(rFu(k).Mitarbeiter, maxL)  & "'",true : rFu(k).Mitarbeiter = Left$(rFu(k).Mitarbeiter, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFu.QS: '" & rFu(k).QS & "' -> '" & Left$(rFu(k).QS, maxL)  & "'",true : rFu(k).QS = Left$(rFu(k).QS, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFu.QT: '" & rFu(k).QT & "' -> '" & Left$(rFu(k).QT, maxL)  & "'",true : rFu(k).QT = Left$(rFu(k).QT, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' fussSpeichern

Public FUNCTION roUlZuw(i&, j&)
 roUl(i).FID = rUl(j).FID
 roUl(i).Pat_ID = rUl(j).Pat_ID
 roUl(i).ZeitPunkt = rUl(j).ZeitPunkt
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
End FUNCTION ' roUlZuw

Public FUNCTION UlZUnt%(i&, j&)
 IF roUl(i).FID <> rUl(j).FID THEN gosub unter
 IF roUl(i).Pat_ID <> rUl(j).Pat_ID THEN gosub unter
 IF roUl(i).ZeitPunkt <> rUl(j).ZeitPunkt THEN gosub unter
 IF roUl(i).Lokalisation <> rUl(j).Lokalisation THEN gosub unter
 IF roUl(i).Seite <> rUl(j).Seite THEN gosub unter
 IF roUl(i).Größe <> rUl(j).Größe THEN gosub unter
 IF roUl(i).Beläge <> rUl(j).Beläge THEN gosub unter
 IF roUl(i).Exsudat <> rUl(j).Exsudat THEN gosub unter
 IF roUl(i).Geruch <> rUl(j).Geruch THEN gosub unter
 IF roUl(i).Wundrand <> rUl(j).Wundrand THEN gosub unter
 IF roUl(i).Wundumgebung <> rUl(j).Wundumgebung THEN gosub unter
 IF roUl(i).Temperatur <> rUl(j).Temperatur THEN gosub unter
 IF roUl(i).Fotodoku <> rUl(j).Fotodoku THEN gosub unter
 IF roUl(i).Wundversorgung <> rUl(j).Wundversorgung THEN gosub unter
 IF roUl(i).Mitarbeiter <> rUl(j).Mitarbeiter THEN gosub unter
 IF roUl(i).absPos <> rUl(j).absPos THEN gosub unter
 IF roUl(i).AktZeit <> rUl(j).AktZeit THEN gosub unter
 IF roUl(i).StByte <> rUl(j).StByte THEN gosub unter
 Exit Function
unter:
 UlZUnt = UlZUnt + 1
 Return
End FUNCTION ' UlZUnt

Public FUNCTION ulcusLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roUl(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Lokalisation,'') Lokalisation" & _
",COALESCE(Seite,'') Seite,COALESCE(Größe,'') Größe,COALESCE(Beläge,'') Beläge,COALESCE(Exsudat,'') Exsudat" & _
",COALESCE(Geruch,'') Geruch,COALESCE(Wundrand,'') Wundrand,COALESCE(Wundumgebung,'') Wundumgebung,COALESCE(Temperatur,'') Temperatur" & _
",COALESCE(Fotodoku,'') Fotodoku,COALESCE(Wundversorgung,'') Wundversorgung,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos" & _
",IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte FROM `ulcus` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roUl)
  roUl(akt).FID = rs!FID
  roUl(akt).Pat_ID = rs!Pat_ID
  roUl(akt).ZeitPunkt = rs!ZeitPunkt
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
  IF Not rs.EOF THEN ReDim Preserve roUl(UBound(roUl) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' ulcusLaden

Function ulcusEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rUl) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rUl)
   IF rUl(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roUl)
    IF roUl(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roUl(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roUl(robeg + UBound(rUl) - rbeg)
   For ri = rbeg To UBound(rUl)
    Call roUlZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rUl = roUl
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' ulcusEinf

Public FUNCTION ulcusSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere ulcus"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `ulcus` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `ulcus` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte)  VALUES"))
 For i = 1 to ubound(rUl)
'  rUl(i).AktZeit = now()
  rUl(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rUl(i).FID, "," , rUl(i).Pat_ID, "," , DatFor_k(rUl(i).ZeitPunkt), ",'" , rUl(i).Lokalisation, "','" , rUl(i).Seite, "','" , rUl(i).Größe, "','" , rUl(i).Beläge, "','" , rUl(i).Exsudat, "','" ,  _
   rUl(i).Geruch, "','" , rUl(i).Wundrand, "','" , rUl(i).Wundumgebung, "','" , rUl(i).Temperatur, "','" , rUl(i).Fotodoku, "','" , rUl(i).Wundversorgung, "','" , rUl(i).Mitarbeiter, "'," ,  _
   rUl(i).absPos, "," , DatFor_k(rUl(i).AktZeit), "," , rUl(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rUl) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rUl) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(11)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rUl),i)
  IF Len(rUl(k).Lokalisation) > maxi(0) THEN maxi(0) = Len(rUl(k).Lokalisation)
  IF Len(rUl(k).Seite) > maxi(1) THEN maxi(1) = Len(rUl(k).Seite)
  IF Len(rUl(k).Größe) > maxi(2) THEN maxi(2) = Len(rUl(k).Größe)
  IF Len(rUl(k).Beläge) > maxi(3) THEN maxi(3) = Len(rUl(k).Beläge)
  IF Len(rUl(k).Exsudat) > maxi(4) THEN maxi(4) = Len(rUl(k).Exsudat)
  IF Len(rUl(k).Geruch) > maxi(5) THEN maxi(5) = Len(rUl(k).Geruch)
  IF Len(rUl(k).Wundrand) > maxi(6) THEN maxi(6) = Len(rUl(k).Wundrand)
  IF Len(rUl(k).Wundumgebung) > maxi(7) THEN maxi(7) = Len(rUl(k).Wundumgebung)
  IF Len(rUl(k).Temperatur) > maxi(8) THEN maxi(8) = Len(rUl(k).Temperatur)
  IF Len(rUl(k).Fotodoku) > maxi(9) THEN maxi(9) = Len(rUl(k).Fotodoku)
  IF Len(rUl(k).Wundversorgung) > maxi(10) THEN maxi(10) = Len(rUl(k).Wundversorgung)
  IF Len(rUl(k).Mitarbeiter) > maxi(11) THEN maxi(11) = Len(rUl(k).Mitarbeiter)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "ulcus", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "ulcus", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rUl), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUl.Lokalisation: '" & rUl(k).Lokalisation & "' -> '" & Left$(rUl(k).Lokalisation, maxL)  & "'",true : rUl(k).Lokalisation = Left$(rUl(k).Lokalisation, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUl.Seite: '" & rUl(k).Seite & "' -> '" & Left$(rUl(k).Seite, maxL)  & "'",true : rUl(k).Seite = Left$(rUl(k).Seite, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUl.Größe: '" & rUl(k).Größe & "' -> '" & Left$(rUl(k).Größe, maxL)  & "'",true : rUl(k).Größe = Left$(rUl(k).Größe, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUl.Beläge: '" & rUl(k).Beläge & "' -> '" & Left$(rUl(k).Beläge, maxL)  & "'",true : rUl(k).Beläge = Left$(rUl(k).Beläge, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUl.Exsudat: '" & rUl(k).Exsudat & "' -> '" & Left$(rUl(k).Exsudat, maxL)  & "'",true : rUl(k).Exsudat = Left$(rUl(k).Exsudat, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUl.Geruch: '" & rUl(k).Geruch & "' -> '" & Left$(rUl(k).Geruch, maxL)  & "'",true : rUl(k).Geruch = Left$(rUl(k).Geruch, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundrand: '" & rUl(k).Wundrand & "' -> '" & Left$(rUl(k).Wundrand, maxL)  & "'",true : rUl(k).Wundrand = Left$(rUl(k).Wundrand, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundumgebung: '" & rUl(k).Wundumgebung & "' -> '" & Left$(rUl(k).Wundumgebung, maxL)  & "'",true : rUl(k).Wundumgebung = Left$(rUl(k).Wundumgebung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUl.Temperatur: '" & rUl(k).Temperatur & "' -> '" & Left$(rUl(k).Temperatur, maxL)  & "'",true : rUl(k).Temperatur = Left$(rUl(k).Temperatur, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUl.Fotodoku: '" & rUl(k).Fotodoku & "' -> '" & Left$(rUl(k).Fotodoku, maxL)  & "'",true : rUl(k).Fotodoku = Left$(rUl(k).Fotodoku, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundversorgung: '" & rUl(k).Wundversorgung & "' -> '" & Left$(rUl(k).Wundversorgung, maxL)  & "'",true : rUl(k).Wundversorgung = Left$(rUl(k).Wundversorgung, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUl.Mitarbeiter: '" & rUl(k).Mitarbeiter & "' -> '" & Left$(rUl(k).Mitarbeiter, maxL)  & "'",true : rUl(k).Mitarbeiter = Left$(rUl(k).Mitarbeiter, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' ulcusSpeichern

Public FUNCTION roVkZuw(i&, j&)
 roVk(i).FID = rVk(j).FID
 roVk(i).Pat_ID = rVk(j).Pat_ID
 roVk(i).ZeitPunkt = rVk(j).ZeitPunkt
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
End FUNCTION ' roVkZuw

Public FUNCTION VkZUnt%(i&, j&)
 IF roVk(i).FID <> rVk(j).FID THEN gosub unter
 IF roVk(i).Pat_ID <> rVk(j).Pat_ID THEN gosub unter
 IF roVk(i).ZeitPunkt <> rVk(j).ZeitPunkt THEN gosub unter
 IF roVk(i).Wohlempfinden <> rVk(j).Wohlempfinden THEN gosub unter
 IF roVk(i).Saettigung <> rVk(j).Saettigung THEN gosub unter
 IF roVk(i).Zielwerterreichung <> rVk(j).Zielwerterreichung THEN gosub unter
 IF roVk(i).Ketonkörper <> rVk(j).Ketonkörper THEN gosub unter
 IF roVk(i).Gynaekologenbefund <> rVk(j).Gynaekologenbefund THEN gosub unter
 IF roVk(i).Gewichtsentwicklung <> rVk(j).Gewichtsentwicklung THEN gosub unter
 IF roVk(i).HbA1c <> rVk(j).HbA1c THEN gosub unter
 IF roVk(i).Bewegung <> rVk(j).Bewegung THEN gosub unter
 IF roVk(i).Minuten <> rVk(j).Minuten THEN gosub unter
 IF roVk(i).Blutdruck <> rVk(j).Blutdruck THEN gosub unter
 IF roVk(i).Puls <> rVk(j).Puls THEN gosub unter
 IF roVk(i).Mitarbeiter <> rVk(j).Mitarbeiter THEN gosub unter
 IF roVk(i).absPos <> rVk(j).absPos THEN gosub unter
 IF roVk(i).AktZeit <> rVk(j).AktZeit THEN gosub unter
 IF roVk(i).StByte <> rVk(j).StByte THEN gosub unter
 Exit Function
unter:
 VkZUnt = VkZUnt + 1
 Return
End FUNCTION ' VkZUnt

Public FUNCTION vkgdLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roVk(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,COALESCE(Wohlempfinden,'') Wohlempfinden" & _
",COALESCE(Saettigung,'') Saettigung,COALESCE(Zielwerterreichung,'') Zielwerterreichung,COALESCE(Ketonkörper,'') Ketonkörper,COALESCE(Gynaekologenbefund,'') Gynaekologenbefund" & _
",COALESCE(Gewichtsentwicklung,'') Gewichtsentwicklung,COALESCE(HbA1c,'') HbA1c,COALESCE(Bewegung,'') Bewegung,COALESCE(Minuten,'') Minuten" & _
",COALESCE(Blutdruck,'') Blutdruck,COALESCE(Puls,'') Puls,COALESCE(Mitarbeiter,'') Mitarbeiter,COALESCE(absPos,0) absPos" & _
",IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte FROM `vkgd` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roVk)
  roVk(akt).FID = rs!FID
  roVk(akt).Pat_ID = rs!Pat_ID
  roVk(akt).ZeitPunkt = rs!ZeitPunkt
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
  IF Not rs.EOF THEN ReDim Preserve roVk(UBound(roVk) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' vkgdLaden

Function vkgdEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rVk) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rVk)
   IF rVk(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roVk)
    IF roVk(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roVk(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roVk(robeg + UBound(rVk) - rbeg)
   For ri = rbeg To UBound(rVk)
    Call roVkZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rVk = roVk
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' vkgdEinf

Public FUNCTION vkgdSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere vkgd"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `vkgd` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `vkgd` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte)            VALUES"))
 For i = 1 to ubound(rVk)
'  rVk(i).AktZeit = now()
  rVk(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rVk(i).FID, "," , rVk(i).Pat_ID, "," , DatFor_k(rVk(i).ZeitPunkt), ",'" , rVk(i).Wohlempfinden, "','" , rVk(i).Saettigung, "','" , rVk(i).Zielwerterreichung, "','" , rVk(i).Ketonkörper, "','" ,  _
   rVk(i).Gynaekologenbefund, "','" , rVk(i).Gewichtsentwicklung, "','" , rVk(i).HbA1c, "','" , rVk(i).Bewegung, "','" , rVk(i).Minuten, "','" , rVk(i).Blutdruck, "','" , rVk(i).Puls, "','" ,  _
   rVk(i).Mitarbeiter, "'," , rVk(i).absPos, "," , DatFor_k(rVk(i).AktZeit), "," , rVk(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rVk) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rVk) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(11)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rVk),i)
  IF Len(rVk(k).Wohlempfinden) > maxi(0) THEN maxi(0) = Len(rVk(k).Wohlempfinden)
  IF Len(rVk(k).Saettigung) > maxi(1) THEN maxi(1) = Len(rVk(k).Saettigung)
  IF Len(rVk(k).Zielwerterreichung) > maxi(2) THEN maxi(2) = Len(rVk(k).Zielwerterreichung)
  IF Len(rVk(k).Ketonkörper) > maxi(3) THEN maxi(3) = Len(rVk(k).Ketonkörper)
  IF Len(rVk(k).Gynaekologenbefund) > maxi(4) THEN maxi(4) = Len(rVk(k).Gynaekologenbefund)
  IF Len(rVk(k).Gewichtsentwicklung) > maxi(5) THEN maxi(5) = Len(rVk(k).Gewichtsentwicklung)
  IF Len(rVk(k).HbA1c) > maxi(6) THEN maxi(6) = Len(rVk(k).HbA1c)
  IF Len(rVk(k).Bewegung) > maxi(7) THEN maxi(7) = Len(rVk(k).Bewegung)
  IF Len(rVk(k).Minuten) > maxi(8) THEN maxi(8) = Len(rVk(k).Minuten)
  IF Len(rVk(k).Blutdruck) > maxi(9) THEN maxi(9) = Len(rVk(k).Blutdruck)
  IF Len(rVk(k).Puls) > maxi(10) THEN maxi(10) = Len(rVk(k).Puls)
  IF Len(rVk(k).Mitarbeiter) > maxi(11) THEN maxi(11) = Len(rVk(k).Mitarbeiter)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "vkgd", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "vkgd", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rVk), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rVk.Wohlempfinden: '" & rVk(k).Wohlempfinden & "' -> '" & Left$(rVk(k).Wohlempfinden, maxL)  & "'",true : rVk(k).Wohlempfinden = Left$(rVk(k).Wohlempfinden, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rVk.Saettigung: '" & rVk(k).Saettigung & "' -> '" & Left$(rVk(k).Saettigung, maxL)  & "'",true : rVk(k).Saettigung = Left$(rVk(k).Saettigung, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rVk.Zielwerterreichung: '" & rVk(k).Zielwerterreichung & "' -> '" & Left$(rVk(k).Zielwerterreichung, maxL)  & "'",true : rVk(k).Zielwerterreichung = Left$(rVk(k).Zielwerterreichung, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rVk.Ketonkörper: '" & rVk(k).Ketonkörper & "' -> '" & Left$(rVk(k).Ketonkörper, maxL)  & "'",true : rVk(k).Ketonkörper = Left$(rVk(k).Ketonkörper, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gynaekologenbefund: '" & rVk(k).Gynaekologenbefund & "' -> '" & Left$(rVk(k).Gynaekologenbefund, maxL)  & "'",true : rVk(k).Gynaekologenbefund = Left$(rVk(k).Gynaekologenbefund, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gewichtsentwicklung: '" & rVk(k).Gewichtsentwicklung & "' -> '" & Left$(rVk(k).Gewichtsentwicklung, maxL)  & "'",true : rVk(k).Gewichtsentwicklung = Left$(rVk(k).Gewichtsentwicklung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rVk.HbA1c: '" & rVk(k).HbA1c & "' -> '" & Left$(rVk(k).HbA1c, maxL)  & "'",true : rVk(k).HbA1c = Left$(rVk(k).HbA1c, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rVk.Bewegung: '" & rVk(k).Bewegung & "' -> '" & Left$(rVk(k).Bewegung, maxL)  & "'",true : rVk(k).Bewegung = Left$(rVk(k).Bewegung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rVk.Minuten: '" & rVk(k).Minuten & "' -> '" & Left$(rVk(k).Minuten, maxL)  & "'",true : rVk(k).Minuten = Left$(rVk(k).Minuten, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rVk.Blutdruck: '" & rVk(k).Blutdruck & "' -> '" & Left$(rVk(k).Blutdruck, maxL)  & "'",true : rVk(k).Blutdruck = Left$(rVk(k).Blutdruck, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rVk.Puls: '" & rVk(k).Puls & "' -> '" & Left$(rVk(k).Puls, maxL)  & "'",true : rVk(k).Puls = Left$(rVk(k).Puls, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rVk.Mitarbeiter: '" & rVk(k).Mitarbeiter & "' -> '" & Left$(rVk(k).Mitarbeiter, maxL)  & "'",true : rVk(k).Mitarbeiter = Left$(rVk(k).Mitarbeiter, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' vkgdSpeichern

Public FUNCTION roSwZuw(i&, j&)
 roSw(i).FID = rSw(j).FID
 roSw(i).Pat_ID = rSw(j).Pat_ID
 roSw(i).ZeitPunkt = rSw(j).ZeitPunkt
 roSw(i).LR = rSw(j).LR
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
End FUNCTION ' roSwZuw

Public FUNCTION SwZUnt%(i&, j&)
 IF roSw(i).FID <> rSw(j).FID THEN gosub unter
 IF roSw(i).Pat_ID <> rSw(j).Pat_ID THEN gosub unter
 IF roSw(i).ZeitPunkt <> rSw(j).ZeitPunkt THEN gosub unter
 IF roSw(i).LR <> rSw(j).LR THEN gosub unter
 IF roSw(i).vorET <> rSw(j).vorET THEN gosub unter
 IF roSw(i).ET <> rSw(j).ET THEN gosub unter
 IF roSw(i).efLR <> rSw(j).efLR THEN gosub unter
 IF roSw(i).erLR <> rSw(j).erLR THEN gosub unter
 IF roSw(i).kGT <> rSw(j).kGT THEN gosub unter
 IF roSw(i).MB <> rSw(j).MB THEN gosub unter
 IF roSw(i).EndeArt <> rSw(j).EndeArt THEN gosub unter
 IF roSw(i).ED <> rSw(j).ED THEN gosub unter
 IF roSw(i).absPos <> rSw(j).absPos THEN gosub unter
 IF roSw(i).AktZeit <> rSw(j).AktZeit THEN gosub unter
 IF roSw(i).StByte <> rSw(j).StByte THEN gosub unter
 Exit Function
unter:
 SwZUnt = SwZUnt + 1
 Return
End FUNCTION ' SwZUnt

Public FUNCTION swsLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roSw(1)
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,IF(ZeitPunkt IS NULL OR ZeitPunkt=0,CONVERT('18991230',DATE),ZeitPunkt) ZeitPunkt,IF(LR IS NULL OR LR=0,CONVERT('18991230',DATE),LR) LR" & _
",IF(vorET IS NULL OR vorET=0,CONVERT('18991230',DATE),vorET) vorET,IF(ET IS NULL OR ET=0,CONVERT('18991230',DATE),ET) ET,IF(efLR IS NULL OR efLR=0,CONVERT('18991230',DATE),efLR) efLR,IF(erLR IS NULL OR erLR=0,CONVERT('18991230',DATE),erLR) erLR" & _
",IF(kGT IS NULL OR kGT=0,CONVERT('18991230',DATE),kGT) kGT,IF(MB IS NULL OR MB=0,CONVERT('18991230',DATE),MB) MB,COALESCE(EndeArt,'') EndeArt,IF(ED IS NULL OR ED=0,CONVERT('18991230',DATE),ED) ED" & _
",COALESCE(absPos,0) absPos,IF(AktZeit IS NULL OR AktZeit=0,CONVERT('18991230',DATE),AktZeit) AktZeit,COALESCE(StByte,0) StByte FROM `sws` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 myfrag rs, sql
 Do While Not rs.EOF
  akt = UBound(roSw)
  roSw(akt).FID = rs!FID
  roSw(akt).Pat_ID = rs!Pat_ID
  roSw(akt).ZeitPunkt = rs!ZeitPunkt
  roSw(akt).LR = rs!LR
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
  IF Not rs.EOF THEN ReDim Preserve roSw(UBound(roSw) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' swsLaden

Function swsEinf
 Dim rbeg&, robeg&, ri&, roi&
 On Error GoTo fehler
 IF UBound(rSw) > 0 THEN
  rbeg = 0
  For ri = 1 To UBound(rSw)
   IF rSw(ri).ZeitPunkt >= qbeg THEN
    rbeg = ri
    Exit For
   END IF
  Next ri
  IF rbeg <> 0 THEN
   For robeg = 0 To UBound(roSw)
    IF roSw(robeg).ZeitPunkt >= qbeg THEN
     Exit For
    END IF
   Next robeg
   IF robeg <= UBound(roFa) THEN
    ReDim Preserve roSw(robeg - 1)
   END IF ' IF robeg <= UBound(roFa) THEN
   ReDim Preserve roSw(robeg + UBound(rSw) - rbeg)
   For ri = rbeg To UBound(rSw)
    Call roSwZuw(robeg + ri - rbeg, ri)
   Next ri
  END IF ' IF rbeg <> 0 THEN
 END IF ' IF UBound(rFa) > 0 THEN
 rSw = roSw
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 THEN
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #END IF
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END Select
End FUNCTION ' swsEinf

Public FUNCTION swsSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere sws"
 IF not Allepat THEN
  sql = "SELECT pat_id FROM `sws` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  myfrag rs, sql
  IF Not rs.BOF THEN
   SET rs = nothing
   sql = "DELETE FROM `sws` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  END IF
 END IF ' not allePat
' sql0 = " INSERT " & sqlignore &  "INTO `sws` (FID,Pat_ID,ZeitPunkt," & _
     "LR,vorET,ET,efLR,erLR,kGT,MB,EndeArt,ED,absPos," & _
     "AktZeit,StByte) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `sws` (FID,Pat_ID,ZeitPunkt," & _
     "LR,vorET,ET,efLR,erLR,kGT,MB,EndeArt,ED,absPos," & _
     "AktZeit,StByte)        VALUES"))
 For i = 1 to ubound(rSw)
'  rSw(i).AktZeit = now()
  rSw(i).StByte = CStr(AktByte)
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rSw(i).FID, "," , rSw(i).Pat_ID, "," , DatFor_k(rSw(i).ZeitPunkt), "," , DatFor_k(rSw(i).LR), "," , DatFor_k(rSw(i).vorET), "," , DatFor_k(rSw(i).ET), "," , DatFor_k(rSw(i).efLR), "," , DatFor_k( _
   rSw(i).erLR), "," , DatFor_k(rSw(i).kGT), "," , DatFor_k(rSw(i).MB), ",'" , rSw(i).EndeArt, "'," , DatFor_k(rSw(i).ED), "," , rSw(i).absPos, "," , DatFor_k(rSw(i).AktZeit), "," ,  _
   rSw(i).StByte, ")")
  IF SammelInsert <> 0 AND i < ubound(rSw) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rSw) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rSw),i)
  IF Len(rSw(k).EndeArt) > maxi(0) THEN maxi(0) = Len(rSw(k).EndeArt)
 next k
 IF obTrans <>0 THEN DBCn.CommitTrans: obtrans = 0
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "sws", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "sws", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rSw), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rSw.EndeArt: '" & rSw(k).EndeArt & "' -> '" & Left$(rSw(k).EndeArt, maxL)  & "'",true : rSw(k).EndeArt = Left$(rSw(k).EndeArt, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' swsSpeichern

Public FUNCTION laborxsaetzeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxsaetze"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             VALUES"))
 For i = 0 to ubound(rLs)
'  rLs(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 0 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLs(i).DatID, ",'" , rLs(i).Satzart, "','" , rLs(i).Satzlänge, "','" , rLs(i).SatzlängeSchluss, "','" , rLs(i).VersionSatzb, "','" , rLs(i).Arztnr, "','" , rLs(i).Arztname, "','" ,  _
   rLs(i).StraßePraxis, "','" , rLs(i).Arzt, "','" , rLs(i).LANR, "','" , rLs(i).PLZPraxis, "','" , rLs(i).OrtPraxis, "','" , rLs(i).Labor, "','" , rLs(i).StraßeLabor, "','" , rLs(i).PLZLabor, "','" ,  _
   rLs(i).OrtLabor, "','" , rLs(i).KBVPrüfnr, "','" , rLs(i).Zeichensatz, "','" , rLs(i).Kundenarztnr, "','" , rLs(i).Erstellungsdatum, "','" , rLs(i).Gesamtlänge, "')")
  IF SammelInsert <> 0 AND i < ubound(rLs) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLs) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(19)
 for k = iif(SammelInsert<>0,0,i) to iif(SammelInsert<>0,ubound(rLs),i)
  IF Len(rLs(k).Satzart) > maxi(0) THEN maxi(0) = Len(rLs(k).Satzart)
  IF Len(rLs(k).Satzlänge) > maxi(1) THEN maxi(1) = Len(rLs(k).Satzlänge)
  IF Len(rLs(k).SatzlängeSchluss) > maxi(2) THEN maxi(2) = Len(rLs(k).SatzlängeSchluss)
  IF Len(rLs(k).VersionSatzb) > maxi(3) THEN maxi(3) = Len(rLs(k).VersionSatzb)
  IF Len(rLs(k).Arztnr) > maxi(4) THEN maxi(4) = Len(rLs(k).Arztnr)
  IF Len(rLs(k).Arztname) > maxi(5) THEN maxi(5) = Len(rLs(k).Arztname)
  IF Len(rLs(k).StraßePraxis) > maxi(6) THEN maxi(6) = Len(rLs(k).StraßePraxis)
  IF Len(rLs(k).Arzt) > maxi(7) THEN maxi(7) = Len(rLs(k).Arzt)
  IF Len(rLs(k).LANR) > maxi(8) THEN maxi(8) = Len(rLs(k).LANR)
  IF Len(rLs(k).PLZPraxis) > maxi(9) THEN maxi(9) = Len(rLs(k).PLZPraxis)
  IF Len(rLs(k).OrtPraxis) > maxi(10) THEN maxi(10) = Len(rLs(k).OrtPraxis)
  IF Len(rLs(k).Labor) > maxi(11) THEN maxi(11) = Len(rLs(k).Labor)
  IF Len(rLs(k).StraßeLabor) > maxi(12) THEN maxi(12) = Len(rLs(k).StraßeLabor)
  IF Len(rLs(k).PLZLabor) > maxi(13) THEN maxi(13) = Len(rLs(k).PLZLabor)
  IF Len(rLs(k).OrtLabor) > maxi(14) THEN maxi(14) = Len(rLs(k).OrtLabor)
  IF Len(rLs(k).KBVPrüfnr) > maxi(15) THEN maxi(15) = Len(rLs(k).KBVPrüfnr)
  IF Len(rLs(k).Zeichensatz) > maxi(16) THEN maxi(16) = Len(rLs(k).Zeichensatz)
  IF Len(rLs(k).Kundenarztnr) > maxi(17) THEN maxi(17) = Len(rLs(k).Kundenarztnr)
  IF Len(rLs(k).Erstellungsdatum) > maxi(18) THEN maxi(18) = Len(rLs(k).Erstellungsdatum)
  IF Len(rLs(k).Gesamtlänge) > maxi(19) THEN maxi(19) = Len(rLs(k).Gesamtlänge)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxsaetze", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,0, i) To IIf(SammelInsert <> 0,ubound(rLs), i)
      SELECT CASE m
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
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxsaetzeSpeichern

Public FUNCTION laborxeingelSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxeingel"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  VALUES"))
 For i = 1 to ubound(rLg)
'  rLg(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rLg(i).Pfad, "','" , rLg(i).Name, "'," , DatFor_k(rLg(i).Zp), "," , cstr(-(rLg(i).fertig<>0)) , ")")
  IF SammelInsert <> 0 AND i < ubound(rLg) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLg) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLg),i)
  IF Len(rLg(k).Pfad) > maxi(0) THEN maxi(0) = Len(rLg(k).Pfad)
  IF Len(rLg(k).Name) > maxi(1) THEN maxi(1) = Len(rLg(k).Name)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLg), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL)  & "'",true : rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).Name & "' -> '" & Left$(rLg(k).Name, maxL)  & "'",true : rLg(k).Name = Left$(rLg(k).Name, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxeingelSpeichern

Public FUNCTION laborxusSpeichern(SammelInsert%, BezfSp%, j&)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxus"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxus` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,Nachname,Vorname,GebDat,Titel," & _
     "NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp,Pat_idErwVNG," & _
     "Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte,verglichen," & _
     "AfN)     VALUES"))
 For i = j to j
'  rLu(i).AktZeit = now()
  IF SammelInsert = 0 OR i = j THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLu(i).DatID, "," , rLu(i).SatzID, ",'" , rLu(i).Satzart, "','" , rLu(i).Satzlänge, "','" , rLu(i).Auftragsnummer, "','" , rLu(i).Auftragsschlüssel, "'," , DatFor_k(rLu(i).Eingang), ",'" ,  _
   rLu(i).Berichtsdatum, "'," , rLu(i).Pat_id, ",'" , rLu(i).Nachname, "','" , rLu(i).Vorname, "','" , rLu(i).GebDat, "','" , rLu(i).Titel, "','" , rLu(i).NVorsatz, "','" , rLu(i).BefArt, "','" ,  _
   rLu(i).Abrechnungstyp, "','" , rLu(i).GebüOrd, "','" , rLu(i).Auftraggeber, "','" , rLu(i).Patienteninformation, "','" , rLu(i).Geschlecht, "','" , rLu(i).AuftrHinw, "','" , rLu(i).Pat_idUrsp, "','" ,  _
   rLu(i).Pat_idErwVNG, "','" , rLu(i).Pat_idErwVN, "','" , rLu(i).Pat_idErwG, "','" , rLu(i).Pat_idErwGB, "','" , rLu(i).Pat_idErwGL, "','" , rLu(i).Pat_idLaborNeu, "'," , DatFor_k(rLu(i).ZeitpunktLaborneu), "," ,  _
   rLu(i).ZdüP, "," , rLu(i).ZdiP, ",'" , rLu(i).LWerte, "'," , DatFor_k(rLu(i).verglichen), "," , rLu(i).AfN, ")")
  IF SammelInsert <> 0 AND i < j THEN csql.Append ","
  IF SammelInsert = 0 OR i = j THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(24)
 for k = iif(SammelInsert<>0,j,i) to iif(SammelInsert<>0,j,i)
  IF Len(rLu(k).Satzart) > maxi(0) THEN maxi(0) = Len(rLu(k).Satzart)
  IF Len(rLu(k).Satzlänge) > maxi(1) THEN maxi(1) = Len(rLu(k).Satzlänge)
  IF Len(rLu(k).Auftragsnummer) > maxi(2) THEN maxi(2) = Len(rLu(k).Auftragsnummer)
  IF Len(rLu(k).Auftragsschlüssel) > maxi(3) THEN maxi(3) = Len(rLu(k).Auftragsschlüssel)
  IF Len(rLu(k).Berichtsdatum) > maxi(4) THEN maxi(4) = Len(rLu(k).Berichtsdatum)
  IF Len(rLu(k).Nachname) > maxi(5) THEN maxi(5) = Len(rLu(k).Nachname)
  IF Len(rLu(k).Vorname) > maxi(6) THEN maxi(6) = Len(rLu(k).Vorname)
  IF Len(rLu(k).GebDat) > maxi(7) THEN maxi(7) = Len(rLu(k).GebDat)
  IF Len(rLu(k).Titel) > maxi(8) THEN maxi(8) = Len(rLu(k).Titel)
  IF Len(rLu(k).NVorsatz) > maxi(9) THEN maxi(9) = Len(rLu(k).NVorsatz)
  IF Len(rLu(k).BefArt) > maxi(10) THEN maxi(10) = Len(rLu(k).BefArt)
  IF Len(rLu(k).Abrechnungstyp) > maxi(11) THEN maxi(11) = Len(rLu(k).Abrechnungstyp)
  IF Len(rLu(k).GebüOrd) > maxi(12) THEN maxi(12) = Len(rLu(k).GebüOrd)
  IF Len(rLu(k).Auftraggeber) > maxi(13) THEN maxi(13) = Len(rLu(k).Auftraggeber)
  IF Len(rLu(k).Patienteninformation) > maxi(14) THEN maxi(14) = Len(rLu(k).Patienteninformation)
  IF Len(rLu(k).Geschlecht) > maxi(15) THEN maxi(15) = Len(rLu(k).Geschlecht)
  IF Len(rLu(k).AuftrHinw) > maxi(16) THEN maxi(16) = Len(rLu(k).AuftrHinw)
  IF Len(rLu(k).Pat_idUrsp) > maxi(17) THEN maxi(17) = Len(rLu(k).Pat_idUrsp)
  IF Len(rLu(k).Pat_idErwVNG) > maxi(18) THEN maxi(18) = Len(rLu(k).Pat_idErwVNG)
  IF Len(rLu(k).Pat_idErwVN) > maxi(19) THEN maxi(19) = Len(rLu(k).Pat_idErwVN)
  IF Len(rLu(k).Pat_idErwG) > maxi(20) THEN maxi(20) = Len(rLu(k).Pat_idErwG)
  IF Len(rLu(k).Pat_idErwGB) > maxi(21) THEN maxi(21) = Len(rLu(k).Pat_idErwGB)
  IF Len(rLu(k).Pat_idErwGL) > maxi(22) THEN maxi(22) = Len(rLu(k).Pat_idErwGL)
  IF Len(rLu(k).Pat_idLaborNeu) > maxi(23) THEN maxi(23) = Len(rLu(k).Pat_idLaborNeu)
  IF Len(rLu(k).LWerte) > maxi(24) THEN maxi(24) = Len(rLu(k).LWerte)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxus", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,j, i) To IIf(SammelInsert <> 0,j, i)
      SELECT CASE m
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
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxusSpeichern

Public FUNCTION laborxbaktSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxbakt"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  VALUES"))
 For i = 1 to ubound(rLo)
'  rLo(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLo(i).RefNr, ",'" , rLo(i).Verf, "','" , rLo(i).KuQu, "','" , rLo(i).Quelle, "','" , rLo(i).QSpez, "'," , DatFor_k(rLo(i).AbnDat), ",'" , rLo(i).Kommentar, "','" , rLo(i).Erklärung, "','" ,  _
   rLo(i).Keimzahl, "')")
  IF SammelInsert <> 0 AND i < ubound(rLo) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLo) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(6)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLo),i)
  IF Len(rLo(k).Verf) > maxi(0) THEN maxi(0) = Len(rLo(k).Verf)
  IF Len(rLo(k).KuQu) > maxi(1) THEN maxi(1) = Len(rLo(k).KuQu)
  IF Len(rLo(k).Quelle) > maxi(2) THEN maxi(2) = Len(rLo(k).Quelle)
  IF Len(rLo(k).QSpez) > maxi(3) THEN maxi(3) = Len(rLo(k).QSpez)
  IF Len(rLo(k).Kommentar) > maxi(4) THEN maxi(4) = Len(rLo(k).Kommentar)
  IF Len(rLo(k).Erklärung) > maxi(5) THEN maxi(5) = Len(rLo(k).Erklärung)
  IF Len(rLo(k).Keimzahl) > maxi(6) THEN maxi(6) = Len(rLo(k).Keimzahl)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxbakt", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & Left$(rLo(k).Verf, maxL)  & "'",true : rLo(k).Verf = Left$(rLo(k).Verf, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & Left$(rLo(k).KuQu, maxL)  & "'",true : rLo(k).KuQu = Left$(rLo(k).KuQu, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & Left$(rLo(k).Quelle, maxL)  & "'",true : rLo(k).Quelle = Left$(rLo(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & Left$(rLo(k).QSpez, maxL)  & "'",true : rLo(k).QSpez = Left$(rLo(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & Left$(rLo(k).Kommentar, maxL)  & "'",true : rLo(k).Kommentar = Left$(rLo(k).Kommentar, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & Left$(rLo(k).Erklärung, maxL)  & "'",true : rLo(k).Erklärung = Left$(rLo(k).Erklärung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & Left$(rLo(k).Keimzahl, maxL)  & "'",true : rLo(k).Keimzahl = Left$(rLo(k).Keimzahl, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxbaktSpeichern

Public FUNCTION laborxwertSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxwert"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    VALUES"))
 For i = 1 to ubound(rLw)
'  rLw(i).AktZeit = now()
    Dim j&
    For j = 1 To i - 1
    IF rLw(i).RefNr <> rLw(j).RefNr THEN GoTo nextj
    IF rLw(i).Abkü <> rLw(j).Abkü THEN GoTo nextj
    IF rLw(i).Langname <> rLw(j).Langname THEN GoTo nextj
    IF rLw(i).Quelle <> rLw(j).Quelle THEN GoTo nextj
    IF rLw(i).QSpez <> rLw(j).QSpez THEN GoTo nextj
    IF rLw(i).AbnDat <> rLw(j).AbnDat THEN GoTo nextj
    IF rLw(i).Wert <> rLw(j).Wert THEN GoTo nextj
    IF rLw(i).Einheit <> rLw(j).Einheit THEN GoTo nextj
    IF rLw(i).Grenzwerti <> rLw(j).Grenzwerti THEN GoTo nextj
    IF rLw(i).Kommentar <> rLw(j).Kommentar THEN GoTo nextj
    IF rLw(i).Teststatus <> rLw(j).Teststatus THEN GoTo nextj
    IF rLw(i).Erklärung <> rLw(j).Erklärung THEN GoTo nextj
    IF rLw(i).AuftrHinw <> rLw(j).AuftrHinw THEN GoTo nextj
    IF rLw(i).nbid <> rLw(j).nbid THEN GoTo nextj
     GoTo nexti
nextj:
     Dim rsdop As New ADODB.Recordset
     SET rsdop = Nothing
     myfrag rsdop, "SELECT 0 FROM `laborxwert` WHERE `RefNr` = " & rLw(i).RefNr & " AND `Abkü` = '" & rLw(i).Abkü & "' AND `Langname` = '" & rLw(i).Langname & "' AND `Quelle` = '" &  _
   rLw(i).Quelle & "' AND `QSpez` = '" & rLw(i).QSpez & "' AND `AbnDat` = " & DatFor_k(rLw(i).AbnDat) & " AND `Wert` = '" &  _
   rLw(i).Wert & "' AND `Einheit` = '" & rLw(i).Einheit & "' AND `Grenzwerti` = '" & rLw(i).Grenzwerti & "' AND `Kommentar` = '" & rLw(i).Kommentar & "' AND `Teststatus` = '" &  _
   rLw(i).Teststatus & "' AND `Erklärung` = '" & rLw(i).Erklärung & "' AND `AuftrHinw` = '" & rLw(i).AuftrHinw & "' AND `nbid` = " &  _
   rLw(i).nbid & ""
     IF Not rsdop.EOF THEN GoTo nexti
    Next j
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLw(i).RefNr, ",'" , rLw(i).Abkü, "','" , rLw(i).Langname, "','" , rLw(i).Quelle, "','" , rLw(i).QSpez, "'," , DatFor_k(rLw(i).AbnDat), ",'" , rLw(i).Wert, "','" , rLw(i).Einheit, "','" ,  _
   rLw(i).Grenzwerti, "','" , rLw(i).Kommentar, "','" , rLw(i).Teststatus, "','" , rLw(i).Erklärung, "','" , rLw(i).AuftrHinw, "'," , rLw(i).nbid, ")")
  IF SammelInsert <> 0 AND i < ubound(rLw) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLw) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
nexti:
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(10)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLw),i)
  IF Len(rLw(k).Abkü) > maxi(0) THEN maxi(0) = Len(rLw(k).Abkü)
  IF Len(rLw(k).Langname) > maxi(1) THEN maxi(1) = Len(rLw(k).Langname)
  IF Len(rLw(k).Quelle) > maxi(2) THEN maxi(2) = Len(rLw(k).Quelle)
  IF Len(rLw(k).QSpez) > maxi(3) THEN maxi(3) = Len(rLw(k).QSpez)
  IF Len(rLw(k).Wert) > maxi(4) THEN maxi(4) = Len(rLw(k).Wert)
  IF Len(rLw(k).Einheit) > maxi(5) THEN maxi(5) = Len(rLw(k).Einheit)
  IF Len(rLw(k).Grenzwerti) > maxi(6) THEN maxi(6) = Len(rLw(k).Grenzwerti)
  IF Len(rLw(k).Kommentar) > maxi(7) THEN maxi(7) = Len(rLw(k).Kommentar)
  IF Len(rLw(k).Teststatus) > maxi(8) THEN maxi(8) = Len(rLw(k).Teststatus)
  IF Len(rLw(k).Erklärung) > maxi(9) THEN maxi(9) = Len(rLw(k).Erklärung)
  IF Len(rLw(k).AuftrHinw) > maxi(10) THEN maxi(10) = Len(rLw(k).AuftrHinw)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxwert", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLw), i)
      SELECT CASE m
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
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxwertSpeichern

Public FUNCTION laborxleistSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxleist"
' sql0 = " INSERT " & sqlignore &  "INTO `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   VALUES"))
 For i = 1 to ubound(rLL)
'  rLL(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("(" , rLL(i).RefNr, ",'" , rLL(i).Abkü, "','" , rLL(i).Verf, "','" , rLL(i).EBM, "','" , rLL(i).goä, "','" , rLL(i).Anzahl, "','" , rLL(i).abrd, "')")
  IF SammelInsert <> 0 AND i < ubound(rLL) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLL) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(5)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLL),i)
  IF Len(rLL(k).Abkü) > maxi(0) THEN maxi(0) = Len(rLL(k).Abkü)
  IF Len(rLL(k).Verf) > maxi(1) THEN maxi(1) = Len(rLL(k).Verf)
  IF Len(rLL(k).EBM) > maxi(2) THEN maxi(2) = Len(rLL(k).EBM)
  IF Len(rLL(k).goä) > maxi(3) THEN maxi(3) = Len(rLL(k).goä)
  IF Len(rLL(k).Anzahl) > maxi(4) THEN maxi(4) = Len(rLL(k).Anzahl)
  IF Len(rLL(k).abrd) > maxi(5) THEN maxi(5) = Len(rLL(k).abrd)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxleist", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLL), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & Left$(rLL(k).Abkü, maxL)  & "'",true : rLL(k).Abkü = Left$(rLL(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & Left$(rLL(k).Verf, maxL)  & "'",true : rLL(k).Verf = Left$(rLL(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & Left$(rLL(k).EBM, maxL)  & "'",true : rLL(k).EBM = Left$(rLL(k).EBM, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & Left$(rLL(k).goä, maxL)  & "'",true : rLL(k).goä = Left$(rLL(k).goä, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & Left$(rLL(k).Anzahl, maxL)  & "'",true : rLL(k).Anzahl = Left$(rLL(k).Anzahl, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.abrd: '" & rLL(k).abrd & "' -> '" & Left$(rLL(k).abrd, maxL)  & "'",true : rLL(k).abrd = Left$(rLL(k).abrd, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' laborxleistSpeichern

Public FUNCTION liuezSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere liuez"
' sql0 = " INSERT " & sqlignore &  "INTO `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel,zusatz,ursp," & _
     "aktzeit) VALUES
 Call csql0.AppVar(Array(" INSERT ", sqlIgnore, " INTO `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel,zusatz,ursp," & _
     "aktzeit)               VALUES"))
 For i = 1 to ubound(rLi)
'  rLi(i).AktZeit = now()
  IF SammelInsert = 0 OR i = 1 THEN
   csql.Append csql0
  END IF
  csql.AppVar Array("('" , rLi(i).name, "','" , rLi(i).vorname, "','" , rLi(i).titelt, "','" , rLi(i).fachgruppe, "','" , rLi(i).strasse, "','" , rLi(i).plz, "','" , rLi(i).ort, "','" , rLi(i).telefon, "','" ,  _
   rLi(i).fax, "','" , rLi(i).kvnr, "'," , DatFor_k(rLi(i).aktdat), ",'" , rLi(i).überschrift, "','" , rLi(i).dbnr, "','" , rLi(i).bstelle, "','" , rLi(i).anrede, "','" , rLi(i).tel1, "','" ,  _
   rLi(i).tel2, "','" , rLi(i).tel3, "','" , rLi(i).tel4, "','" , rLi(i).fax1, "','" , rLi(i).fax2, "','" , rLi(i).fax3, "','" , rLi(i).email, "','" , rLi(i).zulg, "','" , rLi(i).arzttyp, "','" ,  _
   rLi(i).gemmit, "','" , rLi(i).beme, "'," , rLi(i).dmpt2, "," , rLi(i).dmpt1, ",'" , rLi(i).geschlecht, "','" , rLi(i).titel, "','" , rLi(i).zusatz, "','" , rLi(i).ursp, "'," , DatFor_k(rLi(i).aktzeit), ")")
  IF SammelInsert <> 0 AND i < ubound(rLi) THEN csql.Append ","
  IF SammelInsert = 0 OR i = ubound(rLi) THEN
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   IF obfor THEN
    Call ForeignYes0
    Call ForeignYes1
   END IF
   IF sammelinsert = 0 THEN csql.m_len = 0
   IF lese.obmysql AND obmitAlterTab THEN
    SET rs = DBCn.Execute("SHOW WARNINGS")
    IF not rs.BOF() THEN
     IF rs!code = 1265 THEN
      Err.Raise -2147217833
     END IF
    END IF
   END IF
   END IF
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
ErrDescription = Err.Description
IF Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 THEN
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 THEN ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a foreign key constraint fails
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 END IF
 Resume
ElseIf Err.Number = -2147217833 OR InStrB(ErrDescription, "Daten zu lang") <> 0 THEN
 Dim rsc As Adodb.Recordset, maxi%(), k%
 redim maxi(29)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLi),i)
  IF Len(rLi(k).name) > maxi(0) THEN maxi(0) = Len(rLi(k).name)
  IF Len(rLi(k).vorname) > maxi(1) THEN maxi(1) = Len(rLi(k).vorname)
  IF Len(rLi(k).titelt) > maxi(2) THEN maxi(2) = Len(rLi(k).titelt)
  IF Len(rLi(k).fachgruppe) > maxi(3) THEN maxi(3) = Len(rLi(k).fachgruppe)
  IF Len(rLi(k).strasse) > maxi(4) THEN maxi(4) = Len(rLi(k).strasse)
  IF Len(rLi(k).plz) > maxi(5) THEN maxi(5) = Len(rLi(k).plz)
  IF Len(rLi(k).ort) > maxi(6) THEN maxi(6) = Len(rLi(k).ort)
  IF Len(rLi(k).telefon) > maxi(7) THEN maxi(7) = Len(rLi(k).telefon)
  IF Len(rLi(k).fax) > maxi(8) THEN maxi(8) = Len(rLi(k).fax)
  IF Len(rLi(k).kvnr) > maxi(9) THEN maxi(9) = Len(rLi(k).kvnr)
  IF Len(rLi(k).überschrift) > maxi(10) THEN maxi(10) = Len(rLi(k).überschrift)
  IF Len(rLi(k).dbnr) > maxi(11) THEN maxi(11) = Len(rLi(k).dbnr)
  IF Len(rLi(k).bstelle) > maxi(12) THEN maxi(12) = Len(rLi(k).bstelle)
  IF Len(rLi(k).anrede) > maxi(13) THEN maxi(13) = Len(rLi(k).anrede)
  IF Len(rLi(k).tel1) > maxi(14) THEN maxi(14) = Len(rLi(k).tel1)
  IF Len(rLi(k).tel2) > maxi(15) THEN maxi(15) = Len(rLi(k).tel2)
  IF Len(rLi(k).tel3) > maxi(16) THEN maxi(16) = Len(rLi(k).tel3)
  IF Len(rLi(k).tel4) > maxi(17) THEN maxi(17) = Len(rLi(k).tel4)
  IF Len(rLi(k).fax1) > maxi(18) THEN maxi(18) = Len(rLi(k).fax1)
  IF Len(rLi(k).fax2) > maxi(19) THEN maxi(19) = Len(rLi(k).fax2)
  IF Len(rLi(k).fax3) > maxi(20) THEN maxi(20) = Len(rLi(k).fax3)
  IF Len(rLi(k).email) > maxi(21) THEN maxi(21) = Len(rLi(k).email)
  IF Len(rLi(k).zulg) > maxi(22) THEN maxi(22) = Len(rLi(k).zulg)
  IF Len(rLi(k).arzttyp) > maxi(23) THEN maxi(23) = Len(rLi(k).arzttyp)
  IF Len(rLi(k).gemmit) > maxi(24) THEN maxi(24) = Len(rLi(k).gemmit)
  IF Len(rLi(k).beme) > maxi(25) THEN maxi(25) = Len(rLi(k).beme)
  IF Len(rLi(k).geschlecht) > maxi(26) THEN maxi(26) = Len(rLi(k).geschlecht)
  IF Len(rLi(k).titel) > maxi(27) THEN maxi(27) = Len(rLi(k).titel)
  IF Len(rLi(k).zusatz) > maxi(28) THEN maxi(28) = Len(rLi(k).zusatz)
  IF Len(rLi(k).ursp) > maxi(29) THEN maxi(29) = Len(rLi(k).ursp)
 next k
 nochmal:
 SET rsc = New ADODB.Recordset
 SET rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "liuez", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "liuez", rsc)
    IF maxL > 0 THEN
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLi.name: '" & rLi(k).name & "' -> '" & Left$(rLi(k).name, maxL)  & "'",true : rLi(k).name = Left$(rLi(k).name, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLi.vorname: '" & rLi(k).vorname & "' -> '" & Left$(rLi(k).vorname, maxL)  & "'",true : rLi(k).vorname = Left$(rLi(k).vorname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLi.titelt: '" & rLi(k).titelt & "' -> '" & Left$(rLi(k).titelt, maxL)  & "'",true : rLi(k).titelt = Left$(rLi(k).titelt, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLi.fachgruppe: '" & rLi(k).fachgruppe & "' -> '" & Left$(rLi(k).fachgruppe, maxL)  & "'",true : rLi(k).fachgruppe = Left$(rLi(k).fachgruppe, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLi.strasse: '" & rLi(k).strasse & "' -> '" & Left$(rLi(k).strasse, maxL)  & "'",true : rLi(k).strasse = Left$(rLi(k).strasse, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLi.plz: '" & rLi(k).plz & "' -> '" & Left$(rLi(k).plz, maxL)  & "'",true : rLi(k).plz = Left$(rLi(k).plz, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLi.ort: '" & rLi(k).ort & "' -> '" & Left$(rLi(k).ort, maxL)  & "'",true : rLi(k).ort = Left$(rLi(k).ort, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLi.telefon: '" & rLi(k).telefon & "' -> '" & Left$(rLi(k).telefon, maxL)  & "'",true : rLi(k).telefon = Left$(rLi(k).telefon, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax: '" & rLi(k).fax & "' -> '" & Left$(rLi(k).fax, maxL)  & "'",true : rLi(k).fax = Left$(rLi(k).fax, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLi.kvnr: '" & rLi(k).kvnr & "' -> '" & Left$(rLi(k).kvnr, maxL)  & "'",true : rLi(k).kvnr = Left$(rLi(k).kvnr, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLi.überschrift: '" & rLi(k).überschrift & "' -> '" & Left$(rLi(k).überschrift, maxL)  & "'",true : rLi(k).überschrift = Left$(rLi(k).überschrift, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLi.dbnr: '" & rLi(k).dbnr & "' -> '" & Left$(rLi(k).dbnr, maxL)  & "'",true : rLi(k).dbnr = Left$(rLi(k).dbnr, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLi.bstelle: '" & rLi(k).bstelle & "' -> '" & Left$(rLi(k).bstelle, maxL)  & "'",true : rLi(k).bstelle = Left$(rLi(k).bstelle, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLi.anrede: '" & rLi(k).anrede & "' -> '" & Left$(rLi(k).anrede, maxL)  & "'",true : rLi(k).anrede = Left$(rLi(k).anrede, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel1: '" & rLi(k).tel1 & "' -> '" & Left$(rLi(k).tel1, maxL)  & "'",true : rLi(k).tel1 = Left$(rLi(k).tel1, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel2: '" & rLi(k).tel2 & "' -> '" & Left$(rLi(k).tel2, maxL)  & "'",true : rLi(k).tel2 = Left$(rLi(k).tel2, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel3: '" & rLi(k).tel3 & "' -> '" & Left$(rLi(k).tel3, maxL)  & "'",true : rLi(k).tel3 = Left$(rLi(k).tel3, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel4: '" & rLi(k).tel4 & "' -> '" & Left$(rLi(k).tel4, maxL)  & "'",true : rLi(k).tel4 = Left$(rLi(k).tel4, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax1: '" & rLi(k).fax1 & "' -> '" & Left$(rLi(k).fax1, maxL)  & "'",true : rLi(k).fax1 = Left$(rLi(k).fax1, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax2: '" & rLi(k).fax2 & "' -> '" & Left$(rLi(k).fax2, maxL)  & "'",true : rLi(k).fax2 = Left$(rLi(k).fax2, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax3: '" & rLi(k).fax3 & "' -> '" & Left$(rLi(k).fax3, maxL)  & "'",true : rLi(k).fax3 = Left$(rLi(k).fax3, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLi.email: '" & rLi(k).email & "' -> '" & Left$(rLi(k).email, maxL)  & "'",true : rLi(k).email = Left$(rLi(k).email, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLi.zulg: '" & rLi(k).zulg & "' -> '" & Left$(rLi(k).zulg, maxL)  & "'",true : rLi(k).zulg = Left$(rLi(k).zulg, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLi.arzttyp: '" & rLi(k).arzttyp & "' -> '" & Left$(rLi(k).arzttyp, maxL)  & "'",true : rLi(k).arzttyp = Left$(rLi(k).arzttyp, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLi.gemmit: '" & rLi(k).gemmit & "' -> '" & Left$(rLi(k).gemmit, maxL)  & "'",true : rLi(k).gemmit = Left$(rLi(k).gemmit, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rLi.beme: '" & rLi(k).beme & "' -> '" & Left$(rLi(k).beme, maxL)  & "'",true : rLi(k).beme = Left$(rLi(k).beme, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rLi.geschlecht: '" & rLi(k).geschlecht & "' -> '" & Left$(rLi(k).geschlecht, maxL)  & "'",true : rLi(k).geschlecht = Left$(rLi(k).geschlecht, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rLi.titel: '" & rLi(k).titel & "' -> '" & Left$(rLi(k).titel, maxL)  & "'",true : rLi(k).titel = Left$(rLi(k).titel, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rLi.zusatz: '" & rLi(k).zusatz & "' -> '" & Left$(rLi(k).zusatz, maxL)  & "'",true : rLi(k).zusatz = Left$(rLi(k).zusatz, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rLi.ursp: '" & rLi(k).ursp & "' -> '" & Left$(rLi(k).ursp, maxL)  & "'",true : rLi(k).ursp = Left$(rLi(k).ursp, maxL)
      END SELECT
     Next
    elseif maxl < 0 THEN
     goto nochmal
    END IF
    m = m + 1
  END SELECT
  IF rsc.State = 0 THEN Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 IF lese.obMysql THEN Resume next ELSE resume
END IF ' Err.Number = -2147217833 THEN
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in liuezSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' liuezSpeichern

Public FUNCTION tuLaden
 call namenLaden
 call faelleLaden
   IF not lese.obmysql THEN
    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   END IF
 call auLaden
 call briefeLaden
 call diagnosenLaden
 call dokumenteLaden
 call eintraegeLaden
   IF not lese.obmysql THEN
    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   END IF
 call forminhkopfLaden
 call kheinweisLaden
 call lbanforderungenLaden
 call laborneuLaden
 call leistungenLaden
 call medplanLaden
 call rezepteintraegeLaden
 call rrLaden
 call dmpreiheLaden
 call desktopLaden
 call usdmLaden
 call fussLaden
 call ulcusLaden
 call vkgdLaden
 call swsLaden
End FUNCTION ' tuLaden

Public FUNCTION tuSpeichern(frm AS Lese, SI%, BfS%) ' frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern
 Dim rAf&
 ON Error GoTo fehler
 call namenSpeichern(SI, BfS)
 call faelleSpeichern(SI, BfS)
   IF not lese.obmysql THEN
    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   END IF
 call auSpeichern(SI, BfS)
 call briefeSpeichern(SI, BfS)
 call diagnosenSpeichern(SI, BfS)
 call dokumenteSpeichern(SI, BfS)
 call eintraegeSpeichern(SI, BfS)
 call forminhaltform_abkSpeichern(SI, BfS)
   IF not lese.obmysql THEN
    IF obTrans <> 0 THEN Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   END IF
 call formulareSpeichern(SI, BfS)
 call forminhkopfSpeichern(SI, BfS)
 call forminhfeldSpeichern(SI, BfS)
 call kheinweisSpeichern(SI, BfS)
 call lbanforderungenSpeichern(SI, BfS)
 call laborneuSpeichern(SI, BfS)
 call leistungenSpeichern(SI, BfS)
 call medplanSpeichern(SI, BfS)
 call rezepteintraegeSpeichern(SI, BfS)
 call rrSpeichern(SI, BfS)
 call kvnrueSpeichern(SI, BfS)
 call unbekannte_kennungenSpeichern(SI, BfS)
 call dmpreiheSpeichern(SI, BfS)
 call desktopSpeichern(SI, BfS)
 call usdmSpeichern(SI, BfS)
 call fussSpeichern(SI, BfS)
 call ulcusSpeichern(SI, BfS)
 call vkgdSpeichern(SI, BfS)
 call swsSpeichern(SI, BfS)
 Call DBCn.Execute("UPDATE `namen` SET aktZeit = " & DatFor_k(rNa(0).AktZeit) & " WHERE pat_id = " & rNa(0).Pat_ID,rAf)
 IF rAf <> 1 THEN 
  frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_ID & " " & rNa(0).Nachname & " " & rNa(0).Vorname, true
 END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 THEN
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#END IF
 ErrDescription = Err.Description
 IF InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 THEN
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
  Resume
 END IF
 SELECT CASE MsgBox("FNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in tuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Progende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End FUNCTION ' tuSpeichern

