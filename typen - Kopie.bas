Option Explicit
Public obFor%
Dim sql$, T1!, T2!, rs AS ADODB.Recordset, maxL%

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
 f3800 AS string 'f3800 varchar '3800
 dmpklass AS long 'dmpklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier, 4 = DMP ausgeschrieben
 dmpbeg AS date 'dmpbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass AS long 'dmpkhkklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg AS date 'dmpkhkbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass AS long 'dmpcopdklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg AS date 'dmpcopdbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpabklass AS long 'dmpabklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpabbeg AS date 'dmpabbeg date 'Datum der aktuellen DMP-Klassifikation
 dakab AS date 'dakab date 'DAK-Einverständnis-Datum
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
 Quelldatum AS date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 Typ AS string 'Typ varchar '
 AktZeit AS date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe AS long 'DokGroe int 'Größe der Datei
 DokAenD AS date 'DokAenD datetime 'Dokument-letzte Änderung
 QS AS string 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT AS string 'QT varchar 'Quartal des Behandlungsfallbeginns
 absPos AS long 'absPos int 'Zeile in der BDT-Datei
 StByte AS long 'StByte int 'Ordnungsnummer der Datenübertragung
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
 f6010 AS byte 'f6010 tinyint '6010 Diagnose gelöscht (Karteikarteneintrag: bdd)
 f6011 AS string 'f6011 varchar '6011 8.12.10: bisher nur ""TM#?""
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
 Medikament AS string 'Medikament varchar 'Referenz auf medarten.Medikament
 MedAnfang AS string 'MedAnfang varchar '
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
 ok AS integer 'ok bit 'ob ""ok""
 ausgedruckt AS integer 'ausgedruckt bit 'ob ""ausgedruckt""
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
 Spritzst AS string 'Spritzst varchar 'Spritzstellen/Liphyper./~atr.
 Fußbef_re AS string 'Fußbef_re varchar 'Fußbefund rechts
 Fußbef_li AS string 'Fußbef_li varchar 'Fußbefund links
 Hyperk_re AS string 'Hyperk_re varchar 'Hyperkeratosen rechts
 Hyperk_li AS string 'Hyperk_li varchar 'Hyperkeratosen links
 Ulcera_re AS string 'Ulcera_re varchar 'Ulcera rechts
 Ulcera_li AS string 'Ulcera_li varchar 'Ulcera links
 Kraft_Zh_re AS string 'Kraft_Zh_re varchar 'Kraft Zehenheber re
 Kraft_Zh_li AS string 'Kraft_Zh_li varchar 'Kraft Zehenheber li
 Kraft_Zb_re AS string 'Kraft_Zb_re varchar 'Kraft Zehenbeuger re
 Kraft_Zb_li AS string 'Kraft_Zb_li varchar 'Kraft Zehenbeuger li
 Kraft_Knie_re AS string 'Kraft_Knie_re varchar 'Kraft Knie re
 Kraft_Knie_li AS string 'Kraft_Knie_li varchar 'Kraft Knie li
 ASR_re AS string 'ASR_re varchar 'Achillessehnenreflex rechts
 ASR_li AS string 'ASR_li varchar 'Achillessehnenreflex links
 PSR_re AS string 'PSR_re varchar 'Patellarsehnenreflex rechts
 PSR_li AS string 'PSR_li varchar 'Patellarsehnenreflex links
 Oberfl_re AS string 'Oberfl_re varchar 'Oberflächensensibilität rechts
 Oberfl_li AS string 'Oberfl_li varchar 'Oberflächensensibilität links
 MF_re AS string 'MF_re varchar 'Monofilament rechts
 MF_li AS string 'MF_li varchar 'Monofilament links
 KW_re AS string 'KW_re varchar 'Kalt-warm rechts
 KW_li AS string 'KW_li varchar 'Kalt-warm links
 Vibr_IK_re AS string 'Vibr_IK_re varchar 'Vibration Innenknöchel rechts
 Vibr_IK_li AS string 'Vibr_IK_li varchar 'Vibration Innenknöchel links
 Vibr_GZ_re AS string 'Vibr_GZ_re varchar 'Vibration Großzehe rechts
 Vibr_GZ_li AS string 'Vibr_GZ_li varchar 'Vibration Großzehe links
 PulsL_re AS string 'PulsL_re varchar 'Puls Leiste rechts
 PulsL_li AS string 'PulsL_li varchar 'Puls Leiste links
 PulsKK_re AS string 'PulsKK_re varchar 'Puls Kniekehle rechts
 PulsKK_li AS string 'PulsKK_li varchar 'Puls Kniekehle links
 PulsAtp_re AS string 'PulsAtp_re varchar 'Puls Arteria tibialis posterior rechts
 PulsAtp_li AS string 'PulsAtp_li varchar 'Puls Arteria tibialis posterior links
 PulsAdp_re AS string 'PulsAdp_re varchar 'Puls Arteria dorsalis pedis rechts
 PulsAdp_li AS string 'PulsAdp_li varchar 'Puls Arteria dorsalis pedis links
 Mitarbeiter AS string 'Mitarbeiter varchar 'Mitarbeiter
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
 BefArt AS string 'BefArt varchar '8401 Befundart (Turbomed) / Fertigstellungsgrad (""E""=Endbefund, ""T"" = Teilbefund)
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
 Diabetes_seit AS string 'Diabetes_seit varchar '<seit
 Tabletten_seit AS string 'Tabletten_seit varchar ', Tabletten seit
 Insulin_seit AS string 'Insulin_seit varchar ', Insulin seit
 Grund_für_Vorstellung AS string 'Grund_für_Vorstellung varchar '^:
 Familienanamnese AS string 'Familienanamnese varchar '^:
 Größe AS double 'Größe double '^:
 Gewicht AS double 'Gewicht double ',:
 bmi AS double 'bmi decimal '
 Tendenz AS string 'Tendenz varchar '<, Tendenz
 DiabetesMedikament_1 AS string 'DiabetesMedikament_1 varchar '^Letzte Diabetesmedikation:
 DiabetesMedikament_1_Menge AS string 'DiabetesMedikament_1_Menge varchar '<
 DiabetesMedikament_2 AS string 'DiabetesMedikament_2 varchar '<,
 DiabetesMedikament_2_Menge AS string 'DiabetesMedikament_2_Menge varchar '<
 DiabetesMedikament_3 AS string 'DiabetesMedikament_3 varchar '<,
 DiabetesMedikament_3_Menge AS string 'DiabetesMedikament_3_Menge varchar '<
 DiabetesMedikament_4 AS string 'DiabetesMedikament_4 varchar '<,
 DiabetesMedikament_4_Menge AS string 'DiabetesMedikament_4_Menge varchar '<,
 Insulinpumpe AS byte 'Insulinpumpe tinyint '^:
 Insulinpumpe_seit AS string 'Insulinpumpe_seit varchar '<seit
 Insulinpumpe_Marke AS string 'Insulinpumpe_Marke varchar '<, Marke:
 Broteinheiten_gesamt AS string 'Broteinheiten_gesamt varchar '^Broteinheiten:gesamt
 Broteinheiten_früh AS string 'Broteinheiten_früh varchar '<, früh
 Broteinheiten_ZM_früh AS string 'Broteinheiten_ZM_früh varchar '<, Zwischenmahlzeit vormittags
 Broteinheiten_mittags AS string 'Broteinheiten_mittags varchar '<, mittags
 Broteinheiten_nachmittags AS string 'Broteinheiten_nachmittags varchar '<, nachmittags
 Broteinheiten_abends AS string 'Broteinheiten_abends varchar '<, abends
 Broteinheiten_nachts AS string 'Broteinheiten_nachts varchar '<, nachts
 Essenszeit_früh AS string 'Essenszeit_früh varchar '^Essenszeiten:früh
 Essenszeit_vormittags AS string 'Essenszeit_vormittags varchar '<, vormittags
 Essenszeit_mittags AS string 'Essenszeit_mittags varchar '<, mittags
 Essenszeit_nachmittags AS string 'Essenszeit_nachmittags varchar '<, nachmittags
 Essenszeit_abends AS string 'Essenszeit_abends varchar '<, abends
 Essenszeit_spät AS string 'Essenszeit_spät varchar '<, spät
 Spritz_Eß_Abstand_früh AS string 'Spritz_Eß_Abstand_früh varchar '^Spritz-Eß-Abstand:früh
 Spritz_Eß_Abstand_mittags AS string 'Spritz_Eß_Abstand_mittags varchar '<, mittags
 Spritz_Eß_Abstand_abends AS string 'Spritz_Eß_Abstand_abends varchar '<, abends
 Spritzstelle_früh AS string 'Spritzstelle_früh varchar '^Spritzstellen:früh
 Spritzstelle_mittags AS string 'Spritzstelle_mittags varchar '<, mittags
 Spritzstelle_abends AS string 'Spritzstelle_abends varchar '<, abends
 Spritzstelle_nachts AS string 'Spritzstelle_nachts varchar '<, nachts
 Ernährung AS string 'Ernährung varchar '^:
 Jahr_letzte_Diabetesschulung AS string 'Jahr_letzte_Diabetesschulung varchar '^Letzte Diabetesschulung:
 Ort_Schulung AS string 'Ort_Schulung varchar '<in
 letztes_HbA1c AS string 'letztes_HbA1c varchar '^Letztes HbA1c:
 gemessen_am AS string 'gemessen_am varchar '<, gemessen
 vorherige_Werte AS string 'vorherige_Werte varchar '<, vorher:
 subcutane_Zuckermessung AS string 'subcutane_Zuckermessung varchar '^:
 CGM_seit AS string 'CGM_seit varchar '<
 BZMessungen_selbst AS string 'BZMessungen_selbst varchar '^Blutzuckermessung:Selbstmessung?
 Gerät AS string 'Gerät varchar '<:
 BZMessungen_pW AS string 'BZMessungen_pW varchar '<Zahl d.Messungen pro Woche:
 BZMessungen_pW_ndE AS string 'BZMessungen_pW_ndE varchar '<, davon nach dem Essen:
 BZMessungen_p_W_nachts AS string 'BZMessungen_p_W_nachts varchar '<, nachts:
 Aufschreiben AS string 'Aufschreiben varchar '<, Dokumentation:
 BZWerte_v_d_Essen AS string 'BZWerte_v_d_Essen varchar '^Blutzuckerwerte vor dem Essen:
 BZWerte_n_d_Essen AS string 'BZWerte_n_d_Essen varchar '<, nach dem Essen:
 UZ_Tageszeit AS string 'UZ_Tageszeit varchar '^Unterzucker:Bevorzugte Tages-/Uhrzeit
 Unterzucker_pM AS string 'Unterzucker_pM varchar '<Zahl der schweren (<50 mg/dl) pro Monat:
 UZ_rechtzeitig AS string 'UZ_rechtzeitig varchar '<, rechtzeitig bemerkt:
 Schwere_Uzu AS string 'Schwere_Uzu varchar 'Schwere Unterzucker
 Fremde_Hilfe_pa AS string 'Fremde_Hilfe_pa varchar '<, fremde Hilfe deshalb nötig:
 Bewußtlos_pa AS string 'Bewußtlos_pa varchar '<, bewußtlos deshalb:
 Keto AS string 'Keto varchar '^Bisher Ketoazidosen mit Krankenhauseinweisung:
 Keto_pa AS string 'Keto_pa varchar '^Zahl der Ketoazidosen pro Jahr:
 BZgr300_pM AS string 'BZgr300_pM varchar ', Zahl der Blutzucker > 300 mg/dl pro Monat:
 Bluthochdruck AS string 'Bluthochdruck varchar '^Bluthochdruck:
 BHD_seit AS string 'BHD_seit varchar '<seit:
 BHD_beh_mit AS string 'BHD_beh_mit varchar '<, behandelt mit:
 Blutdruckwerte AS string 'Blutdruckwerte varchar '^Blutdruckwerte:
 BDselbst AS string 'BDselbst varchar '^Blutdruckselbstmessung:
 Schwanger AS string 'Schwanger varchar '^Aktuelle Schwangerschaft:
 Schwanger_seit AS string 'Schwanger_seit varchar '<, seit:
 Augensp_zuletzt AS string 'Augensp_zuletzt varchar '^Letzte Augenspiegelung:
 Augensp_Befund AS string 'Augensp_Befund varchar '<, Befund:
 Netzhaut_gelasert AS string 'Netzhaut_gelasert varchar ', Netzhaut schon gelasert:
 Sehminderung_unbehebbar AS string 'Sehminderung_unbehebbar varchar ', mit Brille nicht behebbare Sehminderung:
 Diabet_Nierenschaden AS string 'Diabet_Nierenschaden varchar '^Diabetischer Nierenschaden:
 Albumin_zuletzt AS string 'Albumin_zuletzt varchar ', letztes Albumin:
 erhöht AS string 'erhöht varchar '<, Befund:
 Kreatinin AS string 'Kreatinin varchar ',:
 Dialyse AS byte 'Dialyse tinyint ',:
 Dialyse_seit AS string 'Dialyse_seit varchar '<seit
 andere_Nierenerkrankung AS string 'andere_Nierenerkrankung varchar ', andere Nierenerkrankung:
 Herzkrankheit AS string 'Herzkrankheit varchar '^Herzkrankheit:
 Angina_pectoris AS string 'Angina_pectoris varchar ',:
 Herzinfarkt AS string 'Herzinfarkt varchar ',:
 Herzinfarkt_wann AS string 'Herzinfarkt_wann varchar '<, wann:
 PTCA_oder_Stent AS string 'PTCA_oder_Stent varchar ',:
 Bypass_kardial AS byte 'Bypass_kardial tinyint ',:
 Bypass_wann AS string 'Bypass_wann varchar '<, wann:
 Herzschwäche AS string 'Herzschwäche varchar ',:
 Herzkrankheit_Beschreibung AS string 'Herzkrankheit_Beschreibung varchar ', Beschreibung:
 Hirndurchblutungsstörung AS string 'Hirndurchblutungsstörung varchar '^:
 Schlaganfall AS string 'Schlaganfall varchar ',:
 Beindurchblutungsstörung AS string 'Beindurchblutungsstörung varchar '^:
 Schaufensterkrankheit AS string 'Schaufensterkrankheit varchar ',:
 Bypaß_peripher AS byte 'Bypaß_peripher tinyint ',:
 Geschwür AS string 'Geschwür varchar ',:
 Amputation AS string 'Amputation varchar ',:
 pAVK_Beschreibung AS string 'pAVK_Beschreibung varchar ', Beschreibung der Beinbeschwerden:
 Ameisenlaufen AS string 'Ameisenlaufen varchar '^:
 Ameisen_Ausmaß AS string 'Ameisen_Ausmaß varchar '<, Ausmaß:
 Druckstellen AS string 'Druckstellen varchar ',:
 Verformungen AS string 'Verformungen varchar ',:
 Verformungen_Beschreibung AS string 'Verformungen_Beschreibung varchar '<Beschreibung:
 Fußpflege AS string 'Fußpflege varchar '^:
 Podologie AS string 'Podologie varchar ',:
 Einlagen AS string 'Einlagen varchar ', diabetesgerechte orthopädische Einlagen/Schuhe:
 Neue_Fußkomplikationen AS string 'Neue_Fußkomplikationen varchar '^Neue Fußkomplikationen in den letzten 12 Monaten:
 Entleerungsstörungen_Magen AS string 'Entleerungsstörungen_Magen varchar '^:
 Entleerungsstörungen_Harnblase AS string 'Entleerungsstörungen_Harnblase varchar ',:
 Schwindel_Aufstehen AS string 'Schwindel_Aufstehen varchar ',:
 Folgeerkrankungen_Haut AS string 'Folgeerkrankungen_Haut varchar '^:
 Bewegungseinschränkungen AS string 'Bewegungseinschränkungen varchar ',:
 Sexualstörung AS string 'Sexualstörung varchar '^:
 Sexualstörung_seit AS string 'Sexualstörung_seit varchar '<seit
 Weitere_Anamnese AS string 'Weitere_Anamnese varchar '^:
 Tabak AS string 'Tabak varchar '^Tabak:
 tabakex AS string 'tabakex varchar '<, früher:
 tabakbis AS string 'tabakbis varchar '<, bis:
 tabakakt AS string 'tabakakt varchar '<, aktuell:
 tabakmenge AS string 'tabakmenge varchar '<, Menge:
 Alkohol AS string 'Alkohol varchar '^Alkohol pro Woche:
 Mitarbeiter AS string 'Mitarbeiter varchar '<, Mitarbeiter:
 Weitere_Medikation AS string 'Weitere_Medikation varchar '^:
 Liphypertrophien_Abdomen AS string 'Liphypertrophien_Abdomen varchar '^Liphypertrophien:Abdomen
 Liphypertrophien_Beine AS string 'Liphypertrophien_Beine varchar '<, Beine:
 Liphypertrophien_Arme AS string 'Liphypertrophien_Arme varchar '<, Arme:
 Beinbefund AS string 'Beinbefund varchar '^:
 Hyperkeratosen AS string 'Hyperkeratosen varchar ',:
 Ulcera AS string 'Ulcera varchar ',:
 Kraft_Zehenheber AS string 'Kraft_Zehenheber varchar '^Kraft:Zehenheber
 Kraft_Zehenbeuger AS string 'Kraft_Zehenbeuger varchar '<, Zehenbeuger:
 Kraft_Knie AS string 'Kraft_Knie varchar '<, Knie:
 ASR AS string 'ASR varchar ',:
 PSR AS string 'PSR varchar ',:
 Oberflächensensibilität AS string 'Oberflächensensibilität varchar '^:
 Monofilamenttest AS string 'Monofilamenttest varchar ',:
 Kalt_Warm AS string 'Kalt_Warm varchar ', Kalt-Warm-Diskrimination:
 Vibration_IK AS string 'Vibration_IK varchar ', Vibrationsempfinden Innenknöchel:
 Vibration_Großzehe AS string 'Vibration_Großzehe varchar '<, Großzehe:
 Puls_Leiste AS string 'Puls_Leiste varchar '^Pulse:Leiste
 Puls_Kniekehle AS string 'Puls_Kniekehle varchar '<,Kniekehle:
 Puls_Atp AS string 'Puls_Atp varchar '<,Innenknöchel:
 Puls_Adp AS string 'Puls_Adp varchar '<,Fußrücken:
 RR AS string 'RR varchar '^Blutdruck:
 RRTurboMed AS string 'RRTurboMed varchar '
 Herz AS string 'Herz varchar '^:
 Lunge AS string 'Lunge varchar ',:
 Bauch AS string 'Bauch varchar ', Abdomen:
 WS AS string 'WS varchar ', Wirbelsäule:
 NL AS string 'NL varchar ', Nierenlager:
 SD AS string 'SD varchar ', Schilddrüse:
 Carotiden AS string 'Carotiden varchar ', Halsschlagadern:
 NNH AS string 'NNH varchar ', Nasennebenhöhlen:
 Zähne AS string 'Zähne varchar ',:
 Parodontitis AS string 'Parodontitis varchar ',:
 Mundhöhle AS string 'Mundhöhle varchar ',:
 LK AS string 'LK varchar ', Lymphknoten:
 BeinödVen AS string 'BeinödVen varchar ', Beinödeme/ Venenkrankheiten:
 Neuro_sonst AS string 'Neuro_sonst varchar '^Sonstige neurologische Befunde:
 Weitere_Befunde AS string 'Weitere_Befunde varchar ', weitere Befunde:
 Schulung AS string 'Schulung varchar 'ob Schulungsbedarf
 DMP AS string 'DMP varchar 'ob Pat. bei HA im DMP
 DMSchulz AS integer 'DMSchulz smallint 'Zahl der DMP-Schulungen hier
 DMSchL AS integer 'DMSchL smallint 'Zahl der abgerechneten DMP-Schulungen hier
 RRSchulz AS integer 'RRSchulz smallint 'Zahl der Hypertonie-Schulungen hier
 DMPhier AS date 'DMPhier datetime 'ob Pat hier im DMP
 HANr AS string 'HANr varchar 'mit ""/""
 HANr2 AS string 'HANr2 varchar 'mit ""/""
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
 Hausarzt AS string 'Hausarzt varchar 'Hausarzt laut Anamnesebogen
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
Public roLs() AS laborxsaetze
Public roLg() AS laborxeingel
Public roLu() AS laborxus
Public roLo() AS laborxbakt
Public roLw() AS laborxwert
Public roLL() AS laborxleist
Public roLi() AS liuez
Public roAna() AS anamnesebogen

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
 if wdh = 0 then ReDim rFi(0)
 if wdh = 0 then ReDim rFo(0)
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
 ReDim rLi(0)
 wdh = -1
End Function ' Tinit

Public Function doEntleer(frm AS lese, Tbl$)
 Dim rs AS New ADODB.Recordset
 Set rs = DBCn.Execute("SELECT COUNT(0) AS ct FROM `" & Tbl & "`")
 frm.Ausgeb "Lösche: `" & Tbl & "` (" & rs!ct & " Datensätze)", True
 sql = sqlDeletefrom & "`" & Tbl & "`"
 Call DBCn.Execute(sql) ' ,,adasyncexecute
 DoEvents
End Function ' doEntleer

Public Function AllesLösch(frm AS lese)
 Dim ct&, rs AS new ADODB.recordset
 ON Error GoTo fehler
 call ForeignNo0
 call ForeignNo1
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
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in allesLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end SELECT
End Function ' AllesLösch

Public Function LabLösch(frm AS lese)
 Dim ct&, rs AS new ADODB.recordset
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
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabLösch/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 end SELECT
End Function ' LabLösch

Function doBezFeh(csqlVal$, obSpei%, ErrDes$)
 Call ForeignNo0
 Call ForeignNo1
 obfor = True
 If obSpei <> 0 Then
  Open BezFeh For Append AS #299
  Print #299, vbCrLf & vbCrLf & Now() & ": " & csqlVal
  Print #299, vbCrLf & ErrDes
  Close #299
 End If
End Function 'doBezFeh

Public Function roNaZuw(i&, j&)
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
 roNa(i).f3800 = rNa(j).f3800
 roNa(i).dmpklass = rNa(j).dmpklass
 roNa(i).dmpbeg = rNa(j).dmpbeg
 roNa(i).dmpkhkklass = rNa(j).dmpkhkklass
 roNa(i).dmpkhkbeg = rNa(j).dmpkhkbeg
 roNa(i).dmpcopdklass = rNa(j).dmpcopdklass
 roNa(i).dmpcopdbeg = rNa(j).dmpcopdbeg
 roNa(i).dmpabklass = rNa(j).dmpabklass
 roNa(i).dmpabbeg = rNa(j).dmpabbeg
 roNa(i).dakab = rNa(j).dakab
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
 if roNa(i).Pat_ID <> rNa(j).Pat_ID Then gosub unter
 if roNa(i).lfdnr <> rNa(j).lfdnr Then gosub unter
 if roNa(i).NVorsatz <> rNa(j).NVorsatz Then gosub unter
 if roNa(i).Nachname <> rNa(j).Nachname Then gosub unter
 if roNa(i).Vorname <> rNa(j).Vorname Then gosub unter
 if roNa(i).GebDat <> rNa(j).GebDat Then gosub unter
 if roNa(i).f3004 <> rNa(j).f3004 Then gosub unter
 if roNa(i).f3006 <> rNa(j).f3006 Then gosub unter
 if roNa(i).Straße <> rNa(j).Straße Then gosub unter
 if roNa(i).KVKStatus <> rNa(j).KVKStatus Then gosub unter
 if roNa(i).Hausnr <> rNa(j).Hausnr Then gosub unter
 if roNa(i).Geschlecht <> rNa(j).Geschlecht Then gosub unter
 if roNa(i).Plz <> rNa(j).Plz Then gosub unter
 if roNa(i).Ort <> rNa(j).Ort Then gosub unter
 if roNa(i).Lkz <> rNa(j).Lkz Then gosub unter
 if roNa(i).Anschrzus <> rNa(j).Anschrzus Then gosub unter
 if roNa(i).NVors <> rNa(j).NVors Then gosub unter
 if roNa(i).PFPlz <> rNa(j).PFPlz Then gosub unter
 if roNa(i).PFOrt <> rNa(j).PFOrt Then gosub unter
 if roNa(i).PFNr <> rNa(j).PFNr Then gosub unter
 if roNa(i).f3124 <> rNa(j).f3124 Then gosub unter
 if roNa(i).AnschrZus_2 <> rNa(j).AnschrZus_2 Then gosub unter
 if roNa(i).Postfach_2 <> rNa(j).Postfach_2 Then gosub unter
 if roNa(i).LK_2 <> rNa(j).LK_2 Then gosub unter
 if roNa(i).Postfach <> rNa(j).Postfach Then gosub unter
 if roNa(i).Beruf <> rNa(j).Beruf Then gosub unter
 if roNa(i).Weggeldzone <> rNa(j).Weggeldzone Then gosub unter
 if roNa(i).WeggzZahl <> rNa(j).WeggzZahl Then gosub unter
 if roNa(i).AufnDat <> rNa(j).AufnDat Then gosub unter
 if roNa(i).kAufDat <> rNa(j).kAufDat Then gosub unter
 if roNa(i).LANR <> rNa(j).LANR Then gosub unter
 if roNa(i).BStNr <> rNa(j).BStNr Then gosub unter
 if roNa(i).Titel <> rNa(j).Titel Then gosub unter
 if roNa(i).Versichertennummer <> rNa(j).Versichertennummer Then gosub unter
 if roNa(i).PrivatTel <> rNa(j).PrivatTel Then gosub unter
 if roNa(i).KVNr <> rNa(j).KVNr Then gosub unter
 if roNa(i).KVNr2 <> rNa(j).KVNr2 Then gosub unter
 if roNa(i).KVNr3 <> rNa(j).KVNr3 Then gosub unter
 if roNa(i).KVNr4 <> rNa(j).KVNr4 Then gosub unter
 if roNa(i).PrivatTel_2 <> rNa(j).PrivatTel_2 Then gosub unter
 if roNa(i).PrivatFax <> rNa(j).PrivatFax Then gosub unter
 if roNa(i).DienstTel <> rNa(j).DienstTel Then gosub unter
 if roNa(i).PrivatMobil <> rNa(j).PrivatMobil Then gosub unter
 if roNa(i).Email <> rNa(j).Email Then gosub unter
 if roNa(i).Arbeitgeber <> rNa(j).Arbeitgeber Then gosub unter
 if roNa(i).AnAllgda <> rNa(j).AnAllgda Then gosub unter
 if roNa(i).An1da <> rNa(j).An1da Then gosub unter
 if roNa(i).An2da <> rNa(j).An2da Then gosub unter
 if roNa(i).Checkda <> rNa(j).Checkda Then gosub unter
 if roNa(i).DMTypaD <> rNa(j).DMTypaD Then gosub unter
 if roNa(i).AktZeit <> rNa(j).AktZeit Then gosub unter
 if roNa(i).absPos <> rNa(j).absPos Then gosub unter
 if roNa(i).StByte <> rNa(j).StByte Then gosub unter
 if roNa(i).StByteA <> rNa(j).StByteA Then gosub unter
 if roNa(i).Cave <> rNa(j).Cave Then gosub unter
 if roNa(i).Notiz <> rNa(j).Notiz Then gosub unter
 if roNa(i).f3800 <> rNa(j).f3800 Then gosub unter
 if roNa(i).dmpklass <> rNa(j).dmpklass Then gosub unter
 if roNa(i).dmpbeg <> rNa(j).dmpbeg Then gosub unter
 if roNa(i).dmpkhkklass <> rNa(j).dmpkhkklass Then gosub unter
 if roNa(i).dmpkhkbeg <> rNa(j).dmpkhkbeg Then gosub unter
 if roNa(i).dmpcopdklass <> rNa(j).dmpcopdklass Then gosub unter
 if roNa(i).dmpcopdbeg <> rNa(j).dmpcopdbeg Then gosub unter
 if roNa(i).dmpabklass <> rNa(j).dmpabklass Then gosub unter
 if roNa(i).dmpabbeg <> rNa(j).dmpabbeg Then gosub unter
 if roNa(i).dakab <> rNa(j).dakab Then gosub unter
 if roNa(i).getHA0 <> rNa(j).getHA0 Then gosub unter
 if roNa(i).fnHA0 <> rNa(j).fnHA0 Then gosub unter
 if roNa(i).getHA1 <> rNa(j).getHA1 Then gosub unter
 if roNa(i).fnHA1 <> rNa(j).fnHA1 Then gosub unter
 if roNa(i).getHA2 <> rNa(j).getHA2 Then gosub unter
 if roNa(i).fnHA2 <> rNa(j).fnHA2 Then gosub unter
 if roNa(i).zubenach <> rNa(j).zubenach Then gosub unter
 if roNa(i).Verwandt <> rNa(j).Verwandt Then gosub unter
 if roNa(i).Sprache <> rNa(j).Sprache Then gosub unter
 if roNa(i).lAktTM <> rNa(j).lAktTM Then gosub unter
 if roNa(i).Mitarbeiter <> rNa(j).Mitarbeiter Then gosub unter
 Exit Function
unter:
 NaZUnt = NaZUnt + 1
 Return
End Function ' NaZUnt

Public Function namenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roNa(1)
 sql = "SELECT Pat_ID,lfdnr,NVorsatz,Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz,Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2,LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel,Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte,StByteA,Cave,Notiz,f3800,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg,dmpabklass,dmpabbeg,dakab,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt,Sprache,lAktTM,Mitarbeiter FROM `namen` WHERE Pat_ID=" & pid & " ORDER BY `kAufDat`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roNa)
  roNa(akt).Pat_ID = rs!Pat_ID
  roNa(akt).lfdnr = rs!lfdnr
  roNa(akt).NVorsatz = rs!NVorsatz
  roNa(akt).Nachname = rs!Nachname
  roNa(akt).Vorname = rs!Vorname
  roNa(akt).GebDat = rs!GebDat
  roNa(akt).f3004 = rs!f3004
  roNa(akt).f3006 = rs!f3006
  roNa(akt).Straße = rs!Straße
  roNa(akt).KVKStatus = rs!KVKStatus
  roNa(akt).Hausnr = rs!Hausnr
  roNa(akt).Geschlecht = rs!Geschlecht
  roNa(akt).Plz = rs!Plz
  roNa(akt).Ort = rs!Ort
  roNa(akt).Lkz = rs!Lkz
  roNa(akt).Anschrzus = rs!Anschrzus
  roNa(akt).NVors = rs!NVors
  roNa(akt).PFPlz = rs!PFPlz
  roNa(akt).PFOrt = rs!PFOrt
  roNa(akt).PFNr = rs!PFNr
  roNa(akt).f3124 = rs!f3124
  roNa(akt).AnschrZus_2 = rs!AnschrZus_2
  roNa(akt).Postfach_2 = rs!Postfach_2
  roNa(akt).LK_2 = rs!LK_2
  roNa(akt).Postfach = rs!Postfach
  roNa(akt).Beruf = rs!Beruf
  roNa(akt).Weggeldzone = rs!Weggeldzone
  roNa(akt).WeggzZahl = rs!WeggzZahl
  roNa(akt).AufnDat = rs!AufnDat
  roNa(akt).kAufDat = rs!kAufDat
  roNa(akt).LANR = rs!LANR
  roNa(akt).BStNr = rs!BStNr
  roNa(akt).Titel = rs!Titel
  roNa(akt).Versichertennummer = rs!Versichertennummer
  roNa(akt).PrivatTel = rs!PrivatTel
  roNa(akt).KVNr = rs!KVNr
  roNa(akt).KVNr2 = rs!KVNr2
  roNa(akt).KVNr3 = rs!KVNr3
  roNa(akt).KVNr4 = rs!KVNr4
  roNa(akt).PrivatTel_2 = rs!PrivatTel_2
  roNa(akt).PrivatFax = rs!PrivatFax
  roNa(akt).DienstTel = rs!DienstTel
  roNa(akt).PrivatMobil = rs!PrivatMobil
  roNa(akt).Email = rs!Email
  roNa(akt).Arbeitgeber = rs!Arbeitgeber
  roNa(akt).AnAllgda = rs!AnAllgda
  roNa(akt).An1da = rs!An1da
  roNa(akt).An2da = rs!An2da
  roNa(akt).Checkda = rs!Checkda
  roNa(akt).DMTypaD = rs!DMTypaD
  roNa(akt).AktZeit = rs!AktZeit
  roNa(akt).absPos = rs!absPos
  roNa(akt).StByte = rs!StByte
  roNa(akt).StByteA = rs!StByteA
  roNa(akt).Cave = rs!Cave
  roNa(akt).Notiz = rs!Notiz
  roNa(akt).f3800 = rs!f3800
  roNa(akt).dmpklass = rs!dmpklass
  roNa(akt).dmpbeg = rs!dmpbeg
  roNa(akt).dmpkhkklass = rs!dmpkhkklass
  roNa(akt).dmpkhkbeg = rs!dmpkhkbeg
  roNa(akt).dmpcopdklass = rs!dmpcopdklass
  roNa(akt).dmpcopdbeg = rs!dmpcopdbeg
  roNa(akt).dmpabklass = rs!dmpabklass
  roNa(akt).dmpabbeg = rs!dmpabbeg
  roNa(akt).dakab = rs!dakab
  roNa(akt).getHA0 = rs!getHA0
  roNa(akt).fnHA0 = rs!fnHA0
  roNa(akt).getHA1 = rs!getHA1
  roNa(akt).fnHA1 = rs!fnHA1
  roNa(akt).getHA2 = rs!getHA2
  roNa(akt).fnHA2 = rs!fnHA2
  roNa(akt).zubenach = rs!zubenach
  roNa(akt).Verwandt = rs!Verwandt
  roNa(akt).Sprache = rs!Sprache
  roNa(akt).lAktTM = rs!lAktTM
  roNa(akt).Mitarbeiter = rs!Mitarbeiter
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roNa(UBound(roNa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' namenLaden

Function namenEinf
 Dim ibeg&, i&, j&
 If UBound(rNa) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rNa)
   If rNa(i).kAufDat >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roNa)
    If roNa(i).kAufDat >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roNa(i - 1)
   ReDim Preserve roNa(i + UBound(rNa) - ibeg)
   For j = ibeg To UBound(rNa)
    Call roNaZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rNa = roNa
End Function ' namenEinf

Public Function namenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere namen"
 if not Allepat then
   sql = "DELETE FROM `namen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,f3800,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach," & _
     "Verwandt,Sprache,lAktTM,Mitarbeiter) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `namen` (Pat_ID,lfdnr,NVorsatz," & _
     "Nachname,Vorname,GebDat,f3004,f3006,Straße,KVKStatus,Hausnr,Geschlecht,Plz," & _
     "Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,f3124,AnschrZus_2,Postfach_2," & _
     "LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr,Titel," & _
     "Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel,PrivatMobil," & _
     "Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos,StByte," & _
     "StByteA,Cave,Notiz,f3800,dmpklass,dmpbeg,dmpkhkklass,dmpkhkbeg,dmpcopdklass,dmpcopdbeg," & _
     "dmpabklass,dmpabbeg,dakab,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach," & _
     "Verwandt,Sprache,lAktTM,Mitarbeiter)               values"))
 For i = 0 to ubound(rNa)
'  rNa(i).AktZeit = now()
  rNa(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rNa(i).Pat_ID, "," , rNa(i).lfdnr, ",'" , rNa(i).NVorsatz, "','" , rNa(i).Nachname, "','" , rNa(i).Vorname, "'," , DatFor_k(rNa(i).GebDat), ",'" , rNa(i).f3004, "','" , rNa(i).f3006, "','" ,  _
   rNa(i).Straße, "','" , rNa(i).KVKStatus, "','" , rNa(i).Hausnr, "','" , rNa(i).Geschlecht, "','" , rNa(i).Plz, "','" , rNa(i).Ort, "','" , rNa(i).Lkz, "','" , rNa(i).Anschrzus, "','" ,  _
   rNa(i).NVors, "','" , rNa(i).PFPlz, "','" , rNa(i).PFOrt, "','" , rNa(i).PFNr, "','" , rNa(i).f3124, "','" , rNa(i).AnschrZus_2, "','" , rNa(i).Postfach_2, "','" , rNa(i).LK_2, "','" , rNa(i).Postfach, "','" ,  _
   rNa(i).Beruf, "','" , rNa(i).Weggeldzone, "'," , rNa(i).WeggzZahl, "," , DatFor_k(rNa(i).AufnDat), "," , DatFor_k(rNa(i).kAufDat), ",'" , rNa(i).LANR, "','" , rNa(i).BStNr, "','" ,  _
   rNa(i).Titel, "','" , rNa(i).Versichertennummer, "','" , rNa(i).PrivatTel, "','" , rNa(i).KVNr, "','" , rNa(i).KVNr2, "','" , rNa(i).KVNr3, "','" , rNa(i).KVNr4, "','" , rNa(i).PrivatTel_2, "','" ,  _
   rNa(i).PrivatFax, "','" , rNa(i).DienstTel, "','" , rNa(i).PrivatMobil, "','" , rNa(i).Email, "','" , rNa(i).Arbeitgeber, "'," , cstr(-(rNa(i).AnAllgda<>0)) , "," , cstr(-(rNa(i).An1da<>0)) , "," , cstr(-( _
   rNa(i).An2da<>0)) , "," , cstr(-(rNa(i).Checkda<>0)) , ",'" , rNa(i).DMTypaD, "'," , DatFor_k( 0 ), "," , rNa(i).absPos, "," , rNa(i).StByte, "," , rNa(i).StByteA, ",'" ,  _
   rNa(i).Cave, "','" , rNa(i).Notiz, "','" , rNa(i).f3800, "'," , rNa(i).dmpklass, "," , DatFor_k(rNa(i).dmpbeg), "," , rNa(i).dmpkhkklass, "," , DatFor_k(rNa(i).dmpkhkbeg), "," , rNa(i).dmpcopdklass, "," , DatFor_k( _
   rNa(i).dmpcopdbeg), "," , rNa(i).dmpabklass, "," , DatFor_k(rNa(i).dmpabbeg), "," , DatFor_k(rNa(i).dakab), "," , rNa(i).getHA0, ",'" , rNa(i).fnHA0, "'," , rNa(i).getHA1, ",'" ,  _
   rNa(i).fnHA1, "'," , rNa(i).getHA2, ",'" , rNa(i).fnHA2, "','" , rNa(i).zubenach, "','" , rNa(i).Verwandt, "','" , rNa(i).Sprache, "'," , DatFor_k(rNa(i).lAktTM), "," , rNa(i).Mitarbeiter, ")")
  If SammelInsert <> 0 AND i < ubound(rNa) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rNa) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(48)
 for k = iif(SammelInsert<>0,0,i) to iif(SammelInsert<>0,ubound(rNa),i)
  If Len(rNa(k).NVorsatz) > maxi(0) Then maxi(0) = Len(rNa(k).NVorsatz)
  If Len(rNa(k).Nachname) > maxi(1) Then maxi(1) = Len(rNa(k).Nachname)
  If Len(rNa(k).Vorname) > maxi(2) Then maxi(2) = Len(rNa(k).Vorname)
  If Len(rNa(k).f3004) > maxi(3) Then maxi(3) = Len(rNa(k).f3004)
  If Len(rNa(k).f3006) > maxi(4) Then maxi(4) = Len(rNa(k).f3006)
  If Len(rNa(k).Straße) > maxi(5) Then maxi(5) = Len(rNa(k).Straße)
  If Len(rNa(k).KVKStatus) > maxi(6) Then maxi(6) = Len(rNa(k).KVKStatus)
  If Len(rNa(k).Hausnr) > maxi(7) Then maxi(7) = Len(rNa(k).Hausnr)
  If Len(rNa(k).Geschlecht) > maxi(8) Then maxi(8) = Len(rNa(k).Geschlecht)
  If Len(rNa(k).Plz) > maxi(9) Then maxi(9) = Len(rNa(k).Plz)
  If Len(rNa(k).Ort) > maxi(10) Then maxi(10) = Len(rNa(k).Ort)
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
  If Len(rNa(k).LANR) > maxi(24) Then maxi(24) = Len(rNa(k).LANR)
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
  If Len(rNa(k).Email) > maxi(37) Then maxi(37) = Len(rNa(k).Email)
  If Len(rNa(k).Arbeitgeber) > maxi(38) Then maxi(38) = Len(rNa(k).Arbeitgeber)
  If Len(rNa(k).DMTypaD) > maxi(39) Then maxi(39) = Len(rNa(k).DMTypaD)
  If Len(rNa(k).Cave) > maxi(40) Then maxi(40) = Len(rNa(k).Cave)
  If Len(rNa(k).Notiz) > maxi(41) Then maxi(41) = Len(rNa(k).Notiz)
  If Len(rNa(k).f3800) > maxi(42) Then maxi(42) = Len(rNa(k).f3800)
  If Len(rNa(k).fnHA0) > maxi(43) Then maxi(43) = Len(rNa(k).fnHA0)
  If Len(rNa(k).fnHA1) > maxi(44) Then maxi(44) = Len(rNa(k).fnHA1)
  If Len(rNa(k).fnHA2) > maxi(45) Then maxi(45) = Len(rNa(k).fnHA2)
  If Len(rNa(k).zubenach) > maxi(46) Then maxi(46) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(47) Then maxi(47) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(48) Then maxi(48) = Len(rNa(k).Sprache)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "namen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "namen", rsc)
    If maxL > 0 Then
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
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rNa.f3800: '" & rNa(k).f3800 & "' -> '" & Left$(rNa(k).f3800, maxL)  & "'",true : rNa(k).f3800 = Left$(rNa(k).f3800, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & Left$(rNa(k).fnHA0, maxL)  & "'",true : rNa(k).fnHA0 = Left$(rNa(k).fnHA0, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & Left$(rNa(k).fnHA1, maxL)  & "'",true : rNa(k).fnHA1 = Left$(rNa(k).fnHA1, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & Left$(rNa(k).fnHA2, maxL)  & "'",true : rNa(k).fnHA2 = Left$(rNa(k).fnHA2, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & Left$(rNa(k).zubenach, maxL)  & "'",true : rNa(k).zubenach = Left$(rNa(k).zubenach, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & Left$(rNa(k).Verwandt, maxL)  & "'",true : rNa(k).Verwandt = Left$(rNa(k).Verwandt, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & Left$(rNa(k).Sprache, maxL)  & "'",true : rNa(k).Sprache = Left$(rNa(k).Sprache, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
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
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' namenSpeichern

Public Function roFaZuw(i&, j&)
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
End Function ' roFaZuw

Public Function FaZUnt%(i&, j&)
 if roFa(i).FID <> rFa(j).FID Then gosub unter
 if roFa(i).Pat_ID <> rFa(j).Pat_ID Then gosub unter
 if roFa(i).Quartal <> rFa(j).Quartal Then gosub unter
 if roFa(i).Nachname <> rFa(j).Nachname Then gosub unter
 if roFa(i).Vorname <> rFa(j).Vorname Then gosub unter
 if roFa(i).lfdnr <> rFa(j).lfdnr Then gosub unter
 if roFa(i).TMFNr <> rFa(j).TMFNr Then gosub unter
 if roFa(i).VKNr <> rFa(j).VKNr Then gosub unter
 if roFa(i).f4131 <> rFa(j).f4131 Then gosub unter
 if roFa(i).f4132 <> rFa(j).f4132 Then gosub unter
 if roFa(i).VschBeg <> rFa(j).VschBeg Then gosub unter
 if roFa(i).KKasse_2 <> rFa(j).KKasse_2 Then gosub unter
 if roFa(i).FaktPers <> rFa(j).FaktPers Then gosub unter
 if roFa(i).FaktTechn <> rFa(j).FaktTechn Then gosub unter
 if roFa(i).FaktLabor <> rFa(j).FaktLabor Then gosub unter
 if roFa(i).BhFB <> rFa(j).BhFB Then gosub unter
 if roFa(i).BhFE1 <> rFa(j).BhFE1 Then gosub unter
 if roFa(i).BhFE2 <> rFa(j).BhFE2 Then gosub unter
 if roFa(i).f4202 <> rFa(j).f4202 Then gosub unter
 if roFa(i).ausgst <> rFa(j).ausgst Then gosub unter
 if roFa(i).KtrAbrB <> rFa(j).KtrAbrB Then gosub unter
 if roFa(i).AbrAr <> rFa(j).AbrAr Then gosub unter
 if roFa(i).lVorl <> rFa(j).lVorl Then gosub unter
 if roFa(i).IK <> rFa(j).IK Then gosub unter
 if roFa(i).KVKs <> rFa(j).KVKs Then gosub unter
 if roFa(i).KVKserg <> rFa(j).KVKserg Then gosub unter
 if roFa(i).Status <> rFa(j).Status Then gosub unter
 if roFa(i).Kasse <> rFa(j).Kasse Then gosub unter
 if roFa(i).KID <> rFa(j).KID Then gosub unter
 if roFa(i).GebOr <> rFa(j).GebOr Then gosub unter
 if roFa(i).AbrGb <> rFa(j).AbrGb Then gosub unter
 if roFa(i).PersKreis <> rFa(j).PersKreis Then gosub unter
 if roFa(i).SKtZusatz <> rFa(j).SKtZusatz Then gosub unter
 if roFa(i).letzteRegel <> rFa(j).letzteRegel Then gosub unter
 if roFa(i).ÜwText <> rFa(j).ÜwText Then gosub unter
 if roFa(i).f4210 <> rFa(j).f4210 Then gosub unter
 if roFa(i).AkfHAH <> rFa(j).AkfHAH Then gosub unter
 if roFa(i).AkfAB0 <> rFa(j).AkfAB0 Then gosub unter
 if roFa(i).AkfAK <> rFa(j).AkfAK Then gosub unter
 if roFa(i).statNuller <> rFa(j).statNuller Then gosub unter
 if roFa(i).ÜbwV <> rFa(j).ÜbwV Then gosub unter
 if roFa(i).ÜbWVLANR <> rFa(j).ÜbWVLANR Then gosub unter
 if roFa(i).ÜbWVBSNR <> rFa(j).ÜbWVBSNR Then gosub unter
 if roFa(i).ÜbWVKVNR <> rFa(j).ÜbWVKVNR Then gosub unter
 if roFa(i).AndÜw <> rFa(j).AndÜw Then gosub unter
 if roFa(i).Übwr <> rFa(j).Übwr Then gosub unter
 if roFa(i).ÜbwLANR <> rFa(j).ÜbwLANR Then gosub unter
 if roFa(i).ÜWZiel <> rFa(j).ÜWZiel Then gosub unter
 if roFa(i).ÜWNNr <> rFa(j).ÜWNNr Then gosub unter
 if roFa(i).ÜWNaN <> rFa(j).ÜWNaN Then gosub unter
 if roFa(i).ÜWTit <> rFa(j).ÜWTit Then gosub unter
 if roFa(i).ÜWVor <> rFa(j).ÜWVor Then gosub unter
 if roFa(i).ÜWVsw <> rFa(j).ÜWVsw Then gosub unter
 if roFa(i).üwvid <> rFa(j).üwvid Then gosub unter
 if roFa(i).Auftrag <> rFa(j).Auftrag Then gosub unter
 if roFa(i).Verdacht <> rFa(j).Verdacht Then gosub unter
 if roFa(i).Befund <> rFa(j).Befund Then gosub unter
 if roFa(i).statKlasse <> rFa(j).statKlasse Then gosub unter
 if roFa(i).f4237 <> rFa(j).f4237 Then gosub unter
 if roFa(i).statBehTage <> rFa(j).statBehTage Then gosub unter
 if roFa(i).SchGr <> rFa(j).SchGr Then gosub unter
 if roFa(i).Weiterbeh <> rFa(j).Weiterbeh Then gosub unter
 if roFa(i).f4266 <> rFa(j).f4266 Then gosub unter
 if roFa(i).PGeb <> rFa(j).PGeb Then gosub unter
 if roFa(i).PGebErg <> rFa(j).PGebErg Then gosub unter
 if roFa(i).Mahnfrist <> rFa(j).Mahnfrist Then gosub unter
 if roFa(i).Unfallort <> rFa(j).Unfallort Then gosub unter
 if roFa(i).BeschAls <> rFa(j).BeschAls Then gosub unter
 if roFa(i).BeschSeit <> rFa(j).BeschSeit Then gosub unter
 if roFa(i).Unfallbetrieb <> rFa(j).Unfallbetrieb Then gosub unter
 if roFa(i).f4570 <> rFa(j).f4570 Then gosub unter
 if roFa(i).GOÄKatNr <> rFa(j).GOÄKatNr Then gosub unter
 if roFa(i).GOÄKatName <> rFa(j).GOÄKatName Then gosub unter
 if roFa(i).abrArzt <> rFa(j).abrArzt Then gosub unter
 if roFa(i).privVers <> rFa(j).privVers Then gosub unter
 if roFa(i).AdNam <> rFa(j).AdNam Then gosub unter
 if roFa(i).AdStr <> rFa(j).AdStr Then gosub unter
 if roFa(i).AdPlz <> rFa(j).AdPlz Then gosub unter
 if roFa(i).AdOrt <> rFa(j).AdOrt Then gosub unter
 if roFa(i).ÜwBG <> rFa(j).ÜwBG Then gosub unter
 if roFa(i).BhFE <> rFa(j).BhFE Then gosub unter
 if roFa(i).s8000 <> rFa(j).s8000 Then gosub unter
 if roFa(i).s8100 <> rFa(j).s8100 Then gosub unter
 if roFa(i).AktZeit <> rFa(j).AktZeit Then gosub unter
 if roFa(i).Fanf <> rFa(j).Fanf Then gosub unter
 if roFa(i).altQuart <> rFa(j).altQuart Then gosub unter
 if roFa(i).QAnf <> rFa(j).QAnf Then gosub unter
 if roFa(i).QEnd <> rFa(j).QEnd Then gosub unter
 if roFa(i).QS <> rFa(j).QS Then gosub unter
 if roFa(i).QT <> rFa(j).QT Then gosub unter
 if roFa(i).StByte <> rFa(j).StByte Then gosub unter
 if roFa(i).absPos <> rFa(j).absPos Then gosub unter
 if roFa(i).LANRid <> rFa(j).LANRid Then gosub unter
 if roFa(i).f4108 <> rFa(j).f4108 Then gosub unter
 if roFa(i).BGFallNr <> rFa(j).BGFallNr Then gosub unter
 if roFa(i).lGewicht <> rFa(j).lGewicht Then gosub unter
 if roFa(i).vorET <> rFa(j).vorET Then gosub unter
 if roFa(i).dmpVertret <> rFa(j).dmpVertret Then gosub unter
 if roFa(i).dmpArztw <> rFa(j).dmpArztw Then gosub unter
 if roFa(i).dmpHypos <> rFa(j).dmpHypos Then gosub unter
 if roFa(i).dmpKhsA <> rFa(j).dmpKhsA Then gosub unter
 if roFa(i).dmpDMSchulEmpf <> rFa(j).dmpDMSchulEmpf Then gosub unter
 if roFa(i).dmpDMSchulWahrg <> rFa(j).dmpDMSchulWahrg Then gosub unter
 if roFa(i).dmpHypertSchulEmpf <> rFa(j).dmpHypertSchulEmpf Then gosub unter
 if roFa(i).dmpHypertSchulWahrg <> rFa(j).dmpHypertSchulWahrg Then gosub unter
 if roFa(i).dmpKKTabakEmpf <> rFa(j).dmpKKTabakEmpf Then gosub unter
 if roFa(i).dmpKKErnEmpf <> rFa(j).dmpKKErnEmpf Then gosub unter
 if roFa(i).dmpKKkTrainEmpf <> rFa(j).dmpKKkTrainEmpf Then gosub unter
 if roFa(i).dmpHbA1cZiel <> rFa(j).dmpHbA1cZiel Then gosub unter
 if roFa(i).dmpUewFuss <> rFa(j).dmpUewFuss Then gosub unter
 if roFa(i).dmpEinwDM <> rFa(j).dmpEinwDM Then gosub unter
 if roFa(i).dmphalbj <> rFa(j).dmphalbj Then gosub unter
 if roFa(i).dmpMA <> rFa(j).dmpMA Then gosub unter
 Exit Function
unter:
 FaZUnt = FaZUnt + 1
 Return
End Function ' FaZUnt

Public Function faelleLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFa(1)
 sql = "SELECT FID,Pat_ID,Quartal,Nachname,Vorname,lfdnr,TMFNr,VKNr,f4131,f4132,VschBeg,KKasse_2,FaktPers,FaktTechn,FaktLabor,BhFB,BhFE1,BhFE2,f4202,ausgst,KtrAbrB,AbrAr,lVorl,IK,KVKs,KVKserg,Status,Kasse,KID,GebOr,AbrGb,PersKreis,SKtZusatz,letzteRegel,ÜwText,f4210,AkfHAH,AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR,ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,Auftrag,Verdacht,Befund,statKlasse,f4237,statBehTage,SchGr,Weiterbeh,f4266,PGeb,PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,f4570,GOÄKatNr,GOÄKatName,abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100,AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid,f4108,BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg,dmpHypertSchulEmpf,dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA FROM `faelle` WHERE Pat_ID=" & pid & " ORDER BY `fanf`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roFa)
  roFa(akt).FID = rs!FID
  roFa(akt).Pat_ID = rs!Pat_ID
  roFa(akt).Quartal = rs!Quartal
  roFa(akt).Nachname = rs!Nachname
  roFa(akt).Vorname = rs!Vorname
  roFa(akt).lfdnr = rs!lfdnr
  roFa(akt).TMFNr = rs!TMFNr
  roFa(akt).VKNr = rs!VKNr
  roFa(akt).f4131 = rs!f4131
  roFa(akt).f4132 = rs!f4132
  roFa(akt).VschBeg = rs!VschBeg
  roFa(akt).KKasse_2 = rs!KKasse_2
  roFa(akt).FaktPers = rs!FaktPers
  roFa(akt).FaktTechn = rs!FaktTechn
  roFa(akt).FaktLabor = rs!FaktLabor
  roFa(akt).BhFB = rs!BhFB
  roFa(akt).BhFE1 = rs!BhFE1
  roFa(akt).BhFE2 = rs!BhFE2
  roFa(akt).f4202 = rs!f4202
  roFa(akt).ausgst = rs!ausgst
  roFa(akt).KtrAbrB = rs!KtrAbrB
  roFa(akt).AbrAr = rs!AbrAr
  roFa(akt).lVorl = rs!lVorl
  roFa(akt).IK = rs!IK
  roFa(akt).KVKs = rs!KVKs
  roFa(akt).KVKserg = rs!KVKserg
  roFa(akt).Status = rs!Status
  roFa(akt).Kasse = rs!Kasse
  roFa(akt).KID = rs!KID
  roFa(akt).GebOr = rs!GebOr
  roFa(akt).AbrGb = rs!AbrGb
  roFa(akt).PersKreis = rs!PersKreis
  roFa(akt).SKtZusatz = rs!SKtZusatz
  roFa(akt).letzteRegel = rs!letzteRegel
  roFa(akt).ÜwText = rs!ÜwText
  roFa(akt).f4210 = rs!f4210
  roFa(akt).AkfHAH = rs!AkfHAH
  roFa(akt).AkfAB0 = rs!AkfAB0
  roFa(akt).AkfAK = rs!AkfAK
  roFa(akt).statNuller = rs!statNuller
  roFa(akt).ÜbwV = rs!ÜbwV
  roFa(akt).ÜbWVLANR = rs!ÜbWVLANR
  roFa(akt).ÜbWVBSNR = rs!ÜbWVBSNR
  roFa(akt).ÜbWVKVNR = rs!ÜbWVKVNR
  roFa(akt).AndÜw = rs!AndÜw
  roFa(akt).Übwr = rs!Übwr
  roFa(akt).ÜbwLANR = rs!ÜbwLANR
  roFa(akt).ÜWZiel = rs!ÜWZiel
  roFa(akt).ÜWNNr = rs!ÜWNNr
  roFa(akt).ÜWNaN = rs!ÜWNaN
  roFa(akt).ÜWTit = rs!ÜWTit
  roFa(akt).ÜWVor = rs!ÜWVor
  roFa(akt).ÜWVsw = rs!ÜWVsw
  roFa(akt).üwvid = rs!üwvid
  roFa(akt).Auftrag = rs!Auftrag
  roFa(akt).Verdacht = rs!Verdacht
  roFa(akt).Befund = rs!Befund
  roFa(akt).statKlasse = rs!statKlasse
  roFa(akt).f4237 = rs!f4237
  roFa(akt).statBehTage = rs!statBehTage
  roFa(akt).SchGr = rs!SchGr
  roFa(akt).Weiterbeh = rs!Weiterbeh
  roFa(akt).f4266 = rs!f4266
  roFa(akt).PGeb = rs!PGeb
  roFa(akt).PGebErg = rs!PGebErg
  roFa(akt).Mahnfrist = rs!Mahnfrist
  roFa(akt).Unfallort = rs!Unfallort
  roFa(akt).BeschAls = rs!BeschAls
  roFa(akt).BeschSeit = rs!BeschSeit
  roFa(akt).Unfallbetrieb = rs!Unfallbetrieb
  roFa(akt).f4570 = rs!f4570
  roFa(akt).GOÄKatNr = rs!GOÄKatNr
  roFa(akt).GOÄKatName = rs!GOÄKatName
  roFa(akt).abrArzt = rs!abrArzt
  roFa(akt).privVers = rs!privVers
  roFa(akt).AdNam = rs!AdNam
  roFa(akt).AdStr = rs!AdStr
  roFa(akt).AdPlz = rs!AdPlz
  roFa(akt).AdOrt = rs!AdOrt
  roFa(akt).ÜwBG = rs!ÜwBG
  roFa(akt).BhFE = rs!BhFE
  roFa(akt).s8000 = rs!s8000
  roFa(akt).s8100 = rs!s8100
  roFa(akt).AktZeit = rs!AktZeit
  roFa(akt).Fanf = rs!Fanf
  roFa(akt).altQuart = rs!altQuart
  roFa(akt).QAnf = rs!QAnf
  roFa(akt).QEnd = rs!QEnd
  roFa(akt).QS = rs!QS
  roFa(akt).QT = rs!QT
  roFa(akt).StByte = rs!StByte
  roFa(akt).absPos = rs!absPos
  roFa(akt).LANRid = rs!LANRid
  roFa(akt).f4108 = rs!f4108
  roFa(akt).BGFallNr = rs!BGFallNr
  roFa(akt).lGewicht = rs!lGewicht
  roFa(akt).vorET = rs!vorET
  roFa(akt).dmpVertret = rs!dmpVertret
  roFa(akt).dmpArztw = rs!dmpArztw
  roFa(akt).dmpHypos = rs!dmpHypos
  roFa(akt).dmpKhsA = rs!dmpKhsA
  roFa(akt).dmpDMSchulEmpf = rs!dmpDMSchulEmpf
  roFa(akt).dmpDMSchulWahrg = rs!dmpDMSchulWahrg
  roFa(akt).dmpHypertSchulEmpf = rs!dmpHypertSchulEmpf
  roFa(akt).dmpHypertSchulWahrg = rs!dmpHypertSchulWahrg
  roFa(akt).dmpKKTabakEmpf = rs!dmpKKTabakEmpf
  roFa(akt).dmpKKErnEmpf = rs!dmpKKErnEmpf
  roFa(akt).dmpKKkTrainEmpf = rs!dmpKKkTrainEmpf
  roFa(akt).dmpHbA1cZiel = rs!dmpHbA1cZiel
  roFa(akt).dmpUewFuss = rs!dmpUewFuss
  roFa(akt).dmpEinwDM = rs!dmpEinwDM
  roFa(akt).dmphalbj = rs!dmphalbj
  roFa(akt).dmpMA = rs!dmpMA
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roFa(UBound(roFa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' faelleLaden

Function faelleEinf
 Dim ibeg&, i&, j&
 If UBound(rFa) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rFa)
   If rFa(i).fanf >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roFa)
    If roFa(i).fanf >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roFa(i - 1)
   ReDim Preserve roFa(i + UBound(rFa) - ibeg)
   For j = ibeg To UBound(rFa)
    Call roFaZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rFa = roFa
End Function ' faelleEinf

Public Function faelleSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere faelle"
 if not Allepat then
  sql = "SELECT pat_id FROM `faelle` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `faelle` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `faelle` (Pat_ID,Quartal,Nachname," & _
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
     "dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `faelle` (Pat_ID,Quartal,Nachname," & _
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
     "dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)        values"))
 For i = 1 to ubound(rFa)
'  rFa(i).AktZeit = now()
  rFa(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFa(i).Pat_ID, ",'" , rFa(i).Quartal, "','" , rFa(i).Nachname, "','" , rFa(i).Vorname, "'," , rFa(i).lfdnr, ",'" , rFa(i).TMFNr, "','" , rFa(i).VKNr, "','" , rFa(i).f4131, "','" ,  _
   rFa(i).f4132, "'," , DatFor_k(rFa(i).VschBeg), ",'" , rFa(i).KKasse_2, "'," , Replace$(rFa(i).FaktPers,",","."), "," , Replace$(rFa(i).FaktTechn,",","."), "," , Replace$(rFa(i).FaktLabor,",","."), "," , DatFor_k( _
   rFa(i).BhFB), "," , DatFor_k(rFa(i).BhFE1), "," , DatFor_k(rFa(i).BhFE2), ",'" , rFa(i).f4202, "'," , DatFor_k(rFa(i).ausgst), ",'" , rFa(i).KtrAbrB, "','" , rFa(i).AbrAr, "'," , DatFor_k( _
   rFa(i).lVorl), ",'" , rFa(i).IK, "','" , rFa(i).KVKs, "','" , rFa(i).KVKserg, "','" , rFa(i).Status, "','" , rFa(i).Kasse, "'," , rFa(i).KID, ",'" , rFa(i).GebOr, "','" , rFa(i).AbrGb, "','" ,  _
   rFa(i).PersKreis, "','" , rFa(i).SKtZusatz, "','" , rFa(i).letzteRegel, "','" , rFa(i).ÜwText, "'," , rFa(i).f4210, "," , rFa(i).AkfHAH, "," , rFa(i).AkfAB0, "," , rFa(i).AkfAK, ",'" ,  _
   rFa(i).statNuller, "','" , rFa(i).ÜbwV, "','" , rFa(i).ÜbWVLANR, "','" , rFa(i).ÜbWVBSNR, "','" , rFa(i).ÜbWVKVNR, "','" , rFa(i).AndÜw, "','" , rFa(i).Übwr, "','" , rFa(i).ÜbwLANR, "','" ,  _
   rFa(i).ÜWZiel, "','" , rFa(i).ÜWNNr, "','" , rFa(i).ÜWNaN, "','" , rFa(i).ÜWTit, "','" , rFa(i).ÜWVor, "','" , rFa(i).ÜWVsw, "'," , rFa(i).üwvid, ",'" , rFa(i).Auftrag, "','" , rFa(i).Verdacht, "','" ,  _
   rFa(i).Befund, "','" , rFa(i).statKlasse, "','" , rFa(i).f4237, "'," , rFa(i).statBehTage, "," , rFa(i).SchGr, ",'" , rFa(i).Weiterbeh, "'," , Replace$(rFa(i).f4266,",","."), ",'" ,  _
   rFa(i).PGeb, "','" , rFa(i).PGebErg, "','" , rFa(i).Mahnfrist, "','" , rFa(i).Unfallort, "','" , rFa(i).BeschAls, "'," , DatFor_k(rFa(i).BeschSeit), ",'" , rFa(i).Unfallbetrieb, "','" , rFa(i).f4570, "','" ,  _
   rFa(i).GOÄKatNr, "','" , rFa(i).GOÄKatName, "','" , rFa(i).abrArzt, "','" , rFa(i).privVers, "','" , rFa(i).AdNam, "','" , rFa(i).AdStr, "','" , rFa(i).AdPlz, "','" , rFa(i).AdOrt, "','" ,  _
   rFa(i).ÜwBG, "'," , DatFor_k(rFa(i).BhFE), ",'" , rFa(i).s8000, "','" , rFa(i).s8100, "'," , DatFor_k(rFa(i).AktZeit), "," , DatFor_k(rFa(i).Fanf), ",'" , rFa(i).altQuart, "'," , DatFor_k(rFa(i).QAnf), "," , DatFor_k( _
   rFa(i).QEnd), ",'" , rFa(i).QS, "','" , rFa(i).QT, "'," , rFa(i).StByte, "," , rFa(i).absPos, "," , rFa(i).LANRid, ",'" , rFa(i).f4108, "','" , rFa(i).BGFallNr, "'," ,  _
   rFa(i).lGewicht, "," , DatFor_k(rFa(i).vorET), ",'" , rFa(i).dmpVertret, "','" , rFa(i).dmpArztw, "','" , rFa(i).dmpHypos, "','" , rFa(i).dmpKhsA, "','" , rFa(i).dmpDMSchulEmpf, "','" ,  _
   rFa(i).dmpDMSchulWahrg, "','" , rFa(i).dmpHypertSchulEmpf, "','" , rFa(i).dmpHypertSchulWahrg, "','" , rFa(i).dmpKKTabakEmpf, "','" , rFa(i).dmpKKErnEmpf, "','" , rFa(i).dmpKKkTrainEmpf, "','" ,  _
   rFa(i).dmpHbA1cZiel, "','" , rFa(i).dmpUewFuss, "','" , rFa(i).dmpEinwDM, "','" , rFa(i).dmphalbj, "','" , rFa(i).dmpMA, "')")
  If SammelInsert <> 0 AND i < ubound(rFa) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFa) Then
   If Not obFor Then ForeignNo0
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If Not obFor Then ForeignYes0
   If rAF = 0 Then
    Err.Raise 998, , "Fehler in faelleSpeichern b.Pat. " & rFa(i).Pat_id & "Err.Number " & Err.Number & ", err.description: " & Err.Description
   End If
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
    If not rs.BOF() then
     If rs!code = 1265 Then
      Err.Raise -2147217833
     End If
    End If
   End If
 'Hier gibts mit Sammelins noch ein Problem ...
  Set rs = DBCn.Execute("SELECT FID FROM `faelle` WHERE pat_id = " & rFa(i).Pat_ID & " AND quartal = '" & rFa(i).Quartal & "' AND bhfb = " & DatFor_k(rFa(i).BhFB) & " AND bhfe1 = " & DatFor_k(rFa(i).BhFE1) & " AND ausgst = " & DatFor_k(rFa(i).ausgst))
  If rs.BOF Then
   Err.Raise 999, , "Fehler bei der Fallaktualisierung b.Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
  Else
   If rs!FID <> rFa(i).FID Then
    Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & rs!FID, true
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
     For jjj = 1 To UBound(rUs)
      If rUs(jjj).FID = rFa(i).FID Then
       rUs(jjj).FID = rs!FID
      End If
     Next jjj
     For jjj = 1 To UBound(rFu)
      If rFu(jjj).FID = rFa(i).FID Then
       rFu(jjj).FID = rs!FID
      End If
     Next jjj
     For jjj = 1 To UBound(rUl)
      If rUl(jjj).FID = rFa(i).FID Then
       rUl(jjj).FID = rs!FID
      End If
     Next jjj
     For jjj = 1 To UBound(rVk)
      If rVk(jjj).FID = rFa(i).FID Then
       rVk(jjj).FID = rs!FID
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(80)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFa),i)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "faelle", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "faelle", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
If Err.Number = -2147467259 Then
 Dim sqlquer$
 sqlquer = "INSERT INTO `kassenliste`(name,kurzname,`GO`,`VKNR`,`IK`,`eingef`,pid) values (" & "'" & rFa(I).kasse & "', '" & rFa(I).kkasse_2 & "', '" & rFa(I).GOÄKatName & "', '" & rFa(I).VKNr & "', '" & rFa(I).IK & "'," & Format(Now(), "yyyymmddHHMMSS") & "," & rFa(I).Pat_id & ")"
 InsKorr DBCn, DBCnS, sqlquer, rAF
 Resume
end if ' Err.Number = -2147467259 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in faelleSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' faelleSpeichern

Public Function roAuZuw(i&, j&)
 roAu(i).FID = rAu(j).FID
 roAu(i).Pat_ID = rAu(j).Pat_ID
 roAu(i).ZeitPunkt = rAu(j).ZeitPunkt
 roAu(i).Beginn = rAu(j).Beginn
 roAu(i).Ende = rAu(j).Ende
 roAu(i).ICDs = rAu(j).ICDs
 roAu(i).absPos = rAu(j).absPos
 roAu(i).AktZeit = rAu(j).AktZeit
 roAu(i).StByte = rAu(j).StByte
End Function ' roAuZuw

Public Function AuZUnt%(i&, j&)
 if roAu(i).FID <> rAu(j).FID Then gosub unter
 if roAu(i).Pat_ID <> rAu(j).Pat_ID Then gosub unter
 if roAu(i).ZeitPunkt <> rAu(j).ZeitPunkt Then gosub unter
 if roAu(i).Beginn <> rAu(j).Beginn Then gosub unter
 if roAu(i).Ende <> rAu(j).Ende Then gosub unter
 if roAu(i).ICDs <> rAu(j).ICDs Then gosub unter
 if roAu(i).absPos <> rAu(j).absPos Then gosub unter
 if roAu(i).AktZeit <> rAu(j).AktZeit Then gosub unter
 if roAu(i).StByte <> rAu(j).StByte Then gosub unter
 Exit Function
unter:
 AuZUnt = AuZUnt + 1
 Return
End Function ' AuZUnt

Public Function auLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roAu(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Beginn,Ende,ICDs,absPos,AktZeit,StByte FROM `au` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roAu)
  roAu(akt).FID = rs!FID
  roAu(akt).Pat_ID = rs!Pat_ID
  roAu(akt).ZeitPunkt = rs!ZeitPunkt
  roAu(akt).Beginn = rs!Beginn
  roAu(akt).Ende = rs!Ende
  roAu(akt).ICDs = rs!ICDs
  roAu(akt).absPos = rs!absPos
  roAu(akt).AktZeit = rs!AktZeit
  roAu(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roAu(UBound(roAu) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' auLaden

Function auEinf
 Dim ibeg&, i&, j&
 If UBound(rAu) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rAu)
   If rAu(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roAu)
    If roAu(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roAu(i - 1)
   ReDim Preserve roAu(i + UBound(rAu) - ibeg)
   For j = ibeg To UBound(rAu)
    Call roAuZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rAu = roAu
End Function ' auEinf

Public Function auSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere au"
 if not Allepat then
  sql = "SELECT pat_id FROM `au` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `au` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rAu(i).FID, "," , rAu(i).Pat_ID, "," , DatFor_k(rAu(i).ZeitPunkt), ",'" , rAu(i).Beginn, "','" , rAu(i).Ende, "','" , rAu(i).ICDs, "'," , rAu(i).absPos, "," , DatFor_k(rAu(i).AktZeit), "," ,  _
   rAu(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rAu) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rAu) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rAu),i)
  If Len(rAu(k).Beginn) > maxi(0) Then maxi(0) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(1) Then maxi(1) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(2) Then maxi(2) = Len(rAu(k).ICDs)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "au", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "au", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rAu), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & Left$(rAu(k).Beginn, maxL)  & "'",true : rAu(k).Beginn = Left$(rAu(k).Beginn, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & Left$(rAu(k).Ende, maxL)  & "'",true : rAu(k).Ende = Left$(rAu(k).Ende, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & Left$(rAu(k).ICDs, maxL)  & "'",true : rAu(k).ICDs = Left$(rAu(k).ICDs, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' auSpeichern

Public Function roBrZuw(i&, j&)
 roBr(i).FID = rBr(j).FID
 roBr(i).Pat_ID = rBr(j).Pat_ID
 roBr(i).ZeitPunkt = rBr(j).ZeitPunkt
 roBr(i).Pfad = rBr(j).Pfad
 roBr(i).Art = rBr(j).Art
 roBr(i).Name = rBr(j).Name
 roBr(i).Quelldatum = rBr(j).Quelldatum
 roBr(i).Typ = rBr(j).Typ
 roBr(i).AktZeit = rBr(j).AktZeit
 roBr(i).DokGroe = rBr(j).DokGroe
 roBr(i).DokAenD = rBr(j).DokAenD
 roBr(i).QS = rBr(j).QS
 roBr(i).QT = rBr(j).QT
 roBr(i).absPos = rBr(j).absPos
 roBr(i).StByte = rBr(j).StByte
End Function ' roBrZuw

Public Function BrZUnt%(i&, j&)
 if roBr(i).FID <> rBr(j).FID Then gosub unter
 if roBr(i).Pat_ID <> rBr(j).Pat_ID Then gosub unter
 if roBr(i).ZeitPunkt <> rBr(j).ZeitPunkt Then gosub unter
 if roBr(i).Pfad <> rBr(j).Pfad Then gosub unter
 if roBr(i).Art <> rBr(j).Art Then gosub unter
 if roBr(i).Name <> rBr(j).Name Then gosub unter
 if roBr(i).Quelldatum <> rBr(j).Quelldatum Then gosub unter
 if roBr(i).Typ <> rBr(j).Typ Then gosub unter
 if roBr(i).AktZeit <> rBr(j).AktZeit Then gosub unter
 if roBr(i).DokGroe <> rBr(j).DokGroe Then gosub unter
 if roBr(i).DokAenD <> rBr(j).DokAenD Then gosub unter
 if roBr(i).QS <> rBr(j).QS Then gosub unter
 if roBr(i).QT <> rBr(j).QT Then gosub unter
 if roBr(i).absPos <> rBr(j).absPos Then gosub unter
 if roBr(i).StByte <> rBr(j).StByte Then gosub unter
 Exit Function
unter:
 BrZUnt = BrZUnt + 1
 Return
End Function ' BrZUnt

Public Function briefeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roBr(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Pfad,Art,Name,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS,QT,absPos,StByte FROM `briefe` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roBr)
  roBr(akt).FID = rs!FID
  roBr(akt).Pat_ID = rs!Pat_ID
  roBr(akt).ZeitPunkt = rs!ZeitPunkt
  roBr(akt).Pfad = rs!Pfad
  roBr(akt).Art = rs!Art
  roBr(akt).Name = rs!Name
  roBr(akt).Quelldatum = rs!Quelldatum
  roBr(akt).Typ = rs!Typ
  roBr(akt).AktZeit = rs!AktZeit
  roBr(akt).DokGroe = rs!DokGroe
  roBr(akt).DokAenD = rs!DokAenD
  roBr(akt).QS = rs!QS
  roBr(akt).QT = rs!QT
  roBr(akt).absPos = rs!absPos
  roBr(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roBr(UBound(roBr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' briefeLaden

Function briefeEinf
 Dim ibeg&, i&, j&
 If UBound(rBr) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rBr)
   If rBr(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roBr)
    If roBr(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roBr(i - 1)
   ReDim Preserve roBr(i + UBound(rBr) - ibeg)
   For j = ibeg To UBound(rBr)
    Call roBrZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rBr = roBr
End Function ' briefeEinf

Public Function briefeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere briefe"
 if not Allepat then
  sql = "SELECT pat_id FROM `briefe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `briefe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "absPos,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `briefe` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "absPos,StByte)         values"))
 For i = 1 to ubound(rBr)
'  rBr(i).AktZeit = now()
  rBr(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rBr(i).FID, "," , rBr(i).Pat_ID, "," , DatFor_k(rBr(i).ZeitPunkt), ",'" , rBr(i).Pfad, "','" , rBr(i).Art, "','" , rBr(i).Name, "'," , DatFor_k(rBr(i).Quelldatum), ",'" , rBr(i).Typ, "'," , DatFor_k( _
   rBr(i).AktZeit), "," , rBr(i).DokGroe, "," , DatFor_k(rBr(i).DokAenD), ",'" , rBr(i).QS, "','" , rBr(i).QT, "'," , rBr(i).absPos, "," , rBr(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rBr) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rBr) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(5)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rBr),i)
  If Len(rBr(k).Pfad) > maxi(0) Then maxi(0) = Len(rBr(k).Pfad)
  If Len(rBr(k).Art) > maxi(1) Then maxi(1) = Len(rBr(k).Art)
  If Len(rBr(k).Name) > maxi(2) Then maxi(2) = Len(rBr(k).Name)
  If Len(rBr(k).Typ) > maxi(3) Then maxi(3) = Len(rBr(k).Typ)
  If Len(rBr(k).QS) > maxi(4) Then maxi(4) = Len(rBr(k).QS)
  If Len(rBr(k).QT) > maxi(5) Then maxi(5) = Len(rBr(k).QT)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "briefe", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "briefe", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rBr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & Left$(rBr(k).Pfad, maxL)  & "'",true : rBr(k).Pfad = Left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).Art & "' -> '" & Left$(rBr(k).Art, maxL)  & "'",true : rBr(k).Art = Left$(rBr(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).Name & "' -> '" & Left$(rBr(k).Name, maxL)  & "'",true : rBr(k).Name = Left$(rBr(k).Name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & Left$(rBr(k).Typ, maxL)  & "'",true : rBr(k).Typ = Left$(rBr(k).Typ, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & Left$(rBr(k).QS, maxL)  & "'",true : rBr(k).QS = Left$(rBr(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & Left$(rBr(k).QT, maxL)  & "'",true : rBr(k).QT = Left$(rBr(k).QT, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
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
 roDi(i).f6011 = rDi(j).f6011
End Function ' roDiZuw

Public Function DiZUnt%(i&, j&)
 if roDi(i).ID1 <> rDi(j).ID1 Then gosub unter
 if roDi(i).FID <> rDi(j).FID Then gosub unter
 if roDi(i).Pat_id <> rDi(j).Pat_id Then gosub unter
 if roDi(i).DiagDatum <> rDi(j).DiagDatum Then gosub unter
 if roDi(i).DiagSicherheit <> rDi(j).DiagSicherheit Then gosub unter
 if roDi(i).DiagText <> rDi(j).DiagText Then gosub unter
 if roDi(i).DiagSeite <> rDi(j).DiagSeite Then gosub unter
 if roDi(i).DiagAttr <> rDi(j).DiagAttr Then gosub unter
 if roDi(i).ICD <> rDi(j).ICD Then gosub unter
 if roDi(i).obDauer <> rDi(j).obDauer Then gosub unter
 if roDi(i).intBemerk <> rDi(j).intBemerk Then gosub unter
 if roDi(i).absPos <> rDi(j).absPos Then gosub unter
 if roDi(i).AktZeit <> rDi(j).AktZeit Then gosub unter
 if roDi(i).StByte <> rDi(j).StByte Then gosub unter
 if roDi(i).AusnBegr <> rDi(j).AusnBegr Then gosub unter
 if roDi(i).f6010 <> rDi(j).f6010 Then gosub unter
 if roDi(i).f6011 <> rDi(j).f6011 Then gosub unter
 Exit Function
unter:
 DiZUnt = DiZUnt + 1
 Return
End Function ' DiZUnt

Public Function diagnosenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDi(1)
 sql = "SELECT ID1,FID,Pat_id,DiagDatum,DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte,AusnBegr,f6010,f6011 FROM `diagnosen` WHERE Pat_ID=" & pid & " ORDER BY `DiagDatum`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roDi)
  roDi(akt).ID1 = rs!ID1
  roDi(akt).FID = rs!FID
  roDi(akt).Pat_id = rs!Pat_id
  roDi(akt).DiagDatum = rs!DiagDatum
  roDi(akt).DiagSicherheit = rs!DiagSicherheit
  roDi(akt).DiagText = rs!DiagText
  roDi(akt).DiagSeite = rs!DiagSeite
  roDi(akt).DiagAttr = rs!DiagAttr
  roDi(akt).ICD = rs!ICD
  roDi(akt).obDauer = rs!obDauer
  roDi(akt).intBemerk = rs!intBemerk
  roDi(akt).absPos = rs!absPos
  roDi(akt).AktZeit = rs!AktZeit
  roDi(akt).StByte = rs!StByte
  roDi(akt).AusnBegr = rs!AusnBegr
  roDi(akt).f6010 = rs!f6010
  roDi(akt).f6011 = rs!f6011
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roDi(UBound(roDi) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' diagnosenLaden

Function diagnosenEinf
 Dim ibeg&, i&, j&
 If UBound(rDi) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rDi)
   If rDi(i).DiagDatum >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roDi)
    If roDi(i).DiagDatum >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roDi(i - 1)
   ReDim Preserve roDi(i + UBound(rDi) - ibeg)
   For j = ibeg To UBound(rDi)
    Call roDiZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rDi = roDi
End Function ' diagnosenEinf

Public Function diagnosenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere diagnosen"
 if not Allepat then
  sql = "SELECT pat_id FROM `diagnosen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `diagnosen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,f6011) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `diagnosen` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,f6010,f6011)  values"))
 For i = 1 to ubound(rDi)
'  rDi(i).AktZeit = now()
  rDi(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDi(i).FID, "," , rDi(i).Pat_id, "," , DatFor_k(rDi(i).DiagDatum), ",'" , rDi(i).DiagSicherheit, "','" , rDi(i).DiagText, "','" , rDi(i).DiagSeite, "','" , rDi(i).DiagAttr, "','" ,  _
   rDi(i).ICD, "'," , rDi(i).obDauer, ",'" , rDi(i).intBemerk, "'," , rDi(i).absPos, "," , DatFor_k(rDi(i).AktZeit), "," , rDi(i).StByte, ",'" , rDi(i).AusnBegr, "'," , rDi(i).f6010, ",'" , rDi(i).f6011, "')")
  If SammelInsert <> 0 AND i < ubound(rDi) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rDi) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(7)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDi),i)
  If Len(rDi(k).DiagSicherheit) > maxi(0) Then maxi(0) = Len(rDi(k).DiagSicherheit)
  If Len(rDi(k).DiagText) > maxi(1) Then maxi(1) = Len(rDi(k).DiagText)
  If Len(rDi(k).DiagSeite) > maxi(2) Then maxi(2) = Len(rDi(k).DiagSeite)
  If Len(rDi(k).DiagAttr) > maxi(3) Then maxi(3) = Len(rDi(k).DiagAttr)
  If Len(rDi(k).ICD) > maxi(4) Then maxi(4) = Len(rDi(k).ICD)
  If Len(rDi(k).intBemerk) > maxi(5) Then maxi(5) = Len(rDi(k).intBemerk)
  If Len(rDi(k).AusnBegr) > maxi(6) Then maxi(6) = Len(rDi(k).AusnBegr)
  If Len(rDi(k).f6011) > maxi(7) Then maxi(7) = Len(rDi(k).f6011)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "diagnosen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "diagnosen", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' diagnosenSpeichern

Public Function roDoZuw(i&, j&)
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
End Function ' roDoZuw

Public Function DoZUnt%(i&, j&)
 if roDo(i).FID <> rDo(j).FID Then gosub unter
 if roDo(i).Pat_ID <> rDo(j).Pat_ID Then gosub unter
 if roDo(i).ZeitPunkt <> rDo(j).ZeitPunkt Then gosub unter
 if roDo(i).DokPfad <> rDo(j).DokPfad Then gosub unter
 if roDo(i).DokArt <> rDo(j).DokArt Then gosub unter
 if roDo(i).DokName <> rDo(j).DokName Then gosub unter
 if roDo(i).Quelldatum <> rDo(j).Quelldatum Then gosub unter
 if roDo(i).absPos <> rDo(j).absPos Then gosub unter
 if roDo(i).AktZeit <> rDo(j).AktZeit Then gosub unter
 if roDo(i).DokGroe <> rDo(j).DokGroe Then gosub unter
 if roDo(i).DokAenD <> rDo(j).DokAenD Then gosub unter
 if roDo(i).QS <> rDo(j).QS Then gosub unter
 if roDo(i).QT <> rDo(j).QT Then gosub unter
 if roDo(i).StByte <> rDo(j).StByte Then gosub unter
 Exit Function
unter:
 DoZUnt = DoZUnt + 1
 Return
End Function ' DoZUnt

Public Function dokumenteLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDo(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT,StByte FROM `dokumente` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roDo)
  roDo(akt).FID = rs!FID
  roDo(akt).Pat_ID = rs!Pat_ID
  roDo(akt).ZeitPunkt = rs!ZeitPunkt
  roDo(akt).DokPfad = rs!DokPfad
  roDo(akt).DokArt = rs!DokArt
  roDo(akt).DokName = rs!DokName
  roDo(akt).Quelldatum = rs!Quelldatum
  roDo(akt).absPos = rs!absPos
  roDo(akt).AktZeit = rs!AktZeit
  roDo(akt).DokGroe = rs!DokGroe
  roDo(akt).DokAenD = rs!DokAenD
  roDo(akt).QS = rs!QS
  roDo(akt).QT = rs!QT
  roDo(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roDo(UBound(roDo) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' dokumenteLaden

Function dokumenteEinf
 Dim ibeg&, i&, j&
 If UBound(rDo) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rDo)
   If rDo(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roDo)
    If roDo(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roDo(i - 1)
   ReDim Preserve roDo(i + UBound(rDo) - ibeg)
   For j = ibeg To UBound(rDo)
    Call roDoZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rDo = roDo
End Function ' dokumenteEinf

Public Function dokumenteSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere dokumente"
 if not Allepat then
  sql = "SELECT pat_id FROM `dokumente` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `dokumente` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rDo(i).FID, "," , rDo(i).Pat_ID, "," , DatFor_k(rDo(i).ZeitPunkt), ",'" , rDo(i).DokPfad, "','" , rDo(i).DokArt, "','" , rDo(i).DokName, "'," , DatFor_k(rDo(i).Quelldatum), "," ,  _
   rDo(i).absPos, "," , DatFor_k(rDo(i).AktZeit), "," , rDo(i).DokGroe, "," , DatFor_k(rDo(i).DokAenD), ",'" , rDo(i).QS, "','" , rDo(i).QT, "'," , rDo(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rDo) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rDo) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(4)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDo),i)
  If Len(rDo(k).DokPfad) > maxi(0) Then maxi(0) = Len(rDo(k).DokPfad)
  If Len(rDo(k).DokArt) > maxi(1) Then maxi(1) = Len(rDo(k).DokArt)
  If Len(rDo(k).DokName) > maxi(2) Then maxi(2) = Len(rDo(k).DokName)
  If Len(rDo(k).QS) > maxi(3) Then maxi(3) = Len(rDo(k).QS)
  If Len(rDo(k).QT) > maxi(4) Then maxi(4) = Len(rDo(k).QT)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dokumente", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dokumente", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & Left$(rDo(k).DokPfad, maxL)  & "'",true : rDo(k).DokPfad = Left$(rDo(k).DokPfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & Left$(rDo(k).DokArt, maxL)  & "'",true : rDo(k).DokArt = Left$(rDo(k).DokArt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & Left$(rDo(k).DokName, maxL)  & "'",true : rDo(k).DokName = Left$(rDo(k).DokName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & Left$(rDo(k).QS, maxL)  & "'",true : rDo(k).QS = Left$(rDo(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & Left$(rDo(k).QT, maxL)  & "'",true : rDo(k).QT = Left$(rDo(k).QT, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' dokumenteSpeichern

Public Function roEiZuw(i&, j&)
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
End Function ' roEiZuw

Public Function EiZUnt%(i&, j&)
 if roEi(i).FID <> rEi(j).FID Then gosub unter
 if roEi(i).Pat_ID <> rEi(j).Pat_ID Then gosub unter
 if roEi(i).ZeitPunkt <> rEi(j).ZeitPunkt Then gosub unter
 if roEi(i).Art <> rEi(j).Art Then gosub unter
 if roEi(i).Inhalt <> rEi(j).Inhalt Then gosub unter
 if roEi(i).absPos <> rEi(j).absPos Then gosub unter
 if roEi(i).AktZeit <> rEi(j).AktZeit Then gosub unter
 if roEi(i).QS <> rEi(j).QS Then gosub unter
 if roEi(i).QT <> rEi(j).QT Then gosub unter
 if roEi(i).StByte <> rEi(j).StByte Then gosub unter
 if roEi(i).id <> rEi(j).id Then gosub unter
 if roEi(i).inhNum <> rEi(j).inhNum Then gosub unter
 Exit Function
unter:
 EiZUnt = EiZUnt + 1
 Return
End Function ' EiZUnt

Public Function eintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roEi(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Art,Inhalt,absPos,AktZeit,QS,QT,StByte,id,inhNum FROM `eintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roEi)
  roEi(akt).FID = rs!FID
  roEi(akt).Pat_ID = rs!Pat_ID
  roEi(akt).ZeitPunkt = rs!ZeitPunkt
  roEi(akt).Art = rs!Art
  roEi(akt).Inhalt = rs!Inhalt
  roEi(akt).absPos = rs!absPos
  roEi(akt).AktZeit = rs!AktZeit
  roEi(akt).QS = rs!QS
  roEi(akt).QT = rs!QT
  roEi(akt).StByte = rs!StByte
  roEi(akt).id = rs!id
  roEi(akt).inhNum = rs!inhNum
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roEi(UBound(roEi) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' eintraegeLaden

Function eintraegeEinf
 Dim ibeg&, i&, j&
 If UBound(rEi) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rEi)
   If rEi(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roEi)
    If roEi(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roEi(i - 1)
   ReDim Preserve roEi(i + UBound(rEi) - ibeg)
   For j = ibeg To UBound(rEi)
    Call roEiZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rEi = roEi
End Function ' eintraegeEinf

Public Function eintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere eintraege"
 if not Allepat then
  sql = "SELECT pat_id FROM `eintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `eintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rEi(i).FID, "," , rEi(i).Pat_ID, "," , DatFor_k(rEi(i).ZeitPunkt), ",'" , rEi(i).Art, "','" , rEi(i).Inhalt, "'," , rEi(i).absPos, "," , DatFor_k(rEi(i).AktZeit), ",'" , rEi(i).QS, "','" ,  _
   rEi(i).QT, "'," , rEi(i).StByte, "," , Replace$(rEi(i).inhNum,",","."), ")")
  If SammelInsert <> 0 AND i < ubound(rEi) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rEi) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(3)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rEi),i)
  If Len(rEi(k).Art) > maxi(0) Then maxi(0) = Len(rEi(k).Art)
  If Len(rEi(k).Inhalt) > maxi(1) Then maxi(1) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(2) Then maxi(2) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(3) Then maxi(3) = Len(rEi(k).QT)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "eintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "eintraege", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rEi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).Art & "' -> '" & Left$(rEi(k).Art, maxL)  & "'",true : rEi(k).Art = Left$(rEi(k).Art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & Left$(rEi(k).Inhalt, maxL)  & "'",true : rEi(k).Inhalt = Left$(rEi(k).Inhalt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & Left$(rEi(k).QS, maxL)  & "'",true : rEi(k).QS = Left$(rEi(k).QS, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & Left$(rEi(k).QT, maxL)  & "'",true : rEi(k).QT = Left$(rEi(k).QT, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' eintraegeSpeichern
Public Function forminhaltform_abkSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhaltform_abk"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhaltform_abk` (Form_Abk) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhaltform_abk` (Form_Abk)             values"))
 For i = rFi1 + 1 to ubound(rFi)
'  rFi(i).AktZeit = now()
  If SammelInsert = 0 Or i = rFi1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rFi(i).Form_Abk, "')")
  If SammelInsert <> 0 AND i < ubound(rFi) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFi) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,rFi1 + 1,i) to iif(SammelInsert<>0,ubound(rFi),i)
  If Len(rFi(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFi(k).Form_Abk)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhaltform_abk", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhaltform_abk", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,rFi1 + 1, i) To IIf(SammelInsert <> 0,ubound(rFi), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFi.Form_Abk: '" & rFi(k).Form_Abk & "' -> '" & Left$(rFi(k).Form_Abk, maxL)  & "'",true : rFi(k).Form_Abk = Left$(rFi(k).Form_Abk, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhaltform_abkSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' forminhaltform_abkSpeichern
Public Function formulareSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
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
  If SammelInsert = 0 Or i = rFo1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFo(i).FormID, ",'" , rFo(i).Form_Abk, "','" , rFo(i).FormBez, "','" , rFo(i).FormVorl, "'," , DatFor_k(rFo(i).AktZeit), "," , rFo(i).absPos, "," , rFo(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rFo) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFo) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,rFo1 + 1,i) to iif(SammelInsert<>0,ubound(rFo),i)
  If Len(rFo(k).Form_Abk) > maxi(0) Then maxi(0) = Len(rFo(k).Form_Abk)
  If Len(rFo(k).FormBez) > maxi(1) Then maxi(1) = Len(rFo(k).FormBez)
  If Len(rFo(k).FormVorl) > maxi(2) Then maxi(2) = Len(rFo(k).FormVorl)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "formulare", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "formulare", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,rFo1 + 1, i) To IIf(SammelInsert <> 0,ubound(rFo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & Left$(rFo(k).Form_Abk, maxL)  & "'",true : rFo(k).Form_Abk = Left$(rFo(k).Form_Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & Left$(rFo(k).FormBez, maxL)  & "'",true : rFo(k).FormBez = Left$(rFo(k).FormBez, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & Left$(rFo(k).FormVorl, maxL)  & "'",true : rFo(k).FormVorl = Left$(rFo(k).FormVorl, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' formulareSpeichern

Public Function roFrZuw(i&, j&)
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
End Function ' roFrZuw

Public Function FrZUnt%(i&, j&)
 if roFr(i).FoID <> rFr(j).FoID Then gosub unter
 if roFr(i).FID <> rFr(j).FID Then gosub unter
 if roFr(i).Pat_ID <> rFr(j).Pat_ID Then gosub unter
 if roFr(i).Form_ID <> rFr(j).Form_ID Then gosub unter
 if roFr(i).ZeitPunkt <> rFr(j).ZeitPunkt Then gosub unter
 if roFr(i).AbsPos <> rFr(j).AbsPos Then gosub unter
 if roFr(i).AktZeit <> rFr(j).AktZeit Then gosub unter
 if roFr(i).StByte <> rFr(j).StByte Then gosub unter
 if roFr(i).Satzart <> rFr(j).Satzart Then gosub unter
 if roFr(i).Satzlänge <> rFr(j).Satzlänge Then gosub unter
 if roFr(i).LANRid <> rFr(j).LANRid Then gosub unter
 Exit Function
unter:
 FrZUnt = FrZUnt + 1
 Return
End Function ' FrZUnt

Public Function forminhkopfLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFr(1)
 sql = "SELECT FoID,FID,Pat_ID,Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid FROM `forminhkopf` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
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
  roFr(akt).Satzart = rs!Satzart
  roFr(akt).Satzlänge = rs!Satzlänge
  roFr(akt).LANRid = rs!LANRid
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roFr(UBound(roFr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' forminhkopfLaden

Function forminhkopfEinf
 Dim ibeg&, i&, j&
 If UBound(rFr) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rFr)
   If rFr(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roFr)
    If roFr(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roFr(i - 1)
   ReDim Preserve roFr(i + UBound(rFr) - ibeg)
   For j = ibeg To UBound(rFr)
    Call roFrZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rFr = roFr
End Function ' forminhkopfEinf

Public Function forminhkopfSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhkopf"
 if not Allepat then
'  sql = "DELETE FROM `forminhfeld` WHERE foid in (SELECT foid FROM `forminhkopf` WHERE pat_id = " & CStr(rNa(0).Pat_ID) & ")"
'  Call DBCn.Execute(sql)
  sql = "SELECT pat_id FROM `forminhkopf` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `forminhkopf` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFr(i).FoID, "," , rFr(i).FID, "," , rFr(i).Pat_ID, "," , rFr(i).Form_ID, "," , DatFor_k(rFr(i).ZeitPunkt), "," , rFr(i).AbsPos, "," , DatFor_k(rFr(i).AktZeit), "," , rFr(i).StByte, ",'" ,  _
   rFr(i).Satzart, "','" , rFr(i).Satzlänge, "'," , rFr(i).LANRid, ")")
  If SammelInsert <> 0 AND i < ubound(rFr) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFr) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFr),i)
  If Len(rFr(k).Satzart) > maxi(0) Then maxi(0) = Len(rFr(k).Satzart)
  If Len(rFr(k).Satzlänge) > maxi(1) Then maxi(1) = Len(rFr(k).Satzlänge)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhkopf", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhkopf", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & Left$(rFr(k).Satzart, maxL)  & "'",true : rFr(k).Satzart = Left$(rFr(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & Left$(rFr(k).Satzlänge, maxL)  & "'",true : rFr(k).Satzlänge = Left$(rFr(k).Satzlänge, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' forminhkopfSpeichern
Public Function forminhfeldSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere forminhfeld"
 if not Allepat then
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `forminhfeld` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      values"))
 For i = 1 to ubound(rFm)
'  rFm(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFm(i).FoID, "," , rFm(i).Nr, "," , rFm(i).FeldNr, "," , rFm(i).FeldVW, "," , rFm(i).FeldInhVW, ")")
  If SammelInsert <> 0 AND i < ubound(rFm) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFm) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(-1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFm),i)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "forminhfeld", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "forminhfeld", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rFm), i)
      SELECT CASE m
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' forminhfeldSpeichern

Public Function roKhZuw(i&, j&)
 roKh(i).FID = rKh(j).FID
 roKh(i).Pat_ID = rKh(j).Pat_ID
 roKh(i).ZeitPunkt = rKh(j).ZeitPunkt
 roKh(i).Ziel = rKh(j).Ziel
 roKh(i).Diagnose = rKh(j).Diagnose
 roKh(i).absPos = rKh(j).absPos
 roKh(i).AktZeit = rKh(j).AktZeit
 roKh(i).StByte = rKh(j).StByte
End Function ' roKhZuw

Public Function KhZUnt%(i&, j&)
 if roKh(i).FID <> rKh(j).FID Then gosub unter
 if roKh(i).Pat_ID <> rKh(j).Pat_ID Then gosub unter
 if roKh(i).ZeitPunkt <> rKh(j).ZeitPunkt Then gosub unter
 if roKh(i).Ziel <> rKh(j).Ziel Then gosub unter
 if roKh(i).Diagnose <> rKh(j).Diagnose Then gosub unter
 if roKh(i).absPos <> rKh(j).absPos Then gosub unter
 if roKh(i).AktZeit <> rKh(j).AktZeit Then gosub unter
 if roKh(i).StByte <> rKh(j).StByte Then gosub unter
 Exit Function
unter:
 KhZUnt = KhZUnt + 1
 Return
End Function ' KhZUnt

Public Function kheinweisLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roKh(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Ziel,Diagnose,absPos,AktZeit,StByte FROM `kheinweis` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roKh)
  roKh(akt).FID = rs!FID
  roKh(akt).Pat_ID = rs!Pat_ID
  roKh(akt).ZeitPunkt = rs!ZeitPunkt
  roKh(akt).Ziel = rs!Ziel
  roKh(akt).Diagnose = rs!Diagnose
  roKh(akt).absPos = rs!absPos
  roKh(akt).AktZeit = rs!AktZeit
  roKh(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roKh(UBound(roKh) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' kheinweisLaden

Function kheinweisEinf
 Dim ibeg&, i&, j&
 If UBound(rKh) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rKh)
   If rKh(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roKh)
    If roKh(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roKh(i - 1)
   ReDim Preserve roKh(i + UBound(rKh) - ibeg)
   For j = ibeg To UBound(rKh)
    Call roKhZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rKh = roKh
End Function ' kheinweisEinf

Public Function kheinweisSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere kheinweis"
 if not Allepat then
  sql = "SELECT pat_id FROM `kheinweis` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `kheinweis` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKh(i).FID, "," , rKh(i).Pat_ID, "," , DatFor_k(rKh(i).ZeitPunkt), ",'" , rKh(i).Ziel, "','" , rKh(i).Diagnose, "'," , rKh(i).absPos, "," , DatFor_k(rKh(i).AktZeit), "," , rKh(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rKh) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rKh) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rKh),i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kheinweis", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kheinweis", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rKh), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & Left$(rKh(k).Ziel, maxL)  & "'",true : rKh(k).Ziel = Left$(rKh(k).Ziel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & Left$(rKh(k).Diagnose, maxL)  & "'",true : rKh(k).Diagnose = Left$(rKh(k).Diagnose, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' kheinweisSpeichern

Public Function roLbZuw(i&, j&)
 roLb(i).FID = rLb(j).FID
 roLb(i).Pat_ID = rLb(j).Pat_ID
 roLb(i).ZeitPunkt = rLb(j).ZeitPunkt
 roLb(i).AnfText = rLb(j).AnfText
 roLb(i).absPos = rLb(j).absPos
 roLb(i).AktZeit = rLb(j).AktZeit
 roLb(i).StByte = rLb(j).StByte
End Function ' roLbZuw

Public Function LbZUnt%(i&, j&)
 if roLb(i).FID <> rLb(j).FID Then gosub unter
 if roLb(i).Pat_ID <> rLb(j).Pat_ID Then gosub unter
 if roLb(i).ZeitPunkt <> rLb(j).ZeitPunkt Then gosub unter
 if roLb(i).AnfText <> rLb(j).AnfText Then gosub unter
 if roLb(i).absPos <> rLb(j).absPos Then gosub unter
 if roLb(i).AktZeit <> rLb(j).AktZeit Then gosub unter
 if roLb(i).StByte <> rLb(j).StByte Then gosub unter
 Exit Function
unter:
 LbZUnt = LbZUnt + 1
 Return
End Function ' LbZUnt

Public Function lbanforderungenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLb(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,AnfText,absPos,AktZeit,StByte FROM `lbanforderungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roLb)
  roLb(akt).FID = rs!FID
  roLb(akt).Pat_ID = rs!Pat_ID
  roLb(akt).ZeitPunkt = rs!ZeitPunkt
  roLb(akt).AnfText = rs!AnfText
  roLb(akt).absPos = rs!absPos
  roLb(akt).AktZeit = rs!AktZeit
  roLb(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roLb(UBound(roLb) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' lbanforderungenLaden

Function lbanforderungenEinf
 Dim ibeg&, i&, j&
 If UBound(rLb) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rLb)
   If rLb(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roLb)
    If roLb(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roLb(i - 1)
   ReDim Preserve roLb(i + UBound(rLb) - ibeg)
   For j = ibeg To UBound(rLb)
    Call roLbZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rLb = roLb
End Function ' lbanforderungenEinf

Public Function lbanforderungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere lbanforderungen"
 if not Allepat then
  sql = "SELECT pat_id FROM `lbanforderungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `lbanforderungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLb(i).FID, "," , rLb(i).Pat_ID, "," , DatFor_k(rLb(i).ZeitPunkt), ",'" , rLb(i).AnfText, "'," , rLb(i).absPos, "," , DatFor_k(rLb(i).AktZeit), "," , rLb(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rLb) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLb) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLb),i)
  If Len(rLb(k).AnfText) > maxi(0) Then maxi(0) = Len(rLb(k).AnfText)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "lbanforderungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "lbanforderungen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLb), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & Left$(rLb(k).AnfText, maxL)  & "'",true : rLb(k).AnfText = Left$(rLb(k).AnfText, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' lbanforderungenSpeichern

Public Function roLaZuw(i&, j&)
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
End Function ' roLaZuw

Public Function LaZUnt%(i&, j&)
 if roLa(i).FID <> rLa(j).FID Then gosub unter
 if roLa(i).Pat_ID <> rLa(j).Pat_ID Then gosub unter
 if roLa(i).ZeitPunkt <> rLa(j).ZeitPunkt Then gosub unter
 if roLa(i).FertigStGrad <> rLa(j).FertigStGrad Then gosub unter
 if roLa(i).Abkü <> rLa(j).Abkü Then gosub unter
 if roLa(i).LangtextVW <> rLa(j).LangtextVW Then gosub unter
 if roLa(i).Wert <> rLa(j).Wert Then gosub unter
 if roLa(i).Einheit <> rLa(j).Einheit Then gosub unter
 if roLa(i).KommentarVW <> rLa(j).KommentarVW Then gosub unter
 if roLa(i).AbsPos <> rLa(j).AbsPos Then gosub unter
 if roLa(i).AktZeit <> rLa(j).AktZeit Then gosub unter
 if roLa(i).Refnr <> rLa(j).Refnr Then gosub unter
 if roLa(i).StByte <> rLa(j).StByte Then gosub unter
 if roLa(i).ID <> rLa(j).ID Then gosub unter
 Exit Function
unter:
 LaZUnt = LaZUnt + 1
 Return
End Function ' LaZUnt

Public Function laborneuLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLa(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,FertigStGrad,Abkü,LangtextVW,Wert,Einheit,KommentarVW,AbsPos,AktZeit,Refnr,StByte,ID FROM `laborneu` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roLa)
  roLa(akt).FID = rs!FID
  roLa(akt).Pat_ID = rs!Pat_ID
  roLa(akt).ZeitPunkt = rs!ZeitPunkt
  roLa(akt).FertigStGrad = rs!FertigStGrad
  roLa(akt).Abkü = rs!Abkü
  roLa(akt).LangtextVW = rs!LangtextVW
  roLa(akt).Wert = rs!Wert
  roLa(akt).Einheit = rs!Einheit
  roLa(akt).KommentarVW = rs!KommentarVW
  roLa(akt).AbsPos = rs!AbsPos
  roLa(akt).AktZeit = rs!AktZeit
  roLa(akt).Refnr = rs!Refnr
  roLa(akt).StByte = rs!StByte
  roLa(akt).ID = rs!ID
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roLa(UBound(roLa) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' laborneuLaden

Function laborneuEinf
 Dim ibeg&, i&, j&
 If UBound(rLa) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rLa)
   If rLa(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roLa)
    If roLa(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roLa(i - 1)
   ReDim Preserve roLa(i + UBound(rLa) - ibeg)
   For j = ibeg To UBound(rLa)
    Call roLaZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rLa = roLa
End Function ' laborneuEinf

Public Function laborneuSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborneu"
 if not Allepat then
  sql = "SELECT pat_id FROM `laborneu` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `laborneu` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLa(i).FID, "," , rLa(i).Pat_ID, "," , DatFor_k(rLa(i).ZeitPunkt), ",'" , rLa(i).FertigStGrad, "','" , rLa(i).Abkü, "'," , rLa(i).LangtextVW, ",'" , rLa(i).Wert, "','" , rLa(i).Einheit, "'," ,  _
   rLa(i).KommentarVW, "," , rLa(i).AbsPos, "," , DatFor_k(rLa(i).AktZeit), "," , rLa(i).Refnr, "," , rLa(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rLa) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLa) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(3)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLa),i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborneu", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborneu", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLa), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & Left$(rLa(k).FertigStGrad, maxL)  & "'",true : rLa(k).FertigStGrad = Left$(rLa(k).FertigStGrad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & Left$(rLa(k).Abkü, maxL)  & "'",true : rLa(k).Abkü = Left$(rLa(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & Left$(rLa(k).Wert, maxL)  & "'",true : rLa(k).Wert = Left$(rLa(k).Wert, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & Left$(rLa(k).Einheit, maxL)  & "'",true : rLa(k).Einheit = Left$(rLa(k).Einheit, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborneuSpeichern

Public Function roLeZuw(i&, j&)
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
End Function ' roLeZuw

Public Function LeZUnt%(i&, j&)
 if roLe(i).id <> rLe(j).id Then gosub unter
 if roLe(i).FID <> rLe(j).FID Then gosub unter
 if roLe(i).Pat_ID <> rLe(j).Pat_ID Then gosub unter
 if roLe(i).ZeitPunkt <> rLe(j).ZeitPunkt Then gosub unter
 if roLe(i).Leistung <> rLe(j).Leistung Then gosub unter
 if roLe(i).f5002 <> rLe(j).f5002 Then gosub unter
 if roLe(i).f5005 <> rLe(j).f5005 Then gosub unter
 if roLe(i).f5006 <> rLe(j).f5006 Then gosub unter
 if roLe(i).f5009 <> rLe(j).f5009 Then gosub unter
 if roLe(i).Med <> rLe(j).Med Then gosub unter
 if roLe(i).f5015 <> rLe(j).f5015 Then gosub unter
 if roLe(i).f5016 <> rLe(j).f5016 Then gosub unter
 if roLe(i).f5021 <> rLe(j).f5021 Then gosub unter
 if roLe(i).f5026 <> rLe(j).f5026 Then gosub unter
 if roLe(i).Faktor <> rLe(j).Faktor Then gosub unter
 if roLe(i).f5098 <> rLe(j).f5098 Then gosub unter
 if roLe(i).LANR <> rLe(j).LANR Then gosub unter
 if roLe(i).letzVorg <> rLe(j).letzVorg Then gosub unter
 if roLe(i).Ausn <> rLe(j).Ausn Then gosub unter
 if roLe(i).Beme <> rLe(j).Beme Then gosub unter
 if roLe(i).absPos <> rLe(j).absPos Then gosub unter
 if roLe(i).AktZeit <> rLe(j).AktZeit Then gosub unter
 if roLe(i).QS <> rLe(j).QS Then gosub unter
 if roLe(i).QT <> rLe(j).QT Then gosub unter
 if roLe(i).StByte <> rLe(j).StByte Then gosub unter
 if roLe(i).LANRid <> rLe(j).LANRid Then gosub unter
 if roLe(i).Sachkbez <> rLe(j).Sachkbez Then gosub unter
 if roLe(i).Sachkct <> rLe(j).Sachkct Then gosub unter
 if roLe(i).Zone <> rLe(j).Zone Then gosub unter
 Exit Function
unter:
 LeZUnt = LeZUnt + 1
 Return
End Function ' LeZUnt

Public Function leistungenLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roLe(1)
 sql = "SELECT id,FID,Pat_ID,ZeitPunkt,Leistung,f5002,f5005,f5006,f5009,Med,f5015,f5016,f5021,f5026,Faktor,f5098,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS,QT,StByte,LANRid,Sachkbez,Sachkct,Zone FROM `leistungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roLe)
  roLe(akt).id = rs!id
  roLe(akt).FID = rs!FID
  roLe(akt).Pat_ID = rs!Pat_ID
  roLe(akt).ZeitPunkt = rs!ZeitPunkt
  roLe(akt).Leistung = rs!Leistung
  roLe(akt).f5002 = rs!f5002
  roLe(akt).f5005 = rs!f5005
  roLe(akt).f5006 = rs!f5006
  roLe(akt).f5009 = rs!f5009
  roLe(akt).Med = rs!Med
  roLe(akt).f5015 = rs!f5015
  roLe(akt).f5016 = rs!f5016
  roLe(akt).f5021 = rs!f5021
  roLe(akt).f5026 = rs!f5026
  roLe(akt).Faktor = rs!Faktor
  roLe(akt).f5098 = rs!f5098
  roLe(akt).LANR = rs!LANR
  roLe(akt).letzVorg = rs!letzVorg
  roLe(akt).Ausn = rs!Ausn
  roLe(akt).Beme = rs!Beme
  roLe(akt).absPos = rs!absPos
  roLe(akt).AktZeit = rs!AktZeit
  roLe(akt).QS = rs!QS
  roLe(akt).QT = rs!QT
  roLe(akt).StByte = rs!StByte
  roLe(akt).LANRid = rs!LANRid
  roLe(akt).Sachkbez = rs!Sachkbez
  roLe(akt).Sachkct = rs!Sachkct
  roLe(akt).Zone = rs!Zone
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roLe(UBound(roLe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' leistungenLaden

Function leistungenEinf
 Dim ibeg&, i&, j&
 If UBound(rLe) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rLe)
   If rLe(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roLe)
    If roLe(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roLe(i - 1)
   ReDim Preserve roLe(i + UBound(rLe) - ibeg)
   For j = ibeg To UBound(rLe)
    Call roLeZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rLe = roLe
End Function ' leistungenEinf

Public Function leistungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere leistungen"
 if not Allepat then
  sql = "SELECT pat_id FROM `leistungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `leistungen` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLe(i).FID, "," , rLe(i).Pat_ID, "," , DatFor_k(rLe(i).ZeitPunkt), ",'" , rLe(i).Leistung, "','" , rLe(i).f5002, "','" , rLe(i).f5005, "','" , rLe(i).f5006, "','" , rLe(i).f5009, "','" ,  _
   rLe(i).Med, "','" , rLe(i).f5015, "','" , rLe(i).f5016, "','" , rLe(i).f5021, "','" , rLe(i).f5026, "','" , rLe(i).Faktor, "','" , rLe(i).f5098, "','" , rLe(i).LANR, "'," , DatFor_k(rLe(i).letzVorg), ",'" ,  _
   rLe(i).Ausn, "','" , rLe(i).Beme, "'," , rLe(i).absPos, "," , DatFor_k(rLe(i).AktZeit), ",'" , rLe(i).QS, "','" , rLe(i).QT, "'," , rLe(i).StByte, "," , rLe(i).LANRid, ",'" ,  _
   rLe(i).Sachkbez, "'," , rLe(i).Sachkct, ",'" , rLe(i).Zone, "')")
  If SammelInsert <> 0 AND i < ubound(rLe) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLe) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(18)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLe),i)
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
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "leistungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "leistungen", rsc)
    If maxL > 0 Then
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
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).LANR & "' -> '" & Left$(rLe(k).LANR, maxL)  & "'",true : rLe(k).LANR = Left$(rLe(k).LANR, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & Left$(rLe(k).Ausn, maxL)  & "'",true : rLe(k).Ausn = Left$(rLe(k).Ausn, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).Beme & "' -> '" & Left$(rLe(k).Beme, maxL)  & "'",true : rLe(k).Beme = Left$(rLe(k).Beme, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & Left$(rLe(k).QS, maxL)  & "'",true : rLe(k).QS = Left$(rLe(k).QS, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & Left$(rLe(k).QT, maxL)  & "'",true : rLe(k).QT = Left$(rLe(k).QT, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & Left$(rLe(k).Sachkbez, maxL)  & "'",true : rLe(k).Sachkbez = Left$(rLe(k).Sachkbez, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & Left$(rLe(k).Zone, maxL)  & "'",true : rLe(k).Zone = Left$(rLe(k).Zone, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' leistungenSpeichern

Public Function roMeZuw(i&, j&)
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
End Function ' roMeZuw

Public Function MeZUnt%(i&, j&)
 if roMe(i).FID <> rMe(j).FID Then gosub unter
 if roMe(i).Pat_ID <> rMe(j).Pat_ID Then gosub unter
 if roMe(i).MPNr <> rMe(j).MPNr Then gosub unter
 if roMe(i).ZeitPunkt <> rMe(j).ZeitPunkt Then gosub unter
 if roMe(i).Datum <> rMe(j).Datum Then gosub unter
 if roMe(i).Medikament <> rMe(j).Medikament Then gosub unter
 if roMe(i).MedAnfang <> rMe(j).MedAnfang Then gosub unter
 if roMe(i).Wirkstoff <> rMe(j).Wirkstoff Then gosub unter
 if roMe(i).PZN <> rMe(j).PZN Then gosub unter
 if roMe(i).FeldNr <> rMe(j).FeldNr Then gosub unter
 if roMe(i).mo <> rMe(j).mo Then gosub unter
 if roMe(i).mi <> rMe(j).mi Then gosub unter
 if roMe(i).nm <> rMe(j).nm Then gosub unter
 if roMe(i).ab <> rMe(j).ab Then gosub unter
 if roMe(i).zn <> rMe(j).zn Then gosub unter
 if roMe(i).bBed <> rMe(j).bBed Then gosub unter
 if roMe(i).Bemerkung <> rMe(j).Bemerkung Then gosub unter
 if roMe(i).Grund <> rMe(j).Grund Then gosub unter
 if roMe(i).Stärke <> rMe(j).Stärke Then gosub unter
 if roMe(i).Einheit <> rMe(j).Einheit Then gosub unter
 if roMe(i).Form <> rMe(j).Form Then gosub unter
 if roMe(i).AbsPos <> rMe(j).AbsPos Then gosub unter
 if roMe(i).AktZeit <> rMe(j).AktZeit Then gosub unter
 if roMe(i).StByte <> rMe(j).StByte Then gosub unter
 Exit Function
unter:
 MeZUnt = MeZUnt + 1
 Return
End Function ' MeZUnt

Public Function medplanLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roMe(1)
 sql = "SELECT FID,Pat_ID,MPNr,ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm,ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit,StByte FROM `medplan` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roMe)
  roMe(akt).FID = rs!FID
  roMe(akt).Pat_ID = rs!Pat_ID
  roMe(akt).MPNr = rs!MPNr
  roMe(akt).ZeitPunkt = rs!ZeitPunkt
  roMe(akt).Datum = rs!Datum
  roMe(akt).Medikament = rs!Medikament
  roMe(akt).MedAnfang = rs!MedAnfang
  roMe(akt).Wirkstoff = rs!Wirkstoff
  roMe(akt).PZN = rs!PZN
  roMe(akt).FeldNr = rs!FeldNr
  roMe(akt).mo = rs!mo
  roMe(akt).mi = rs!mi
  roMe(akt).nm = rs!nm
  roMe(akt).ab = rs!ab
  roMe(akt).zn = rs!zn
  roMe(akt).bBed = rs!bBed
  roMe(akt).Bemerkung = rs!Bemerkung
  roMe(akt).Grund = rs!Grund
  roMe(akt).Stärke = rs!Stärke
  roMe(akt).Einheit = rs!Einheit
  roMe(akt).Form = rs!Form
  roMe(akt).AbsPos = rs!AbsPos
  roMe(akt).AktZeit = rs!AktZeit
  roMe(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roMe(UBound(roMe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' medplanLaden

Function medplanEinf
 Dim ibeg&, i&, j&
 If UBound(rMe) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rMe)
   If rMe(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roMe)
    If roMe(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roMe(i - 1)
   ReDim Preserve roMe(i + UBound(rMe) - ibeg)
   For j = ibeg To UBound(rMe)
    Call roMeZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rMe = roMe
End Function ' medplanEinf

Public Function medplanSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere medplan"
 if not Allepat then
  sql = "SELECT pat_id FROM `medplan` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `medplan` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `medplan` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,AbsPos,AktZeit," & _
     "StByte)  values"))
 For i = 1 to ubound(rMe)
'  rMe(i).AktZeit = now()
  rMe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rMe(i).FID, "," , rMe(i).Pat_ID, "," , rMe(i).MPNr, "," , DatFor_k(rMe(i).ZeitPunkt), "," , DatFor_k(rMe(i).Datum), ",'" , rMe(i).Medikament, "','" , rMe(i).MedAnfang, "','" , rMe(i).Wirkstoff, "'," ,  _
   rMe(i).PZN, "," , rMe(i).FeldNr, ",'" , rMe(i).mo, "','" , rMe(i).mi, "','" , rMe(i).nm, "','" , rMe(i).ab, "','" , rMe(i).zn, "'," , cstr(-(rMe(i).bBed<>0)) , ",'" , rMe(i).Bemerkung, "','" ,  _
   rMe(i).Grund, "','" , rMe(i).Stärke, "','" , rMe(i).Einheit, "','" , rMe(i).Form, "'," , rMe(i).AbsPos, "," , DatFor_k(rMe(i).AktZeit), "," , rMe(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rMe) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rMe) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(12)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rMe),i)
  If Len(rMe(k).Medikament) > maxi(0) Then maxi(0) = Len(rMe(k).Medikament)
  If Len(rMe(k).MedAnfang) > maxi(1) Then maxi(1) = Len(rMe(k).MedAnfang)
  If Len(rMe(k).Wirkstoff) > maxi(2) Then maxi(2) = Len(rMe(k).Wirkstoff)
  If Len(rMe(k).mo) > maxi(3) Then maxi(3) = Len(rMe(k).mo)
  If Len(rMe(k).mi) > maxi(4) Then maxi(4) = Len(rMe(k).mi)
  If Len(rMe(k).nm) > maxi(5) Then maxi(5) = Len(rMe(k).nm)
  If Len(rMe(k).ab) > maxi(6) Then maxi(6) = Len(rMe(k).ab)
  If Len(rMe(k).zn) > maxi(7) Then maxi(7) = Len(rMe(k).zn)
  If Len(rMe(k).Bemerkung) > maxi(8) Then maxi(8) = Len(rMe(k).Bemerkung)
  If Len(rMe(k).Grund) > maxi(9) Then maxi(9) = Len(rMe(k).Grund)
  If Len(rMe(k).Stärke) > maxi(10) Then maxi(10) = Len(rMe(k).Stärke)
  If Len(rMe(k).Einheit) > maxi(11) Then maxi(11) = Len(rMe(k).Einheit)
  If Len(rMe(k).Form) > maxi(12) Then maxi(12) = Len(rMe(k).Form)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "medplan", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "medplan", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' medplanSpeichern

Public Function roReZuw(i&, j&)
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
End Function ' roReZuw

Public Function ReZUnt%(i&, j&)
 if roRe(i).FID <> rRe(j).FID Then gosub unter
 if roRe(i).Pat_ID <> rRe(j).Pat_ID Then gosub unter
 if roRe(i).ZeitPunkt <> rRe(j).ZeitPunkt Then gosub unter
 if roRe(i).Rezept <> rRe(j).Rezept Then gosub unter
 if roRe(i).RKlnm <> rRe(j).RKlnm Then gosub unter
 if roRe(i).Rezeptklasse <> rRe(j).Rezeptklasse Then gosub unter
 if roRe(i).Rezklkurz <> rRe(j).Rezklkurz Then gosub unter
 if roRe(i).Rezkllang <> rRe(j).Rezkllang Then gosub unter
 if roRe(i).kbez <> rRe(j).kbez Then gosub unter
 if roRe(i).Medikament <> rRe(j).Medikament Then gosub unter
 if roRe(i).auti <> rRe(j).auti Then gosub unter
 if roRe(i).anzl <> rRe(j).anzl Then gosub unter
 if roRe(i).PZN <> rRe(j).PZN Then gosub unter
 if roRe(i).absPos <> rRe(j).absPos Then gosub unter
 if roRe(i).AktZeit <> rRe(j).AktZeit Then gosub unter
 if roRe(i).QS <> rRe(j).QS Then gosub unter
 if roRe(i).QT <> rRe(j).QT Then gosub unter
 if roRe(i).StByte <> rRe(j).StByte Then gosub unter
 if roRe(i).LANRid <> rRe(j).LANRid Then gosub unter
 if roRe(i).id <> rRe(j).id Then gosub unter
 Exit Function
unter:
 ReZUnt = ReZUnt + 1
 Return
End Function ' ReZUnt

Public Function rezepteintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roRe(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN,absPos,AktZeit,QS,QT,StByte,LANRid,id FROM `rezepteintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roRe)
  roRe(akt).FID = rs!FID
  roRe(akt).Pat_ID = rs!Pat_ID
  roRe(akt).ZeitPunkt = rs!ZeitPunkt
  roRe(akt).Rezept = rs!Rezept
  roRe(akt).RKlnm = rs!RKlnm
  roRe(akt).Rezeptklasse = rs!Rezeptklasse
  roRe(akt).Rezklkurz = rs!Rezklkurz
  roRe(akt).Rezkllang = rs!Rezkllang
  roRe(akt).kbez = rs!kbez
  roRe(akt).Medikament = rs!Medikament
  roRe(akt).auti = rs!auti
  roRe(akt).anzl = rs!anzl
  roRe(akt).PZN = rs!PZN
  roRe(akt).absPos = rs!absPos
  roRe(akt).AktZeit = rs!AktZeit
  roRe(akt).QS = rs!QS
  roRe(akt).QT = rs!QT
  roRe(akt).StByte = rs!StByte
  roRe(akt).LANRid = rs!LANRid
  roRe(akt).id = rs!id
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roRe(UBound(roRe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' rezepteintraegeLaden

Function rezepteintraegeEinf
 Dim ibeg&, i&, j&
 If UBound(rRe) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rRe)
   If rRe(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roRe)
    If roRe(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roRe(i - 1)
   ReDim Preserve roRe(i + UBound(rRe) - ibeg)
   For j = ibeg To UBound(rRe)
    Call roReZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rRe = roRe
End Function ' rezepteintraegeEinf

Public Function rezepteintraegeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere rezepteintraege"
 if not Allepat then
  sql = "SELECT pat_id FROM `rezepteintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `rezepteintraege` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rezepteintraege` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,Rezkllang,kbez,Medikament,auti,anzl,PZN," & _
     "absPos,AktZeit,QS,QT,StByte,LANRid)  values"))
 For i = 1 to ubound(rRe)
'  rRe(i).AktZeit = now()
  rRe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRe(i).FID, "," , rRe(i).Pat_ID, "," , DatFor_k(rRe(i).ZeitPunkt), ",'" , rRe(i).Rezept, "','" , rRe(i).RKlnm, "','" , rRe(i).Rezeptklasse, "','" , rRe(i).Rezklkurz, "','" , rRe(i).Rezkllang, "','" ,  _
   rRe(i).kbez, "','" , rRe(i).Medikament, "'," , rRe(i).auti, "," , rRe(i).anzl, ",'" , rRe(i).PZN, "'," , rRe(i).absPos, "," , DatFor_k(rRe(i).AktZeit), ",'" , rRe(i).QS, "','" ,  _
   rRe(i).QT, "'," , rRe(i).StByte, "," , rRe(i).LANRid, ")")
  If SammelInsert <> 0 AND i < ubound(rRe) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rRe) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(9)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rRe),i)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rezepteintraege", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rezepteintraege", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' rezepteintraegeSpeichern

Public Function roRrZuw(i&, j&)
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
End Function ' roRrZuw

Public Function RrZUnt%(i&, j&)
 if roRr(i).FID <> rRr(j).FID Then gosub unter
 if roRr(i).Pat_ID <> rRr(j).Pat_ID Then gosub unter
 if roRr(i).ZeitPunkt <> rRr(j).ZeitPunkt Then gosub unter
 if roRr(i).RR <> rRr(j).RR Then gosub unter
 if roRr(i).Puls <> rRr(j).Puls Then gosub unter
 if roRr(i).RRsyst <> rRr(j).RRsyst Then gosub unter
 if roRr(i).RRdiast <> rRr(j).RRdiast Then gosub unter
 if roRr(i).RRzahl <> rRr(j).RRzahl Then gosub unter
 if roRr(i).Quelle <> rRr(j).Quelle Then gosub unter
 if roRr(i).Bemerkung <> rRr(j).Bemerkung Then gosub unter
 if roRr(i).absPos <> rRr(j).absPos Then gosub unter
 if roRr(i).AktZeit <> rRr(j).AktZeit Then gosub unter
 if roRr(i).StByte <> rRr(j).StByte Then gosub unter
 Exit Function
unter:
 RrZUnt = RrZUnt + 1
 Return
End Function ' RrZUnt

Public Function rrLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roRr(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte FROM `rr` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roRr)
  roRr(akt).FID = rs!FID
  roRr(akt).Pat_ID = rs!Pat_ID
  roRr(akt).ZeitPunkt = rs!ZeitPunkt
  roRr(akt).RR = rs!RR
  roRr(akt).Puls = rs!Puls
  roRr(akt).RRsyst = rs!RRsyst
  roRr(akt).RRdiast = rs!RRdiast
  roRr(akt).RRzahl = rs!RRzahl
  roRr(akt).Quelle = rs!Quelle
  roRr(akt).Bemerkung = rs!Bemerkung
  roRr(akt).absPos = rs!absPos
  roRr(akt).AktZeit = rs!AktZeit
  roRr(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roRr(UBound(roRr) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' rrLaden

Function rrEinf
 Dim ibeg&, i&, j&
 If UBound(rRr) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rRr)
   If rRr(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roRr)
    If roRr(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roRr(i - 1)
   ReDim Preserve roRr(i + UBound(rRr) - ibeg)
   For j = ibeg To UBound(rRr)
    Call roRrZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rRr = roRr
End Function ' rrEinf

Public Function rrSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere rr"
 if not Allepat then
  sql = "SELECT pat_id FROM `rr` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `rr` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `rr` (FID,Pat_ID,ZeitPunkt," & _
     "RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos,AktZeit,StByte)          values"))
 For i = 1 to ubound(rRr)
'  rRr(i).AktZeit = now()
  rRr(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rRr(i).FID, "," , rRr(i).Pat_ID, "," , DatFor_k(rRr(i).ZeitPunkt), ",'" , rRr(i).RR, "'," , rRr(i).Puls, "," , rRr(i).RRsyst, "," , rRr(i).RRdiast, "," , rRr(i).RRzahl, ",'" , rRr(i).Quelle, "','" ,  _
   rRr(i).Bemerkung, "'," , rRr(i).absPos, "," , DatFor_k(rRr(i).AktZeit), "," , rRr(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rRr) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rRr) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(2)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rRr),i)
  If Len(rRr(k).RR) > maxi(0) Then maxi(0) = Len(rRr(k).RR)
  If Len(rRr(k).Quelle) > maxi(1) Then maxi(1) = Len(rRr(k).Quelle)
  If Len(rRr(k).Bemerkung) > maxi(2) Then maxi(2) = Len(rRr(k).Bemerkung)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "rr", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "rr", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rRr), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & Left$(rRr(k).RR, maxL)  & "'",true : rRr(k).RR = Left$(rRr(k).RR, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.Quelle: '" & rRr(k).Quelle & "' -> '" & Left$(rRr(k).Quelle, maxL)  & "'",true : rRr(k).Quelle = Left$(rRr(k).Quelle, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRr.Bemerkung: '" & rRr(k).Bemerkung & "' -> '" & Left$(rRr(k).Bemerkung, maxL)  & "'",true : rRr(k).Bemerkung = Left$(rRr(k).Bemerkung, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' rrSpeichern
Public Function kvnrueSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere kvnrue"
 if not Allepat then
  sql = "SELECT pat_id FROM `kvnrue` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `kvnrue` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
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
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rKv(i).Pat_ID, ",'" , rKv(i).KVNr, "'," , rKv(i).absPos, "," , DatFor_k(rKv(i).AktZeit), "," , rKv(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rKv) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rKv) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(0)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rKv),i)
  If Len(rKv(k).KVNr) > maxi(0) Then maxi(0) = Len(rKv(k).KVNr)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "kvnrue", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "kvnrue", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rKv), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & Left$(rKv(k).KVNr, maxL)  & "'",true : rKv(k).KVNr = Left$(rKv(k).KVNr, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' kvnrueSpeichern
Public Function unbekannte_kennungenSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
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
  If SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rUn(i).Kennung, "'," , rUn(i).absPos, "," , rUn(i).StByte, "," , rUn(i).Pat_id, ",'" , rUn(i).Inhalt, "')")
  If SammelInsert <> 0 AND i < ubound(rUn) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rUn) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,rUn1 + 1,i) to iif(SammelInsert<>0,ubound(rUn),i)
  If Len(rUn(k).Kennung) > maxi(0) Then maxi(0) = Len(rUn(k).Kennung)
  If Len(rUn(k).Inhalt) > maxi(1) Then maxi(1) = Len(rUn(k).Inhalt)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "unbekannte kennungen", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "unbekannte kennungen", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,rUn1 + 1, i) To IIf(SammelInsert <> 0,ubound(rUn), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & Left$(rUn(k).Kennung, maxL)  & "'",true : rUn(k).Kennung = Left$(rUn(k).Kennung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & Left$(rUn(k).Inhalt, maxL)  & "'",true : rUn(k).Inhalt = Left$(rUn(k).Inhalt, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in unbekannte_kennungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' unbekannte_kennungenSpeichern

Public Function roDmZuw(i&, j&)
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
End Function ' roDmZuw

Public Function DmZUnt%(i&, j&)
 if roDm(i).Abk <> rDm(j).Abk Then gosub unter
 if roDm(i).Art <> rDm(j).Art Then gosub unter
 if roDm(i).KarteiDatum <> rDm(j).KarteiDatum Then gosub unter
 if roDm(i).exportiert <> rDm(j).exportiert Then gosub unter
 if roDm(i).DokuDatum <> rDm(j).DokuDatum Then gosub unter
 if roDm(i).obvoll <> rDm(j).obvoll Then gosub unter
 if roDm(i).ok <> rDm(j).ok Then gosub unter
 if roDm(i).ausgedruckt <> rDm(j).ausgedruckt Then gosub unter
 if roDm(i).NachName <> rDm(j).NachName Then gosub unter
 if roDm(i).VorName <> rDm(j).VorName Then gosub unter
 if roDm(i).GebDat <> rDm(j).GebDat Then gosub unter
 if roDm(i).Pat_id <> rDm(j).Pat_id Then gosub unter
 if roDm(i).StByte <> rDm(j).StByte Then gosub unter
 if roDm(i).AktZeit <> rDm(j).AktZeit Then gosub unter
 if roDm(i).lanrid <> rDm(j).lanrid Then gosub unter
 if roDm(i).Zusatzdaten <> rDm(j).Zusatzdaten Then gosub unter
 Exit Function
unter:
 DmZUnt = DmZUnt + 1
 Return
End Function ' DmZUnt

Public Function dmpreiheLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDm(1)
 sql = "SELECT Abk,Art,KarteiDatum,exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte,AktZeit,lanrid,Zusatzdaten FROM `dmpreihe` WHERE Pat_ID=" & pid & " ORDER BY `Dokudatum`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roDm)
  roDm(akt).Abk = rs!Abk
  roDm(akt).Art = rs!Art
  roDm(akt).KarteiDatum = rs!KarteiDatum
  roDm(akt).exportiert = rs!exportiert
  roDm(akt).DokuDatum = rs!DokuDatum
  roDm(akt).obvoll = rs!obvoll
  roDm(akt).ok = rs!ok
  roDm(akt).ausgedruckt = rs!ausgedruckt
  roDm(akt).NachName = rs!NachName
  roDm(akt).VorName = rs!VorName
  roDm(akt).GebDat = rs!GebDat
  roDm(akt).Pat_id = rs!Pat_id
  roDm(akt).StByte = rs!StByte
  roDm(akt).AktZeit = rs!AktZeit
  roDm(akt).lanrid = rs!lanrid
  roDm(akt).Zusatzdaten = rs!Zusatzdaten
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roDm(UBound(roDm) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' dmpreiheLaden

Function dmpreiheEinf
 Dim ibeg&, i&, j&
 If UBound(rDm) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rDm)
   If rDm(i).Dokudatum >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roDm)
    If roDm(i).Dokudatum >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roDm(i - 1)
   ReDim Preserve roDm(i + UBound(rDm) - ibeg)
   For j = ibeg To UBound(rDm)
    Call roDmZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rDm = roDm
End Function ' dmpreiheEinf

Public Function dmpreiheSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere dmpreihe"
 if not Allepat then
  sql = "SELECT pat_id FROM `dmpreihe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `dmpreihe` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `dmpreihe` (Abk,Art,KarteiDatum," & _
     "exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id,StByte," & _
     "AktZeit,lanrid,Zusatzdaten)          values"))
 For i = 1 to ubound(rDm)
'  rDm(i).AktZeit = now()
  rDm(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rDm(i).Abk, "','" , rDm(i).Art, "'," , DatFor_k(rDm(i).KarteiDatum), "," , DatFor_k(rDm(i).exportiert), "," , DatFor_k(rDm(i).DokuDatum), "," , cstr(-(rDm(i).obvoll<>0)) , "," , cstr(-( _
   rDm(i).ok<>0)) , "," , cstr(-(rDm(i).ausgedruckt<>0)) , ",'" , rDm(i).NachName, "','" , rDm(i).VorName, "'," , DatFor_k(rDm(i).GebDat), "," , rDm(i).Pat_id, "," , rDm(i).StByte, "," , DatFor_k(rDm(i).AktZeit), "," ,  _
   rDm(i).lanrid, ",'" , rDm(i).Zusatzdaten, "')")
  If SammelInsert <> 0 AND i < ubound(rDm) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rDm) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(4)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDm),i)
  If Len(rDm(k).Abk) > maxi(0) Then maxi(0) = Len(rDm(k).Abk)
  If Len(rDm(k).Art) > maxi(1) Then maxi(1) = Len(rDm(k).Art)
  If Len(rDm(k).NachName) > maxi(2) Then maxi(2) = Len(rDm(k).NachName)
  If Len(rDm(k).VorName) > maxi(3) Then maxi(3) = Len(rDm(k).VorName)
  If Len(rDm(k).Zusatzdaten) > maxi(4) Then maxi(4) = Len(rDm(k).Zusatzdaten)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "dmpreihe", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "dmpreihe", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDm), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & Left$(rDm(k).Abk, maxL)  & "'",true : rDm(k).Abk = Left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).Art & "' -> '" & Left$(rDm(k).Art, maxL)  & "'",true : rDm(k).Art = Left$(rDm(k).Art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).NachName & "' -> '" & Left$(rDm(k).NachName, maxL)  & "'",true : rDm(k).NachName = Left$(rDm(k).NachName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).VorName & "' -> '" & Left$(rDm(k).VorName, maxL)  & "'",true : rDm(k).VorName = Left$(rDm(k).VorName, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDm.Zusatzdaten: '" & rDm(k).Zusatzdaten & "' -> '" & Left$(rDm(k).Zusatzdaten, maxL)  & "'",true : rDm(k).Zusatzdaten = Left$(rDm(k).Zusatzdaten, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' dmpreiheSpeichern

Public Function roDeZuw(i&, j&)
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
End Function ' roDeZuw

Public Function DeZUnt%(i&, j&)
 if roDe(i).id <> rDe(j).id Then gosub unter
 if roDe(i).IDS <> rDe(j).IDS Then gosub unter
 if roDe(i).Pat_ID <> rDe(j).Pat_ID Then gosub unter
 if roDe(i).erstZP <> rDe(j).erstZP Then gosub unter
 if roDe(i).exoL <> rDe(j).exoL Then gosub unter
 if roDe(i).hideT <> rDe(j).hideT Then gosub unter
 if roDe(i).iconPath <> rDe(j).iconPath Then gosub unter
 if roDe(i).noteBkColor <> rDe(j).noteBkColor Then gosub unter
 if roDe(i).noteFgColor <> rDe(j).noteFgColor Then gosub unter
 if roDe(i).positionBottom <> rDe(j).positionBottom Then gosub unter
 if roDe(i).positionLeft <> rDe(j).positionLeft Then gosub unter
 if roDe(i).positionRight <> rDe(j).positionRight Then gosub unter
 if roDe(i).positionTop <> rDe(j).positionTop Then gosub unter
 if roDe(i).showAsNote <> rDe(j).showAsNote Then gosub unter
 if roDe(i).syncInfoList <> rDe(j).syncInfoList Then gosub unter
 if roDe(i).titel <> rDe(j).titel Then gosub unter
 if roDe(i).toolTipText <> rDe(j).toolTipText Then gosub unter
 if roDe(i).verankert <> rDe(j).verankert Then gosub unter
 if roDe(i).absPos <> rDe(j).absPos Then gosub unter
 if roDe(i).AktZeit <> rDe(j).AktZeit Then gosub unter
 if roDe(i).StByte <> rDe(j).StByte Then gosub unter
 Exit Function
unter:
 DeZUnt = DeZUnt + 1
 Return
End Function ' DeZUnt

Public Function desktopLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roDe(1)
 sql = "SELECT id,IDS,Pat_ID,erstZP,exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote,syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte FROM `desktop` WHERE Pat_ID=" & pid & " ORDER BY `erstZP`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roDe)
  roDe(akt).id = rs!id
  roDe(akt).IDS = rs!IDS
  roDe(akt).Pat_ID = rs!Pat_ID
  roDe(akt).erstZP = rs!erstZP
  roDe(akt).exoL = rs!exoL
  roDe(akt).hideT = rs!hideT
  roDe(akt).iconPath = rs!iconPath
  roDe(akt).noteBkColor = rs!noteBkColor
  roDe(akt).noteFgColor = rs!noteFgColor
  roDe(akt).positionBottom = rs!positionBottom
  roDe(akt).positionLeft = rs!positionLeft
  roDe(akt).positionRight = rs!positionRight
  roDe(akt).positionTop = rs!positionTop
  roDe(akt).showAsNote = rs!showAsNote
  roDe(akt).syncInfoList = rs!syncInfoList
  roDe(akt).titel = rs!titel
  roDe(akt).toolTipText = rs!toolTipText
  roDe(akt).verankert = rs!verankert
  roDe(akt).absPos = rs!absPos
  roDe(akt).AktZeit = rs!AktZeit
  roDe(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roDe(UBound(roDe) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in desktopLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' desktopLaden

Function desktopEinf
 Dim ibeg&, i&, j&
 If UBound(rDe) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rDe)
   If rDe(i).erstZP >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roDe)
    If roDe(i).erstZP >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roDe(i - 1)
   ReDim Preserve roDe(i + UBound(rDe) - ibeg)
   For j = ibeg To UBound(rDe)
    Call roDeZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rDe = roDe
End Function ' desktopEinf

Public Function desktopSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere desktop"
 if not Allepat then
  sql = "SELECT pat_id FROM `desktop` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `desktop` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `desktop` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte)  values"))
 For i = 1 to ubound(rDe)
'  rDe(i).AktZeit = now()
  rDe(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rDe(i).IDS, "'," , rDe(i).Pat_ID, "," , DatFor_k(rDe(i).erstZP), ",'" , rDe(i).exoL, "'," , rDe(i).hideT, ",'" , rDe(i).iconPath, "'," , rDe(i).noteBkColor, "," , rDe(i).noteFgColor, "," ,  _
   rDe(i).positionBottom, "," , rDe(i).positionLeft, "," , rDe(i).positionRight, "," , rDe(i).positionTop, "," , rDe(i).showAsNote, ",'" , rDe(i).syncInfoList, "','" , rDe(i).titel, "','" ,  _
   rDe(i).toolTipText, "'," , rDe(i).verankert, "," , rDe(i).absPos, "," , DatFor_k(rDe(i).AktZeit), "," , rDe(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rDe) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rDe) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(5)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rDe),i)
  If Len(rDe(k).IDS) > maxi(0) Then maxi(0) = Len(rDe(k).IDS)
  If Len(rDe(k).exoL) > maxi(1) Then maxi(1) = Len(rDe(k).exoL)
  If Len(rDe(k).iconPath) > maxi(2) Then maxi(2) = Len(rDe(k).iconPath)
  If Len(rDe(k).syncInfoList) > maxi(3) Then maxi(3) = Len(rDe(k).syncInfoList)
  If Len(rDe(k).titel) > maxi(4) Then maxi(4) = Len(rDe(k).titel)
  If Len(rDe(k).toolTipText) > maxi(5) Then maxi(5) = Len(rDe(k).toolTipText)
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "desktop", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "desktop", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rDe), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDe.IDS: '" & rDe(k).IDS & "' -> '" & Left$(rDe(k).IDS, maxL)  & "'",true : rDe(k).IDS = Left$(rDe(k).IDS, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDe.exoL: '" & rDe(k).exoL & "' -> '" & Left$(rDe(k).exoL, maxL)  & "'",true : rDe(k).exoL = Left$(rDe(k).exoL, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDe.iconPath: '" & rDe(k).iconPath & "' -> '" & Left$(rDe(k).iconPath, maxL)  & "'",true : rDe(k).iconPath = Left$(rDe(k).iconPath, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDe.syncInfoList: '" & rDe(k).syncInfoList & "' -> '" & Left$(rDe(k).syncInfoList, maxL)  & "'",true : rDe(k).syncInfoList = Left$(rDe(k).syncInfoList, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDe.titel: '" & rDe(k).titel & "' -> '" & Left$(rDe(k).titel, maxL)  & "'",true : rDe(k).titel = Left$(rDe(k).titel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDe.toolTipText: '" & rDe(k).toolTipText & "' -> '" & Left$(rDe(k).toolTipText, maxL)  & "'",true : rDe(k).toolTipText = Left$(rDe(k).toolTipText, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in desktopSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' desktopSpeichern

Public Function roUsZuw(i&, j&)
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
End Function ' roUsZuw

Public Function UsZUnt%(i&, j&)
 if roUs(i).FID <> rUs(j).FID Then gosub unter
 if roUs(i).Pat_ID <> rUs(j).Pat_ID Then gosub unter
 if roUs(i).ZeitPunkt <> rUs(j).ZeitPunkt Then gosub unter
 if roUs(i).Art <> rUs(j).Art Then gosub unter
 if roUs(i).Spritzst <> rUs(j).Spritzst Then gosub unter
 if roUs(i).Fußbef_re <> rUs(j).Fußbef_re Then gosub unter
 if roUs(i).Fußbef_li <> rUs(j).Fußbef_li Then gosub unter
 if roUs(i).Hyperk_re <> rUs(j).Hyperk_re Then gosub unter
 if roUs(i).Hyperk_li <> rUs(j).Hyperk_li Then gosub unter
 if roUs(i).Ulcera_re <> rUs(j).Ulcera_re Then gosub unter
 if roUs(i).Ulcera_li <> rUs(j).Ulcera_li Then gosub unter
 if roUs(i).Kraft_Zh_re <> rUs(j).Kraft_Zh_re Then gosub unter
 if roUs(i).Kraft_Zh_li <> rUs(j).Kraft_Zh_li Then gosub unter
 if roUs(i).Kraft_Zb_re <> rUs(j).Kraft_Zb_re Then gosub unter
 if roUs(i).Kraft_Zb_li <> rUs(j).Kraft_Zb_li Then gosub unter
 if roUs(i).Kraft_Knie_re <> rUs(j).Kraft_Knie_re Then gosub unter
 if roUs(i).Kraft_Knie_li <> rUs(j).Kraft_Knie_li Then gosub unter
 if roUs(i).ASR_re <> rUs(j).ASR_re Then gosub unter
 if roUs(i).ASR_li <> rUs(j).ASR_li Then gosub unter
 if roUs(i).PSR_re <> rUs(j).PSR_re Then gosub unter
 if roUs(i).PSR_li <> rUs(j).PSR_li Then gosub unter
 if roUs(i).Oberfl_re <> rUs(j).Oberfl_re Then gosub unter
 if roUs(i).Oberfl_li <> rUs(j).Oberfl_li Then gosub unter
 if roUs(i).MF_re <> rUs(j).MF_re Then gosub unter
 if roUs(i).MF_li <> rUs(j).MF_li Then gosub unter
 if roUs(i).KW_re <> rUs(j).KW_re Then gosub unter
 if roUs(i).KW_li <> rUs(j).KW_li Then gosub unter
 if roUs(i).Vibr_IK_re <> rUs(j).Vibr_IK_re Then gosub unter
 if roUs(i).Vibr_IK_li <> rUs(j).Vibr_IK_li Then gosub unter
 if roUs(i).Vibr_GZ_re <> rUs(j).Vibr_GZ_re Then gosub unter
 if roUs(i).Vibr_GZ_li <> rUs(j).Vibr_GZ_li Then gosub unter
 if roUs(i).PulsL_re <> rUs(j).PulsL_re Then gosub unter
 if roUs(i).PulsL_li <> rUs(j).PulsL_li Then gosub unter
 if roUs(i).PulsKK_re <> rUs(j).PulsKK_re Then gosub unter
 if roUs(i).PulsKK_li <> rUs(j).PulsKK_li Then gosub unter
 if roUs(i).PulsAtp_re <> rUs(j).PulsAtp_re Then gosub unter
 if roUs(i).PulsAtp_li <> rUs(j).PulsAtp_li Then gosub unter
 if roUs(i).PulsAdp_re <> rUs(j).PulsAdp_re Then gosub unter
 if roUs(i).PulsAdp_li <> rUs(j).PulsAdp_li Then gosub unter
 if roUs(i).Mitarbeiter <> rUs(j).Mitarbeiter Then gosub unter
 if roUs(i).absPos <> rUs(j).absPos Then gosub unter
 if roUs(i).AktZeit <> rUs(j).AktZeit Then gosub unter
 if roUs(i).QS <> rUs(j).QS Then gosub unter
 if roUs(i).QT <> rUs(j).QT Then gosub unter
 if roUs(i).StByte <> rUs(j).StByte Then gosub unter
 if roUs(i).id <> rUs(j).id Then gosub unter
 Exit Function
unter:
 UsZUnt = UsZUnt + 1
 Return
End Function ' UsZUnt

Public Function usdmLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roUs(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li,Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li,MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li,PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS,QT,StByte,id FROM `usdm` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roUs)
  roUs(akt).FID = rs!FID
  roUs(akt).Pat_ID = rs!Pat_ID
  roUs(akt).ZeitPunkt = rs!ZeitPunkt
  roUs(akt).Art = rs!Art
  roUs(akt).Spritzst = rs!Spritzst
  roUs(akt).Fußbef_re = rs!Fußbef_re
  roUs(akt).Fußbef_li = rs!Fußbef_li
  roUs(akt).Hyperk_re = rs!Hyperk_re
  roUs(akt).Hyperk_li = rs!Hyperk_li
  roUs(akt).Ulcera_re = rs!Ulcera_re
  roUs(akt).Ulcera_li = rs!Ulcera_li
  roUs(akt).Kraft_Zh_re = rs!Kraft_Zh_re
  roUs(akt).Kraft_Zh_li = rs!Kraft_Zh_li
  roUs(akt).Kraft_Zb_re = rs!Kraft_Zb_re
  roUs(akt).Kraft_Zb_li = rs!Kraft_Zb_li
  roUs(akt).Kraft_Knie_re = rs!Kraft_Knie_re
  roUs(akt).Kraft_Knie_li = rs!Kraft_Knie_li
  roUs(akt).ASR_re = rs!ASR_re
  roUs(akt).ASR_li = rs!ASR_li
  roUs(akt).PSR_re = rs!PSR_re
  roUs(akt).PSR_li = rs!PSR_li
  roUs(akt).Oberfl_re = rs!Oberfl_re
  roUs(akt).Oberfl_li = rs!Oberfl_li
  roUs(akt).MF_re = rs!MF_re
  roUs(akt).MF_li = rs!MF_li
  roUs(akt).KW_re = rs!KW_re
  roUs(akt).KW_li = rs!KW_li
  roUs(akt).Vibr_IK_re = rs!Vibr_IK_re
  roUs(akt).Vibr_IK_li = rs!Vibr_IK_li
  roUs(akt).Vibr_GZ_re = rs!Vibr_GZ_re
  roUs(akt).Vibr_GZ_li = rs!Vibr_GZ_li
  roUs(akt).PulsL_re = rs!PulsL_re
  roUs(akt).PulsL_li = rs!PulsL_li
  roUs(akt).PulsKK_re = rs!PulsKK_re
  roUs(akt).PulsKK_li = rs!PulsKK_li
  roUs(akt).PulsAtp_re = rs!PulsAtp_re
  roUs(akt).PulsAtp_li = rs!PulsAtp_li
  roUs(akt).PulsAdp_re = rs!PulsAdp_re
  roUs(akt).PulsAdp_li = rs!PulsAdp_li
  roUs(akt).Mitarbeiter = rs!Mitarbeiter
  roUs(akt).absPos = rs!absPos
  roUs(akt).AktZeit = rs!AktZeit
  roUs(akt).QS = rs!QS
  roUs(akt).QT = rs!QT
  roUs(akt).StByte = rs!StByte
  roUs(akt).id = rs!id
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roUs(UBound(roUs) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdmLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' usdmLaden

Function usdmEinf
 Dim ibeg&, i&, j&
 If UBound(rUs) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rUs)
   If rUs(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roUs)
    If roUs(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roUs(i - 1)
   ReDim Preserve roUs(i + UBound(rUs) - ibeg)
   For j = ibeg To UBound(rUs)
    Call roUsZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rUs = roUs
End Function ' usdmEinf

Public Function usdmSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere usdm"
 if not Allepat then
  sql = "SELECT pat_id FROM `usdm` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `usdm` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `usdm` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte)             values"))
 For i = 1 to ubound(rUs)
'  rUs(i).AktZeit = now()
  rUs(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rUs(i).FID, "," , rUs(i).Pat_ID, "," , DatFor_k(rUs(i).ZeitPunkt), ",'" , rUs(i).Art, "','" , rUs(i).Spritzst, "','" , rUs(i).Fußbef_re, "','" , rUs(i).Fußbef_li, "','" , rUs(i).Hyperk_re, "','" ,  _
   rUs(i).Hyperk_li, "','" , rUs(i).Ulcera_re, "','" , rUs(i).Ulcera_li, "','" , rUs(i).Kraft_Zh_re, "','" , rUs(i).Kraft_Zh_li, "','" , rUs(i).Kraft_Zb_re, "','" , rUs(i).Kraft_Zb_li, "','" ,  _
   rUs(i).Kraft_Knie_re, "','" , rUs(i).Kraft_Knie_li, "','" , rUs(i).ASR_re, "','" , rUs(i).ASR_li, "','" , rUs(i).PSR_re, "','" , rUs(i).PSR_li, "','" , rUs(i).Oberfl_re, "','" , rUs(i).Oberfl_li, "','" ,  _
   rUs(i).MF_re, "','" , rUs(i).MF_li, "','" , rUs(i).KW_re, "','" , rUs(i).KW_li, "','" , rUs(i).Vibr_IK_re, "','" , rUs(i).Vibr_IK_li, "','" , rUs(i).Vibr_GZ_re, "','" ,  _
   rUs(i).Vibr_GZ_li, "','" , rUs(i).PulsL_re, "','" , rUs(i).PulsL_li, "','" , rUs(i).PulsKK_re, "','" , rUs(i).PulsKK_li, "','" , rUs(i).PulsAtp_re, "','" , rUs(i).PulsAtp_li, "','" , rUs(i).PulsAdp_re, "','" ,  _
   rUs(i).PulsAdp_li, "','" , rUs(i).Mitarbeiter, "'," , rUs(i).absPos, "," , DatFor_k(rUs(i).AktZeit), ",'" , rUs(i).QS, "','" , rUs(i).QT, "'," , rUs(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rUs) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rUs) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(38)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rUs),i)
  If Len(rUs(k).Art) > maxi(0) Then maxi(0) = Len(rUs(k).Art)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "usdm", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "usdm", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in usdmSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' usdmSpeichern

Public Function roFuZuw(i&, j&)
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
End Function ' roFuZuw

Public Function FuZUnt%(i&, j&)
 if roFu(i).FID <> rFu(j).FID Then gosub unter
 if roFu(i).Pat_ID <> rFu(j).Pat_ID Then gosub unter
 if roFu(i).ZeitPunkt <> rFu(j).ZeitPunkt Then gosub unter
 if roFu(i).Art <> rFu(j).Art Then gosub unter
 if roFu(i).Fußdeform <> rFu(j).Fußdeform Then gosub unter
 if roFu(i).Hyper_mEin <> rFu(j).Hyper_mEin Then gosub unter
 if roFu(i).Weiteres <> rFu(j).Weiteres Then gosub unter
 if roFu(i).Zn_Ulcus <> rFu(j).Zn_Ulcus Then gosub unter
 if roFu(i).Zn_Amput <> rFu(j).Zn_Amput Then gosub unter
 if roFu(i).Fuß_ang <> rFu(j).Fuß_ang Then gosub unter
 if roFu(i).Ulcera <> rFu(j).Ulcera Then gosub unter
 if roFu(i).Wundinfektion <> rFu(j).Wundinfektion Then gosub unter
 if roFu(i).nae_US <> rFu(j).nae_US Then gosub unter
 if roFu(i).Mitarbeiter <> rFu(j).Mitarbeiter Then gosub unter
 if roFu(i).absPos <> rFu(j).absPos Then gosub unter
 if roFu(i).AktZeit <> rFu(j).AktZeit Then gosub unter
 if roFu(i).QS <> rFu(j).QS Then gosub unter
 if roFu(i).QT <> rFu(j).QT Then gosub unter
 if roFu(i).StByte <> rFu(j).StByte Then gosub unter
 if roFu(i).id <> rFu(j).id Then gosub unter
 Exit Function
unter:
 FuZUnt = FuZUnt + 1
 Return
End Function ' FuZUnt

Public Function fussLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roFu(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US,Mitarbeiter,absPos,AktZeit,QS,QT,StByte,id FROM `fuss` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roFu)
  roFu(akt).FID = rs!FID
  roFu(akt).Pat_ID = rs!Pat_ID
  roFu(akt).ZeitPunkt = rs!ZeitPunkt
  roFu(akt).Art = rs!Art
  roFu(akt).Fußdeform = rs!Fußdeform
  roFu(akt).Hyper_mEin = rs!Hyper_mEin
  roFu(akt).Weiteres = rs!Weiteres
  roFu(akt).Zn_Ulcus = rs!Zn_Ulcus
  roFu(akt).Zn_Amput = rs!Zn_Amput
  roFu(akt).Fuß_ang = rs!Fuß_ang
  roFu(akt).Ulcera = rs!Ulcera
  roFu(akt).Wundinfektion = rs!Wundinfektion
  roFu(akt).nae_US = rs!nae_US
  roFu(akt).Mitarbeiter = rs!Mitarbeiter
  roFu(akt).absPos = rs!absPos
  roFu(akt).AktZeit = rs!AktZeit
  roFu(akt).QS = rs!QS
  roFu(akt).QT = rs!QT
  roFu(akt).StByte = rs!StByte
  roFu(akt).id = rs!id
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roFu(UBound(roFu) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fussLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' fussLaden

Function fussEinf
 Dim ibeg&, i&, j&
 If UBound(rFu) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rFu)
   If rFu(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roFu)
    If roFu(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roFu(i - 1)
   ReDim Preserve roFu(i + UBound(rFu) - ibeg)
   For j = ibeg To UBound(rFu)
    Call roFuZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rFu = roFu
End Function ' fussEinf

Public Function fussSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere fuss"
 if not Allepat then
  sql = "SELECT pat_id FROM `fuss` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `fuss` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `fuss` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte)           values"))
 For i = 1 to ubound(rFu)
'  rFu(i).AktZeit = now()
  rFu(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rFu(i).FID, "," , rFu(i).Pat_ID, "," , DatFor_k(rFu(i).ZeitPunkt), ",'" , rFu(i).Art, "','" , rFu(i).Fußdeform, "','" , rFu(i).Hyper_mEin, "','" , rFu(i).Weiteres, "','" , rFu(i).Zn_Ulcus, "','" ,  _
   rFu(i).Zn_Amput, "','" , rFu(i).Fuß_ang, "','" , rFu(i).Ulcera, "','" , rFu(i).Wundinfektion, "','" , rFu(i).nae_US, "','" , rFu(i).Mitarbeiter, "'," , rFu(i).absPos, "," , DatFor_k(rFu(i).AktZeit), ",'" ,  _
   rFu(i).QS, "','" , rFu(i).QT, "'," , rFu(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rFu) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rFu) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(12)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rFu),i)
  If Len(rFu(k).Art) > maxi(0) Then maxi(0) = Len(rFu(k).Art)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "fuss", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "fuss", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fussSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' fussSpeichern

Public Function roUlZuw(i&, j&)
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
End Function ' roUlZuw

Public Function UlZUnt%(i&, j&)
 if roUl(i).FID <> rUl(j).FID Then gosub unter
 if roUl(i).Pat_ID <> rUl(j).Pat_ID Then gosub unter
 if roUl(i).ZeitPunkt <> rUl(j).ZeitPunkt Then gosub unter
 if roUl(i).Lokalisation <> rUl(j).Lokalisation Then gosub unter
 if roUl(i).Seite <> rUl(j).Seite Then gosub unter
 if roUl(i).Größe <> rUl(j).Größe Then gosub unter
 if roUl(i).Beläge <> rUl(j).Beläge Then gosub unter
 if roUl(i).Exsudat <> rUl(j).Exsudat Then gosub unter
 if roUl(i).Geruch <> rUl(j).Geruch Then gosub unter
 if roUl(i).Wundrand <> rUl(j).Wundrand Then gosub unter
 if roUl(i).Wundumgebung <> rUl(j).Wundumgebung Then gosub unter
 if roUl(i).Temperatur <> rUl(j).Temperatur Then gosub unter
 if roUl(i).Fotodoku <> rUl(j).Fotodoku Then gosub unter
 if roUl(i).Wundversorgung <> rUl(j).Wundversorgung Then gosub unter
 if roUl(i).Mitarbeiter <> rUl(j).Mitarbeiter Then gosub unter
 if roUl(i).absPos <> rUl(j).absPos Then gosub unter
 if roUl(i).AktZeit <> rUl(j).AktZeit Then gosub unter
 if roUl(i).StByte <> rUl(j).StByte Then gosub unter
 Exit Function
unter:
 UlZUnt = UlZUnt + 1
 Return
End Function ' UlZUnt

Public Function ulcusLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roUl(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku,Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte FROM `ulcus` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roUl)
  roUl(akt).FID = rs!FID
  roUl(akt).Pat_ID = rs!Pat_ID
  roUl(akt).ZeitPunkt = rs!ZeitPunkt
  roUl(akt).Lokalisation = rs!Lokalisation
  roUl(akt).Seite = rs!Seite
  roUl(akt).Größe = rs!Größe
  roUl(akt).Beläge = rs!Beläge
  roUl(akt).Exsudat = rs!Exsudat
  roUl(akt).Geruch = rs!Geruch
  roUl(akt).Wundrand = rs!Wundrand
  roUl(akt).Wundumgebung = rs!Wundumgebung
  roUl(akt).Temperatur = rs!Temperatur
  roUl(akt).Fotodoku = rs!Fotodoku
  roUl(akt).Wundversorgung = rs!Wundversorgung
  roUl(akt).Mitarbeiter = rs!Mitarbeiter
  roUl(akt).absPos = rs!absPos
  roUl(akt).AktZeit = rs!AktZeit
  roUl(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roUl(UBound(roUl) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ulcusLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' ulcusLaden

Function ulcusEinf
 Dim ibeg&, i&, j&
 If UBound(rUl) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rUl)
   If rUl(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roUl)
    If roUl(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roUl(i - 1)
   ReDim Preserve roUl(i + UBound(rUl) - ibeg)
   For j = ibeg To UBound(rUl)
    Call roUlZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rUl = roUl
End Function ' ulcusEinf

Public Function ulcusSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere ulcus"
 if not Allepat then
  sql = "SELECT pat_id FROM `ulcus` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `ulcus` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `ulcus` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte)  values"))
 For i = 1 to ubound(rUl)
'  rUl(i).AktZeit = now()
  rUl(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rUl(i).FID, "," , rUl(i).Pat_ID, "," , DatFor_k(rUl(i).ZeitPunkt), ",'" , rUl(i).Lokalisation, "','" , rUl(i).Seite, "','" , rUl(i).Größe, "','" , rUl(i).Beläge, "','" , rUl(i).Exsudat, "','" ,  _
   rUl(i).Geruch, "','" , rUl(i).Wundrand, "','" , rUl(i).Wundumgebung, "','" , rUl(i).Temperatur, "','" , rUl(i).Fotodoku, "','" , rUl(i).Wundversorgung, "','" , rUl(i).Mitarbeiter, "'," ,  _
   rUl(i).absPos, "," , DatFor_k(rUl(i).AktZeit), "," , rUl(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rUl) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rUl) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(11)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rUl),i)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "ulcus", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "ulcus", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ulcusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' ulcusSpeichern

Public Function roVkZuw(i&, j&)
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
End Function ' roVkZuw

Public Function VkZUnt%(i&, j&)
 if roVk(i).FID <> rVk(j).FID Then gosub unter
 if roVk(i).Pat_ID <> rVk(j).Pat_ID Then gosub unter
 if roVk(i).ZeitPunkt <> rVk(j).ZeitPunkt Then gosub unter
 if roVk(i).Wohlempfinden <> rVk(j).Wohlempfinden Then gosub unter
 if roVk(i).Saettigung <> rVk(j).Saettigung Then gosub unter
 if roVk(i).Zielwerterreichung <> rVk(j).Zielwerterreichung Then gosub unter
 if roVk(i).Ketonkörper <> rVk(j).Ketonkörper Then gosub unter
 if roVk(i).Gynaekologenbefund <> rVk(j).Gynaekologenbefund Then gosub unter
 if roVk(i).Gewichtsentwicklung <> rVk(j).Gewichtsentwicklung Then gosub unter
 if roVk(i).HbA1c <> rVk(j).HbA1c Then gosub unter
 if roVk(i).Bewegung <> rVk(j).Bewegung Then gosub unter
 if roVk(i).Minuten <> rVk(j).Minuten Then gosub unter
 if roVk(i).Blutdruck <> rVk(j).Blutdruck Then gosub unter
 if roVk(i).Puls <> rVk(j).Puls Then gosub unter
 if roVk(i).Mitarbeiter <> rVk(j).Mitarbeiter Then gosub unter
 if roVk(i).absPos <> rVk(j).absPos Then gosub unter
 if roVk(i).AktZeit <> rVk(j).AktZeit Then gosub unter
 if roVk(i).StByte <> rVk(j).StByte Then gosub unter
 Exit Function
unter:
 VkZUnt = VkZUnt + 1
 Return
End Function ' VkZUnt

Public Function vkgdLaden()
 Dim pid$, rs As New Recordset, akt&
 ON Error GoTo fehler
 pid = rNa(0).Pat_id
 ReDim roVk(1)
 sql = "SELECT FID,Pat_ID,ZeitPunkt,Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck,Puls,Mitarbeiter,absPos,AktZeit,StByte FROM `vkgd` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  akt = UBound(roVk)
  roVk(akt).FID = rs!FID
  roVk(akt).Pat_ID = rs!Pat_ID
  roVk(akt).ZeitPunkt = rs!ZeitPunkt
  roVk(akt).Wohlempfinden = rs!Wohlempfinden
  roVk(akt).Saettigung = rs!Saettigung
  roVk(akt).Zielwerterreichung = rs!Zielwerterreichung
  roVk(akt).Ketonkörper = rs!Ketonkörper
  roVk(akt).Gynaekologenbefund = rs!Gynaekologenbefund
  roVk(akt).Gewichtsentwicklung = rs!Gewichtsentwicklung
  roVk(akt).HbA1c = rs!HbA1c
  roVk(akt).Bewegung = rs!Bewegung
  roVk(akt).Minuten = rs!Minuten
  roVk(akt).Blutdruck = rs!Blutdruck
  roVk(akt).Puls = rs!Puls
  roVk(akt).Mitarbeiter = rs!Mitarbeiter
  roVk(akt).absPos = rs!absPos
  roVk(akt).AktZeit = rs!AktZeit
  roVk(akt).StByte = rs!StByte
  rs.MoveNext
  If Not rs.EOF Then ReDim Preserve roVk(UBound(roVk) + 1)
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = currentDB.Name
 #Else
 AnwPfad = App.Path
 #End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in vkgdLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 END SELECT
End Function ' vkgdLaden

Function vkgdEinf
 Dim ibeg&, i&, j&
 If UBound(rVk) > 0 Then
  ibeg = 0
  For i = 1 To UBound(rVk)
   If rVk(i).ZeitPunkt >= qbeg Then
    ibeg = i
    Exit For
   End If
  Next i
  If ibeg <> 0 Then
   For i = 0 To UBound(roVk)
    If roVk(i).ZeitPunkt >= qbeg Then
     Exit For
    End If
   Next i
   ReDim Preserve roVk(i - 1)
   ReDim Preserve roVk(i + UBound(rVk) - ibeg)
   For j = ibeg To UBound(rVk)
    Call roVkZuw(i + j - ibeg, j)
   Next j
  End If
 End If
 rVk = roVk
End Function ' vkgdEinf

Public Function vkgdSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere vkgd"
 if not Allepat then
  sql = "SELECT pat_id FROM `vkgd` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   set rs = nothing
   sql = "DELETE FROM `vkgd` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call DBCn.Execute(sql)
  End If
 End if ' not allePat
' sql0 = " insert " & sqlignore &  "into `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `vkgd` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte)            values"))
 For i = 1 to ubound(rVk)
'  rVk(i).AktZeit = now()
  rVk(i).StByte = CStr(AktByte)
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rVk(i).FID, "," , rVk(i).Pat_ID, "," , DatFor_k(rVk(i).ZeitPunkt), ",'" , rVk(i).Wohlempfinden, "','" , rVk(i).Saettigung, "','" , rVk(i).Zielwerterreichung, "','" , rVk(i).Ketonkörper, "','" ,  _
   rVk(i).Gynaekologenbefund, "','" , rVk(i).Gewichtsentwicklung, "','" , rVk(i).HbA1c, "','" , rVk(i).Bewegung, "','" , rVk(i).Minuten, "','" , rVk(i).Blutdruck, "','" , rVk(i).Puls, "','" ,  _
   rVk(i).Mitarbeiter, "'," , rVk(i).absPos, "," , DatFor_k(rVk(i).AktZeit), "," , rVk(i).StByte, ")")
  If SammelInsert <> 0 AND i < ubound(rVk) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rVk) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(11)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rVk),i)
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
 next k
 if obTrans <>0 then DBCn.CommitTrans: obtrans = 0
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "vkgd", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "vkgd", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in vkgdSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' vkgdSpeichern
Public Function laborxsaetzeSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxsaetze"
' sql0 = " insert " & sqlignore &  "into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxsaetze` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             values"))
 For i = 0 to ubound(rLs)
'  rLs(i).AktZeit = now()
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLs(i).DatID, ",'" , rLs(i).Satzart, "','" , rLs(i).Satzlänge, "','" , rLs(i).SatzlängeSchluss, "','" , rLs(i).VersionSatzb, "','" , rLs(i).Arztnr, "','" , rLs(i).Arztname, "','" ,  _
   rLs(i).StraßePraxis, "','" , rLs(i).Arzt, "','" , rLs(i).LANR, "','" , rLs(i).PLZPraxis, "','" , rLs(i).OrtPraxis, "','" , rLs(i).Labor, "','" , rLs(i).StraßeLabor, "','" , rLs(i).PLZLabor, "','" ,  _
   rLs(i).OrtLabor, "','" , rLs(i).KBVPrüfnr, "','" , rLs(i).Zeichensatz, "','" , rLs(i).Kundenarztnr, "','" , rLs(i).Erstellungsdatum, "','" , rLs(i).Gesamtlänge, "')")
  If SammelInsert <> 0 AND i < ubound(rLs) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLs) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(19)
 for k = iif(SammelInsert<>0,0,i) to iif(SammelInsert<>0,ubound(rLs),i)
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
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxsaetze", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxsaetzeSpeichern
Public Function laborxeingelSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxeingel"
' sql0 = " insert " & sqlignore &  "into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxeingel` (Pfad,Name,Zp," & _
     "fertig)  values"))
 For i = 1 to ubound(rLg)
'  rLg(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rLg(i).Pfad, "','" , rLg(i).Name, "'," , DatFor_k(rLg(i).Zp), "," , cstr(-(rLg(i).fertig<>0)) , ")")
  If SammelInsert <> 0 AND i < ubound(rLg) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLg) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(1)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLg),i)
  If Len(rLg(k).Pfad) > maxi(0) Then maxi(0) = Len(rLg(k).Pfad)
  If Len(rLg(k).Name) > maxi(1) Then maxi(1) = Len(rLg(k).Name)
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "laborxeingel", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxeingel", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLg), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & Left$(rLg(k).Pfad, maxL)  & "'",true : rLg(k).Pfad = Left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).Name & "' -> '" & Left$(rLg(k).Name, maxL)  & "'",true : rLg(k).Name = Left$(rLg(k).Name, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxeingelSpeichern
Public Function laborxusSpeichern(SammelInsert%, BezfSp%, j&)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
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
  If SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLu(i).DatID, "," , rLu(i).SatzID, ",'" , rLu(i).Satzart, "','" , rLu(i).Satzlänge, "','" , rLu(i).Auftragsnummer, "','" , rLu(i).Auftragsschlüssel, "'," , DatFor_k(rLu(i).Eingang), ",'" ,  _
   rLu(i).Berichtsdatum, "'," , rLu(i).Pat_id, ",'" , rLu(i).Nachname, "','" , rLu(i).Vorname, "','" , rLu(i).GebDat, "','" , rLu(i).Titel, "','" , rLu(i).NVorsatz, "','" , rLu(i).BefArt, "','" ,  _
   rLu(i).Abrechnungstyp, "','" , rLu(i).GebüOrd, "','" , rLu(i).Auftraggeber, "','" , rLu(i).Patienteninformation, "','" , rLu(i).Geschlecht, "','" , rLu(i).AuftrHinw, "','" , rLu(i).Pat_idUrsp, "','" ,  _
   rLu(i).Pat_idErwVNG, "','" , rLu(i).Pat_idErwVN, "','" , rLu(i).Pat_idErwG, "','" , rLu(i).Pat_idErwGB, "','" , rLu(i).Pat_idErwGL, "','" , rLu(i).Pat_idLaborNeu, "'," , DatFor_k(rLu(i).ZeitpunktLaborneu), "," ,  _
   rLu(i).ZdüP, "," , rLu(i).ZdiP, ",'" , rLu(i).LWerte, "'," , DatFor_k(rLu(i).verglichen), "," , rLu(i).AfN, ")")
  If SammelInsert <> 0 AND i < j Then csql.Append ","
  If SammelInsert = 0 Or i = j Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(24)
 for k = iif(SammelInsert<>0,j,i) to iif(SammelInsert<>0,j,i)
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
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxus", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxusSpeichern
Public Function laborxbaktSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxbakt"
' sql0 = " insert " & sqlignore &  "into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxbakt` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  values"))
 For i = 1 to ubound(rLo)
'  rLo(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLo(i).RefNr, ",'" , rLo(i).Verf, "','" , rLo(i).KuQu, "','" , rLo(i).Quelle, "','" , rLo(i).QSpez, "'," , DatFor_k(rLo(i).AbnDat), ",'" , rLo(i).Kommentar, "','" , rLo(i).Erklärung, "','" ,  _
   rLo(i).Keimzahl, "')")
  If SammelInsert <> 0 AND i < ubound(rLo) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLo) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(6)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLo),i)
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
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxbakt", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLo), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & Left$(rLo(k).Verf, maxL)  & "'",true : rLo(k).Verf = Left$(rLo(k).Verf, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & Left$(rLo(k).KuQu, maxL)  & "'",true : rLo(k).KuQu = Left$(rLo(k).KuQu, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & Left$(rLo(k).Quelle, maxL)  & "'",true : rLo(k).Quelle = Left$(rLo(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & Left$(rLo(k).QSpez, maxL)  & "'",true : rLo(k).QSpez = Left$(rLo(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & Left$(rLo(k).Kommentar, maxL)  & "'",true : rLo(k).Kommentar = Left$(rLo(k).Kommentar, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & Left$(rLo(k).Erklärung, maxL)  & "'",true : rLo(k).Erklärung = Left$(rLo(k).Erklärung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & Left$(rLo(k).Keimzahl, maxL)  & "'",true : rLo(k).Keimzahl = Left$(rLo(k).Keimzahl, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxbaktSpeichern
Public Function laborxwertSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxwert"
' sql0 = " insert " & sqlignore &  "into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxwert` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    values"))
 For i = 1 to ubound(rLw)
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
     Dim rsdop AS New ADODB.Recordset
     Set rsdop = Nothing
     rsdop.Open "SELECT 0 FROM `laborxwert` WHERE `RefNr` = " & rLw(i).RefNr & " AND `Abkü` = '" & rLw(i).Abkü & "' AND `Langname` = '" & rLw(i).Langname & "' AND `Quelle` = '" &  _
   rLw(i).Quelle & "' AND `QSpez` = '" & rLw(i).QSpez & "' AND `AbnDat` = " & DatFor_k(rLw(i).AbnDat) & " AND `Wert` = '" &  _
   rLw(i).Wert & "' AND `Einheit` = '" & rLw(i).Einheit & "' AND `Grenzwerti` = '" & rLw(i).Grenzwerti & "' AND `Kommentar` = '" & rLw(i).Kommentar & "' AND `Teststatus` = '" &  _
   rLw(i).Teststatus & "' AND `Erklärung` = '" & rLw(i).Erklärung & "' AND `AuftrHinw` = '" & rLw(i).AuftrHinw & "' AND `nbid` = " &  _
   rLw(i).nbid & "", DBCn, adOpenStatic, adLockReadOnly
     If Not rsdop.EOF Then GoTo nexti
    Next j
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLw(i).RefNr, ",'" , rLw(i).Abkü, "','" , rLw(i).Langname, "','" , rLw(i).Quelle, "','" , rLw(i).QSpez, "'," , DatFor_k(rLw(i).AbnDat), ",'" , rLw(i).Wert, "','" , rLw(i).Einheit, "','" ,  _
   rLw(i).Grenzwerti, "','" , rLw(i).Kommentar, "','" , rLw(i).Teststatus, "','" , rLw(i).Erklärung, "','" , rLw(i).AuftrHinw, "'," , rLw(i).nbid, ")")
  If SammelInsert <> 0 AND i < ubound(rLw) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLw) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(10)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLw),i)
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
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxwert", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxwertSpeichern
Public Function laborxleistSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere laborxleist"
' sql0 = " insert " & sqlignore &  "into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `laborxleist` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   values"))
 For i = 1 to ubound(rLL)
'  rLL(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("(" , rLL(i).RefNr, ",'" , rLL(i).Abkü, "','" , rLL(i).Verf, "','" , rLL(i).EBM, "','" , rLL(i).goä, "','" , rLL(i).Anzahl, "','" , rLL(i).abrd, "')")
  If SammelInsert <> 0 AND i < ubound(rLL) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLL) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(5)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLL),i)
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
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "laborxleist", rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0,1, i) To IIf(SammelInsert <> 0,ubound(rLL), i)
      SELECT CASE m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & Left$(rLL(k).Abkü, maxL)  & "'",true : rLL(k).Abkü = Left$(rLL(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & Left$(rLL(k).Verf, maxL)  & "'",true : rLL(k).Verf = Left$(rLL(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & Left$(rLL(k).EBM, maxL)  & "'",true : rLL(k).EBM = Left$(rLL(k).EBM, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & Left$(rLL(k).goä, maxL)  & "'",true : rLL(k).goä = Left$(rLL(k).goä, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & Left$(rLL(k).Anzahl, maxL)  & "'",true : rLL(k).Anzahl = Left$(rLL(k).Anzahl, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.abrd: '" & rLL(k).abrd & "' -> '" & Left$(rLL(k).abrd, maxL)  & "'",true : rLL(k).abrd = Left$(rLL(k).abrd, maxL)
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' laborxleistSpeichern
Public Function liuezSpeichern(SammelInsert%, BezfSp%)
 Dim i&, rAF&, Pid$, m% ',sql0$
 Dim csql0 AS New CString, csql AS New CString
 Dim rs AS new ADODB.recordset
 T1 = Timer
 ON error resume next
 Pid = rNa(0).Pat_id
 ON Error GoTo fehler
 syscmd 4, pid & ": Speichere liuez"
' sql0 = " insert " & sqlignore &  "into `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel) values
 Call csql0.AppVar(Array(" Insert ", sqlIgnore, " Into `liuez` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,aktdat,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel)           values"))
 For i = 1 to ubound(rLi)
'  rLi(i).AktZeit = now()
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If
  csql.AppVar Array("('" , rLi(i).name, "','" , rLi(i).vorname, "','" , rLi(i).titelt, "','" , rLi(i).fachgruppe, "','" , rLi(i).strasse, "','" , rLi(i).plz, "','" , rLi(i).ort, "','" , rLi(i).telefon, "','" ,  _
   rLi(i).fax, "','" , rLi(i).kvnr, "'," , DatFor_k(rLi(i).aktdat), ",'" , rLi(i).überschrift, "','" , rLi(i).dbnr, "','" , rLi(i).bstelle, "','" , rLi(i).anrede, "','" , rLi(i).tel1, "','" ,  _
   rLi(i).tel2, "','" , rLi(i).tel3, "','" , rLi(i).tel4, "','" , rLi(i).fax1, "','" , rLi(i).fax2, "','" , rLi(i).fax3, "','" , rLi(i).email, "','" , rLi(i).zulg, "','" , rLi(i).arzttyp, "','" ,  _
   rLi(i).gemmit, "','" , rLi(i).beme, "'," , rLi(i).dmpt2, "," , rLi(i).dmpt1, ",'" , rLi(i).geschlecht, "','" , rLi(i).titel, "')")
  If SammelInsert <> 0 AND i < ubound(rLi) Then csql.Append ","
  If SammelInsert = 0 Or i = ubound(rLi) Then
   Call DBCn.Execute(csql.value,rAf)', , adAsyncExecute)
   If obfor Then
    Call ForeignYes0
    Call ForeignYes1
   End If
   if sammelinsert = 0 then csql.m_len = 0
   If lese.obmysql AND obmitAlterTab Then
    Set rs = DBCn.Execute("SHOW WARNINGS")
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
If Err.Number = -2147217900 AND InStrB(ErrDescription, " Doppelter; Eintrag; ") <> 0 Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf Err.Number = -2147467259 AND InStrB(ErrDescription, "Daten zu lang") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add Or update a child row: a foreign key constraint fails
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescription)
 End If
 Resume
ElseIf Err.Number = -2147217833 Or InStrB(ErrDescription, "Daten zu lang") <> 0 Then
 Dim rsc AS ADODB.Recordset, maxi%(), k%
 redim maxi(27)
 for k = iif(SammelInsert<>0,1,i) to iif(SammelInsert<>0,ubound(rLi),i)
  If Len(rLi(k).name) > maxi(0) Then maxi(0) = Len(rLi(k).name)
  If Len(rLi(k).vorname) > maxi(1) Then maxi(1) = Len(rLi(k).vorname)
  If Len(rLi(k).titelt) > maxi(2) Then maxi(2) = Len(rLi(k).titelt)
  If Len(rLi(k).fachgruppe) > maxi(3) Then maxi(3) = Len(rLi(k).fachgruppe)
  If Len(rLi(k).strasse) > maxi(4) Then maxi(4) = Len(rLi(k).strasse)
  If Len(rLi(k).plz) > maxi(5) Then maxi(5) = Len(rLi(k).plz)
  If Len(rLi(k).ort) > maxi(6) Then maxi(6) = Len(rLi(k).ort)
  If Len(rLi(k).telefon) > maxi(7) Then maxi(7) = Len(rLi(k).telefon)
  If Len(rLi(k).fax) > maxi(8) Then maxi(8) = Len(rLi(k).fax)
  If Len(rLi(k).kvnr) > maxi(9) Then maxi(9) = Len(rLi(k).kvnr)
  If Len(rLi(k).überschrift) > maxi(10) Then maxi(10) = Len(rLi(k).überschrift)
  If Len(rLi(k).dbnr) > maxi(11) Then maxi(11) = Len(rLi(k).dbnr)
  If Len(rLi(k).bstelle) > maxi(12) Then maxi(12) = Len(rLi(k).bstelle)
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
  If Len(rLi(k).titel) > maxi(27) Then maxi(27) = Len(rLi(k).titel)
 next k
 nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, "liuez", Empty))
 m = 0
 Do While Not rsc.EOF
  SELECT CASE rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), "liuez", rsc)
    If maxL > 0 Then
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
      End SELECT
     Next
    elseif maxl < 0 then
     goto nochmal
    End If
    m = m + 1
  End SELECT
  If rsc.State = 0 Then Exit Do
  rsc.Move 1
 Loop
 Call ForeignNo0
 Call ForeignNo1
 if lese.obMysql then Resume next else resume
End If ' Err.Number = -2147217833 Then
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.Name
#Else
 AnwPfad = App.Path
#End If
 SELECT CASE MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in liuezSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' liuezSpeichern

Public Function tuLaden
 call namenLaden
 call faelleLaden
   if not lese.obmysql then
    if obTrans <> 0 then Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   end if
 call auLaden
 call briefeLaden
 call diagnosenLaden
 call dokumenteLaden
 call eintraegeLaden
   if not lese.obmysql then
    if obTrans <> 0 then Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   end if
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
End Function ' tuLaden

Public Function tuSpeichern(frm AS Lese, SI%, BfS%) ' frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern
 Dim rAf&
 ON Error GoTo fehler
 call namenSpeichern(SI, BfS)
 call faelleSpeichern(SI, BfS)
   if not lese.obmysql then
    if obTrans <> 0 then Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   end if
 call auSpeichern(SI, BfS)
 call briefeSpeichern(SI, BfS)
 call diagnosenSpeichern(SI, BfS)
 call dokumenteSpeichern(SI, BfS)
 call eintraegeSpeichern(SI, BfS)
 call forminhaltform_abkSpeichern(SI, BfS)
   if not lese.obmysql then
    if obTrans <> 0 then Call DBCn.CommitTrans: obtrans = 0
    Call DBCn.BeginTrans: obTrans = 1
   end if
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
 Call DBCn.Execute("UPDATE `namen` SET aktZeit = " & DatFor_k(rNa(0).AktZeit) & " WHERE pat_id = " & rNa(0).Pat_ID,rAf)
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
 ErrDescription = Err.Description
 If InStrB(ErrDescription, "'READ-COMMITTED'") <> 0 Then
  DBCn.Execute "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
  Resume
 End If
 SELECT CASE MsgBox("FNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(ISNULL(Err.Source), vNS, CStr(Err.Source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): Ende
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End SELECT
End Function ' tuSpeichern

