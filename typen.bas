Attribute VB_Name = "Typen"
Option Explicit
Public obForK%
Declare Function GetTickCount& Lib "kernel32" ()
Dim sql$, t1!, t2!, maxL%

Public Type namen
 Pat_ID As Long 'Pat_ID int '3000
 TM_Pat_ID As Long 'TM_Pat_ID int 'Pat_ID aus Turbomed, wenn abweichend
 lfdnr As Long 'lfdnr int 'laufende Patientennummer
 NVorsatz As String 'NVorsatz varchar '3100
 Nachname As String 'Nachname varchar '3101
 Vorname As String 'Vorname varchar '3102
 GebDat As Date 'GebDat datetime '3103
 KarGen As String 'KarGen varchar '3004 Kartentyp/-generation, 2019 gelöscht, bis jetzt 0 oder 2
 eGKSchVer As String 'eGKSchVer varchar '3006 eGK-Versichertenstammdaten-Schema-Version, bis jetzt 5.1.0 oder 5.2.0
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
 PFWsLC As String 'PFWsLC varchar '3124 Postfach Wohnsitzländercode, bisher immer leer
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 StByteA As Long 'StByteA int 'Ordnungsnummer der Datenübertragung, Beginn der Übertragung
 Cave As String 'Cave varchar '3654
 notiz As String 'notiz varchar 'Notiz-Feld in MO (aus TM importiert und frei eingebbar)
 info As String 'info varchar 'Info-Feld in MO (Eintrag ti)
 obChk As String 'obChk varchar 'obChroniker (Feld 3800)
 NZNr As Long 'NZNr int 'Notiz-Zeile Nr. (in welcher Zeile auf dem Desktop steht der unter 'Notiz' eingetragene Rest
 dmpklass As Long 'dmpklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier, 4 = DMP ausgeschrieben
 dmpbeg As Date 'dmpbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpkhkklass As Long 'dmpkhkklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpkhkbeg As Date 'dmpkhkbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpcopdklass As Long 'dmpcopdklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpcopdbeg As Date 'dmpcopdbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpabklass As Long 'dmpabklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpabbeg As Date 'dmpabbeg date 'Datum der aktuellen DMP-Klassifikation
 dmposteoklass As Long 'dmposteoklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmposteobeg As Date 'dmposteobeg date 'Datum der aktuellen DMP-Klassifikation
 dmpraklass As Long 'dmpraklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmprabeg As Date 'dmprabeg date 'Datum der aktuellen DMP-Klassifikation
 dmpchiklass As Long 'dmpchiklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpchibeg As Date 'dmpchibeg date 'Datum der aktuellen DMP-Klassifikation
 dmprsklass As Long 'dmprsklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmprsbeg As Date 'dmprsbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpadklass As Long 'dmpadklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpadbeg As Date 'dmpadbeg date 'Datum der aktuellen DMP-Klassifikation
 dmpdepklass As Long 'dmpdepklass int '1 = DMP nein, 2 = DMP HA, 3 = DMP hier
 dmpdepbeg As Date 'dmpdepbeg date 'Datum der aktuellen DMP-Klassifikation
 dakab As Date 'dakab date 'DAK-Einverständnis-Datum
 HzV As Long 'HzV int '1 = HzV-Eintrag im Notizfeld
 HzVbeg As Date 'HzVbeg date 'HzV ab
 Ds As Long 'DS int '1 = Datenschutzerklärung laut Notizfeld da
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
 SDatum As Date 'SDatum datetime 'Sterbedatum
 inaktiv As Byte 'inaktiv tinyint '1=tot, 2=verzogen, 4=Sonstige
 lAktTM As Date 'lAktTM datetime 'letzte Aktualisierung in Turbomed
 Mitarbeiter As Long 'Mitarbeiter int 'ob Pat. Mitarbeiter ist
 Swz As Integer 'Swz smallint 'Schwangerschaftszahl aus MO
 Gbz As Integer 'Gbz smallint 'Geburtenzahl aus MO
 Kiz As Integer 'Kiz smallint 'Kinderzahl aus MO
 ZdeK As Integer 'ZdeK smallint 'Zahl der eingetragenen Kinder (Schwangerschaften) aus MO
 obk As Integer 'obk bit 'ob Kothny zugeordnet
 obs As Integer 'obs bit 'ob Schade zugeornet
 obh As Integer 'obh bit 'ob Hammerschmidt zugeordnet
 antikoag As Integer 'antikoag bit 'Antikoagulation
 dmt1 As Integer 'dmt1 bit 'Diabetes mellitus Typ 1
 gdm As Integer 'gdm bit 'Gestationsdiabetes
 kdm As Integer 'kdm bit 'kein Diabetes
 cgm As Byte 'cgm tinyint '1=Libre Handy, 2=Libre Gerät, 3=Dexcom Handy, 4=Dexcom Gerät, 5=Simplera, 6=Eversense
 insdat As Date 'insdat datetime 'Zeitpunkt der Festlegung von insanw
 insanw As Byte 'insanw tinyint '1=Novopen, 2=Combo, 3=Insight, 4=Kaleido, 5=Medt.780, 6=Omnipod 5, 7=Dash, 8=TSlim, 9=Ypsopump
End Type

Public Type Faelle
 FID As Long 'FID int '
 Pat_ID As Long 'Pat_ID int '3000 Pat_id
 Quartal As String 'Quartal varchar '4101 Quartal der Ausstellung
 Nachname As String 'Nachname varchar '3101 Nachname
 Vorname As String 'Vorname varchar '3102 Vorname
 DtlOnlPfg As Date 'DtlOnlPfg datetime '3010 Datum der letzten Online-Prüfung
 ErgbdOnlP As Long 'ErgbdOnlP int '3011 Ergebnis der Onlineprüfung ('2')
 ErrorCode As Long 'ErrorCode int '3012 Error-Code (bis 5-stellig)
 PrüfZdFd As String 'PrüfZdFd varchar '3013 Prüfziffer des Fachdienstes (bis 88 Stellen)
 lfdnr As Long 'lfdnr int 'laufende Fallnummer
 TMFNr As String 'TMFNr varchar '4144 Fallnummer in Turbomed
 VKNr As String 'VKNr varchar '4104 VK-Nr.
 bPerG As String 'bPerG varchar '4131 besondere Personengruppe, ' ', 0, 4, 6, 7, 8, Sozialhilfempf., SVA, Asyl
 DMPKnZ As String 'DMPKnZ varchar '4132 DMP_Kennzeichnung, ' ', 0-9
 VschBeg As Date 'VschBeg datetime '4133 Versichternschutzbeginn
 VschEnd As Date 'VschEnd datetime '4110 Versichternschutzende
 KKasse_2 As String 'KKasse_2 varchar '4134 Krankenkasse aus Fall
 FaktPers As Single 'FaktPers float '4136 Faktor persönlich
 FaktTechn As Single 'FaktTechn float '4137 Faktor technisch
 FaktLabor As Single 'FaktLabor float '4138 Faktor Labor
 BhFB As Date 'BhFB datetime '4150 Behandlungsfall: Beginn
 BhFE1 As Date 'BhFE1 datetime '4151 Behandlungsfall: Ende (Musterwoman) / Wohl Ende des Behandlungsfallbeginnquartals
 BhFE2 As Date 'BhFE2 datetime '4152 Behandlungsfall: Ende (Musterwoman),bei offenem Fall 00000000, sonst z.B. 30092006 für 3/06
 UnfFlg As String 'UnfFlg varchar '4202 Unfall, Unfallfolgen nach 4152
 ausgst As Date 'ausgst datetime '4102 ('ausgestellt am')
 KtrAbrB As String 'KtrAbrB varchar '4106, Kostenträgerabrechnungsbereich (00 = Primärabrechnung (immer))
 AbrAr As String 'AbrAr varchar '4107, Abrechnungsart (1 = Primärkassen)
 lVorl As Date 'lVorl datetime '4109, letzte Vorlage
 KartBes As Byte 'KartBes tinyint '0=alles in Ordnung, 1=Schein fehlt, 2=Ersatzverfahren, 3=Schein fehlt
 IK As String 'IK varchar '4111 Krankenkassennummer (IK)
 KVKs As String 'KVKs varchar '4112 Versichertenstatus VK
 KVKserg As String 'KVKserg varchar '4113 Ost/West-Status VK
 Status As String 'Status varchar '6299 => Feld Status in verschiedenen Formularen
 Kasse As String 'Kasse varchar '6299 Kasse (aus Formularen)
 KID As Long 'KID int 'Bezug auf kassenliste.id
 GebOr As String 'GebOr varchar '4121, Gebührenordnung (1 = BMÄ, 2)
 AbrGb As String 'AbrGb varchar '4122, Abrechnungsgebiet (07 = Diabetes)
 PersKreis As String 'PersKreis varchar '4123 Personenkreis/Untersuchungskategorie
 SKtZusatz As String 'SKtZusatz varchar '4124 SKT-Zusatzangaben
 SktBem As String 'SktBem varchar '4126 SKT-Bemerkung
 letzteRegel As String 'letzteRegel varchar '4206, letzter Tag der Regel
 ÜwText As String 'ÜwText longtext '4209: Auftrags- / erläuternder Text zur Überweisung
 AkfLues As Byte 'AkfLues tinyint '4210, Ankreuzfeld Mutterschaftsvorsorge Lues-Suchreaktion
 AkfHAH As Byte 'AkfHAH tinyint '4211 Ankreuzfeld Muvo HAH
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
 KHNam As String 'KHNam varchar '4237 Krankenhausname
 statBehTage As Long 'statBehTage int '4238 Krankenhausaufenthalt
 SchGr As Double 'SchGr decimal '4239, Schein(unter)gruppe
 Weiterbeh As String 'Weiterbeh varchar '4243, Weiterbehandelnder
 KurAbb As Single 'KurAbb float '4266 Kurabbruch,  Datumsfeld, bisher nur bei Musterwoman
 VermiArt As String 'VermiArt varchar '4301 Vermittlungsart: 0=-, 1=TSS-Terminfall, 2=TSS-Akutfall, 3=HA-Vermittlungsfall, 4=offene Sprechstunde,  5=Neuopatient,  6=TSS-Routinetermin
 VermiCode As Double 'VermiCode bigint 'Vermittlungscode, 12-stellig, nicht Turbomed-BDT-Datei
 VermiDatum As Date 'VermiDatum date 'Tag der Terminvermittlung, nicht Turbomed-BDT-Datei
 VermiZusatz As String 'VermiZusatz varchar 'Zusatzinfo zur Terminvermittlung, nicht in Turbomed-BDT-Datei
 PGeb As String 'PGeb varchar '4401, Praxisgebühr
 PGebErg As String 'PGebErg varchar '4402, Array
 Mahnfrist As String 'Mahnfrist varchar '4403, Mahnfrist bis
 Unfallort As String 'Unfallort varchar '4505 Unfallort
 BeschAls As String 'BeschAls varchar '4506 Beschäftigt als
 BeschSeit As Date 'BeschSeit datetime '4506 Beschäftigt seit
 Unfallbetrieb As String 'Unfallbetrieb varchar '4509 Unfallbetrieb
 bHeilb As String 'bHeilb varchar '4570 Besondere Heilbehandlung, nur einamal 0
 GOÄKatNr As String 'GOÄKatNr varchar '4580 (1): Katalog-Nummer
 GOÄKatName As String 'GOÄKatName varchar '4580 (2): Privat-Abrechnungskatalog
 abrArzt As String 'abrArzt varchar '4585 abrechnender Arzt
 privVers As String 'privVers varchar '4586 private Versicherung
 AdNam As String 'AdNam varchar '4602(1) Name Rechnungsanschrift
 AdStr As String 'AdStr varchar '4602(2) Straße Rechnungsanschrift
 AdPlz As String 'AdPlz varchar '4602(3) PLZ Rechnungsanschrift
 AdOrt As String 'AdOrt varchar '4602(4) Ort Rechnungsanschrift
 ÜwBG As String 'ÜwBG varchar '4603 Überweiser BG
 BhFE As Date 'BhFE datetime '4604, Behandlungsfall: Ende, bei Privatpatienten
 s8000 As String 's8000 varchar '8000, Satzidentifikation
 s8100 As String 's8100 varchar '8100 Satzlänge
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 Fanf As Date 'Fanf datetime 'Fallanfang
 altQuart As String 'altQuart varchar '
 QAnf As Date 'QAnf datetime '4101, 5000, 6200 Quartalsanfang
 QEnd As Date 'QEnd datetime '4101, 5000, 6200 Quartalsende
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 ZnrMLes As String 'ZnrMLes varchar '4108 Zulassungsnummer mobiles Lesegerät
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
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Ersteller As String 'Ersteller varchar 'aus Medical Office
 Änderer As String 'Änderer varchar 'aus Medical Office
 art As String 'Art varchar 'aus Medical Office: E=Erst, F=Folge
 Beginn As String 'Beginn varchar '6285 1. Hälfte
 Ende As String 'Ende varchar '6285 2. Hälfte
 ICDs As String 'ICDs varchar '6286
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type briefe
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '
 Zeitpunkt As Date 'ZeitPunkt datetime '
 Pfad As String 'Pfad varchar '
 art As String 'Art varchar '
 name As String 'Name varchar '
 autor As String 'autor varchar 'Autor
 Quelldatum As Date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 Typ As String 'Typ varchar '
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
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
 Pat_ID As Long 'Pat_id int 'Bezug auf Anamneseblattt
 DiagDatum As Date 'DiagDatum datetime '5999 Datum, 6301 Uhrzeit
 DiagSicherheit As String 'DiagSicherheit varchar '3674, 6003 akut
 DiagText As String 'DiagText longtext '3650 Dauer, 6000 akut
 DiagSeite As String 'DiagSeite varchar '3675, 6004 akut
 DiagAttr As String 'DiagAttr varchar '6006 Diagnosenattribut (optionale Erläuterung)
 ICD As String 'ICD varchar '3673, 6001 akut
 obDauer As Byte 'obDauer tinyint 'ob Dauerdiagnose
 intBemerk As String 'intBemerk varchar '6009 interne Bemerkung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 AusnBegr As String 'AusnBegr varchar '6008 Ausnahmebegründung
 Dggel As Byte 'Dggel tinyint '6010 Falsch=dd-Eintrag, Wahr=bdd-Eintragurspr.: 6010 Diagnose gelöscht
 obKasse As Byte 'obKasse tinyint 'ob nach Kodierrichtlinien an Kasse zu übermitteln
 lKasse As Date 'lKasse datetime 'wann zuletzt Kassenübermittlung nach Kodierrichtlinien eingetragen (bd/bdd, f6010 Wahr)
 KFdFA As String 'KFdFA varchar 'Krankheitsfall d. Fallakte 6011 8.12.10: bisher nur """"""""TM#?""""""""
End Type

Public Type Dokumente
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '
 Zeitpunkt As Date 'ZeitPunkt datetime '
 DokPfad As String 'DokPfad varchar '
 DokArt As String 'DokArt varchar '
 DokName As String 'DokName varchar '
 Quelldatum As Date 'Quelldatum datetime 'Datum, auf das sich das Dokument bezieht
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 DokGroe As Long 'DokGroe int 'Dokument-Größe
 DokAenD As Date 'DokAenD datetime 'Dokument-letzte Änderung
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type eintraege
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 art As String 'Art varchar '6330
 Ersteller As String 'Ersteller varchar 'aus Medical Office
 Änderer As String 'Änderer varchar 'aus Medical Office
 Inhalt As String 'Inhalt varchar '8480
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
 inhNum As Double 'inhNum double 'Inhalt numerisch
End Type

Public Type formulare
 FormID As Long 'FormID int 'Primärindex formulare
 Form_Abk As String 'Form_Abk varchar '
 FormBez As String 'FormBez longtext '
 FormVorl As String 'FormVorl varchar '
 aktZeit As Date 'AktZeit datetime 'Zeitpunkt der Aktualisierung
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type forminhkopf
 Foid As Long 'FoID int 'Primär-Index für Formulare
 FID As Long 'FID int '-> faelle.fid
 Pat_ID As Long 'Pat_ID int '-> namen.pat_id
 Form_ID As Long 'Form_ID int '-> formulare.FormID
 Zeitpunkt As Date 'ZeitPunkt datetime '
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 Satzart As String 'Satzart varchar '8000
 Satzlänge As String 'Satzlänge varchar '8100
 lanrid As Long 'LANRid int '-> lanrpraxis.id
End Type

Public Type forminhfeld
 Foid As Long 'FoID int '-> forminhkopf.foid
 nr As Long 'Nr int '
 FeldNr As Integer 'FeldNr smallint '
 FeldVW As Long 'FeldVW int '->forminhaltfeld.FeldVW
 Feld As String ' nur Hilfsfeld, nicht in Datenbank
 FeldInhVW As Long 'FeldInhVW int '-> forminhaltfeldinh.FeldinhVW
 FeldInh As String ' nur Hilfsfeld, nicht in Datenbank
End Type

Public Type kheinweis
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 Ziel As String 'Ziel varchar '6291
 obNot As Integer 'obNot bit 'ob Notfall
 obBeleg As Integer 'obBeleg bit 'obBeleg
 Diagnose As String 'Diagnose longtext '6230
 Befund As String 'Befund longtext 'Untersuchungsergebnisse
 BisMas As String 'BisMas longtext 'Bisherige Maßnahmen
 FraStel As String 'FraStel longtext 'Fragestellung/Hinweise
 MitBef As String 'MitBef longtext 'Mitgegebene Befunde
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type lbanforderungen
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 AnfText As String 'AnfText longtext '6280
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Statusbyte
End Type

Public Type laborneu
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '
 FertigStGrad As String 'FertigStGrad varchar '8401
 Abkü As String 'Abkü varchar '8410
 LangtextVW As Long 'LangtextVW int '8411
 Langtext As String ' nur Hilfsfeld, nicht in Datenbank
 Wert As String 'Wert varchar '8420
 Einheit As String 'Einheit varchar '8421
 obpath As String 'obpath varchar '8462, ggf. zweite Zeile
 AnmkgVW As Long 'AnmkgVW int '8470
 Anmkg As String ' nur Hilfsfeld, nicht in Datenbank
 KommentarVW As Long 'KommentarVW int '8480
 Kommentar As String ' nur Hilfsfeld, nicht in Datenbank
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 RefNr As Long 'Refnr int 'Bezug auf LaborXUS
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 id As Long 'ID int '
 AbschlZlVW As Long 'AbschlZlVW int 'Abschlusszeile 8490
 AbschlZl As String ' nur Hilfsfeld, nicht in Datenbank
 NormberVW As Long 'NormberVW int 'Normbereich MO VW
 Normber As String ' nur Hilfsfeld, nicht in Datenbank
 uNm As String ' nur Hilfsfeld, nicht in Datenbank
 oNm As String ' nur Hilfsfeld, nicht in Datenbank
End Type

Public Type Leistungen
 id As Long 'id int 'eindeutige ID, hinzugefügt 26.3.11
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '5000 + 6201
 Leistung As String 'Leistung varchar '5001 Leistungsziffer
 ArtdUs As String 'ArtdUs varchar '5002 Art der Untersuchung
 LAnzl As String 'LAnzl varchar '5005 Anzahl
 LUhrz As String 'LUhrz varchar '5006 um Uhrzeit
 LfBegr As String 'LfBegr varchar '5009 freier Begründungstext
 Med As String 'Med varchar '5010 Medikament
 LOrgan As String 'LOrgan varchar '5015 Organ
 LArztBf As String 'LArztBf varchar '5016 Name des Arztes (Briefempfänger)
 DtlKbsV As String 'DtlKbsV varchar '5021 Datum letzte Krebsvorsorge
 LEntlDt As String 'LEntlDt varchar '5026 Entlassungsdatum
 Faktor As String 'Faktor varchar '5062 Multiplikator für GOÄ-Rechnung
 LBSNR As String 'LBSNR varchar '5098 Bestriebestättennummer des Ortes der Leistungserbringung 0000000000
 Charge As String 'Charge varchar '9999 Charge
 Lanr As String 'LANR varchar '5099 LANR
 letzVorg As Date 'letzVorg datetime '5101 letzter Vorgang
 Ausn As String 'Ausn mediumtext '3677 Ausnahme/Begründung für abweichendes Geschlecht
 beme As String 'Beme mediumtext '         Bemerkung
 absPos As Long 'absPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 Sachkbez As String 'Sachkbez varchar '5011 Sachkostenbezeichnung
 Sachkct As Long 'Sachkct int '5012 Sach-/Materialkosten in ct
 Zone As String 'Zone varchar '5018 Zone bei Besuchen
 Punkte As Double 'Punkte decimal '5061 aus Medical Office
 Lstgerbnr As Long 'Lstgerbnr int 'Lstgerbnr aus Medical Office
 Position As Long 'Position int 'Position aus Medical Office
 Eignung As Long 'Eignung int 'Eignung aus Medical Office
 Pruefzeit As Long 'Pruefzeit int 'Pruefzeit aus Medical Office
 Kalkzeit As Long 'Kalkzeit int 'Kalkzeit aus Medical Office
 BSNR As Long 'Bsnr int 'Bsnr aus Medical Office
 Abrechnungssperre As Long 'Abrechnungssperre int 'Abrechnungssperre aus Medical Office
 Medikanr As Long 'Medikanr int 'Medikanr aus Medical Office
 Laborkosten As Long 'Laborkosten int 'Laborkosten aus Medical Office
 Sachkostentyp As Long 'Sachkostentyp int 'Sachkostentyp aus Medical Office
 Usegebordtext As String 'Usegebordtext varchar 'aus Medical Office
 Arztliste As String 'Arztliste varchar 'aus Medical Office
 Ersteller As String 'Ersteller varchar 'aus Medical Office
 Änderer As String 'Änderer varchar 'aus Medical Office
 Kennzeichen_arzt_technik As String 'Kennzeichen_arzt_technik varchar 'aus Medical Office
 Auftragsschluessel As String 'Auftragsschluessel varchar 'aus Medical Office
 Auftragstext As String 'Auftragstext varchar 'aus Medical Office
End Type

Public Type medplan
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
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
 Bemerkung As String 'Bemerkung longtext 'Dosiertext, Hinweis und Bemerkung aus BMP
 Grund As String 'Grund varchar 'Grund aus BMP
 Stärke As String 'Stärke varchar 'Stärke aus BMP
 Einheit As String 'Einheit varchar '
 Form As String 'Form varchar 'Form aus BMP
 Menge As Integer 'Menge smallint 'Menge aus BMP
 Nutzer As String 'Nutzer varchar 'zu FNutzernr aus MO
 absPos As Long 'AbsPos int 'Zeile in der BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 ergaenzt As Integer 'ergaenzt bit 'PZN ergaenzt von pznbdt (bdtnachw)
 mpart As Byte 'mpart tinyint '1=alter MP, 2=bmp, 3=emp
End Type

Public Type rezepteintraege
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 Rezept As String 'Rezept varchar '6210, 3652(1), 6218(1)
 RKlnm As String 'RKlnm varchar 'Anfang des Rezeptklassennamens bei Langrezepten (HeilHilfsmittel, LangRezeptEintrag)
 Rezeptklasse As String 'Rezeptklasse varchar '6218(2) N = Sprechstundenbedarf, HI = Heil- und Hilfsmittel (Erklärung = 6218(1)
 Rezklkurz As String 'Rezklkurz varchar 'letztes Split-Feld, z.B. 'rp' oder 'lar'
 erez As Byte 'erez tinyint '1=eRezept, 0=keines
 ered As Byte 'ered tinyint '1=eRezept-Druck, 0=keiner
 Rezkllang As String 'Rezkllang varchar 'erstes Split-Feld, z.B. 'Sprechstundenbedarf': Langrezept, Sprechstundenbedarf, Heilmittel => später durch Rezklkurz ersetzen
 rea As String 'rea varchar 'Rezeptart: akt=aktives Medikament, ina=inaktiv, dau=Dauermed.,his=historisches Med.,hil=Hilfsmittel,hei=Heilmittel,lar=Langrezept
 kbez As String 'kbez varchar 'vorletztes Split-Feld, z.B. 'DTronAnthra'
 Medikament As String 'Medikament varchar '3652(2), 6218(4)
 auti As Long 'auti int 'aut-idem-Eintrag im Rz. (1=kein Ausschluß, 0=Ausschluß)
 obBTM As Byte 'obBTM tinyint 'BTM-Rezept (1=ja, 0=nein)
 noctu As Byte 'noctu tinyint '1=noctu angekreuzt, 0=nicht
 ersatzv As Byte 'ersatzv tinyint '1=Ersatzverordnung, 0=keine
 einzdr As Byte 'einzdr tinyint '1=Einzeldruck, 0=keiner
 freig As Byte 'freig tinyint '1=freigegeben, 0=nicht
 obst As Byte 'obst tinyint '1=storniert, 0=nicht
 anzl As Long 'anzl int 'Anzahl
 PZN As String 'PZN varchar '6210(2), 6218(3)
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Statusbyte
 lanrid As Long 'LANRid int 'Bezug auf lanrpraxis.id
 id As Long 'id int '
 FEintragsart As String 'FEintragsart varchar 'ltag.FEintragsart; 17=Hilfsmittel, 18=Physiotherapie etc.,  2029 = Diga
 Rezeptart As Byte 'Rezeptart tinyint 'Rezeptart in ltag.FDetails
End Type

Public Type RR
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 FormTitel As String 'FormTitel varchar '8409 Formulartitel
 art As String 'art varchar '6330
 RR As String 'RR longtext '6230
 Puls As Long 'Puls int 'Puls
 RRsyst As Integer 'RRsyst smallint '
 RRdiast As Integer 'RRdiast smallint '
 RRzahl As Integer 'RRzahl smallint '
 Quelle As String 'Quelle varchar 'Informationsquelle
 Bemerkung As String 'Bemerkung varchar 'Bemerkung
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type kvnrue
 lfdnr As Long 'lfdnr int '
 Pat_ID As Long 'Pat_ID int '
 KVNr As String 'KVNr varchar '
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Zeit der Aktualisuerung aus der BDT-Datei
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type unbek_kenn
 Kennung As String 'Kennung varchar '
 absPos As Long 'absPos int '
 StByte As Long 'StByte int '
 Pat_ID As Long 'Pat_id int 'zugehöriger Patient für spätere Ermittlungen
 Inhalt As String 'Inhalt varchar 'Inhalt Zeile zum Wiederauffinden
 Zeitpunkt As Date 'Zeitpunkt datetime '
End Type

Public Type dmpreihe
 Abk As String 'Abk varchar 'Abkürzung der DMP-Art
 DMPArt As Byte 'DMPArt tinyint '0=keine, 1=DM1, 2=DM2, 3=KHK, 4=COPD, 5=Asthma, 6=Brustkrebs.7=Osteoporose, 8=Rheumat.A.,9=Herzi, 10=Rückensz, 11=Adipositas, 12=Depression
 art As String 'Art varchar 'ED = Erstdoku, FD = Folgedoku
 KarteiDatum As Date 'KarteiDatum date 'Datum des Karteikarteneintrags der Dokumentation
 exportiert As Date 'exportiert datetime 'Datum des Exports, in Mo Feld 139
 DokuDatum As Date 'DokuDatum datetime 'Datum der Dokumentation, in Mo Feld 96
 obvoll As Integer 'obvoll bit 'ob vollständig
 Ok As Integer 'ok bit 'ob """"""""ok""""""""
 ausgedruckt As Integer 'ausgedruckt bit 'ob """"""""ausgedruckt""""""""
 Nachname As String 'NachName varchar '
 Vorname As String 'VorName varchar '
 GebDat As Date 'GebDat date '
 Pat_ID As Long 'Pat_id int '
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
 aktZeit As Date 'AktZeit datetime 'Aktualisierungzeit
 lanrid As Long 'lanrid int 'Bezug auf lanrpraxis.id
 Zusatzdaten As String 'Zusatzdaten varchar 'Zusatzdaten
 Druckdatum As Date 'Druckdatum datetime 'in MO Feld 137
 uDat As Date 'uDat datetime 'in MO Feld 75
 eGFR As Byte 'eGFR tinyint 'in Mo Feld 121, 1. Byte
End Type

Public Type desktop
 id As Long 'id int 'Primärschlüssel
 IDS As String 'IDS varchar 'id=
 Pat_ID As Long 'Pat_ID int '3000
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type usdm
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
End Type

Public Type fuss
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 QS As String 'QS varchar 'Quartal des Behandlungsfallbeginns sortiert
 QT As String 'QT varchar 'Quartal des Behandlungsfallbeginns
 StByte As Long 'StByte int 'Ordnungsnr. der Datenübertragung
 id As Long 'id int '
End Type

Public Type ulcus
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type vkgd
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type sws
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 FormTitel As String 'FormTitel varchar '8409 Formulartitel
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
 StByte As Long 'StByte int 'Ordnungsnummer der Datenübertragung
End Type

Public Type vopl
 FID As Long 'FID int 'Fall-Bezug
 Pat_ID As Long 'Pat_ID int '3000
 Zeitpunkt As Date 'ZeitPunkt datetime '6200 + 6201
 FormTitel As String 'FormTitel varchar '8409 Formulartitel
 Inhalt As String 'Inhalt longtext 'Inhalt der Verordnung
 absPos As Long 'absPos int 'Zeile in BDT-Datei
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
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
 Pat_ID As Long 'Pat_ID int '
 TM_Pat_ID As Long 'TM_Pat_id int '
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
 Lanr As String 'LANR varchar '
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
 aktZeit As Date 'aktzeit datetime '
End Type

Public Type Anamnesebogen
 Prim As Long 'Prim int 'Primärschlüssel
 Pat_ID As Long 'Pat_ID int '
 TM_Pat_ID As Long 'TM_Pat_ID int '
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
 aktZeit As Date 'AktZeit datetime 'Aktualisierungszeit
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

Type fzu
 falt As Long
 fneu As Long
End Type ' fzu

Public rNa() As namen
Public rFa() As Faelle
Public rAu() As au
Public rBr() As briefe
Public rDi() As Diagnosen
Public rDo() As Dokumente
Public rEi() As eintraege
Public rFo() As formulare ' kommt vor in: formulareSpeichern, doTabVorb, dolies
Public rFr() As forminhkopf
Public rFm() As forminhfeld
Public rKh() As kheinweis
Public rLb() As lbanforderungen
Public rLa() As laborneu
Public rLe() As Leistungen
Public rMe() As medplan
Public rRe() As rezepteintraege
Public rRr() As RR
Public rKv() As kvnrue
Public rUn() As unbek_kenn
Public rDm() As dmpreihe
Public rDe() As desktop
Public rUs() As usdm
Public rFu() As fuss
Public rUl() As ulcus
Public rVk() As vkgd
Public rSw() As sws
Public rVo() As vopl
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
Public roDo() As Dokumente
Public roEi() As eintraege
Public roFo() As formulare
Public roFr() As forminhkopf
Public roFm() As forminhfeld
Public roKh() As kheinweis
Public roLb() As lbanforderungen
Public roLa() As laborneu
Public roLe() As Leistungen
Public roMe() As medplan
Public roRe() As rezepteintraege
Public roRr() As RR
Public roKv() As kvnrue
Public roUn() As unbek_kenn
Public roDm() As dmpreihe
Public roDe() As desktop
Public roUs() As usdm
Public roFu() As fuss
Public roUl() As ulcus
Public roVk() As vkgd
Public roSw() As sws
Public roVo() As vopl
Public roLs() As laborxsaetze
Public roLg() As laborxeingel
Public roLu() As laborxus
Public roLo() As laborxbakt
Public roLw() As laborxwert
Public roLL() As laborxleist
Public roLi() As LIUEZ
Public roAna() As Anamnesebogen

' in Geslies(2x)
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
 ReDim rVo(0)
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
End Function ' LabInit

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

' in Pat_loeschen_Click, doPatvonMO
Public Sub LöschePat(pid&, Optional obAnzeig%, Optional ohneLabor%)
 Dim Tb, tbn, rAf&, ergeb$
 syscmd 4, "Lösche Pat. " & pid
 On Error GoTo fehler
 tbn = Array("namen", "faelle", "au", "briefe", "diagnosen", "dokumente", "eintraege", "forminhkopf", "kheinweis", "lbanforderungen", "leistungen", "medplan", "rezepteintraege", "rr", "kvnrue", "dmpreihe", "desktop", "usdm", "fuss", "ulcus", "vkgd", "sws", "vopl")
 If Not ohneLabor Then
  ReDim Preserve tbn(UBound(tbn) + 1)
  tbn(UBound(tbn)) = "laborneu"
 End If ' Not ohneLabor Then
 myEFrag "DELETE fif FROM forminhfeld fif LEFT JOIN forminhkopf fk USING (foid) WHERE pat_ID=" & pid, rAf
 For Each Tb In tbn
  myEFrag "DELETE FROM `" & Tb & "` WHERE PAT_ID = " & pid, rAf
  ergeb = ergeb & vbCrLf & rAf & " Sätze aus `" & Tb & "` gelöscht."
 Next
 If obAnzeig Then
  MsgBox ergeb
  Debug.Print ergeb
 End If ' obAnzeig Then
 Exit Sub
fehler:
 Dim AnwPfad$
 #If VBA6 Then
  AnwPfad = CurrentDb.name
 #Else
  AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in LöschePat/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Sub ' LöschePat

Public Function AllesLösch(frm As Lese)
 Dim ct&, rs As New ADODB.Recordset
 On Error GoTo fehler
 Call ForeignNo0
 Call ForeignNo1
 Call doEntleer(frm, "vopl")
 Call doEntleer(frm, "sws")
 Call doEntleer(frm, "vkgd")
 Call doEntleer(frm, "ulcus")
 Call doEntleer(frm, "fuss")
 Call doEntleer(frm, "usdm")
 Call doEntleer(frm, "desktop")
 Call doEntleer(frm, "dmpreihe")
 Call doEntleer(frm, "unbek_kenn")
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
 roNa(i).Pat_ID = rNa(j).Pat_ID
 roNa(i).TM_Pat_ID = rNa(j).TM_Pat_ID
 roNa(i).lfdnr = rNa(j).lfdnr
 roNa(i).NVorsatz = rNa(j).NVorsatz
 roNa(i).Nachname = rNa(j).Nachname
 roNa(i).Vorname = rNa(j).Vorname
 roNa(i).GebDat = rNa(j).GebDat
 roNa(i).KarGen = rNa(j).KarGen
 roNa(i).eGKSchVer = rNa(j).eGKSchVer
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
 roNa(i).PFWsLC = rNa(j).PFWsLC
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
 roNa(i).aktZeit = rNa(j).aktZeit
 roNa(i).absPos = rNa(j).absPos
 roNa(i).StByte = rNa(j).StByte
 roNa(i).StByteA = rNa(j).StByteA
 roNa(i).Cave = rNa(j).Cave
 roNa(i).notiz = rNa(j).notiz
 roNa(i).info = rNa(j).info
 roNa(i).obChk = rNa(j).obChk
 roNa(i).NZNr = rNa(j).NZNr
 roNa(i).dmpklass = rNa(j).dmpklass
 roNa(i).dmpbeg = rNa(j).dmpbeg
 roNa(i).dmpkhkklass = rNa(j).dmpkhkklass
 roNa(i).dmpkhkbeg = rNa(j).dmpkhkbeg
 roNa(i).dmpcopdklass = rNa(j).dmpcopdklass
 roNa(i).dmpcopdbeg = rNa(j).dmpcopdbeg
 roNa(i).dmpabklass = rNa(j).dmpabklass
 roNa(i).dmpabbeg = rNa(j).dmpabbeg
 roNa(i).dmposteoklass = rNa(j).dmposteoklass
 roNa(i).dmposteobeg = rNa(j).dmposteobeg
 roNa(i).dmpraklass = rNa(j).dmpraklass
 roNa(i).dmprabeg = rNa(j).dmprabeg
 roNa(i).dmpchiklass = rNa(j).dmpchiklass
 roNa(i).dmpchibeg = rNa(j).dmpchibeg
 roNa(i).dmprsklass = rNa(j).dmprsklass
 roNa(i).dmprsbeg = rNa(j).dmprsbeg
 roNa(i).dmpadklass = rNa(j).dmpadklass
 roNa(i).dmpadbeg = rNa(j).dmpadbeg
 roNa(i).dmpdepklass = rNa(j).dmpdepklass
 roNa(i).dmpdepbeg = rNa(j).dmpdepbeg
 roNa(i).dakab = rNa(j).dakab
 roNa(i).HzV = rNa(j).HzV
 roNa(i).HzVbeg = rNa(j).HzVbeg
 roNa(i).Ds = rNa(j).Ds
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
 roNa(i).SDatum = rNa(j).SDatum
 roNa(i).inaktiv = rNa(j).inaktiv
 roNa(i).lAktTM = rNa(j).lAktTM
 roNa(i).Mitarbeiter = rNa(j).Mitarbeiter
 roNa(i).Swz = rNa(j).Swz
 roNa(i).Gbz = rNa(j).Gbz
 roNa(i).Kiz = rNa(j).Kiz
 roNa(i).ZdeK = rNa(j).ZdeK
 roNa(i).obk = rNa(j).obk
 roNa(i).obs = rNa(j).obs
 roNa(i).obh = rNa(j).obh
 roNa(i).antikoag = rNa(j).antikoag
 roNa(i).dmt1 = rNa(j).dmt1
 roNa(i).gdm = rNa(j).gdm
 roNa(i).kdm = rNa(j).kdm
 roNa(i).cgm = rNa(j).cgm
 roNa(i).insdat = rNa(j).insdat
 roNa(i).insanw = rNa(j).insanw
End Function ' roNaZuw

Public Function NaZUnt%(i&, j&)
 If roNa(i).Pat_ID <> rNa(j).Pat_ID Then GoSub unter
 If roNa(i).TM_Pat_ID <> rNa(j).TM_Pat_ID Then GoSub unter
 If roNa(i).lfdnr <> rNa(j).lfdnr Then GoSub unter
 If roNa(i).NVorsatz <> rNa(j).NVorsatz Then GoSub unter
 If roNa(i).Nachname <> rNa(j).Nachname Then GoSub unter
 If roNa(i).Vorname <> rNa(j).Vorname Then GoSub unter
 If roNa(i).GebDat <> rNa(j).GebDat Then GoSub unter
 If roNa(i).KarGen <> rNa(j).KarGen Then GoSub unter
 If roNa(i).eGKSchVer <> rNa(j).eGKSchVer Then GoSub unter
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
 If roNa(i).PFWsLC <> rNa(j).PFWsLC Then GoSub unter
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
 If roNa(i).aktZeit <> rNa(j).aktZeit Then GoSub unter
 If roNa(i).absPos <> rNa(j).absPos Then GoSub unter
 If roNa(i).StByte <> rNa(j).StByte Then GoSub unter
 If roNa(i).StByteA <> rNa(j).StByteA Then GoSub unter
 If roNa(i).Cave <> rNa(j).Cave Then GoSub unter
 If roNa(i).notiz <> rNa(j).notiz Then GoSub unter
 If roNa(i).info <> rNa(j).info Then GoSub unter
 If roNa(i).obChk <> rNa(j).obChk Then GoSub unter
 If roNa(i).NZNr <> rNa(j).NZNr Then GoSub unter
 If roNa(i).dmpklass <> rNa(j).dmpklass Then GoSub unter
 If roNa(i).dmpbeg <> rNa(j).dmpbeg Then GoSub unter
 If roNa(i).dmpkhkklass <> rNa(j).dmpkhkklass Then GoSub unter
 If roNa(i).dmpkhkbeg <> rNa(j).dmpkhkbeg Then GoSub unter
 If roNa(i).dmpcopdklass <> rNa(j).dmpcopdklass Then GoSub unter
 If roNa(i).dmpcopdbeg <> rNa(j).dmpcopdbeg Then GoSub unter
 If roNa(i).dmpabklass <> rNa(j).dmpabklass Then GoSub unter
 If roNa(i).dmpabbeg <> rNa(j).dmpabbeg Then GoSub unter
 If roNa(i).dmposteoklass <> rNa(j).dmposteoklass Then GoSub unter
 If roNa(i).dmposteobeg <> rNa(j).dmposteobeg Then GoSub unter
 If roNa(i).dmpraklass <> rNa(j).dmpraklass Then GoSub unter
 If roNa(i).dmprabeg <> rNa(j).dmprabeg Then GoSub unter
 If roNa(i).dmpchiklass <> rNa(j).dmpchiklass Then GoSub unter
 If roNa(i).dmpchibeg <> rNa(j).dmpchibeg Then GoSub unter
 If roNa(i).dmprsklass <> rNa(j).dmprsklass Then GoSub unter
 If roNa(i).dmprsbeg <> rNa(j).dmprsbeg Then GoSub unter
 If roNa(i).dmpadklass <> rNa(j).dmpadklass Then GoSub unter
 If roNa(i).dmpadbeg <> rNa(j).dmpadbeg Then GoSub unter
 If roNa(i).dmpdepklass <> rNa(j).dmpdepklass Then GoSub unter
 If roNa(i).dmpdepbeg <> rNa(j).dmpdepbeg Then GoSub unter
 If roNa(i).dakab <> rNa(j).dakab Then GoSub unter
 If roNa(i).HzV <> rNa(j).HzV Then GoSub unter
 If roNa(i).HzVbeg <> rNa(j).HzVbeg Then GoSub unter
 If roNa(i).Ds <> rNa(j).Ds Then GoSub unter
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
 If roNa(i).SDatum <> rNa(j).SDatum Then GoSub unter
 If roNa(i).inaktiv <> rNa(j).inaktiv Then GoSub unter
 If roNa(i).lAktTM <> rNa(j).lAktTM Then GoSub unter
 If roNa(i).Mitarbeiter <> rNa(j).Mitarbeiter Then GoSub unter
 If roNa(i).Swz <> rNa(j).Swz Then GoSub unter
 If roNa(i).Gbz <> rNa(j).Gbz Then GoSub unter
 If roNa(i).Kiz <> rNa(j).Kiz Then GoSub unter
 If roNa(i).ZdeK <> rNa(j).ZdeK Then GoSub unter
 If roNa(i).obk <> rNa(j).obk Then GoSub unter
 If roNa(i).obs <> rNa(j).obs Then GoSub unter
 If roNa(i).obh <> rNa(j).obh Then GoSub unter
 If roNa(i).antikoag <> rNa(j).antikoag Then GoSub unter
 If roNa(i).dmt1 <> rNa(j).dmt1 Then GoSub unter
 If roNa(i).gdm <> rNa(j).gdm Then GoSub unter
 If roNa(i).kdm <> rNa(j).kdm Then GoSub unter
 If roNa(i).cgm <> rNa(j).cgm Then GoSub unter
 If roNa(i).insdat <> rNa(j).insdat Then GoSub unter
 If roNa(i).insanw <> rNa(j).insanw Then GoSub unter
 Exit Function
unter:
 NaZUnt = NaZUnt + 1
 Return
End Function ' NaZUnt

Public Function namenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(Pat_ID,0) Pat_ID,COALESCE(TM_Pat_ID,0) TM_Pat_ID,COALESCE(lfdnr,0) lfdnr,COALESCE(NVorsatz,'') NVorsatz" & _
",COALESCE(Nachname,'') Nachname,COALESCE(Vorname,'') Vorname,COALESCE(GebDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) GebDat,COALESCE(KarGen,'') KarGen" & _
",COALESCE(eGKSchVer,'') eGKSchVer,COALESCE(Straße,'') Straße,COALESCE(KVKStatus,'') KVKStatus,COALESCE(Hausnr,'') Hausnr" & _
",COALESCE(Geschlecht,'') Geschlecht,COALESCE(Plz,'') Plz,COALESCE(Ort,'') Ort,COALESCE(Lkz,'') Lkz" & _
",COALESCE(Anschrzus,'') Anschrzus,COALESCE(NVors,'') NVors,COALESCE(PFPlz,'') PFPlz,COALESCE(PFOrt,'') PFOrt" & _
",COALESCE(PFNr,'') PFNr,COALESCE(PFWsLC,'') PFWsLC,COALESCE(AnschrZus_2,'') AnschrZus_2,COALESCE(Postfach_2,'') Postfach_2" & _
",COALESCE(LK_2,'') LK_2,COALESCE(Postfach,'') Postfach,COALESCE(Beruf,'') Beruf,COALESCE(Weggeldzone,'') Weggeldzone" & _
",COALESCE(WeggzZahl,0) WeggzZahl,COALESCE(AufnDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AufnDat,COALESCE(kAufDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) kAufDat,COALESCE(LANR,'') LANR" & _
",COALESCE(BStNr,'') BStNr,COALESCE(Titel,'') Titel,COALESCE(Versichertennummer,'') Versichertennummer,COALESCE(PrivatTel,'') PrivatTel" & _
",COALESCE(KVNr,'') KVNr,COALESCE(KVNr2,'') KVNr2,COALESCE(KVNr3,'') KVNr3,COALESCE(KVNr4,'') KVNr4" & _
",COALESCE(PrivatTel_2,'') PrivatTel_2,COALESCE(PrivatFax,'') PrivatFax,COALESCE(DienstTel,'') DienstTel,COALESCE(PrivatMobil,'') PrivatMobil" & _
",COALESCE(Email,'') Email,COALESCE(Arbeitgeber,'') Arbeitgeber,COALESCE(AnAllgda,0) AnAllgda,COALESCE(An1da,0) An1da" & _
",COALESCE(An2da,0) An2da,COALESCE(Checkda,0) Checkda,COALESCE(DMTypaD,'') DMTypaD,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(absPos,0) absPos,COALESCE(StByte,0) StByte,COALESCE(StByteA,0) StByteA,COALESCE(Cave,'') Cave" & _
",COALESCE(notiz,'') notiz,COALESCE(info,'') info,COALESCE(obChk,'') obChk,COALESCE(NZNr,0) NZNr" & _
",COALESCE(dmpklass,0) dmpklass,COALESCE(dmpbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpbeg,COALESCE(dmpkhkklass,0) dmpkhkklass,COALESCE(dmpkhkbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpkhkbeg" & _
",COALESCE(dmpcopdklass,0) dmpcopdklass,COALESCE(dmpcopdbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpcopdbeg,COALESCE(dmpabklass,0) dmpabklass,COALESCE(dmpabbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpabbeg" & _
",COALESCE(dmposteoklass,0) dmposteoklass,COALESCE(dmposteobeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmposteobeg,COALESCE(dmpraklass,0) dmpraklass,COALESCE(dmprabeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmprabeg" & _
",COALESCE(dmpchiklass,0) dmpchiklass,COALESCE(dmpchibeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpchibeg,COALESCE(dmprsklass,0) dmprsklass,COALESCE(dmprsbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmprsbeg" & _
",COALESCE(dmpadklass,0) dmpadklass,COALESCE(dmpadbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpadbeg,COALESCE(dmpdepklass,0) dmpdepklass,COALESCE(dmpdepbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dmpdepbeg" & _
",COALESCE(dakab - INTERVAL 0 DAY,CONVERT('18991230',DATE)) dakab,COALESCE(HzV,0) HzV,COALESCE(HzVbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) HzVbeg,COALESCE(DS,0) DS" & _
",COALESCE(DSbeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DSbeg,COALESCE(getHA0,0) getHA0,COALESCE(fnHA0,'') fnHA0,COALESCE(getHA1,0) getHA1" & _
",COALESCE(fnHA1,'') fnHA1,COALESCE(getHA2,0) getHA2,COALESCE(fnHA2,'') fnHA2,COALESCE(zubenach,'') zubenach" & _
",COALESCE(Verwandt,'') Verwandt,COALESCE(Sprache,'') Sprache,COALESCE(SDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) SDatum,COALESCE(inaktiv,0) inaktiv" & _
",COALESCE(lAktTM - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lAktTM,COALESCE(Mitarbeiter,0) Mitarbeiter,COALESCE(Swz,0) Swz,COALESCE(Gbz,0) Gbz"
sql = sql & ",COALESCE(Kiz,0) Kiz,COALESCE(ZdeK,0) ZdeK,COALESCE(obk,0) obk,COALESCE(obs,0) obs" & _
",COALESCE(obh,0) obh,COALESCE(antikoag,0) antikoag,COALESCE(dmt1,0) dmt1,COALESCE(gdm,0) gdm" & _
",COALESCE(kdm,0) kdm,COALESCE(cgm,0) cgm,COALESCE(insdat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) insdat,COALESCE(insanw,0) insanw" & _
" FROM `namen` WHERE Pat_ID=" & pid & " ORDER BY `kAufDat`"
 myFrag rs, sql
 ReDim roNa(1)
 If Not rs.EOF Then
  Do While Not rs.EOF
   akt = UBound(roNa)
   roNa(akt).Pat_ID = rs!Pat_ID
   roNa(akt).TM_Pat_ID = rs!TM_Pat_ID
   roNa(akt).lfdnr = rs!lfdnr
   roNa(akt).NVorsatz = doUmwfSQL(rs!NVorsatz, lies.obMySQL, False)
   roNa(akt).Nachname = doUmwfSQL(rs!Nachname, lies.obMySQL, False)
   roNa(akt).Vorname = doUmwfSQL(rs!Vorname, lies.obMySQL, False)
   roNa(akt).GebDat = rs!GebDat
   roNa(akt).KarGen = doUmwfSQL(rs!KarGen, lies.obMySQL, False)
   roNa(akt).eGKSchVer = doUmwfSQL(rs!eGKSchVer, lies.obMySQL, False)
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
   roNa(akt).PFWsLC = doUmwfSQL(rs!PFWsLC, lies.obMySQL, False)
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
   roNa(akt).aktZeit = rs!aktZeit
   roNa(akt).absPos = rs!absPos
   roNa(akt).StByte = rs!StByte
   roNa(akt).StByteA = rs!StByteA
   roNa(akt).Cave = doUmwfSQL(rs!Cave, lies.obMySQL, False)
   roNa(akt).notiz = doUmwfSQL(rs!notiz, lies.obMySQL, False)
   roNa(akt).info = doUmwfSQL(rs!info, lies.obMySQL, False)
   roNa(akt).obChk = doUmwfSQL(rs!obChk, lies.obMySQL, False)
   roNa(akt).NZNr = rs!NZNr
   roNa(akt).dmpklass = rs!dmpklass
   roNa(akt).dmpbeg = rs!dmpbeg
   roNa(akt).dmpkhkklass = rs!dmpkhkklass
   roNa(akt).dmpkhkbeg = rs!dmpkhkbeg
   roNa(akt).dmpcopdklass = rs!dmpcopdklass
   roNa(akt).dmpcopdbeg = rs!dmpcopdbeg
   roNa(akt).dmpabklass = rs!dmpabklass
   roNa(akt).dmpabbeg = rs!dmpabbeg
   roNa(akt).dmposteoklass = rs!dmposteoklass
   roNa(akt).dmposteobeg = rs!dmposteobeg
   roNa(akt).dmpraklass = rs!dmpraklass
   roNa(akt).dmprabeg = rs!dmprabeg
   roNa(akt).dmpchiklass = rs!dmpchiklass
   roNa(akt).dmpchibeg = rs!dmpchibeg
   roNa(akt).dmprsklass = rs!dmprsklass
   roNa(akt).dmprsbeg = rs!dmprsbeg
   roNa(akt).dmpadklass = rs!dmpadklass
   roNa(akt).dmpadbeg = rs!dmpadbeg
   roNa(akt).dmpdepklass = rs!dmpdepklass
   roNa(akt).dmpdepbeg = rs!dmpdepbeg
   roNa(akt).dakab = rs!dakab
   roNa(akt).HzV = rs!HzV
   roNa(akt).HzVbeg = rs!HzVbeg
   roNa(akt).Ds = rs!Ds
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
   roNa(akt).SDatum = rs!SDatum
   roNa(akt).inaktiv = rs!inaktiv
   roNa(akt).lAktTM = rs!lAktTM
   roNa(akt).Mitarbeiter = rs!Mitarbeiter
   roNa(akt).Swz = rs!Swz
   roNa(akt).Gbz = rs!Gbz
   roNa(akt).Kiz = rs!Kiz
   roNa(akt).ZdeK = rs!ZdeK
   roNa(akt).obk = rs!obk
   roNa(akt).obs = rs!obs
   roNa(akt).obh = rs!obh
   roNa(akt).antikoag = rs!antikoag
   roNa(akt).dmt1 = rs!dmt1
   roNa(akt).gdm = rs!gdm
   roNa(akt).kdm = rs!kdm
   roNa(akt).cgm = rs!cgm
   roNa(akt).insdat = rs!insdat
   roNa(akt).insanw = rs!insanw
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rNa) > 0 Then
  For ri = 1 To UBound(rNa)
   If rNa(ri).kAufDat >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roNa)
    If roNa(roendpe).kAufDat >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roNa(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roNa(roendpe + UBound(rNa) - rbeg)
   For ri = rbeg To UBound(rNa)
    Call roNaZuw(roendpe + ri - rbeg, ri)
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

Public Function rNaDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rNadump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rNa)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rNa(" & i & ").Pat_ID:" & String$(33, "."), 33) & rNa(i).Pat_ID
  Print #200, left$("rNa(" & i & ").TM_Pat_ID:" & String$(33, "."), 33) & rNa(i).TM_Pat_ID
  Print #200, left$("rNa(" & i & ").lfdnr:" & String$(33, "."), 33) & rNa(i).lfdnr
  Print #200, left$("rNa(" & i & ").NVorsatz:" & String$(33, "."), 33) & "'" & rNa(i).NVorsatz & "'"
  Print #200, left$("rNa(" & i & ").Nachname:" & String$(33, "."), 33) & "'" & rNa(i).Nachname & "'"
  Print #200, left$("rNa(" & i & ").Vorname:" & String$(33, "."), 33) & "'" & rNa(i).Vorname & "'"
  Print #200, left$("rNa(" & i & ").GebDat:" & String$(33, "."), 33) & rNa(i).GebDat
  Print #200, left$("rNa(" & i & ").KarGen:" & String$(33, "."), 33) & "'" & rNa(i).KarGen & "'"
  Print #200, left$("rNa(" & i & ").eGKSchVer:" & String$(33, "."), 33) & "'" & rNa(i).eGKSchVer & "'"
  Print #200, left$("rNa(" & i & ").Straße:" & String$(33, "."), 33) & "'" & rNa(i).Straße & "'"
  Print #200, left$("rNa(" & i & ").KVKStatus:" & String$(33, "."), 33) & "'" & rNa(i).KVKStatus & "'"
  Print #200, left$("rNa(" & i & ").Hausnr:" & String$(33, "."), 33) & "'" & rNa(i).Hausnr & "'"
  Print #200, left$("rNa(" & i & ").Geschlecht:" & String$(33, "."), 33) & "'" & rNa(i).geschlecht & "'"
  Print #200, left$("rNa(" & i & ").Plz:" & String$(33, "."), 33) & "'" & rNa(i).plz & "'"
  Print #200, left$("rNa(" & i & ").Ort:" & String$(33, "."), 33) & "'" & rNa(i).ort & "'"
  Print #200, left$("rNa(" & i & ").Lkz:" & String$(33, "."), 33) & "'" & rNa(i).Lkz & "'"
  Print #200, left$("rNa(" & i & ").Anschrzus:" & String$(33, "."), 33) & "'" & rNa(i).Anschrzus & "'"
  Print #200, left$("rNa(" & i & ").NVors:" & String$(33, "."), 33) & "'" & rNa(i).NVors & "'"
  Print #200, left$("rNa(" & i & ").PFPlz:" & String$(33, "."), 33) & "'" & rNa(i).PFPlz & "'"
  Print #200, left$("rNa(" & i & ").PFOrt:" & String$(33, "."), 33) & "'" & rNa(i).PFOrt & "'"
  Print #200, left$("rNa(" & i & ").PFNr:" & String$(33, "."), 33) & "'" & rNa(i).PFNr & "'"
  Print #200, left$("rNa(" & i & ").PFWsLC:" & String$(33, "."), 33) & "'" & rNa(i).PFWsLC & "'"
  Print #200, left$("rNa(" & i & ").AnschrZus_2:" & String$(33, "."), 33) & "'" & rNa(i).AnschrZus_2 & "'"
  Print #200, left$("rNa(" & i & ").Postfach_2:" & String$(33, "."), 33) & "'" & rNa(i).Postfach_2 & "'"
  Print #200, left$("rNa(" & i & ").LK_2:" & String$(33, "."), 33) & "'" & rNa(i).LK_2 & "'"
  Print #200, left$("rNa(" & i & ").Postfach:" & String$(33, "."), 33) & "'" & rNa(i).Postfach & "'"
  Print #200, left$("rNa(" & i & ").Beruf:" & String$(33, "."), 33) & "'" & rNa(i).Beruf & "'"
  Print #200, left$("rNa(" & i & ").Weggeldzone:" & String$(33, "."), 33) & "'" & rNa(i).Weggeldzone & "'"
  Print #200, left$("rNa(" & i & ").WeggzZahl:" & String$(33, "."), 33) & "'" & rNa(i).WeggzZahl & "'"
  Print #200, left$("rNa(" & i & ").AufnDat:" & String$(33, "."), 33) & rNa(i).AufnDat
  Print #200, left$("rNa(" & i & ").kAufDat:" & String$(33, "."), 33) & rNa(i).kAufDat
  Print #200, left$("rNa(" & i & ").LANR:" & String$(33, "."), 33) & "'" & rNa(i).Lanr & "'"
  Print #200, left$("rNa(" & i & ").BStNr:" & String$(33, "."), 33) & "'" & rNa(i).BStNr & "'"
  Print #200, left$("rNa(" & i & ").Titel:" & String$(33, "."), 33) & "'" & rNa(i).Titel & "'"
  Print #200, left$("rNa(" & i & ").Versichertennummer:" & String$(33, "."), 33) & "'" & rNa(i).Versichertennummer & "'"
  Print #200, left$("rNa(" & i & ").PrivatTel:" & String$(33, "."), 33) & "'" & rNa(i).PrivatTel & "'"
  Print #200, left$("rNa(" & i & ").KVNr:" & String$(33, "."), 33) & "'" & rNa(i).KVNr & "'"
  Print #200, left$("rNa(" & i & ").KVNr2:" & String$(33, "."), 33) & "'" & rNa(i).KVNr2 & "'"
  Print #200, left$("rNa(" & i & ").KVNr3:" & String$(33, "."), 33) & "'" & rNa(i).KVNr3 & "'"
  Print #200, left$("rNa(" & i & ").KVNr4:" & String$(33, "."), 33) & "'" & rNa(i).KVNr4 & "'"
  Print #200, left$("rNa(" & i & ").PrivatTel_2:" & String$(33, "."), 33) & "'" & rNa(i).PrivatTel_2 & "'"
  Print #200, left$("rNa(" & i & ").PrivatFax:" & String$(33, "."), 33) & "'" & rNa(i).PrivatFax & "'"
  Print #200, left$("rNa(" & i & ").DienstTel:" & String$(33, "."), 33) & "'" & rNa(i).DienstTel & "'"
  Print #200, left$("rNa(" & i & ").PrivatMobil:" & String$(33, "."), 33) & "'" & rNa(i).PrivatMobil & "'"
  Print #200, left$("rNa(" & i & ").Email:" & String$(33, "."), 33) & "'" & rNa(i).email & "'"
  Print #200, left$("rNa(" & i & ").Arbeitgeber:" & String$(33, "."), 33) & "'" & rNa(i).Arbeitgeber & "'"
  Print #200, left$("rNa(" & i & ").AnAllgda:" & String$(33, "."), 33) & rNa(i).AnAllgda
  Print #200, left$("rNa(" & i & ").An1da:" & String$(33, "."), 33) & rNa(i).An1da
  Print #200, left$("rNa(" & i & ").An2da:" & String$(33, "."), 33) & rNa(i).An2da
  Print #200, left$("rNa(" & i & ").Checkda:" & String$(33, "."), 33) & rNa(i).Checkda
  Print #200, left$("rNa(" & i & ").DMTypaD:" & String$(33, "."), 33) & "'" & rNa(i).DMTypaD & "'"
  Print #200, left$("rNa(" & i & ").AktZeit:" & String$(33, "."), 33) & rNa(i).aktZeit
  Print #200, left$("rNa(" & i & ").absPos:" & String$(33, "."), 33) & rNa(i).absPos
  Print #200, left$("rNa(" & i & ").StByte:" & String$(33, "."), 33) & rNa(i).StByte
  Print #200, left$("rNa(" & i & ").StByteA:" & String$(33, "."), 33) & rNa(i).StByteA
  Print #200, left$("rNa(" & i & ").Cave:" & String$(33, "."), 33) & "'" & rNa(i).Cave & "'"
  Print #200, left$("rNa(" & i & ").notiz:" & String$(33, "."), 33) & "'" & rNa(i).notiz & "'"
  Print #200, left$("rNa(" & i & ").info:" & String$(33, "."), 33) & "'" & rNa(i).info & "'"
  Print #200, left$("rNa(" & i & ").obChk:" & String$(33, "."), 33) & "'" & rNa(i).obChk & "'"
  Print #200, left$("rNa(" & i & ").NZNr:" & String$(33, "."), 33) & rNa(i).NZNr
  Print #200, left$("rNa(" & i & ").dmpklass:" & String$(33, "."), 33) & rNa(i).dmpklass
  Print #200, left$("rNa(" & i & ").dmpbeg:" & String$(33, "."), 33) & rNa(i).dmpbeg
  Print #200, left$("rNa(" & i & ").dmpkhkklass:" & String$(33, "."), 33) & rNa(i).dmpkhkklass
  Print #200, left$("rNa(" & i & ").dmpkhkbeg:" & String$(33, "."), 33) & rNa(i).dmpkhkbeg
  Print #200, left$("rNa(" & i & ").dmpcopdklass:" & String$(33, "."), 33) & rNa(i).dmpcopdklass
  Print #200, left$("rNa(" & i & ").dmpcopdbeg:" & String$(33, "."), 33) & rNa(i).dmpcopdbeg
  Print #200, left$("rNa(" & i & ").dmpabklass:" & String$(33, "."), 33) & rNa(i).dmpabklass
  Print #200, left$("rNa(" & i & ").dmpabbeg:" & String$(33, "."), 33) & rNa(i).dmpabbeg
  Print #200, left$("rNa(" & i & ").dmposteoklass:" & String$(33, "."), 33) & rNa(i).dmposteoklass
  Print #200, left$("rNa(" & i & ").dmposteobeg:" & String$(33, "."), 33) & rNa(i).dmposteobeg
  Print #200, left$("rNa(" & i & ").dmpraklass:" & String$(33, "."), 33) & rNa(i).dmpraklass
  Print #200, left$("rNa(" & i & ").dmprabeg:" & String$(33, "."), 33) & rNa(i).dmprabeg
  Print #200, left$("rNa(" & i & ").dmpchiklass:" & String$(33, "."), 33) & rNa(i).dmpchiklass
  Print #200, left$("rNa(" & i & ").dmpchibeg:" & String$(33, "."), 33) & rNa(i).dmpchibeg
  Print #200, left$("rNa(" & i & ").dmprsklass:" & String$(33, "."), 33) & rNa(i).dmprsklass
  Print #200, left$("rNa(" & i & ").dmprsbeg:" & String$(33, "."), 33) & rNa(i).dmprsbeg
  Print #200, left$("rNa(" & i & ").dmpadklass:" & String$(33, "."), 33) & rNa(i).dmpadklass
  Print #200, left$("rNa(" & i & ").dmpadbeg:" & String$(33, "."), 33) & rNa(i).dmpadbeg
  Print #200, left$("rNa(" & i & ").dmpdepklass:" & String$(33, "."), 33) & rNa(i).dmpdepklass
  Print #200, left$("rNa(" & i & ").dmpdepbeg:" & String$(33, "."), 33) & rNa(i).dmpdepbeg
  Print #200, left$("rNa(" & i & ").dakab:" & String$(33, "."), 33) & rNa(i).dakab
  Print #200, left$("rNa(" & i & ").HzV:" & String$(33, "."), 33) & rNa(i).HzV
  Print #200, left$("rNa(" & i & ").HzVbeg:" & String$(33, "."), 33) & rNa(i).HzVbeg
  Print #200, left$("rNa(" & i & ").DS:" & String$(33, "."), 33) & rNa(i).Ds
  Print #200, left$("rNa(" & i & ").DSbeg:" & String$(33, "."), 33) & rNa(i).DSbeg
  Print #200, left$("rNa(" & i & ").getHA0:" & String$(33, "."), 33) & rNa(i).getHA0
  Print #200, left$("rNa(" & i & ").fnHA0:" & String$(33, "."), 33) & "'" & rNa(i).fnHA0 & "'"
  Print #200, left$("rNa(" & i & ").getHA1:" & String$(33, "."), 33) & rNa(i).getHA1
  Print #200, left$("rNa(" & i & ").fnHA1:" & String$(33, "."), 33) & "'" & rNa(i).fnHA1 & "'"
  Print #200, left$("rNa(" & i & ").getHA2:" & String$(33, "."), 33) & rNa(i).getHA2
  Print #200, left$("rNa(" & i & ").fnHA2:" & String$(33, "."), 33) & "'" & rNa(i).fnHA2 & "'"
  Print #200, left$("rNa(" & i & ").zubenach:" & String$(33, "."), 33) & "'" & rNa(i).zubenach & "'"
  Print #200, left$("rNa(" & i & ").Verwandt:" & String$(33, "."), 33) & "'" & rNa(i).Verwandt & "'"
  Print #200, left$("rNa(" & i & ").Sprache:" & String$(33, "."), 33) & "'" & rNa(i).Sprache & "'"
  Print #200, left$("rNa(" & i & ").SDatum:" & String$(33, "."), 33) & rNa(i).SDatum
  Print #200, left$("rNa(" & i & ").inaktiv:" & String$(33, "."), 33) & rNa(i).inaktiv
  Print #200, left$("rNa(" & i & ").lAktTM:" & String$(33, "."), 33) & rNa(i).lAktTM
  Print #200, left$("rNa(" & i & ").Mitarbeiter:" & String$(33, "."), 33) & rNa(i).Mitarbeiter
  Print #200, left$("rNa(" & i & ").Swz:" & String$(33, "."), 33) & rNa(i).Swz
  Print #200, left$("rNa(" & i & ").Gbz:" & String$(33, "."), 33) & rNa(i).Gbz
  Print #200, left$("rNa(" & i & ").Kiz:" & String$(33, "."), 33) & rNa(i).Kiz
  Print #200, left$("rNa(" & i & ").ZdeK:" & String$(33, "."), 33) & rNa(i).ZdeK
  Print #200, left$("rNa(" & i & ").obk:" & String$(33, "."), 33) & rNa(i).obk
  Print #200, left$("rNa(" & i & ").obs:" & String$(33, "."), 33) & rNa(i).obs
  Print #200, left$("rNa(" & i & ").obh:" & String$(33, "."), 33) & rNa(i).obh
  Print #200, left$("rNa(" & i & ").antikoag:" & String$(33, "."), 33) & rNa(i).antikoag
  Print #200, left$("rNa(" & i & ").dmt1:" & String$(33, "."), 33) & rNa(i).dmt1
  Print #200, left$("rNa(" & i & ").gdm:" & String$(33, "."), 33) & rNa(i).gdm
  Print #200, left$("rNa(" & i & ").kdm:" & String$(33, "."), 33) & rNa(i).kdm
  Print #200, left$("rNa(" & i & ").cgm:" & String$(33, "."), 33) & rNa(i).cgm
  Print #200, left$("rNa(" & i & ").insdat:" & String$(33, "."), 33) & rNa(i).insdat
  Print #200, left$("rNa(" & i & ").insanw:" & String$(33, "."), 33) & rNa(i).insanw
 Next i
 Close #200
 zeigan ffadat
End Function ' namenDump

Public Function namenSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "namen")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rNa) + 1 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Pat_ID,TM_Pat_ID,lfdnr," & _
     "NVorsatz,Nachname,Vorname,GebDat,KarGen,eGKSchVer,Straße,KVKStatus,Hausnr,Geschlecht," & _
     "Plz,Ort,Lkz,Anschrzus,NVors,PFPlz,PFOrt,PFNr,PFWsLC,AnschrZus_2," & _
     "Postfach_2,LK_2,Postfach,Beruf,Weggeldzone,WeggzZahl,AufnDat,kAufDat,LANR,BStNr," & _
     "Titel,Versichertennummer,PrivatTel,KVNr,KVNr2,KVNr3,KVNr4,PrivatTel_2,PrivatFax,DienstTel," & _
     "PrivatMobil,Email,Arbeitgeber,AnAllgda,An1da,An2da,Checkda,DMTypaD,AktZeit,absPos," & _
     "StByte,StByteA,Cave,notiz,info,obChk,NZNr,dmpklass,dmpbeg,dmpkhkklass," & _
     "dmpkhkbeg,dmpcopdklass,dmpcopdbeg,dmpabklass,dmpabbeg,dmposteoklass,dmposteobeg,dmpraklass,dmprabeg,dmpchiklass," & _
     "dmpchibeg,dmprsklass,dmprsbeg,dmpadklass,dmpadbeg,dmpdepklass,dmpdepbeg,dakab,HzV,HzVbeg," & _
     "DS,DSbeg,getHA0,fnHA0,getHA1,fnHA1,getHA2,fnHA2,zubenach,Verwandt," & _
     "Sprache,SDatum,inaktiv,lAktTM,Mitarbeiter,Swz,Gbz,Kiz,ZdeK,obk," & _
     "obs,obh,antikoag,dmt1,gdm,kdm,cgm,insdat,insanw)   VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 0 To UBound(rNa)
  rNa(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 0 Then
  csql.AppVar Array("(", rNa(i).Pat_ID, ",", rNa(i).TM_Pat_ID, ",", rNa(i).lfdnr, ",'", rNa(i).NVorsatz, "','", rNa(i).Nachname, "','", rNa(i).Vorname, "',", DatFor_k(rNa(i).GebDat), ",'", rNa(i).KarGen, "','", _
   rNa(i).eGKSchVer, "','", rNa(i).Straße, "','", rNa(i).KVKStatus, "','", rNa(i).Hausnr, "','", rNa(i).geschlecht, "','", rNa(i).plz, "','", rNa(i).ort, "','", rNa(i).Lkz, "','", _
   rNa(i).Anschrzus, "','", rNa(i).NVors, "','", rNa(i).PFPlz, "','", rNa(i).PFOrt, "','", rNa(i).PFNr, "','", rNa(i).PFWsLC, "','", rNa(i).AnschrZus_2, "','", rNa(i).Postfach_2, "','", _
   rNa(i).LK_2, "','", rNa(i).Postfach, "','", rNa(i).Beruf, "','", rNa(i).Weggeldzone, "',", rNa(i).WeggzZahl, ",", DatFor_k(rNa(i).AufnDat), ",", DatFor_k(rNa(i).kAufDat), ",'", rNa(i).Lanr, "','", _
   rNa(i).BStNr, "','", rNa(i).Titel, "','", rNa(i).Versichertennummer, "','", rNa(i).PrivatTel, "','", rNa(i).KVNr, "','", rNa(i).KVNr2, "','", rNa(i).KVNr3, "','", rNa(i).KVNr4, "','", _
   rNa(i).PrivatTel_2, "','", rNa(i).PrivatFax, "','", rNa(i).DienstTel, "','", rNa(i).PrivatMobil, "','", rNa(i).email, "','", rNa(i).Arbeitgeber, "',", CStr(-(rNa(i).AnAllgda <> 0)), ",", CStr(-( _
   rNa(i).An1da <> 0)), ",", CStr(-(rNa(i).An2da <> 0)), ",", CStr(-(rNa(i).Checkda <> 0)), ",'", rNa(i).DMTypaD, "',", DatFor_k(0), ",", rNa(i).absPos, ",", rNa(i).StByte, ",", _
   rNa(i).StByteA, ",'", rNa(i).Cave, "','", rNa(i).notiz, "','", rNa(i).info, "','", rNa(i).obChk, "',", rNa(i).NZNr, ",", rNa(i).dmpklass, ",", DatFor_k(rNa(i).dmpbeg), ",", rNa(i).dmpkhkklass, ",", DatFor_k( _
   rNa(i).dmpkhkbeg), ",", rNa(i).dmpcopdklass, ",", DatFor_k(rNa(i).dmpcopdbeg), ",", rNa(i).dmpabklass, ",", DatFor_k(rNa(i).dmpabbeg), ",", rNa(i).dmposteoklass, ",", DatFor_k( _
   rNa(i).dmposteobeg), ",", rNa(i).dmpraklass, ",", DatFor_k(rNa(i).dmprabeg), ",", rNa(i).dmpchiklass, ",", DatFor_k(rNa(i).dmpchibeg), ",", rNa(i).dmprsklass, ",", DatFor_k( _
   rNa(i).dmprsbeg), ",", rNa(i).dmpadklass, ",", DatFor_k(rNa(i).dmpadbeg), ",", rNa(i).dmpdepklass, ",", DatFor_k(rNa(i).dmpdepbeg), ",", DatFor_k(rNa(i).dakab), ",", rNa(i).HzV, ",", DatFor_k( _
   rNa(i).HzVbeg), ",", rNa(i).Ds, ",", DatFor_k(rNa(i).DSbeg), ",", rNa(i).getHA0, ",'", rNa(i).fnHA0, "',", rNa(i).getHA1, ",'", rNa(i).fnHA1, "',", rNa(i).getHA2, ",'", rNa(i).fnHA2, "','", _
   rNa(i).zubenach, "','", rNa(i).Verwandt, "','", rNa(i).Sprache, "',", DatFor_k(rNa(i).SDatum), ",", rNa(i).inaktiv, ",", DatFor_k(rNa(i).lAktTM), ",", rNa(i).Mitarbeiter, ",", _
   rNa(i).Swz, ",", rNa(i).Gbz, ",", rNa(i).Kiz, ",", rNa(i).ZdeK, ",", CStr(-(rNa(i).obk <> 0)), ",", CStr(-(rNa(i).obs <> 0)), ",", CStr(-(rNa(i).obh <> 0)), ",", CStr(-(rNa(i).antikoag <> 0)), ",", CStr(-( _
   rNa(i).dmt1 <> 0)), ",", CStr(-(rNa(i).gdm <> 0)), ",", CStr(-(rNa(i).kdm <> 0)), ",", rNa(i).cgm, ",", DatFor_k(rNa(i).insdat), ",", rNa(i).insanw, ")")
  If SammelInsert <> 0 And i < UBound(rNa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rNa) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rNa(i).Pat_ID = myEFrag("SELECT MAX(Pat_ID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rNa)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rNa(" & i & "/" & UBound(rNa) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""namenSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(49)
 For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rNa), i)
  If Len(rNa(k).NVorsatz) > maxi(0) Then maxi(0) = Len(rNa(k).NVorsatz)
  If Len(rNa(k).Nachname) > maxi(1) Then maxi(1) = Len(rNa(k).Nachname)
  If Len(rNa(k).Vorname) > maxi(2) Then maxi(2) = Len(rNa(k).Vorname)
  If Len(rNa(k).KarGen) > maxi(3) Then maxi(3) = Len(rNa(k).KarGen)
  If Len(rNa(k).eGKSchVer) > maxi(4) Then maxi(4) = Len(rNa(k).eGKSchVer)
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
  If Len(rNa(k).PFWsLC) > maxi(17) Then maxi(17) = Len(rNa(k).PFWsLC)
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
  If Len(rNa(k).notiz) > maxi(41) Then maxi(41) = Len(rNa(k).notiz)
  If Len(rNa(k).info) > maxi(42) Then maxi(42) = Len(rNa(k).info)
  If Len(rNa(k).obChk) > maxi(43) Then maxi(43) = Len(rNa(k).obChk)
  If Len(rNa(k).fnHA0) > maxi(44) Then maxi(44) = Len(rNa(k).fnHA0)
  If Len(rNa(k).fnHA1) > maxi(45) Then maxi(45) = Len(rNa(k).fnHA1)
  If Len(rNa(k).fnHA2) > maxi(46) Then maxi(46) = Len(rNa(k).fnHA2)
  If Len(rNa(k).zubenach) > maxi(47) Then maxi(47) = Len(rNa(k).zubenach)
  If Len(rNa(k).Verwandt) > maxi(48) Then maxi(48) = Len(rNa(k).Verwandt)
  If Len(rNa(k).Sprache) > maxi(49) Then maxi(49) = Len(rNa(k).Sprache)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rNa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVorsatz: '" & rNa(k).NVorsatz & "' -> '" & left$(rNa(k).NVorsatz, maxL) & "'", True: rNa(k).NVorsatz = left$(rNa(k).NVorsatz, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rNa.Nachname: '" & rNa(k).Nachname & "' -> '" & left$(rNa(k).Nachname, maxL) & "'", True: rNa(k).Nachname = left$(rNa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rNa.Vorname: '" & rNa(k).Vorname & "' -> '" & left$(rNa(k).Vorname, maxL) & "'", True: rNa(k).Vorname = left$(rNa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rNa.KarGen: '" & rNa(k).KarGen & "' -> '" & left$(rNa(k).KarGen, maxL) & "'", True: rNa(k).KarGen = left$(rNa(k).KarGen, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rNa.eGKSchVer: '" & rNa(k).eGKSchVer & "' -> '" & left$(rNa(k).eGKSchVer, maxL) & "'", True: rNa(k).eGKSchVer = left$(rNa(k).eGKSchVer, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rNa.Straße: '" & rNa(k).Straße & "' -> '" & left$(rNa(k).Straße, maxL) & "'", True: rNa(k).Straße = left$(rNa(k).Straße, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVKStatus: '" & rNa(k).KVKStatus & "' -> '" & left$(rNa(k).KVKStatus, maxL) & "'", True: rNa(k).KVKStatus = left$(rNa(k).KVKStatus, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rNa.Hausnr: '" & rNa(k).Hausnr & "' -> '" & left$(rNa(k).Hausnr, maxL) & "'", True: rNa(k).Hausnr = left$(rNa(k).Hausnr, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rNa.Geschlecht: '" & rNa(k).geschlecht & "' -> '" & left$(rNa(k).geschlecht, maxL) & "'", True: rNa(k).geschlecht = left$(rNa(k).geschlecht, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rNa.Plz: '" & rNa(k).plz & "' -> '" & left$(rNa(k).plz, maxL) & "'", True: rNa(k).plz = left$(rNa(k).plz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rNa.Ort: '" & rNa(k).ort & "' -> '" & left$(rNa(k).ort, maxL) & "'", True: rNa(k).ort = left$(rNa(k).ort, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rNa.Lkz: '" & rNa(k).Lkz & "' -> '" & left$(rNa(k).Lkz, maxL) & "'", True: rNa(k).Lkz = left$(rNa(k).Lkz, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rNa.Anschrzus: '" & rNa(k).Anschrzus & "' -> '" & left$(rNa(k).Anschrzus, maxL) & "'", True: rNa(k).Anschrzus = left$(rNa(k).Anschrzus, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rNa.NVors: '" & rNa(k).NVors & "' -> '" & left$(rNa(k).NVors, maxL) & "'", True: rNa(k).NVors = left$(rNa(k).NVors, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFPlz: '" & rNa(k).PFPlz & "' -> '" & left$(rNa(k).PFPlz, maxL) & "'", True: rNa(k).PFPlz = left$(rNa(k).PFPlz, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFOrt: '" & rNa(k).PFOrt & "' -> '" & left$(rNa(k).PFOrt, maxL) & "'", True: rNa(k).PFOrt = left$(rNa(k).PFOrt, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFNr: '" & rNa(k).PFNr & "' -> '" & left$(rNa(k).PFNr, maxL) & "'", True: rNa(k).PFNr = left$(rNa(k).PFNr, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rNa.PFWsLC: '" & rNa(k).PFWsLC & "' -> '" & left$(rNa(k).PFWsLC, maxL) & "'", True: rNa(k).PFWsLC = left$(rNa(k).PFWsLC, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rNa.AnschrZus_2: '" & rNa(k).AnschrZus_2 & "' -> '" & left$(rNa(k).AnschrZus_2, maxL) & "'", True: rNa(k).AnschrZus_2 = left$(rNa(k).AnschrZus_2, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach_2: '" & rNa(k).Postfach_2 & "' -> '" & left$(rNa(k).Postfach_2, maxL) & "'", True: rNa(k).Postfach_2 = left$(rNa(k).Postfach_2, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rNa.LK_2: '" & rNa(k).LK_2 & "' -> '" & left$(rNa(k).LK_2, maxL) & "'", True: rNa(k).LK_2 = left$(rNa(k).LK_2, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rNa.Postfach: '" & rNa(k).Postfach & "' -> '" & left$(rNa(k).Postfach, maxL) & "'", True: rNa(k).Postfach = left$(rNa(k).Postfach, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rNa.Beruf: '" & rNa(k).Beruf & "' -> '" & left$(rNa(k).Beruf, maxL) & "'", True: rNa(k).Beruf = left$(rNa(k).Beruf, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rNa.Weggeldzone: '" & rNa(k).Weggeldzone & "' -> '" & left$(rNa(k).Weggeldzone, maxL) & "'", True: rNa(k).Weggeldzone = left$(rNa(k).Weggeldzone, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rNa.LANR: '" & rNa(k).Lanr & "' -> '" & left$(rNa(k).Lanr, maxL) & "'", True: rNa(k).Lanr = left$(rNa(k).Lanr, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rNa.BStNr: '" & rNa(k).BStNr & "' -> '" & left$(rNa(k).BStNr, maxL) & "'", True: rNa(k).BStNr = left$(rNa(k).BStNr, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rNa.Titel: '" & rNa(k).Titel & "' -> '" & left$(rNa(k).Titel, maxL) & "'", True: rNa(k).Titel = left$(rNa(k).Titel, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rNa.Versichertennummer: '" & rNa(k).Versichertennummer & "' -> '" & left$(rNa(k).Versichertennummer, maxL) & "'", True: rNa(k).Versichertennummer = left$(rNa(k).Versichertennummer, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel: '" & rNa(k).PrivatTel & "' -> '" & left$(rNa(k).PrivatTel, maxL) & "'", True: rNa(k).PrivatTel = left$(rNa(k).PrivatTel, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr: '" & rNa(k).KVNr & "' -> '" & left$(rNa(k).KVNr, maxL) & "'", True: rNa(k).KVNr = left$(rNa(k).KVNr, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr2: '" & rNa(k).KVNr2 & "' -> '" & left$(rNa(k).KVNr2, maxL) & "'", True: rNa(k).KVNr2 = left$(rNa(k).KVNr2, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr3: '" & rNa(k).KVNr3 & "' -> '" & left$(rNa(k).KVNr3, maxL) & "'", True: rNa(k).KVNr3 = left$(rNa(k).KVNr3, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rNa.KVNr4: '" & rNa(k).KVNr4 & "' -> '" & left$(rNa(k).KVNr4, maxL) & "'", True: rNa(k).KVNr4 = left$(rNa(k).KVNr4, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatTel_2: '" & rNa(k).PrivatTel_2 & "' -> '" & left$(rNa(k).PrivatTel_2, maxL) & "'", True: rNa(k).PrivatTel_2 = left$(rNa(k).PrivatTel_2, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatFax: '" & rNa(k).PrivatFax & "' -> '" & left$(rNa(k).PrivatFax, maxL) & "'", True: rNa(k).PrivatFax = left$(rNa(k).PrivatFax, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rNa.DienstTel: '" & rNa(k).DienstTel & "' -> '" & left$(rNa(k).DienstTel, maxL) & "'", True: rNa(k).DienstTel = left$(rNa(k).DienstTel, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rNa.PrivatMobil: '" & rNa(k).PrivatMobil & "' -> '" & left$(rNa(k).PrivatMobil, maxL) & "'", True: rNa(k).PrivatMobil = left$(rNa(k).PrivatMobil, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rNa.Email: '" & rNa(k).email & "' -> '" & left$(rNa(k).email, maxL) & "'", True: rNa(k).email = left$(rNa(k).email, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rNa.Arbeitgeber: '" & rNa(k).Arbeitgeber & "' -> '" & left$(rNa(k).Arbeitgeber, maxL) & "'", True: rNa(k).Arbeitgeber = left$(rNa(k).Arbeitgeber, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rNa.DMTypaD: '" & rNa(k).DMTypaD & "' -> '" & left$(rNa(k).DMTypaD, maxL) & "'", True: rNa(k).DMTypaD = left$(rNa(k).DMTypaD, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rNa.Cave: '" & rNa(k).Cave & "' -> '" & left$(rNa(k).Cave, maxL) & "'", True: rNa(k).Cave = left$(rNa(k).Cave, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rNa.notiz: '" & rNa(k).notiz & "' -> '" & left$(rNa(k).notiz, maxL) & "'", True: rNa(k).notiz = left$(rNa(k).notiz, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rNa.info: '" & rNa(k).info & "' -> '" & left$(rNa(k).info, maxL) & "'", True: rNa(k).info = left$(rNa(k).info, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rNa.obChk: '" & rNa(k).obChk & "' -> '" & left$(rNa(k).obChk, maxL) & "'", True: rNa(k).obChk = left$(rNa(k).obChk, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA0: '" & rNa(k).fnHA0 & "' -> '" & left$(rNa(k).fnHA0, maxL) & "'", True: rNa(k).fnHA0 = left$(rNa(k).fnHA0, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA1: '" & rNa(k).fnHA1 & "' -> '" & left$(rNa(k).fnHA1, maxL) & "'", True: rNa(k).fnHA1 = left$(rNa(k).fnHA1, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rNa.fnHA2: '" & rNa(k).fnHA2 & "' -> '" & left$(rNa(k).fnHA2, maxL) & "'", True: rNa(k).fnHA2 = left$(rNa(k).fnHA2, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rNa.zubenach: '" & rNa(k).zubenach & "' -> '" & left$(rNa(k).zubenach, maxL) & "'", True: rNa(k).zubenach = left$(rNa(k).zubenach, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rNa.Verwandt: '" & rNa(k).Verwandt & "' -> '" & left$(rNa(k).Verwandt, maxL) & "'", True: rNa(k).Verwandt = left$(rNa(k).Verwandt, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rNa.Sprache: '" & rNa(k).Sprache & "' -> '" & left$(rNa(k).Sprache, maxL) & "'", True: rNa(k).Sprache = left$(rNa(k).Sprache, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
ElseIf ErrNumber = -2147217871 Or ErrNumber = -2147217859 Or ErrNumber = -2147467259 Then
 For i = 0 To 10
  Call ForeignYes0
  Call ForeignYes1
 Next i
 Call ForeignNo0
 Call ForeignNo1
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in namenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' namenSpeichern

Public Function roFaZuw(i&, j&)
 roFa(i).FID = rFa(j).FID
 roFa(i).Pat_ID = rFa(j).Pat_ID
 roFa(i).Quartal = rFa(j).Quartal
 roFa(i).Nachname = rFa(j).Nachname
 roFa(i).Vorname = rFa(j).Vorname
 roFa(i).DtlOnlPfg = rFa(j).DtlOnlPfg
 roFa(i).ErgbdOnlP = rFa(j).ErgbdOnlP
 roFa(i).ErrorCode = rFa(j).ErrorCode
 roFa(i).PrüfZdFd = rFa(j).PrüfZdFd
 roFa(i).lfdnr = rFa(j).lfdnr
 roFa(i).TMFNr = rFa(j).TMFNr
 roFa(i).VKNr = rFa(j).VKNr
 roFa(i).bPerG = rFa(j).bPerG
 roFa(i).DMPKnZ = rFa(j).DMPKnZ
 roFa(i).VschBeg = rFa(j).VschBeg
 roFa(i).VschEnd = rFa(j).VschEnd
 roFa(i).KKasse_2 = rFa(j).KKasse_2
 roFa(i).FaktPers = rFa(j).FaktPers
 roFa(i).FaktTechn = rFa(j).FaktTechn
 roFa(i).FaktLabor = rFa(j).FaktLabor
 roFa(i).BhFB = rFa(j).BhFB
 roFa(i).BhFE1 = rFa(j).BhFE1
 roFa(i).BhFE2 = rFa(j).BhFE2
 roFa(i).UnfFlg = rFa(j).UnfFlg
 roFa(i).ausgst = rFa(j).ausgst
 roFa(i).KtrAbrB = rFa(j).KtrAbrB
 roFa(i).AbrAr = rFa(j).AbrAr
 roFa(i).lVorl = rFa(j).lVorl
 roFa(i).KartBes = rFa(j).KartBes
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
 roFa(i).SktBem = rFa(j).SktBem
 roFa(i).letzteRegel = rFa(j).letzteRegel
 roFa(i).ÜwText = rFa(j).ÜwText
 roFa(i).AkfLues = rFa(j).AkfLues
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
 roFa(i).KHNam = rFa(j).KHNam
 roFa(i).statBehTage = rFa(j).statBehTage
 roFa(i).SchGr = rFa(j).SchGr
 roFa(i).Weiterbeh = rFa(j).Weiterbeh
 roFa(i).KurAbb = rFa(j).KurAbb
 roFa(i).VermiArt = rFa(j).VermiArt
 roFa(i).VermiCode = rFa(j).VermiCode
 roFa(i).VermiDatum = rFa(j).VermiDatum
 roFa(i).VermiZusatz = rFa(j).VermiZusatz
 roFa(i).PGeb = rFa(j).PGeb
 roFa(i).PGebErg = rFa(j).PGebErg
 roFa(i).Mahnfrist = rFa(j).Mahnfrist
 roFa(i).Unfallort = rFa(j).Unfallort
 roFa(i).BeschAls = rFa(j).BeschAls
 roFa(i).BeschSeit = rFa(j).BeschSeit
 roFa(i).Unfallbetrieb = rFa(j).Unfallbetrieb
 roFa(i).bHeilb = rFa(j).bHeilb
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
 roFa(i).aktZeit = rFa(j).aktZeit
 roFa(i).Fanf = rFa(j).Fanf
 roFa(i).altQuart = rFa(j).altQuart
 roFa(i).QAnf = rFa(j).QAnf
 roFa(i).QEnd = rFa(j).QEnd
 roFa(i).QS = rFa(j).QS
 roFa(i).QT = rFa(j).QT
 roFa(i).StByte = rFa(j).StByte
 roFa(i).absPos = rFa(j).absPos
 roFa(i).lanrid = rFa(j).lanrid
 roFa(i).ZnrMLes = rFa(j).ZnrMLes
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
 If roFa(i).Pat_ID <> rFa(j).Pat_ID Then GoSub unter
 If roFa(i).Quartal <> rFa(j).Quartal Then GoSub unter
 If roFa(i).Nachname <> rFa(j).Nachname Then GoSub unter
 If roFa(i).Vorname <> rFa(j).Vorname Then GoSub unter
 If roFa(i).DtlOnlPfg <> rFa(j).DtlOnlPfg Then GoSub unter
 If roFa(i).ErgbdOnlP <> rFa(j).ErgbdOnlP Then GoSub unter
 If roFa(i).ErrorCode <> rFa(j).ErrorCode Then GoSub unter
 If roFa(i).PrüfZdFd <> rFa(j).PrüfZdFd Then GoSub unter
 If roFa(i).lfdnr <> rFa(j).lfdnr Then GoSub unter
 If roFa(i).TMFNr <> rFa(j).TMFNr Then GoSub unter
 If roFa(i).VKNr <> rFa(j).VKNr Then GoSub unter
 If roFa(i).bPerG <> rFa(j).bPerG Then GoSub unter
 If roFa(i).DMPKnZ <> rFa(j).DMPKnZ Then GoSub unter
 If roFa(i).VschBeg <> rFa(j).VschBeg Then GoSub unter
 If roFa(i).VschEnd <> rFa(j).VschEnd Then GoSub unter
 If roFa(i).KKasse_2 <> rFa(j).KKasse_2 Then GoSub unter
 If roFa(i).FaktPers <> rFa(j).FaktPers Then GoSub unter
 If roFa(i).FaktTechn <> rFa(j).FaktTechn Then GoSub unter
 If roFa(i).FaktLabor <> rFa(j).FaktLabor Then GoSub unter
 If roFa(i).BhFB <> rFa(j).BhFB Then GoSub unter
 If roFa(i).BhFE1 <> rFa(j).BhFE1 Then GoSub unter
 If roFa(i).BhFE2 <> rFa(j).BhFE2 Then GoSub unter
 If roFa(i).UnfFlg <> rFa(j).UnfFlg Then GoSub unter
 If roFa(i).ausgst <> rFa(j).ausgst Then GoSub unter
 If roFa(i).KtrAbrB <> rFa(j).KtrAbrB Then GoSub unter
 If roFa(i).AbrAr <> rFa(j).AbrAr Then GoSub unter
 If roFa(i).lVorl <> rFa(j).lVorl Then GoSub unter
 If roFa(i).KartBes <> rFa(j).KartBes Then GoSub unter
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
 If roFa(i).SktBem <> rFa(j).SktBem Then GoSub unter
 If roFa(i).letzteRegel <> rFa(j).letzteRegel Then GoSub unter
 If roFa(i).ÜwText <> rFa(j).ÜwText Then GoSub unter
 If roFa(i).AkfLues <> rFa(j).AkfLues Then GoSub unter
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
 If roFa(i).KHNam <> rFa(j).KHNam Then GoSub unter
 If roFa(i).statBehTage <> rFa(j).statBehTage Then GoSub unter
 If roFa(i).SchGr <> rFa(j).SchGr Then GoSub unter
 If roFa(i).Weiterbeh <> rFa(j).Weiterbeh Then GoSub unter
 If roFa(i).KurAbb <> rFa(j).KurAbb Then GoSub unter
 If roFa(i).VermiArt <> rFa(j).VermiArt Then GoSub unter
 If roFa(i).VermiCode <> rFa(j).VermiCode Then GoSub unter
 If roFa(i).VermiDatum <> rFa(j).VermiDatum Then GoSub unter
 If roFa(i).VermiZusatz <> rFa(j).VermiZusatz Then GoSub unter
 If roFa(i).PGeb <> rFa(j).PGeb Then GoSub unter
 If roFa(i).PGebErg <> rFa(j).PGebErg Then GoSub unter
 If roFa(i).Mahnfrist <> rFa(j).Mahnfrist Then GoSub unter
 If roFa(i).Unfallort <> rFa(j).Unfallort Then GoSub unter
 If roFa(i).BeschAls <> rFa(j).BeschAls Then GoSub unter
 If roFa(i).BeschSeit <> rFa(j).BeschSeit Then GoSub unter
 If roFa(i).Unfallbetrieb <> rFa(j).Unfallbetrieb Then GoSub unter
 If roFa(i).bHeilb <> rFa(j).bHeilb Then GoSub unter
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
 If roFa(i).aktZeit <> rFa(j).aktZeit Then GoSub unter
 If roFa(i).Fanf <> rFa(j).Fanf Then GoSub unter
 If roFa(i).altQuart <> rFa(j).altQuart Then GoSub unter
 If roFa(i).QAnf <> rFa(j).QAnf Then GoSub unter
 If roFa(i).QEnd <> rFa(j).QEnd Then GoSub unter
 If roFa(i).QS <> rFa(j).QS Then GoSub unter
 If roFa(i).QT <> rFa(j).QT Then GoSub unter
 If roFa(i).StByte <> rFa(j).StByte Then GoSub unter
 If roFa(i).absPos <> rFa(j).absPos Then GoSub unter
 If roFa(i).lanrid <> rFa(j).lanrid Then GoSub unter
 If roFa(i).ZnrMLes <> rFa(j).ZnrMLes Then GoSub unter
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
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(Quartal,'') Quartal,COALESCE(Nachname,'') Nachname" & _
",COALESCE(Vorname,'') Vorname,COALESCE(DtlOnlPfg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DtlOnlPfg,COALESCE(ErgbdOnlP,0) ErgbdOnlP,COALESCE(ErrorCode,0) ErrorCode" & _
",COALESCE(PrüfZdFd,'') PrüfZdFd,COALESCE(lfdnr,0) lfdnr,COALESCE(TMFNr,'') TMFNr,COALESCE(VKNr,'') VKNr" & _
",COALESCE(bPerG,'') bPerG,COALESCE(DMPKnZ,'') DMPKnZ,COALESCE(VschBeg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) VschBeg,COALESCE(VschEnd - INTERVAL 0 DAY,CONVERT('18991230',DATE)) VschEnd" & _
",COALESCE(KKasse_2,'') KKasse_2,COALESCE(FaktPers,0) FaktPers,COALESCE(FaktTechn,0) FaktTechn,COALESCE(FaktLabor,0) FaktLabor" & _
",COALESCE(BhFB - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFB,COALESCE(BhFE1 - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE1,COALESCE(BhFE2 - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE2,COALESCE(UnfFlg,'') UnfFlg" & _
",COALESCE(ausgst - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ausgst,COALESCE(KtrAbrB,'') KtrAbrB,COALESCE(AbrAr,'') AbrAr,COALESCE(lVorl - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lVorl" & _
",COALESCE(KartBes,0) KartBes,COALESCE(IK,'') IK,COALESCE(KVKs,'') KVKs,COALESCE(KVKserg,'') KVKserg" & _
",COALESCE(Status,'') Status,COALESCE(Kasse,'') Kasse,COALESCE(KID,0) KID,COALESCE(GebOr,'') GebOr" & _
",COALESCE(AbrGb,'') AbrGb,COALESCE(PersKreis,'') PersKreis,COALESCE(SKtZusatz,'') SKtZusatz,COALESCE(SktBem,'') SktBem" & _
",COALESCE(letzteRegel,'') letzteRegel,COALESCE(ÜwText,'') ÜwText,COALESCE(AkfLues,0) AkfLues,COALESCE(AkfHAH,0) AkfHAH" & _
",COALESCE(AkfAB0,0) AkfAB0,COALESCE(AkfAK,0) AkfAK,COALESCE(statNuller,'') statNuller,COALESCE(ÜbwV,'') ÜbwV" & _
",COALESCE(ÜbWVLANR,'') ÜbWVLANR,COALESCE(ÜbWVBSNR,'') ÜbWVBSNR,COALESCE(ÜbWVKVNR,'') ÜbWVKVNR,COALESCE(AndÜw,'') AndÜw" & _
",COALESCE(Übwr,'') Übwr,COALESCE(ÜbwLANR,'') ÜbwLANR,COALESCE(ÜWZiel,'') ÜWZiel,COALESCE(ÜWNNr,'') ÜWNNr" & _
",COALESCE(ÜWNaN,'') ÜWNaN,COALESCE(ÜWTit,'') ÜWTit,COALESCE(ÜWVor,'') ÜWVor,COALESCE(ÜWVsw,'') ÜWVsw" & _
",COALESCE(üwvid,0) üwvid,COALESCE(Auftrag,'') Auftrag,COALESCE(Verdacht,'') Verdacht,COALESCE(Befund,'') Befund" & _
",COALESCE(statKlasse,'') statKlasse,COALESCE(KHNam,'') KHNam,COALESCE(statBehTage,0) statBehTage,COALESCE(SchGr,0) SchGr" & _
",COALESCE(Weiterbeh,'') Weiterbeh,COALESCE(KurAbb,0) KurAbb,COALESCE(VermiArt,'') VermiArt,COALESCE(VermiCode,0) VermiCode" & _
",COALESCE(VermiDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) VermiDatum,COALESCE(VermiZusatz,'') VermiZusatz,COALESCE(PGeb,'') PGeb,COALESCE(PGebErg,'') PGebErg" & _
",COALESCE(Mahnfrist,'') Mahnfrist,COALESCE(Unfallort,'') Unfallort,COALESCE(BeschAls,'') BeschAls,COALESCE(BeschSeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BeschSeit" & _
",COALESCE(Unfallbetrieb,'') Unfallbetrieb,COALESCE(bHeilb,'') bHeilb,COALESCE(GOÄKatNr,'') GOÄKatNr,COALESCE(GOÄKatName,'') GOÄKatName" & _
",COALESCE(abrArzt,'') abrArzt,COALESCE(privVers,'') privVers,COALESCE(AdNam,'') AdNam,COALESCE(AdStr,'') AdStr" & _
",COALESCE(AdPlz,'') AdPlz,COALESCE(AdOrt,'') AdOrt,COALESCE(ÜwBG,'') ÜwBG,COALESCE(BhFE - INTERVAL 0 DAY,CONVERT('18991230',DATE)) BhFE" & _
",COALESCE(s8000,'') s8000,COALESCE(s8100,'') s8100,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(Fanf - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Fanf" & _
",COALESCE(altQuart,'') altQuart,COALESCE(QAnf - INTERVAL 0 DAY,CONVERT('18991230',DATE)) QAnf,COALESCE(QEnd - INTERVAL 0 DAY,CONVERT('18991230',DATE)) QEnd,COALESCE(QS,'') QS"
sql = sql & ",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(absPos,0) absPos,COALESCE(LANRid,0) LANRid" & _
",COALESCE(ZnrMLes,'') ZnrMLes,COALESCE(BGFallNr,'') BGFallNr,COALESCE(lGewicht,0) lGewicht,COALESCE(vorET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) vorET" & _
",COALESCE(dmpVertret,'') dmpVertret,COALESCE(dmpArztw,'') dmpArztw,COALESCE(dmpHypos,'') dmpHypos,COALESCE(dmpKhsA,'') dmpKhsA" & _
",COALESCE(dmpDMSchulEmpf,'') dmpDMSchulEmpf,COALESCE(dmpDMSchulWahrg,'') dmpDMSchulWahrg,COALESCE(dmpHypertSchulEmpf,'') dmpHypertSchulEmpf,COALESCE(dmpHypertSchulWahrg,'') dmpHypertSchulWahrg" & _
",COALESCE(dmpKKTabakEmpf,'') dmpKKTabakEmpf,COALESCE(dmpKKErnEmpf,'') dmpKKErnEmpf,COALESCE(dmpKKkTrainEmpf,'') dmpKKkTrainEmpf,COALESCE(dmpHbA1cZiel,'') dmpHbA1cZiel" & _
",COALESCE(dmpUewFuss,'') dmpUewFuss,COALESCE(dmpEinwDM,'') dmpEinwDM,COALESCE(dmphalbj,'') dmphalbj,COALESCE(dmpMA,'') dmpMA" & _
" FROM `faelle` WHERE Pat_ID=" & pid & " ORDER BY `fanf`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roFa(0)
 Else ' rs.EOF Then
  ReDim roFa(1)
  Do While Not rs.EOF
   akt = UBound(roFa)
   roFa(akt).FID = rs!FID
   roFa(akt).Pat_ID = rs!Pat_ID
   roFa(akt).Quartal = doUmwfSQL(rs!Quartal, lies.obMySQL, False)
   roFa(akt).Nachname = doUmwfSQL(rs!Nachname, lies.obMySQL, False)
   roFa(akt).Vorname = doUmwfSQL(rs!Vorname, lies.obMySQL, False)
   roFa(akt).DtlOnlPfg = rs!DtlOnlPfg
   roFa(akt).ErgbdOnlP = rs!ErgbdOnlP
   roFa(akt).ErrorCode = rs!ErrorCode
   roFa(akt).PrüfZdFd = doUmwfSQL(rs!PrüfZdFd, lies.obMySQL, False)
   roFa(akt).lfdnr = rs!lfdnr
   roFa(akt).TMFNr = doUmwfSQL(rs!TMFNr, lies.obMySQL, False)
   roFa(akt).VKNr = doUmwfSQL(rs!VKNr, lies.obMySQL, False)
   roFa(akt).bPerG = doUmwfSQL(rs!bPerG, lies.obMySQL, False)
   roFa(akt).DMPKnZ = doUmwfSQL(rs!DMPKnZ, lies.obMySQL, False)
   roFa(akt).VschBeg = rs!VschBeg
   roFa(akt).VschEnd = rs!VschEnd
   roFa(akt).KKasse_2 = doUmwfSQL(rs!KKasse_2, lies.obMySQL, False)
   roFa(akt).FaktPers = rs!FaktPers
   roFa(akt).FaktTechn = rs!FaktTechn
   roFa(akt).FaktLabor = rs!FaktLabor
   roFa(akt).BhFB = rs!BhFB
   roFa(akt).BhFE1 = rs!BhFE1
   roFa(akt).BhFE2 = rs!BhFE2
   roFa(akt).UnfFlg = doUmwfSQL(rs!UnfFlg, lies.obMySQL, False)
   roFa(akt).ausgst = rs!ausgst
   roFa(akt).KtrAbrB = doUmwfSQL(rs!KtrAbrB, lies.obMySQL, False)
   roFa(akt).AbrAr = doUmwfSQL(rs!AbrAr, lies.obMySQL, False)
   roFa(akt).lVorl = rs!lVorl
   roFa(akt).KartBes = rs!KartBes
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
   roFa(akt).SktBem = doUmwfSQL(rs!SktBem, lies.obMySQL, False)
   roFa(akt).letzteRegel = doUmwfSQL(rs!letzteRegel, lies.obMySQL, False)
   roFa(akt).ÜwText = doUmwfSQL(rs!ÜwText, lies.obMySQL, False)
   roFa(akt).AkfLues = rs!AkfLues
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
   roFa(akt).KHNam = doUmwfSQL(rs!KHNam, lies.obMySQL, False)
   roFa(akt).statBehTage = rs!statBehTage
   roFa(akt).SchGr = rs!SchGr
   roFa(akt).Weiterbeh = doUmwfSQL(rs!Weiterbeh, lies.obMySQL, False)
   roFa(akt).KurAbb = rs!KurAbb
   roFa(akt).VermiArt = doUmwfSQL(rs!VermiArt, lies.obMySQL, False)
   roFa(akt).VermiCode = rs!VermiCode
   roFa(akt).VermiDatum = rs!VermiDatum
   roFa(akt).VermiZusatz = doUmwfSQL(rs!VermiZusatz, lies.obMySQL, False)
   roFa(akt).PGeb = doUmwfSQL(rs!PGeb, lies.obMySQL, False)
   roFa(akt).PGebErg = doUmwfSQL(rs!PGebErg, lies.obMySQL, False)
   roFa(akt).Mahnfrist = doUmwfSQL(rs!Mahnfrist, lies.obMySQL, False)
   roFa(akt).Unfallort = doUmwfSQL(rs!Unfallort, lies.obMySQL, False)
   roFa(akt).BeschAls = doUmwfSQL(rs!BeschAls, lies.obMySQL, False)
   roFa(akt).BeschSeit = rs!BeschSeit
   roFa(akt).Unfallbetrieb = doUmwfSQL(rs!Unfallbetrieb, lies.obMySQL, False)
   roFa(akt).bHeilb = doUmwfSQL(rs!bHeilb, lies.obMySQL, False)
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
   roFa(akt).aktZeit = rs!aktZeit
   roFa(akt).Fanf = rs!Fanf
   roFa(akt).altQuart = doUmwfSQL(rs!altQuart, lies.obMySQL, False)
   roFa(akt).QAnf = rs!QAnf
   roFa(akt).QEnd = rs!QEnd
   roFa(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roFa(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roFa(akt).StByte = rs!StByte
   roFa(akt).absPos = rs!absPos
   roFa(akt).lanrid = rs!lanrid
   roFa(akt).ZnrMLes = doUmwfSQL(rs!ZnrMLes, lies.obMySQL, False)
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
 Dim rbeg&, roendpe&, ri&, roi&, jj&, fazu%()
 Dim fzu() As fzu
 On Error GoTo fehler
 If UBound(rFa) > 0 Then ' wenn neue Fälle da
  For ri = 1 To UBound(rFa)
   If rFa(ri).Fanf >= qbeg Then ' aktqanf()
    rbeg = ri ' dann bezeichnet rbeg den ersten zu importierenden Fall
    For roendpe = 0 To UBound(roFa)
     If roFa(roendpe).Fanf >= qbeg Then
      Exit For ' ... und roFa den ersten nicht mehr zu verwendenden bestehenden Fall,
     End If ' roFa(roendpe).Fanf >= qbeg Then
    Next roendpe
    Exit For
   End If ' rFa(ri).Fanf >= qbeg Then ' aktqanf()
  Next ri
 End If ' UBound(rFa) > 0
 If rbeg = 0 Then
  roendpe = UBound(roFa) + 1 ', ggf. alle aus roFa zu verwenden, und nur die,
 Else ' rbeg = 0 Then
 ' ermitteln, welcher Fallnummer aus rFa jeder zu löschende Fall aus rFo entspricht
  If roendpe <= UBound(roFa) Then
   ReDim fazu(roendpe To UBound(roFa))
   For roi = roendpe To UBound(roFa)
    For ri = rbeg To UBound(rFa) - 1
     If roFa(roi).Fanf >= rFa(ri).Fanf And roFa(roi).Fanf < rFa(ri).BhFE Then
      fazu(roi) = ri
      GoTo Fertig
     End If ' roFa(roi).Fanf >= rFa(ri).Fanf And roFa(roi).Fanf < rFa(ri).BhFE
    Next ri
    fazu(roi) = UBound(rFa)
Fertig:
   Next roi
 End If ' If roendpe <= UBound(roFa) Then
' die Fallnummern zu aus roFa zu löschenden Sätze durch  die aus rFa ersetzen
  For roi = roendpe To UBound(roFa)
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
   For jj = 1 To UBound(roVo)
    If roVo(jj).FID = roFa(roi).FID Then roVo(jj).FID = rFa(fazu(roi)).FID
   Next jj
  Next roi
  ReDim Preserve roFa(roendpe - 1) ' roFa kürzen ...
  ReDim Preserve roFa(roendpe + UBound(rFa) - rbeg) ' ... und erweitern ...
  For ri = rbeg To UBound(rFa)
   Call roFaZuw(roendpe + ri - rbeg, ri) ' und die Fälle aus rFa an roFa anhängen
  Next ri
 End If ' rbeg = 0 Then Else
 ReDim fzu(UBound(roFa)) ' dann fzu füllen
 For ri = 1 To UBound(roFa)
  fzu(ri).falt = roFa(ri).FID
  fzu(ri).fneu = patanffid + ri - 1
 Next ri
 For ri = UBound(roFa) To 1 Step -1 ' dann die künftigen Fallnummern statt den aktuellen verwenden
  If roFa(ri).FID <> fzu(ri).falt Then
   MsgBox "Fehler bei " & rNa(0).Pat_ID & ", ri: " & ri & ", " & roFa(ri).FID & " <> " & fzu(ri).falt
  End If ' roFa(ri).FID <> fzu(ri).falt Then
  For jj = 1 To UBound(roAu)
   If roAu(jj).FID = roFa(ri).FID Then roAu(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roBr)
   If roBr(jj).FID = roFa(ri).FID Then roBr(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roDi)
   If roDi(jj).FID = roFa(ri).FID Then roDi(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roDo)
   If roDo(jj).FID = roFa(ri).FID Then roDo(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roEi)
   If roEi(jj).FID = roFa(ri).FID Then roEi(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roFr)
   If roFr(jj).FID = roFa(ri).FID Then roFr(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roKh)
   If roKh(jj).FID = roFa(ri).FID Then roKh(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roLb)
   If roLb(jj).FID = roFa(ri).FID Then roLb(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roLa)
   If roLa(jj).FID = roFa(ri).FID Then roLa(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roLe)
   If roLe(jj).FID = roFa(ri).FID Then roLe(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roMe)
   If roMe(jj).FID = roFa(ri).FID Then roMe(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roRe)
   If roRe(jj).FID = roFa(ri).FID Then roRe(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roRr)
   If roRr(jj).FID = roFa(ri).FID Then roRr(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roUs)
   If roUs(jj).FID = roFa(ri).FID Then roUs(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roFu)
   If roFu(jj).FID = roFa(ri).FID Then roFu(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roUl)
   If roUl(jj).FID = roFa(ri).FID Then roUl(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roVk)
   If roVk(jj).FID = roFa(ri).FID Then roVk(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roSw)
   If roSw(jj).FID = roFa(ri).FID Then roSw(jj).FID = fzu(ri).fneu
  Next jj
  For jj = 1 To UBound(roVo)
   If roVo(jj).FID = roFa(ri).FID Then roVo(jj).FID = fzu(ri).fneu
  Next jj
  roFa(ri).FID = fzu(ri).fneu
 Next ri
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

Public Function rFaDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rFadump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rFa)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rFa(" & i & ").FID:" & String$(33, "."), 33) & rFa(i).FID
  Print #200, left$("rFa(" & i & ").Pat_ID:" & String$(33, "."), 33) & rFa(i).Pat_ID
  Print #200, left$("rFa(" & i & ").Quartal:" & String$(33, "."), 33) & "'" & rFa(i).Quartal & "'"
  Print #200, left$("rFa(" & i & ").Nachname:" & String$(33, "."), 33) & "'" & rFa(i).Nachname & "'"
  Print #200, left$("rFa(" & i & ").Vorname:" & String$(33, "."), 33) & "'" & rFa(i).Vorname & "'"
  Print #200, left$("rFa(" & i & ").DtlOnlPfg:" & String$(33, "."), 33) & rFa(i).DtlOnlPfg
  Print #200, left$("rFa(" & i & ").ErgbdOnlP:" & String$(33, "."), 33) & rFa(i).ErgbdOnlP
  Print #200, left$("rFa(" & i & ").ErrorCode:" & String$(33, "."), 33) & rFa(i).ErrorCode
  Print #200, left$("rFa(" & i & ").PrüfZdFd:" & String$(33, "."), 33) & "'" & rFa(i).PrüfZdFd & "'"
  Print #200, left$("rFa(" & i & ").lfdnr:" & String$(33, "."), 33) & rFa(i).lfdnr
  Print #200, left$("rFa(" & i & ").TMFNr:" & String$(33, "."), 33) & "'" & rFa(i).TMFNr & "'"
  Print #200, left$("rFa(" & i & ").VKNr:" & String$(33, "."), 33) & "'" & rFa(i).VKNr & "'"
  Print #200, left$("rFa(" & i & ").bPerG:" & String$(33, "."), 33) & "'" & rFa(i).bPerG & "'"
  Print #200, left$("rFa(" & i & ").DMPKnZ:" & String$(33, "."), 33) & "'" & rFa(i).DMPKnZ & "'"
  Print #200, left$("rFa(" & i & ").VschBeg:" & String$(33, "."), 33) & rFa(i).VschBeg
  Print #200, left$("rFa(" & i & ").VschEnd:" & String$(33, "."), 33) & rFa(i).VschEnd
  Print #200, left$("rFa(" & i & ").KKasse_2:" & String$(33, "."), 33) & "'" & rFa(i).KKasse_2 & "'"
  Print #200, left$("rFa(" & i & ").FaktPers:" & String$(33, "."), 33) & rFa(i).FaktPers
  Print #200, left$("rFa(" & i & ").FaktTechn:" & String$(33, "."), 33) & rFa(i).FaktTechn
  Print #200, left$("rFa(" & i & ").FaktLabor:" & String$(33, "."), 33) & rFa(i).FaktLabor
  Print #200, left$("rFa(" & i & ").BhFB:" & String$(33, "."), 33) & rFa(i).BhFB
  Print #200, left$("rFa(" & i & ").BhFE1:" & String$(33, "."), 33) & rFa(i).BhFE1
  Print #200, left$("rFa(" & i & ").BhFE2:" & String$(33, "."), 33) & rFa(i).BhFE2
  Print #200, left$("rFa(" & i & ").UnfFlg:" & String$(33, "."), 33) & "'" & rFa(i).UnfFlg & "'"
  Print #200, left$("rFa(" & i & ").ausgst:" & String$(33, "."), 33) & rFa(i).ausgst
  Print #200, left$("rFa(" & i & ").KtrAbrB:" & String$(33, "."), 33) & "'" & rFa(i).KtrAbrB & "'"
  Print #200, left$("rFa(" & i & ").AbrAr:" & String$(33, "."), 33) & "'" & rFa(i).AbrAr & "'"
  Print #200, left$("rFa(" & i & ").lVorl:" & String$(33, "."), 33) & rFa(i).lVorl
  Print #200, left$("rFa(" & i & ").KartBes:" & String$(33, "."), 33) & rFa(i).KartBes
  Print #200, left$("rFa(" & i & ").IK:" & String$(33, "."), 33) & "'" & rFa(i).IK & "'"
  Print #200, left$("rFa(" & i & ").KVKs:" & String$(33, "."), 33) & "'" & rFa(i).KVKs & "'"
  Print #200, left$("rFa(" & i & ").KVKserg:" & String$(33, "."), 33) & "'" & rFa(i).KVKserg & "'"
  Print #200, left$("rFa(" & i & ").Status:" & String$(33, "."), 33) & "'" & rFa(i).Status & "'"
  Print #200, left$("rFa(" & i & ").Kasse:" & String$(33, "."), 33) & "'" & rFa(i).Kasse & "'"
  Print #200, left$("rFa(" & i & ").KID:" & String$(33, "."), 33) & rFa(i).KID
  Print #200, left$("rFa(" & i & ").GebOr:" & String$(33, "."), 33) & "'" & rFa(i).GebOr & "'"
  Print #200, left$("rFa(" & i & ").AbrGb:" & String$(33, "."), 33) & "'" & rFa(i).AbrGb & "'"
  Print #200, left$("rFa(" & i & ").PersKreis:" & String$(33, "."), 33) & "'" & rFa(i).PersKreis & "'"
  Print #200, left$("rFa(" & i & ").SKtZusatz:" & String$(33, "."), 33) & "'" & rFa(i).SKtZusatz & "'"
  Print #200, left$("rFa(" & i & ").SktBem:" & String$(33, "."), 33) & "'" & rFa(i).SktBem & "'"
  Print #200, left$("rFa(" & i & ").letzteRegel:" & String$(33, "."), 33) & "'" & rFa(i).letzteRegel & "'"
  Print #200, left$("rFa(" & i & ").ÜwText:" & String$(33, "."), 33) & "'" & rFa(i).ÜwText & "'"
  Print #200, left$("rFa(" & i & ").AkfLues:" & String$(33, "."), 33) & rFa(i).AkfLues
  Print #200, left$("rFa(" & i & ").AkfHAH:" & String$(33, "."), 33) & rFa(i).AkfHAH
  Print #200, left$("rFa(" & i & ").AkfAB0:" & String$(33, "."), 33) & rFa(i).AkfAB0
  Print #200, left$("rFa(" & i & ").AkfAK:" & String$(33, "."), 33) & rFa(i).AkfAK
  Print #200, left$("rFa(" & i & ").statNuller:" & String$(33, "."), 33) & "'" & rFa(i).statNuller & "'"
  Print #200, left$("rFa(" & i & ").ÜbwV:" & String$(33, "."), 33) & "'" & rFa(i).ÜbwV & "'"
  Print #200, left$("rFa(" & i & ").ÜbWVLANR:" & String$(33, "."), 33) & "'" & rFa(i).ÜbWVLANR & "'"
  Print #200, left$("rFa(" & i & ").ÜbWVBSNR:" & String$(33, "."), 33) & "'" & rFa(i).ÜbWVBSNR & "'"
  Print #200, left$("rFa(" & i & ").ÜbWVKVNR:" & String$(33, "."), 33) & "'" & rFa(i).ÜbWVKVNR & "'"
  Print #200, left$("rFa(" & i & ").AndÜw:" & String$(33, "."), 33) & "'" & rFa(i).AndÜw & "'"
  Print #200, left$("rFa(" & i & ").Übwr:" & String$(33, "."), 33) & "'" & rFa(i).Übwr & "'"
  Print #200, left$("rFa(" & i & ").ÜbwLANR:" & String$(33, "."), 33) & "'" & rFa(i).ÜbwLANR & "'"
  Print #200, left$("rFa(" & i & ").ÜWZiel:" & String$(33, "."), 33) & "'" & rFa(i).ÜWZiel & "'"
  Print #200, left$("rFa(" & i & ").ÜWNNr:" & String$(33, "."), 33) & "'" & rFa(i).ÜWNNr & "'"
  Print #200, left$("rFa(" & i & ").ÜWNaN:" & String$(33, "."), 33) & "'" & rFa(i).ÜWNaN & "'"
  Print #200, left$("rFa(" & i & ").ÜWTit:" & String$(33, "."), 33) & "'" & rFa(i).ÜWTit & "'"
  Print #200, left$("rFa(" & i & ").ÜWVor:" & String$(33, "."), 33) & "'" & rFa(i).ÜWVor & "'"
  Print #200, left$("rFa(" & i & ").ÜWVsw:" & String$(33, "."), 33) & "'" & rFa(i).ÜWVsw & "'"
  Print #200, left$("rFa(" & i & ").üwvid:" & String$(33, "."), 33) & rFa(i).üwvid
  Print #200, left$("rFa(" & i & ").Auftrag:" & String$(33, "."), 33) & "'" & rFa(i).Auftrag & "'"
  Print #200, left$("rFa(" & i & ").Verdacht:" & String$(33, "."), 33) & "'" & rFa(i).Verdacht & "'"
  Print #200, left$("rFa(" & i & ").Befund:" & String$(33, "."), 33) & "'" & rFa(i).Befund & "'"
  Print #200, left$("rFa(" & i & ").statKlasse:" & String$(33, "."), 33) & "'" & rFa(i).statKlasse & "'"
  Print #200, left$("rFa(" & i & ").KHNam:" & String$(33, "."), 33) & "'" & rFa(i).KHNam & "'"
  Print #200, left$("rFa(" & i & ").statBehTage:" & String$(33, "."), 33) & rFa(i).statBehTage
  Print #200, left$("rFa(" & i & ").SchGr:" & String$(33, "."), 33) & "'" & rFa(i).SchGr & "'"
  Print #200, left$("rFa(" & i & ").Weiterbeh:" & String$(33, "."), 33) & "'" & rFa(i).Weiterbeh & "'"
  Print #200, left$("rFa(" & i & ").KurAbb:" & String$(33, "."), 33) & rFa(i).KurAbb
  Print #200, left$("rFa(" & i & ").VermiArt:" & String$(33, "."), 33) & "'" & rFa(i).VermiArt & "'"
  Print #200, left$("rFa(" & i & ").VermiCode:" & String$(33, "."), 33) & rFa(i).VermiCode
  Print #200, left$("rFa(" & i & ").VermiDatum:" & String$(33, "."), 33) & rFa(i).VermiDatum
  Print #200, left$("rFa(" & i & ").VermiZusatz:" & String$(33, "."), 33) & "'" & rFa(i).VermiZusatz & "'"
  Print #200, left$("rFa(" & i & ").PGeb:" & String$(33, "."), 33) & "'" & rFa(i).PGeb & "'"
  Print #200, left$("rFa(" & i & ").PGebErg:" & String$(33, "."), 33) & "'" & rFa(i).PGebErg & "'"
  Print #200, left$("rFa(" & i & ").Mahnfrist:" & String$(33, "."), 33) & "'" & rFa(i).Mahnfrist & "'"
  Print #200, left$("rFa(" & i & ").Unfallort:" & String$(33, "."), 33) & "'" & rFa(i).Unfallort & "'"
  Print #200, left$("rFa(" & i & ").BeschAls:" & String$(33, "."), 33) & "'" & rFa(i).BeschAls & "'"
  Print #200, left$("rFa(" & i & ").BeschSeit:" & String$(33, "."), 33) & rFa(i).BeschSeit
  Print #200, left$("rFa(" & i & ").Unfallbetrieb:" & String$(33, "."), 33) & "'" & rFa(i).Unfallbetrieb & "'"
  Print #200, left$("rFa(" & i & ").bHeilb:" & String$(33, "."), 33) & "'" & rFa(i).bHeilb & "'"
  Print #200, left$("rFa(" & i & ").GOÄKatNr:" & String$(33, "."), 33) & "'" & rFa(i).GOÄKatNr & "'"
  Print #200, left$("rFa(" & i & ").GOÄKatName:" & String$(33, "."), 33) & "'" & rFa(i).GOÄKatName & "'"
  Print #200, left$("rFa(" & i & ").abrArzt:" & String$(33, "."), 33) & "'" & rFa(i).abrArzt & "'"
  Print #200, left$("rFa(" & i & ").privVers:" & String$(33, "."), 33) & "'" & rFa(i).privVers & "'"
  Print #200, left$("rFa(" & i & ").AdNam:" & String$(33, "."), 33) & "'" & rFa(i).AdNam & "'"
  Print #200, left$("rFa(" & i & ").AdStr:" & String$(33, "."), 33) & "'" & rFa(i).AdStr & "'"
  Print #200, left$("rFa(" & i & ").AdPlz:" & String$(33, "."), 33) & "'" & rFa(i).AdPlz & "'"
  Print #200, left$("rFa(" & i & ").AdOrt:" & String$(33, "."), 33) & "'" & rFa(i).AdOrt & "'"
  Print #200, left$("rFa(" & i & ").ÜwBG:" & String$(33, "."), 33) & "'" & rFa(i).ÜwBG & "'"
  Print #200, left$("rFa(" & i & ").BhFE:" & String$(33, "."), 33) & rFa(i).BhFE
  Print #200, left$("rFa(" & i & ").s8000:" & String$(33, "."), 33) & "'" & rFa(i).s8000 & "'"
  Print #200, left$("rFa(" & i & ").s8100:" & String$(33, "."), 33) & "'" & rFa(i).s8100 & "'"
  Print #200, left$("rFa(" & i & ").AktZeit:" & String$(33, "."), 33) & rFa(i).aktZeit
  Print #200, left$("rFa(" & i & ").Fanf:" & String$(33, "."), 33) & rFa(i).Fanf
  Print #200, left$("rFa(" & i & ").altQuart:" & String$(33, "."), 33) & "'" & rFa(i).altQuart & "'"
  Print #200, left$("rFa(" & i & ").QAnf:" & String$(33, "."), 33) & rFa(i).QAnf
  Print #200, left$("rFa(" & i & ").QEnd:" & String$(33, "."), 33) & rFa(i).QEnd
  Print #200, left$("rFa(" & i & ").QS:" & String$(33, "."), 33) & "'" & rFa(i).QS & "'"
  Print #200, left$("rFa(" & i & ").QT:" & String$(33, "."), 33) & "'" & rFa(i).QT & "'"
  Print #200, left$("rFa(" & i & ").StByte:" & String$(33, "."), 33) & rFa(i).StByte
  Print #200, left$("rFa(" & i & ").absPos:" & String$(33, "."), 33) & rFa(i).absPos
  Print #200, left$("rFa(" & i & ").LANRid:" & String$(33, "."), 33) & rFa(i).lanrid
  Print #200, left$("rFa(" & i & ").ZnrMLes:" & String$(33, "."), 33) & "'" & rFa(i).ZnrMLes & "'"
  Print #200, left$("rFa(" & i & ").BGFallNr:" & String$(33, "."), 33) & "'" & rFa(i).BGFallNr & "'"
  Print #200, left$("rFa(" & i & ").lGewicht:" & String$(33, "."), 33) & "'" & rFa(i).lGewicht & "'"
  Print #200, left$("rFa(" & i & ").vorET:" & String$(33, "."), 33) & rFa(i).vorET
  Print #200, left$("rFa(" & i & ").dmpVertret:" & String$(33, "."), 33) & "'" & rFa(i).dmpVertret & "'"
  Print #200, left$("rFa(" & i & ").dmpArztw:" & String$(33, "."), 33) & "'" & rFa(i).dmpArztw & "'"
  Print #200, left$("rFa(" & i & ").dmpHypos:" & String$(33, "."), 33) & "'" & rFa(i).dmpHypos & "'"
  Print #200, left$("rFa(" & i & ").dmpKhsA:" & String$(33, "."), 33) & "'" & rFa(i).dmpKhsA & "'"
  Print #200, left$("rFa(" & i & ").dmpDMSchulEmpf:" & String$(33, "."), 33) & "'" & rFa(i).dmpDMSchulEmpf & "'"
  Print #200, left$("rFa(" & i & ").dmpDMSchulWahrg:" & String$(33, "."), 33) & "'" & rFa(i).dmpDMSchulWahrg & "'"
  Print #200, left$("rFa(" & i & ").dmpHypertSchulEmpf:" & String$(33, "."), 33) & "'" & rFa(i).dmpHypertSchulEmpf & "'"
  Print #200, left$("rFa(" & i & ").dmpHypertSchulWahrg:" & String$(33, "."), 33) & "'" & rFa(i).dmpHypertSchulWahrg & "'"
  Print #200, left$("rFa(" & i & ").dmpKKTabakEmpf:" & String$(33, "."), 33) & "'" & rFa(i).dmpKKTabakEmpf & "'"
  Print #200, left$("rFa(" & i & ").dmpKKErnEmpf:" & String$(33, "."), 33) & "'" & rFa(i).dmpKKErnEmpf & "'"
  Print #200, left$("rFa(" & i & ").dmpKKkTrainEmpf:" & String$(33, "."), 33) & "'" & rFa(i).dmpKKkTrainEmpf & "'"
  Print #200, left$("rFa(" & i & ").dmpHbA1cZiel:" & String$(33, "."), 33) & "'" & rFa(i).dmpHbA1cZiel & "'"
  Print #200, left$("rFa(" & i & ").dmpUewFuss:" & String$(33, "."), 33) & "'" & rFa(i).dmpUewFuss & "'"
  Print #200, left$("rFa(" & i & ").dmpEinwDM:" & String$(33, "."), 33) & "'" & rFa(i).dmpEinwDM & "'"
  Print #200, left$("rFa(" & i & ").dmphalbj:" & String$(33, "."), 33) & "'" & rFa(i).dmphalbj & "'"
  Print #200, left$("rFa(" & i & ").dmpMA:" & String$(33, "."), 33) & "'" & rFa(i).dmpMA & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' faelleDump

Public Function faelleSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "faelle")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
Dim j%
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFa) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Pat_ID,Quartal,Nachname," & _
     "Vorname,DtlOnlPfg,ErgbdOnlP,ErrorCode,PrüfZdFd,lfdnr,TMFNr,VKNr,bPerG,DMPKnZ," & _
     "VschBeg,VschEnd,KKasse_2,FaktPers,FaktTechn,FaktLabor,BhFB,BhFE1,BhFE2,UnfFlg," & _
     "ausgst,KtrAbrB,AbrAr,lVorl,KartBes,IK,KVKs,KVKserg,Status,Kasse," & _
     "KID,GebOr,AbrGb,PersKreis,SKtZusatz,SktBem,letzteRegel,ÜwText,AkfLues,AkfHAH," & _
     "AkfAB0,AkfAK,statNuller,ÜbwV,ÜbWVLANR,ÜbWVBSNR,ÜbWVKVNR,AndÜw,Übwr,ÜbwLANR," & _
     "ÜWZiel,ÜWNNr,ÜWNaN,ÜWTit,ÜWVor,ÜWVsw,üwvid,Auftrag,Verdacht,Befund," & _
     "statKlasse,KHNam,statBehTage,SchGr,Weiterbeh,KurAbb,VermiArt,VermiCode,VermiDatum,VermiZusatz," & _
     "PGeb,PGebErg,Mahnfrist,Unfallort,BeschAls,BeschSeit,Unfallbetrieb,bHeilb,GOÄKatNr,GOÄKatName," & _
     "abrArzt,privVers,AdNam,AdStr,AdPlz,AdOrt,ÜwBG,BhFE,s8000,s8100," & _
     "AktZeit,Fanf,altQuart,QAnf,QEnd,QS,QT,StByte,absPos,LANRid," & _
     "ZnrMLes,BGFallNr,lGewicht,vorET,dmpVertret,dmpArztw,dmpHypos,dmpKhsA,dmpDMSchulEmpf,dmpDMSchulWahrg," & _
     "dmpHypertSchulEmpf,dmpHypertSchulWahrg,dmpKKTabakEmpf,dmpKKErnEmpf,dmpKKkTrainEmpf,dmpHbA1cZiel,dmpUewFuss,dmpEinwDM,dmphalbj,dmpMA)   VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rFa)
  rFa(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rFa(i).Pat_ID, ",'", rFa(i).Quartal, "','", rFa(i).Nachname, "','", rFa(i).Vorname, "',", DatFor_k(rFa(i).DtlOnlPfg), ",", rFa(i).ErgbdOnlP, ",", rFa(i).ErrorCode, ",'", rFa(i).PrüfZdFd, "',", _
   rFa(i).lfdnr, ",'", rFa(i).TMFNr, "','", rFa(i).VKNr, "','", rFa(i).bPerG, "','", rFa(i).DMPKnZ, "',", DatFor_k(rFa(i).VschBeg), ",", DatFor_k(rFa(i).VschEnd), ",'", _
   rFa(i).KKasse_2, "',", REPLACE$(rFa(i).FaktPers, ",", "."), ",", REPLACE$(rFa(i).FaktTechn, ",", "."), ",", REPLACE$(rFa(i).FaktLabor, ",", "."), ",", DatFor_k(rFa(i).BhFB), ",", DatFor_k(rFa(i).BhFE1), ",", DatFor_k( _
   rFa(i).BhFE2), ",'", rFa(i).UnfFlg, "',", DatFor_k(rFa(i).ausgst), ",'", rFa(i).KtrAbrB, "','", rFa(i).AbrAr, "',", DatFor_k(rFa(i).lVorl), ",", rFa(i).KartBes, ",'", _
   rFa(i).IK, "','", rFa(i).KVKs, "','", rFa(i).KVKserg, "','", rFa(i).Status, "','", rFa(i).Kasse, "',", rFa(i).KID, ",'", rFa(i).GebOr, "','", rFa(i).AbrGb, "','", rFa(i).PersKreis, "','", _
   rFa(i).SKtZusatz, "','", rFa(i).SktBem, "','", rFa(i).letzteRegel, "','", rFa(i).ÜwText, "',", rFa(i).AkfLues, ",", rFa(i).AkfHAH, ",", rFa(i).AkfAB0, ",", rFa(i).AkfAK, ",'", _
   rFa(i).statNuller, "','", rFa(i).ÜbwV, "','", rFa(i).ÜbWVLANR, "','", rFa(i).ÜbWVBSNR, "','", rFa(i).ÜbWVKVNR, "','", rFa(i).AndÜw, "','", rFa(i).Übwr, "','", rFa(i).ÜbwLANR, "','", rFa(i).ÜWZiel, "','", _
   rFa(i).ÜWNNr, "','", rFa(i).ÜWNaN, "','", rFa(i).ÜWTit, "','", rFa(i).ÜWVor, "','", rFa(i).ÜWVsw, "',", rFa(i).üwvid, ",'", rFa(i).Auftrag, "','", rFa(i).Verdacht, "','", _
   rFa(i).Befund, "','", rFa(i).statKlasse, "','", rFa(i).KHNam, "',", rFa(i).statBehTage, ",", rFa(i).SchGr, ",'", rFa(i).Weiterbeh, "',", REPLACE$(rFa(i).KurAbb, ",", "."), ",'", rFa(i).VermiArt, "',", _
   rFa(i).VermiCode, ",", DatFor_k(rFa(i).VermiDatum), ",'", rFa(i).VermiZusatz, "','", rFa(i).PGeb, "','", rFa(i).PGebErg, "','", rFa(i).Mahnfrist, "','", rFa(i).Unfallort, "','", _
   rFa(i).BeschAls, "',", DatFor_k(rFa(i).BeschSeit), ",'", rFa(i).Unfallbetrieb, "','", rFa(i).bHeilb, "','", rFa(i).GOÄKatNr, "','", rFa(i).GOÄKatName, "','", rFa(i).abrArzt, "','", _
   rFa(i).privVers, "','", rFa(i).AdNam, "','", rFa(i).AdStr, "','", rFa(i).AdPlz, "','", rFa(i).AdOrt, "','", rFa(i).ÜwBG, "',", DatFor_k(rFa(i).BhFE), ",'", rFa(i).s8000, "','", _
   rFa(i).s8100, "',", DatFor_k(rFa(i).aktZeit), ",", DatFor_k(rFa(i).Fanf), ",'", rFa(i).altQuart, "',", DatFor_k(rFa(i).QAnf), ",", DatFor_k(rFa(i).QEnd), ",'", rFa(i).QS, "','", rFa(i).QT, "',", _
   rFa(i).StByte, ",", rFa(i).absPos, ",", rFa(i).lanrid, ",'", rFa(i).ZnrMLes, "','", rFa(i).BGFallNr, "',", rFa(i).lGewicht, ",", DatFor_k(rFa(i).vorET), ",'", rFa(i).dmpVertret, "','", _
   rFa(i).dmpArztw, "','", rFa(i).dmpHypos, "','", rFa(i).dmpKhsA, "','", rFa(i).dmpDMSchulEmpf, "','", rFa(i).dmpDMSchulWahrg, "','", rFa(i).dmpHypertSchulEmpf, "','", rFa(i).dmpHypertSchulWahrg, "','")
  csql.AppVar Array(rFa(i).dmpKKTabakEmpf, "','", rFa(i).dmpKKErnEmpf, "','", rFa(i).dmpKKkTrainEmpf, "','", rFa(i).dmpHbA1cZiel, "','", rFa(i).dmpUewFuss, "','", rFa(i).dmpEinwDM, "','", _
   rFa(i).dmphalbj, "','", rFa(i).dmpMA, "')")
  If SammelInsert <> 0 And i < UBound(rFa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFa) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rFa(i).FID = myEFrag("SELECT MAX(FID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
'   IF Not obFork THEN ForeignYes0
   If rAf = 0 Then
    Err.Raise 998, , "Fehler in faelleSpeichern b.Pat. " & rFa(i).Pat_ID & ", Err.Number " & Err.Number & ", err.description: " & Err.Description
   End If ' rAF = 0 THEN
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
   If SammelInsert = 0 Then
   'Hier gibts mit Sammelins noch ein Problem ...
'    set rs = nothing
'    For j = 2 To 2
'     If j = 1 Then
'      Set rs = myEFrag("SELECT LAST_INSERT_ID() FID") ' session-spezifisch '27.8.23: liefert in Schleife immer die erste Zahl, auch mit Commit zwischendrin
'     Else ' j = 1 Then
      Set rs = myEFrag("SELECT COALESCE((SELECT MAX(fid) FID FROM `faelle` WHERE pat_id = " & rFa(i).Pat_ID & " AND quartal = '" & rFa(i).Quartal & "' AND bhfb = " & DatFor_k(rFa(i).BhFB) & " AND bhfe1 = " & DatFor_k(rFa(i).BhFE1) & " AND ausgst = " & DatFor_k(rFa(i).ausgst) & "),(SELECT MAX(fid)+1 FID FROM `faelle`)) FID")
'     End If
'     If Not rs.BOF Then If rs.Fields(0) <> 0 Then Exit For
'    Next j
    If rs.BOF Then
     Err.Raise 999, , "Fehler bei der Fallaktualisierung b.Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID
    ElseIf rs!FID = 0 Then
     MsgBox "Fehler in faellespeichern:" & vbCrLf & rs.source
     GoTo sql
    Else ' rs.BOF Then
     neufid = rs!FID
     If neufid <> rFa(i).FID Then
      If rFa(i).FID <> 0 Then
       Lese.Ausgeb "Änderung bei der FallID  bei Pat. " & rFa(i).Pat_ID & ", FID " & rFa(i).FID & " -> " & neufid & " in zu speichernden Tabellen mit fallid", True
      End If ' rFa(i).FID <> 0 Then
      Dim jjj&
      For jjj = 1 To UBound(rAu)
       If rAu(jjj).FID = rFa(i).FID Then
        rAu(jjj).FID = neufid
       End If ' rAu(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rBr)
       If rBr(jjj).FID = rFa(i).FID Then
        rBr(jjj).FID = neufid
       End If ' rBr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rDi)
       If rDi(jjj).FID = rFa(i).FID Then
        rDi(jjj).FID = neufid
       End If ' rDi(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rDo)
       If rDo(jjj).FID = rFa(i).FID Then
        rDo(jjj).FID = neufid
       End If ' rDo(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rEi)
       If rEi(jjj).FID = rFa(i).FID Then
        rEi(jjj).FID = neufid
       End If ' rEi(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rFr)
       If rFr(jjj).FID = rFa(i).FID Then
        rFr(jjj).FID = neufid
       End If ' rFr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rKh)
       If rKh(jjj).FID = rFa(i).FID Then
        rKh(jjj).FID = neufid
       End If ' rKh(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLb)
       If rLb(jjj).FID = rFa(i).FID Then
        rLb(jjj).FID = neufid
       End If ' rLb(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLa)
       If rLa(jjj).FID = rFa(i).FID Then
        rLa(jjj).FID = neufid
       End If ' rLa(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rLe)
       If rLe(jjj).FID = rFa(i).FID Then
        rLe(jjj).FID = neufid
       End If ' rLe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rMe)
       If rMe(jjj).FID = rFa(i).FID Then
        rMe(jjj).FID = neufid
       End If ' rMe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rRe)
       If rRe(jjj).FID = rFa(i).FID Then
        rRe(jjj).FID = neufid
       End If ' rRe(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rRr)
       If rRr(jjj).FID = rFa(i).FID Then
        rRr(jjj).FID = neufid
       End If ' rRr(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rUs)
       If rUs(jjj).FID = rFa(i).FID Then
        rUs(jjj).FID = neufid
       End If ' rUs(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rFu)
       If rFu(jjj).FID = rFa(i).FID Then
        rFu(jjj).FID = neufid
       End If ' rFu(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rUl)
       If rUl(jjj).FID = rFa(i).FID Then
        rUl(jjj).FID = neufid
       End If ' rUl(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rVk)
       If rVk(jjj).FID = rFa(i).FID Then
        rVk(jjj).FID = neufid
       End If ' rVk(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rSw)
       If rSw(jjj).FID = rFa(i).FID Then
        rSw(jjj).FID = neufid
       End If ' rSw(jjj).FID = rFa(i).FID THEN
      Next jjj
      For jjj = 1 To UBound(rVo)
       If rVo(jjj).FID = rFa(i).FID Then
        rVo(jjj).FID = neufid
       End If ' rVo(jjj).FID = rFa(i).FID THEN
      Next jjj
     End If ' neufid <> rFa(i).FID Then
     neufid = neufid + 1 ' für den nächsten Patienten
    End If ' rs.BOF Then
    csql.m_Len = 0
   End If ' IF SammelInsert = 0 Then
  End If ' SammelInsert = 0 OR i = ubound(rFa)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rFa(" & i & "/" & UBound(rFa) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""faelleSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(84)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFa), i)
  If Len(rFa(k).Quartal) > maxi(0) Then maxi(0) = Len(rFa(k).Quartal)
  If Len(rFa(k).Nachname) > maxi(1) Then maxi(1) = Len(rFa(k).Nachname)
  If Len(rFa(k).Vorname) > maxi(2) Then maxi(2) = Len(rFa(k).Vorname)
  If Len(rFa(k).PrüfZdFd) > maxi(3) Then maxi(3) = Len(rFa(k).PrüfZdFd)
  If Len(rFa(k).TMFNr) > maxi(4) Then maxi(4) = Len(rFa(k).TMFNr)
  If Len(rFa(k).VKNr) > maxi(5) Then maxi(5) = Len(rFa(k).VKNr)
  If Len(rFa(k).bPerG) > maxi(6) Then maxi(6) = Len(rFa(k).bPerG)
  If Len(rFa(k).DMPKnZ) > maxi(7) Then maxi(7) = Len(rFa(k).DMPKnZ)
  If Len(rFa(k).KKasse_2) > maxi(8) Then maxi(8) = Len(rFa(k).KKasse_2)
  If Len(rFa(k).UnfFlg) > maxi(9) Then maxi(9) = Len(rFa(k).UnfFlg)
  If Len(rFa(k).KtrAbrB) > maxi(10) Then maxi(10) = Len(rFa(k).KtrAbrB)
  If Len(rFa(k).AbrAr) > maxi(11) Then maxi(11) = Len(rFa(k).AbrAr)
  If Len(rFa(k).IK) > maxi(12) Then maxi(12) = Len(rFa(k).IK)
  If Len(rFa(k).KVKs) > maxi(13) Then maxi(13) = Len(rFa(k).KVKs)
  If Len(rFa(k).KVKserg) > maxi(14) Then maxi(14) = Len(rFa(k).KVKserg)
  If Len(rFa(k).Status) > maxi(15) Then maxi(15) = Len(rFa(k).Status)
  If Len(rFa(k).Kasse) > maxi(16) Then maxi(16) = Len(rFa(k).Kasse)
  If Len(rFa(k).GebOr) > maxi(17) Then maxi(17) = Len(rFa(k).GebOr)
  If Len(rFa(k).AbrGb) > maxi(18) Then maxi(18) = Len(rFa(k).AbrGb)
  If Len(rFa(k).PersKreis) > maxi(19) Then maxi(19) = Len(rFa(k).PersKreis)
  If Len(rFa(k).SKtZusatz) > maxi(20) Then maxi(20) = Len(rFa(k).SKtZusatz)
  If Len(rFa(k).SktBem) > maxi(21) Then maxi(21) = Len(rFa(k).SktBem)
  If Len(rFa(k).letzteRegel) > maxi(22) Then maxi(22) = Len(rFa(k).letzteRegel)
  If Len(rFa(k).ÜwText) > maxi(23) Then maxi(23) = Len(rFa(k).ÜwText)
  If Len(rFa(k).statNuller) > maxi(24) Then maxi(24) = Len(rFa(k).statNuller)
  If Len(rFa(k).ÜbwV) > maxi(25) Then maxi(25) = Len(rFa(k).ÜbwV)
  If Len(rFa(k).ÜbWVLANR) > maxi(26) Then maxi(26) = Len(rFa(k).ÜbWVLANR)
  If Len(rFa(k).ÜbWVBSNR) > maxi(27) Then maxi(27) = Len(rFa(k).ÜbWVBSNR)
  If Len(rFa(k).ÜbWVKVNR) > maxi(28) Then maxi(28) = Len(rFa(k).ÜbWVKVNR)
  If Len(rFa(k).AndÜw) > maxi(29) Then maxi(29) = Len(rFa(k).AndÜw)
  If Len(rFa(k).Übwr) > maxi(30) Then maxi(30) = Len(rFa(k).Übwr)
  If Len(rFa(k).ÜbwLANR) > maxi(31) Then maxi(31) = Len(rFa(k).ÜbwLANR)
  If Len(rFa(k).ÜWZiel) > maxi(32) Then maxi(32) = Len(rFa(k).ÜWZiel)
  If Len(rFa(k).ÜWNNr) > maxi(33) Then maxi(33) = Len(rFa(k).ÜWNNr)
  If Len(rFa(k).ÜWNaN) > maxi(34) Then maxi(34) = Len(rFa(k).ÜWNaN)
  If Len(rFa(k).ÜWTit) > maxi(35) Then maxi(35) = Len(rFa(k).ÜWTit)
  If Len(rFa(k).ÜWVor) > maxi(36) Then maxi(36) = Len(rFa(k).ÜWVor)
  If Len(rFa(k).ÜWVsw) > maxi(37) Then maxi(37) = Len(rFa(k).ÜWVsw)
  If Len(rFa(k).Auftrag) > maxi(38) Then maxi(38) = Len(rFa(k).Auftrag)
  If Len(rFa(k).Verdacht) > maxi(39) Then maxi(39) = Len(rFa(k).Verdacht)
  If Len(rFa(k).Befund) > maxi(40) Then maxi(40) = Len(rFa(k).Befund)
  If Len(rFa(k).statKlasse) > maxi(41) Then maxi(41) = Len(rFa(k).statKlasse)
  If Len(rFa(k).KHNam) > maxi(42) Then maxi(42) = Len(rFa(k).KHNam)
  If Len(rFa(k).Weiterbeh) > maxi(43) Then maxi(43) = Len(rFa(k).Weiterbeh)
  If Len(rFa(k).VermiArt) > maxi(44) Then maxi(44) = Len(rFa(k).VermiArt)
  If Len(rFa(k).VermiZusatz) > maxi(45) Then maxi(45) = Len(rFa(k).VermiZusatz)
  If Len(rFa(k).PGeb) > maxi(46) Then maxi(46) = Len(rFa(k).PGeb)
  If Len(rFa(k).PGebErg) > maxi(47) Then maxi(47) = Len(rFa(k).PGebErg)
  If Len(rFa(k).Mahnfrist) > maxi(48) Then maxi(48) = Len(rFa(k).Mahnfrist)
  If Len(rFa(k).Unfallort) > maxi(49) Then maxi(49) = Len(rFa(k).Unfallort)
  If Len(rFa(k).BeschAls) > maxi(50) Then maxi(50) = Len(rFa(k).BeschAls)
  If Len(rFa(k).Unfallbetrieb) > maxi(51) Then maxi(51) = Len(rFa(k).Unfallbetrieb)
  If Len(rFa(k).bHeilb) > maxi(52) Then maxi(52) = Len(rFa(k).bHeilb)
  If Len(rFa(k).GOÄKatNr) > maxi(53) Then maxi(53) = Len(rFa(k).GOÄKatNr)
  If Len(rFa(k).GOÄKatName) > maxi(54) Then maxi(54) = Len(rFa(k).GOÄKatName)
  If Len(rFa(k).abrArzt) > maxi(55) Then maxi(55) = Len(rFa(k).abrArzt)
  If Len(rFa(k).privVers) > maxi(56) Then maxi(56) = Len(rFa(k).privVers)
  If Len(rFa(k).AdNam) > maxi(57) Then maxi(57) = Len(rFa(k).AdNam)
  If Len(rFa(k).AdStr) > maxi(58) Then maxi(58) = Len(rFa(k).AdStr)
  If Len(rFa(k).AdPlz) > maxi(59) Then maxi(59) = Len(rFa(k).AdPlz)
  If Len(rFa(k).AdOrt) > maxi(60) Then maxi(60) = Len(rFa(k).AdOrt)
  If Len(rFa(k).ÜwBG) > maxi(61) Then maxi(61) = Len(rFa(k).ÜwBG)
  If Len(rFa(k).s8000) > maxi(62) Then maxi(62) = Len(rFa(k).s8000)
  If Len(rFa(k).s8100) > maxi(63) Then maxi(63) = Len(rFa(k).s8100)
  If Len(rFa(k).altQuart) > maxi(64) Then maxi(64) = Len(rFa(k).altQuart)
  If Len(rFa(k).QS) > maxi(65) Then maxi(65) = Len(rFa(k).QS)
  If Len(rFa(k).QT) > maxi(66) Then maxi(66) = Len(rFa(k).QT)
  If Len(rFa(k).ZnrMLes) > maxi(67) Then maxi(67) = Len(rFa(k).ZnrMLes)
  If Len(rFa(k).BGFallNr) > maxi(68) Then maxi(68) = Len(rFa(k).BGFallNr)
  If Len(rFa(k).dmpVertret) > maxi(69) Then maxi(69) = Len(rFa(k).dmpVertret)
  If Len(rFa(k).dmpArztw) > maxi(70) Then maxi(70) = Len(rFa(k).dmpArztw)
  If Len(rFa(k).dmpHypos) > maxi(71) Then maxi(71) = Len(rFa(k).dmpHypos)
  If Len(rFa(k).dmpKhsA) > maxi(72) Then maxi(72) = Len(rFa(k).dmpKhsA)
  If Len(rFa(k).dmpDMSchulEmpf) > maxi(73) Then maxi(73) = Len(rFa(k).dmpDMSchulEmpf)
  If Len(rFa(k).dmpDMSchulWahrg) > maxi(74) Then maxi(74) = Len(rFa(k).dmpDMSchulWahrg)
  If Len(rFa(k).dmpHypertSchulEmpf) > maxi(75) Then maxi(75) = Len(rFa(k).dmpHypertSchulEmpf)
  If Len(rFa(k).dmpHypertSchulWahrg) > maxi(76) Then maxi(76) = Len(rFa(k).dmpHypertSchulWahrg)
  If Len(rFa(k).dmpKKTabakEmpf) > maxi(77) Then maxi(77) = Len(rFa(k).dmpKKTabakEmpf)
  If Len(rFa(k).dmpKKErnEmpf) > maxi(78) Then maxi(78) = Len(rFa(k).dmpKKErnEmpf)
  If Len(rFa(k).dmpKKkTrainEmpf) > maxi(79) Then maxi(79) = Len(rFa(k).dmpKKkTrainEmpf)
  If Len(rFa(k).dmpHbA1cZiel) > maxi(80) Then maxi(80) = Len(rFa(k).dmpHbA1cZiel)
  If Len(rFa(k).dmpUewFuss) > maxi(81) Then maxi(81) = Len(rFa(k).dmpUewFuss)
  If Len(rFa(k).dmpEinwDM) > maxi(82) Then maxi(82) = Len(rFa(k).dmpEinwDM)
  If Len(rFa(k).dmphalbj) > maxi(83) Then maxi(83) = Len(rFa(k).dmphalbj)
  If Len(rFa(k).dmpMA) > maxi(84) Then maxi(84) = Len(rFa(k).dmpMA)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFa.Quartal: '" & rFa(k).Quartal & "' -> '" & left$(rFa(k).Quartal, maxL) & "'", True: rFa(k).Quartal = left$(rFa(k).Quartal, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFa.Nachname: '" & rFa(k).Nachname & "' -> '" & left$(rFa(k).Nachname, maxL) & "'", True: rFa(k).Nachname = left$(rFa(k).Nachname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFa.Vorname: '" & rFa(k).Vorname & "' -> '" & left$(rFa(k).Vorname, maxL) & "'", True: rFa(k).Vorname = left$(rFa(k).Vorname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFa.PrüfZdFd: '" & rFa(k).PrüfZdFd & "' -> '" & left$(rFa(k).PrüfZdFd, maxL) & "'", True: rFa(k).PrüfZdFd = left$(rFa(k).PrüfZdFd, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFa.TMFNr: '" & rFa(k).TMFNr & "' -> '" & left$(rFa(k).TMFNr, maxL) & "'", True: rFa(k).TMFNr = left$(rFa(k).TMFNr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFa.VKNr: '" & rFa(k).VKNr & "' -> '" & left$(rFa(k).VKNr, maxL) & "'", True: rFa(k).VKNr = left$(rFa(k).VKNr, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFa.bPerG: '" & rFa(k).bPerG & "' -> '" & left$(rFa(k).bPerG, maxL) & "'", True: rFa(k).bPerG = left$(rFa(k).bPerG, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFa.DMPKnZ: '" & rFa(k).DMPKnZ & "' -> '" & left$(rFa(k).DMPKnZ, maxL) & "'", True: rFa(k).DMPKnZ = left$(rFa(k).DMPKnZ, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFa.KKasse_2: '" & rFa(k).KKasse_2 & "' -> '" & left$(rFa(k).KKasse_2, maxL) & "'", True: rFa(k).KKasse_2 = left$(rFa(k).KKasse_2, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFa.UnfFlg: '" & rFa(k).UnfFlg & "' -> '" & left$(rFa(k).UnfFlg, maxL) & "'", True: rFa(k).UnfFlg = left$(rFa(k).UnfFlg, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFa.KtrAbrB: '" & rFa(k).KtrAbrB & "' -> '" & left$(rFa(k).KtrAbrB, maxL) & "'", True: rFa(k).KtrAbrB = left$(rFa(k).KtrAbrB, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrAr: '" & rFa(k).AbrAr & "' -> '" & left$(rFa(k).AbrAr, maxL) & "'", True: rFa(k).AbrAr = left$(rFa(k).AbrAr, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFa.IK: '" & rFa(k).IK & "' -> '" & left$(rFa(k).IK, maxL) & "'", True: rFa(k).IK = left$(rFa(k).IK, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKs: '" & rFa(k).KVKs & "' -> '" & left$(rFa(k).KVKs, maxL) & "'", True: rFa(k).KVKs = left$(rFa(k).KVKs, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rFa.KVKserg: '" & rFa(k).KVKserg & "' -> '" & left$(rFa(k).KVKserg, maxL) & "'", True: rFa(k).KVKserg = left$(rFa(k).KVKserg, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rFa.Status: '" & rFa(k).Status & "' -> '" & left$(rFa(k).Status, maxL) & "'", True: rFa(k).Status = left$(rFa(k).Status, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rFa.Kasse: '" & rFa(k).Kasse & "' -> '" & left$(rFa(k).Kasse, maxL) & "'", True: rFa(k).Kasse = left$(rFa(k).Kasse, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rFa.GebOr: '" & rFa(k).GebOr & "' -> '" & left$(rFa(k).GebOr, maxL) & "'", True: rFa(k).GebOr = left$(rFa(k).GebOr, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rFa.AbrGb: '" & rFa(k).AbrGb & "' -> '" & left$(rFa(k).AbrGb, maxL) & "'", True: rFa(k).AbrGb = left$(rFa(k).AbrGb, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rFa.PersKreis: '" & rFa(k).PersKreis & "' -> '" & left$(rFa(k).PersKreis, maxL) & "'", True: rFa(k).PersKreis = left$(rFa(k).PersKreis, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rFa.SKtZusatz: '" & rFa(k).SKtZusatz & "' -> '" & left$(rFa(k).SKtZusatz, maxL) & "'", True: rFa(k).SKtZusatz = left$(rFa(k).SKtZusatz, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rFa.SktBem: '" & rFa(k).SktBem & "' -> '" & left$(rFa(k).SktBem, maxL) & "'", True: rFa(k).SktBem = left$(rFa(k).SktBem, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rFa.letzteRegel: '" & rFa(k).letzteRegel & "' -> '" & left$(rFa(k).letzteRegel, maxL) & "'", True: rFa(k).letzteRegel = left$(rFa(k).letzteRegel, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwText: '" & rFa(k).ÜwText & "' -> '" & left$(rFa(k).ÜwText, maxL) & "'", True: rFa(k).ÜwText = left$(rFa(k).ÜwText, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rFa.statNuller: '" & rFa(k).statNuller & "' -> '" & left$(rFa(k).statNuller, maxL) & "'", True: rFa(k).statNuller = left$(rFa(k).statNuller, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwV: '" & rFa(k).ÜbwV & "' -> '" & left$(rFa(k).ÜbwV, maxL) & "'", True: rFa(k).ÜbwV = left$(rFa(k).ÜbwV, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVLANR: '" & rFa(k).ÜbWVLANR & "' -> '" & left$(rFa(k).ÜbWVLANR, maxL) & "'", True: rFa(k).ÜbWVLANR = left$(rFa(k).ÜbWVLANR, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVBSNR: '" & rFa(k).ÜbWVBSNR & "' -> '" & left$(rFa(k).ÜbWVBSNR, maxL) & "'", True: rFa(k).ÜbWVBSNR = left$(rFa(k).ÜbWVBSNR, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbWVKVNR: '" & rFa(k).ÜbWVKVNR & "' -> '" & left$(rFa(k).ÜbWVKVNR, maxL) & "'", True: rFa(k).ÜbWVKVNR = left$(rFa(k).ÜbWVKVNR, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rFa.AndÜw: '" & rFa(k).AndÜw & "' -> '" & left$(rFa(k).AndÜw, maxL) & "'", True: rFa(k).AndÜw = left$(rFa(k).AndÜw, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rFa.Übwr: '" & rFa(k).Übwr & "' -> '" & left$(rFa(k).Übwr, maxL) & "'", True: rFa(k).Übwr = left$(rFa(k).Übwr, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜbwLANR: '" & rFa(k).ÜbwLANR & "' -> '" & left$(rFa(k).ÜbwLANR, maxL) & "'", True: rFa(k).ÜbwLANR = left$(rFa(k).ÜbwLANR, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWZiel: '" & rFa(k).ÜWZiel & "' -> '" & left$(rFa(k).ÜWZiel, maxL) & "'", True: rFa(k).ÜWZiel = left$(rFa(k).ÜWZiel, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNNr: '" & rFa(k).ÜWNNr & "' -> '" & left$(rFa(k).ÜWNNr, maxL) & "'", True: rFa(k).ÜWNNr = left$(rFa(k).ÜWNNr, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWNaN: '" & rFa(k).ÜWNaN & "' -> '" & left$(rFa(k).ÜWNaN, maxL) & "'", True: rFa(k).ÜWNaN = left$(rFa(k).ÜWNaN, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWTit: '" & rFa(k).ÜWTit & "' -> '" & left$(rFa(k).ÜWTit, maxL) & "'", True: rFa(k).ÜWTit = left$(rFa(k).ÜWTit, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVor: '" & rFa(k).ÜWVor & "' -> '" & left$(rFa(k).ÜWVor, maxL) & "'", True: rFa(k).ÜWVor = left$(rFa(k).ÜWVor, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜWVsw: '" & rFa(k).ÜWVsw & "' -> '" & left$(rFa(k).ÜWVsw, maxL) & "'", True: rFa(k).ÜWVsw = left$(rFa(k).ÜWVsw, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rFa.Auftrag: '" & rFa(k).Auftrag & "' -> '" & left$(rFa(k).Auftrag, maxL) & "'", True: rFa(k).Auftrag = left$(rFa(k).Auftrag, maxL)
       Case 39: Lese.Ausgeb "   Verkürze Inhalt von rFa.Verdacht: '" & rFa(k).Verdacht & "' -> '" & left$(rFa(k).Verdacht, maxL) & "'", True: rFa(k).Verdacht = left$(rFa(k).Verdacht, maxL)
       Case 40: Lese.Ausgeb "   Verkürze Inhalt von rFa.Befund: '" & rFa(k).Befund & "' -> '" & left$(rFa(k).Befund, maxL) & "'", True: rFa(k).Befund = left$(rFa(k).Befund, maxL)
       Case 41: Lese.Ausgeb "   Verkürze Inhalt von rFa.statKlasse: '" & rFa(k).statKlasse & "' -> '" & left$(rFa(k).statKlasse, maxL) & "'", True: rFa(k).statKlasse = left$(rFa(k).statKlasse, maxL)
       Case 42: Lese.Ausgeb "   Verkürze Inhalt von rFa.KHNam: '" & rFa(k).KHNam & "' -> '" & left$(rFa(k).KHNam, maxL) & "'", True: rFa(k).KHNam = left$(rFa(k).KHNam, maxL)
       Case 43: Lese.Ausgeb "   Verkürze Inhalt von rFa.Weiterbeh: '" & rFa(k).Weiterbeh & "' -> '" & left$(rFa(k).Weiterbeh, maxL) & "'", True: rFa(k).Weiterbeh = left$(rFa(k).Weiterbeh, maxL)
       Case 44: Lese.Ausgeb "   Verkürze Inhalt von rFa.VermiArt: '" & rFa(k).VermiArt & "' -> '" & left$(rFa(k).VermiArt, maxL) & "'", True: rFa(k).VermiArt = left$(rFa(k).VermiArt, maxL)
       Case 45: Lese.Ausgeb "   Verkürze Inhalt von rFa.VermiZusatz: '" & rFa(k).VermiZusatz & "' -> '" & left$(rFa(k).VermiZusatz, maxL) & "'", True: rFa(k).VermiZusatz = left$(rFa(k).VermiZusatz, maxL)
       Case 46: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGeb: '" & rFa(k).PGeb & "' -> '" & left$(rFa(k).PGeb, maxL) & "'", True: rFa(k).PGeb = left$(rFa(k).PGeb, maxL)
       Case 47: Lese.Ausgeb "   Verkürze Inhalt von rFa.PGebErg: '" & rFa(k).PGebErg & "' -> '" & left$(rFa(k).PGebErg, maxL) & "'", True: rFa(k).PGebErg = left$(rFa(k).PGebErg, maxL)
       Case 48: Lese.Ausgeb "   Verkürze Inhalt von rFa.Mahnfrist: '" & rFa(k).Mahnfrist & "' -> '" & left$(rFa(k).Mahnfrist, maxL) & "'", True: rFa(k).Mahnfrist = left$(rFa(k).Mahnfrist, maxL)
       Case 49: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallort: '" & rFa(k).Unfallort & "' -> '" & left$(rFa(k).Unfallort, maxL) & "'", True: rFa(k).Unfallort = left$(rFa(k).Unfallort, maxL)
       Case 50: Lese.Ausgeb "   Verkürze Inhalt von rFa.BeschAls: '" & rFa(k).BeschAls & "' -> '" & left$(rFa(k).BeschAls, maxL) & "'", True: rFa(k).BeschAls = left$(rFa(k).BeschAls, maxL)
       Case 51: Lese.Ausgeb "   Verkürze Inhalt von rFa.Unfallbetrieb: '" & rFa(k).Unfallbetrieb & "' -> '" & left$(rFa(k).Unfallbetrieb, maxL) & "'", True: rFa(k).Unfallbetrieb = left$(rFa(k).Unfallbetrieb, maxL)
       Case 52: Lese.Ausgeb "   Verkürze Inhalt von rFa.bHeilb: '" & rFa(k).bHeilb & "' -> '" & left$(rFa(k).bHeilb, maxL) & "'", True: rFa(k).bHeilb = left$(rFa(k).bHeilb, maxL)
       Case 53: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatNr: '" & rFa(k).GOÄKatNr & "' -> '" & left$(rFa(k).GOÄKatNr, maxL) & "'", True: rFa(k).GOÄKatNr = left$(rFa(k).GOÄKatNr, maxL)
       Case 54: Lese.Ausgeb "   Verkürze Inhalt von rFa.GOÄKatName: '" & rFa(k).GOÄKatName & "' -> '" & left$(rFa(k).GOÄKatName, maxL) & "'", True: rFa(k).GOÄKatName = left$(rFa(k).GOÄKatName, maxL)
       Case 55: Lese.Ausgeb "   Verkürze Inhalt von rFa.abrArzt: '" & rFa(k).abrArzt & "' -> '" & left$(rFa(k).abrArzt, maxL) & "'", True: rFa(k).abrArzt = left$(rFa(k).abrArzt, maxL)
       Case 56: Lese.Ausgeb "   Verkürze Inhalt von rFa.privVers: '" & rFa(k).privVers & "' -> '" & left$(rFa(k).privVers, maxL) & "'", True: rFa(k).privVers = left$(rFa(k).privVers, maxL)
       Case 57: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdNam: '" & rFa(k).AdNam & "' -> '" & left$(rFa(k).AdNam, maxL) & "'", True: rFa(k).AdNam = left$(rFa(k).AdNam, maxL)
       Case 58: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdStr: '" & rFa(k).AdStr & "' -> '" & left$(rFa(k).AdStr, maxL) & "'", True: rFa(k).AdStr = left$(rFa(k).AdStr, maxL)
       Case 59: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdPlz: '" & rFa(k).AdPlz & "' -> '" & left$(rFa(k).AdPlz, maxL) & "'", True: rFa(k).AdPlz = left$(rFa(k).AdPlz, maxL)
       Case 60: Lese.Ausgeb "   Verkürze Inhalt von rFa.AdOrt: '" & rFa(k).AdOrt & "' -> '" & left$(rFa(k).AdOrt, maxL) & "'", True: rFa(k).AdOrt = left$(rFa(k).AdOrt, maxL)
       Case 61: Lese.Ausgeb "   Verkürze Inhalt von rFa.ÜwBG: '" & rFa(k).ÜwBG & "' -> '" & left$(rFa(k).ÜwBG, maxL) & "'", True: rFa(k).ÜwBG = left$(rFa(k).ÜwBG, maxL)
       Case 62: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8000: '" & rFa(k).s8000 & "' -> '" & left$(rFa(k).s8000, maxL) & "'", True: rFa(k).s8000 = left$(rFa(k).s8000, maxL)
       Case 63: Lese.Ausgeb "   Verkürze Inhalt von rFa.s8100: '" & rFa(k).s8100 & "' -> '" & left$(rFa(k).s8100, maxL) & "'", True: rFa(k).s8100 = left$(rFa(k).s8100, maxL)
       Case 64: Lese.Ausgeb "   Verkürze Inhalt von rFa.altQuart: '" & rFa(k).altQuart & "' -> '" & left$(rFa(k).altQuart, maxL) & "'", True: rFa(k).altQuart = left$(rFa(k).altQuart, maxL)
       Case 65: Lese.Ausgeb "   Verkürze Inhalt von rFa.QS: '" & rFa(k).QS & "' -> '" & left$(rFa(k).QS, maxL) & "'", True: rFa(k).QS = left$(rFa(k).QS, maxL)
       Case 66: Lese.Ausgeb "   Verkürze Inhalt von rFa.QT: '" & rFa(k).QT & "' -> '" & left$(rFa(k).QT, maxL) & "'", True: rFa(k).QT = left$(rFa(k).QT, maxL)
       Case 67: Lese.Ausgeb "   Verkürze Inhalt von rFa.ZnrMLes: '" & rFa(k).ZnrMLes & "' -> '" & left$(rFa(k).ZnrMLes, maxL) & "'", True: rFa(k).ZnrMLes = left$(rFa(k).ZnrMLes, maxL)
       Case 68: Lese.Ausgeb "   Verkürze Inhalt von rFa.BGFallNr: '" & rFa(k).BGFallNr & "' -> '" & left$(rFa(k).BGFallNr, maxL) & "'", True: rFa(k).BGFallNr = left$(rFa(k).BGFallNr, maxL)
       Case 69: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpVertret: '" & rFa(k).dmpVertret & "' -> '" & left$(rFa(k).dmpVertret, maxL) & "'", True: rFa(k).dmpVertret = left$(rFa(k).dmpVertret, maxL)
       Case 70: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpArztw: '" & rFa(k).dmpArztw & "' -> '" & left$(rFa(k).dmpArztw, maxL) & "'", True: rFa(k).dmpArztw = left$(rFa(k).dmpArztw, maxL)
       Case 71: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypos: '" & rFa(k).dmpHypos & "' -> '" & left$(rFa(k).dmpHypos, maxL) & "'", True: rFa(k).dmpHypos = left$(rFa(k).dmpHypos, maxL)
       Case 72: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKhsA: '" & rFa(k).dmpKhsA & "' -> '" & left$(rFa(k).dmpKhsA, maxL) & "'", True: rFa(k).dmpKhsA = left$(rFa(k).dmpKhsA, maxL)
       Case 73: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulEmpf: '" & rFa(k).dmpDMSchulEmpf & "' -> '" & left$(rFa(k).dmpDMSchulEmpf, maxL) & "'", True: rFa(k).dmpDMSchulEmpf = left$(rFa(k).dmpDMSchulEmpf, maxL)
       Case 74: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpDMSchulWahrg: '" & rFa(k).dmpDMSchulWahrg & "' -> '" & left$(rFa(k).dmpDMSchulWahrg, maxL) & "'", True: rFa(k).dmpDMSchulWahrg = left$(rFa(k).dmpDMSchulWahrg, maxL)
       Case 75: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulEmpf: '" & rFa(k).dmpHypertSchulEmpf & "' -> '" & left$(rFa(k).dmpHypertSchulEmpf, maxL) & "'", True: rFa(k).dmpHypertSchulEmpf = left$(rFa(k).dmpHypertSchulEmpf, maxL)
       Case 76: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHypertSchulWahrg: '" & rFa(k).dmpHypertSchulWahrg & "' -> '" & left$(rFa(k).dmpHypertSchulWahrg, maxL) & "'", True: rFa(k).dmpHypertSchulWahrg = left$(rFa(k).dmpHypertSchulWahrg, maxL)
       Case 77: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKTabakEmpf: '" & rFa(k).dmpKKTabakEmpf & "' -> '" & left$(rFa(k).dmpKKTabakEmpf, maxL) & "'", True: rFa(k).dmpKKTabakEmpf = left$(rFa(k).dmpKKTabakEmpf, maxL)
       Case 78: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKErnEmpf: '" & rFa(k).dmpKKErnEmpf & "' -> '" & left$(rFa(k).dmpKKErnEmpf, maxL) & "'", True: rFa(k).dmpKKErnEmpf = left$(rFa(k).dmpKKErnEmpf, maxL)
       Case 79: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpKKkTrainEmpf: '" & rFa(k).dmpKKkTrainEmpf & "' -> '" & left$(rFa(k).dmpKKkTrainEmpf, maxL) & "'", True: rFa(k).dmpKKkTrainEmpf = left$(rFa(k).dmpKKkTrainEmpf, maxL)
       Case 80: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpHbA1cZiel: '" & rFa(k).dmpHbA1cZiel & "' -> '" & left$(rFa(k).dmpHbA1cZiel, maxL) & "'", True: rFa(k).dmpHbA1cZiel = left$(rFa(k).dmpHbA1cZiel, maxL)
       Case 81: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpUewFuss: '" & rFa(k).dmpUewFuss & "' -> '" & left$(rFa(k).dmpUewFuss, maxL) & "'", True: rFa(k).dmpUewFuss = left$(rFa(k).dmpUewFuss, maxL)
       Case 82: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpEinwDM: '" & rFa(k).dmpEinwDM & "' -> '" & left$(rFa(k).dmpEinwDM, maxL) & "'", True: rFa(k).dmpEinwDM = left$(rFa(k).dmpEinwDM, maxL)
       Case 83: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmphalbj: '" & rFa(k).dmphalbj & "' -> '" & left$(rFa(k).dmphalbj, maxL) & "'", True: rFa(k).dmphalbj = left$(rFa(k).dmphalbj, maxL)
       Case 84: Lese.Ausgeb "   Verkürze Inhalt von rFa.dmpMA: '" & rFa(k).dmpMA & "' -> '" & left$(rFa(k).dmpMA, maxL) & "'", True: rFa(k).dmpMA = left$(rFa(k).dmpMA, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
If ErrNumber = -2147467259 Then
 Dim sqlquer$, uKas$, kat$
 uKas = UCase$(Trim$(rFa(i).Kasse))
 If Len(uKas) < 4 Then uKas = UCase$(Trim$(rFa(i).KKasse_2))
 If rFa(i).SchGr = 90 Then kat = "PRI" Else kat = holKat(uKas)
 sqlquer = "INSERT INTO `kassenliste`(name,kurzname,`GO`,`VKNR`,`IK`,`eingef`,pid,kateg) VALUES (" & "'" & rFa(i).Kasse & "', '" & rFa(i).KKasse_2 & "', '" & rFa(i).GOÄKatName & "', '" & rFa(i).VKNr & "', '" & rFa(i).IK & "'," & Format(Now(), "yyyymmddHHMMSS") & "," & rFa(i).Pat_ID & "," & kat & ")"
 InsKorr DBCn, sqlquer, rAf
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
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' faelleSpeichern

Public Function roAuZuw(i&, j&)
 roAu(i).FID = rAu(j).FID
 roAu(i).Pat_ID = rAu(j).Pat_ID
 roAu(i).Zeitpunkt = rAu(j).Zeitpunkt
 roAu(i).Ersteller = rAu(j).Ersteller
 roAu(i).Änderer = rAu(j).Änderer
 roAu(i).art = rAu(j).art
 roAu(i).Beginn = rAu(j).Beginn
 roAu(i).Ende = rAu(j).Ende
 roAu(i).ICDs = rAu(j).ICDs
 roAu(i).absPos = rAu(j).absPos
 roAu(i).aktZeit = rAu(j).aktZeit
 roAu(i).StByte = rAu(j).StByte
End Function ' roAuZuw

Public Function AuZUnt%(i&, j&)
 If roAu(i).FID <> rAu(j).FID Then GoSub unter
 If roAu(i).Pat_ID <> rAu(j).Pat_ID Then GoSub unter
 If roAu(i).Zeitpunkt <> rAu(j).Zeitpunkt Then GoSub unter
 If roAu(i).Ersteller <> rAu(j).Ersteller Then GoSub unter
 If roAu(i).Änderer <> rAu(j).Änderer Then GoSub unter
 If roAu(i).art <> rAu(j).art Then GoSub unter
 If roAu(i).Beginn <> rAu(j).Beginn Then GoSub unter
 If roAu(i).Ende <> rAu(j).Ende Then GoSub unter
 If roAu(i).ICDs <> rAu(j).ICDs Then GoSub unter
 If roAu(i).absPos <> rAu(j).absPos Then GoSub unter
 If roAu(i).aktZeit <> rAu(j).aktZeit Then GoSub unter
 If roAu(i).StByte <> rAu(j).StByte Then GoSub unter
 Exit Function
unter:
 AuZUnt = AuZUnt + 1
 Return
End Function ' AuZUnt

Public Function auLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Ersteller,'') Ersteller" & _
",COALESCE(Änderer,'') Änderer,COALESCE(Art,'') Art,COALESCE(Beginn,'') Beginn,COALESCE(Ende,'') Ende" & _
",COALESCE(ICDs,'') ICDs,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
" FROM `au` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roAu(0)
 Else ' rs.EOF Then
  ReDim roAu(1)
  Do While Not rs.EOF
   akt = UBound(roAu)
   roAu(akt).FID = rs!FID
   roAu(akt).Pat_ID = rs!Pat_ID
   roAu(akt).Zeitpunkt = rs!Zeitpunkt
   roAu(akt).Ersteller = doUmwfSQL(rs!Ersteller, lies.obMySQL, False)
   roAu(akt).Änderer = doUmwfSQL(rs!Änderer, lies.obMySQL, False)
   roAu(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roAu(akt).Beginn = doUmwfSQL(rs!Beginn, lies.obMySQL, False)
   roAu(akt).Ende = doUmwfSQL(rs!Ende, lies.obMySQL, False)
   roAu(akt).ICDs = doUmwfSQL(rs!ICDs, lies.obMySQL, False)
   roAu(akt).absPos = rs!absPos
   roAu(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rAu) > 0 Then
  For ri = 1 To UBound(rAu)
   If rAu(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roAu)
    If roAu(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roAu(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roAu(roendpe + UBound(rAu) - rbeg)
   For ri = rbeg To UBound(rAu)
    Call roAuZuw(roendpe + ri - rbeg, ri)
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

Public Function rAuDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rAudump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rAu)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rAu(" & i & ").FID:" & String$(33, "."), 33) & rAu(i).FID
  Print #200, left$("rAu(" & i & ").Pat_ID:" & String$(33, "."), 33) & rAu(i).Pat_ID
  Print #200, left$("rAu(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rAu(i).Zeitpunkt
  Print #200, left$("rAu(" & i & ").Ersteller:" & String$(33, "."), 33) & "'" & rAu(i).Ersteller & "'"
  Print #200, left$("rAu(" & i & ").Änderer:" & String$(33, "."), 33) & "'" & rAu(i).Änderer & "'"
  Print #200, left$("rAu(" & i & ").Art:" & String$(33, "."), 33) & "'" & rAu(i).art & "'"
  Print #200, left$("rAu(" & i & ").Beginn:" & String$(33, "."), 33) & "'" & rAu(i).Beginn & "'"
  Print #200, left$("rAu(" & i & ").Ende:" & String$(33, "."), 33) & "'" & rAu(i).Ende & "'"
  Print #200, left$("rAu(" & i & ").ICDs:" & String$(33, "."), 33) & "'" & rAu(i).ICDs & "'"
  Print #200, left$("rAu(" & i & ").absPos:" & String$(33, "."), 33) & rAu(i).absPos
  Print #200, left$("rAu(" & i & ").AktZeit:" & String$(33, "."), 33) & rAu(i).aktZeit
  Print #200, left$("rAu(" & i & ").StByte:" & String$(33, "."), 33) & rAu(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' auDump

Public Function auSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "au")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rAu) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Ersteller,Änderer,Art,Beginn,Ende,ICDs,absPos,AktZeit,StByte)    VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rAu)
  rAu(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rAu(i).FID, ",", rAu(i).Pat_ID, ",", DatFor_k(rAu(i).Zeitpunkt), ",'", rAu(i).Ersteller, "','", rAu(i).Änderer, "','", rAu(i).art, "','", rAu(i).Beginn, "','", rAu(i).Ende, "','", _
   rAu(i).ICDs, "',", rAu(i).absPos, ",", DatFor_k(rAu(i).aktZeit), ",", rAu(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rAu) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rAu) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rAu)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rAu(" & i & "/" & UBound(rAu) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""auSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rAu), i)
  If Len(rAu(k).Ersteller) > maxi(0) Then maxi(0) = Len(rAu(k).Ersteller)
  If Len(rAu(k).Änderer) > maxi(1) Then maxi(1) = Len(rAu(k).Änderer)
  If Len(rAu(k).art) > maxi(2) Then maxi(2) = Len(rAu(k).art)
  If Len(rAu(k).Beginn) > maxi(3) Then maxi(3) = Len(rAu(k).Beginn)
  If Len(rAu(k).Ende) > maxi(4) Then maxi(4) = Len(rAu(k).Ende)
  If Len(rAu(k).ICDs) > maxi(5) Then maxi(5) = Len(rAu(k).ICDs)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rAu), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ersteller: '" & rAu(k).Ersteller & "' -> '" & left$(rAu(k).Ersteller, maxL) & "'", True: rAu(k).Ersteller = left$(rAu(k).Ersteller, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rAu.Änderer: '" & rAu(k).Änderer & "' -> '" & left$(rAu(k).Änderer, maxL) & "'", True: rAu(k).Änderer = left$(rAu(k).Änderer, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rAu.Art: '" & rAu(k).art & "' -> '" & left$(rAu(k).art, maxL) & "'", True: rAu(k).art = left$(rAu(k).art, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rAu.Beginn: '" & rAu(k).Beginn & "' -> '" & left$(rAu(k).Beginn, maxL) & "'", True: rAu(k).Beginn = left$(rAu(k).Beginn, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rAu.Ende: '" & rAu(k).Ende & "' -> '" & left$(rAu(k).Ende, maxL) & "'", True: rAu(k).Ende = left$(rAu(k).Ende, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rAu.ICDs: '" & rAu(k).ICDs & "' -> '" & left$(rAu(k).ICDs, maxL) & "'", True: rAu(k).ICDs = left$(rAu(k).ICDs, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in auSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' auSpeichern

Public Function roBrZuw(i&, j&)
 roBr(i).FID = rBr(j).FID
 roBr(i).Pat_ID = rBr(j).Pat_ID
 roBr(i).Zeitpunkt = rBr(j).Zeitpunkt
 roBr(i).Pfad = rBr(j).Pfad
 roBr(i).art = rBr(j).art
 roBr(i).name = rBr(j).name
 roBr(i).autor = rBr(j).autor
 roBr(i).Quelldatum = rBr(j).Quelldatum
 roBr(i).Typ = rBr(j).Typ
 roBr(i).aktZeit = rBr(j).aktZeit
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
 If roBr(i).Pat_ID <> rBr(j).Pat_ID Then GoSub unter
 If roBr(i).Zeitpunkt <> rBr(j).Zeitpunkt Then GoSub unter
 If roBr(i).Pfad <> rBr(j).Pfad Then GoSub unter
 If roBr(i).art <> rBr(j).art Then GoSub unter
 If roBr(i).name <> rBr(j).name Then GoSub unter
 If roBr(i).autor <> rBr(j).autor Then GoSub unter
 If roBr(i).Quelldatum <> rBr(j).Quelldatum Then GoSub unter
 If roBr(i).Typ <> rBr(j).Typ Then GoSub unter
 If roBr(i).aktZeit <> rBr(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
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
   roBr(akt).Pat_ID = rs!Pat_ID
   roBr(akt).Zeitpunkt = rs!Zeitpunkt
   roBr(akt).Pfad = doUmwfSQL(rs!Pfad, lies.obMySQL, False)
   roBr(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roBr(akt).name = doUmwfSQL(rs!name, lies.obMySQL, False)
   roBr(akt).autor = doUmwfSQL(rs!autor, lies.obMySQL, False)
   roBr(akt).Quelldatum = rs!Quelldatum
   roBr(akt).Typ = doUmwfSQL(rs!Typ, lies.obMySQL, False)
   roBr(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rBr) > 0 Then
  For ri = 1 To UBound(rBr)
   If rBr(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roBr)
    If roBr(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roBr(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roBr(roendpe + UBound(rBr) - rbeg)
   For ri = rbeg To UBound(rBr)
    Call roBrZuw(roendpe + ri - rbeg, ri)
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

Public Function rBrDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rBrdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rBr)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rBr(" & i & ").FID:" & String$(33, "."), 33) & rBr(i).FID
  Print #200, left$("rBr(" & i & ").Pat_ID:" & String$(33, "."), 33) & rBr(i).Pat_ID
  Print #200, left$("rBr(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rBr(i).Zeitpunkt
  Print #200, left$("rBr(" & i & ").Pfad:" & String$(33, "."), 33) & "'" & rBr(i).Pfad & "'"
  Print #200, left$("rBr(" & i & ").Art:" & String$(33, "."), 33) & "'" & rBr(i).art & "'"
  Print #200, left$("rBr(" & i & ").Name:" & String$(33, "."), 33) & "'" & rBr(i).name & "'"
  Print #200, left$("rBr(" & i & ").autor:" & String$(33, "."), 33) & "'" & rBr(i).autor & "'"
  Print #200, left$("rBr(" & i & ").Quelldatum:" & String$(33, "."), 33) & rBr(i).Quelldatum
  Print #200, left$("rBr(" & i & ").Typ:" & String$(33, "."), 33) & "'" & rBr(i).Typ & "'"
  Print #200, left$("rBr(" & i & ").AktZeit:" & String$(33, "."), 33) & rBr(i).aktZeit
  Print #200, left$("rBr(" & i & ").DokGroe:" & String$(33, "."), 33) & rBr(i).DokGroe
  Print #200, left$("rBr(" & i & ").DokAenD:" & String$(33, "."), 33) & rBr(i).DokAenD
  Print #200, left$("rBr(" & i & ").QS:" & String$(33, "."), 33) & "'" & rBr(i).QS & "'"
  Print #200, left$("rBr(" & i & ").QT:" & String$(33, "."), 33) & "'" & rBr(i).QT & "'"
  Print #200, left$("rBr(" & i & ").absPos:" & String$(33, "."), 33) & rBr(i).absPos
  Print #200, left$("rBr(" & i & ").StByte:" & String$(33, "."), 33) & rBr(i).StByte
  Print #200, left$("rBr(" & i & ").ID:" & String$(33, "."), 33) & rBr(i).id
 Next i
 Close #200
 zeigan ffadat
End Function ' briefeDump

Public Function briefeSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "briefe")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rBr) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Pfad,Art,Name,autor,Quelldatum,Typ,AktZeit,DokGroe,DokAenD,QS," & _
     "QT,absPos,StByte)      VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rBr)
  rBr(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rBr(i).FID, ",", rBr(i).Pat_ID, ",", DatFor_k(rBr(i).Zeitpunkt), ",'", rBr(i).Pfad, "','", rBr(i).art, "','", rBr(i).name, "','", rBr(i).autor, "',", DatFor_k(rBr(i).Quelldatum), ",'", _
   rBr(i).Typ, "',", DatFor_k(rBr(i).aktZeit), ",", rBr(i).DokGroe, ",", DatFor_k(rBr(i).DokAenD), ",'", rBr(i).QS, "','", rBr(i).QT, "',", rBr(i).absPos, ",", rBr(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rBr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rBr) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rBr(i).id = myEFrag("SELECT MAX(ID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rBr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rBr(" & i & "/" & UBound(rBr) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""briefeSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rBr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rBr.Pfad: '" & rBr(k).Pfad & "' -> '" & left$(rBr(k).Pfad, maxL) & "'", True: rBr(k).Pfad = left$(rBr(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rBr.Art: '" & rBr(k).art & "' -> '" & left$(rBr(k).art, maxL) & "'", True: rBr(k).art = left$(rBr(k).art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rBr.Name: '" & rBr(k).name & "' -> '" & left$(rBr(k).name, maxL) & "'", True: rBr(k).name = left$(rBr(k).name, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rBr.autor: '" & rBr(k).autor & "' -> '" & left$(rBr(k).autor, maxL) & "'", True: rBr(k).autor = left$(rBr(k).autor, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rBr.Typ: '" & rBr(k).Typ & "' -> '" & left$(rBr(k).Typ, maxL) & "'", True: rBr(k).Typ = left$(rBr(k).Typ, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rBr.QS: '" & rBr(k).QS & "' -> '" & left$(rBr(k).QS, maxL) & "'", True: rBr(k).QS = left$(rBr(k).QS, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rBr.QT: '" & rBr(k).QT & "' -> '" & left$(rBr(k).QT, maxL) & "'", True: rBr(k).QT = left$(rBr(k).QT, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in briefeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' briefeSpeichern

Public Function roDiZuw(i&, j&)
 roDi(i).ID1 = rDi(j).ID1
 roDi(i).FID = rDi(j).FID
 roDi(i).Pat_ID = rDi(j).Pat_ID
 roDi(i).DiagDatum = rDi(j).DiagDatum
 roDi(i).DiagSicherheit = rDi(j).DiagSicherheit
 roDi(i).DiagText = rDi(j).DiagText
 roDi(i).DiagSeite = rDi(j).DiagSeite
 roDi(i).DiagAttr = rDi(j).DiagAttr
 roDi(i).ICD = rDi(j).ICD
 roDi(i).obDauer = rDi(j).obDauer
 roDi(i).intBemerk = rDi(j).intBemerk
 roDi(i).absPos = rDi(j).absPos
 roDi(i).aktZeit = rDi(j).aktZeit
 roDi(i).StByte = rDi(j).StByte
 roDi(i).AusnBegr = rDi(j).AusnBegr
 roDi(i).Dggel = rDi(j).Dggel
 roDi(i).obKasse = rDi(j).obKasse
 roDi(i).lKasse = rDi(j).lKasse
 roDi(i).KFdFA = rDi(j).KFdFA
End Function ' roDiZuw

Public Function DiZUnt%(i&, j&)
 If roDi(i).ID1 <> rDi(j).ID1 Then GoSub unter
 If roDi(i).FID <> rDi(j).FID Then GoSub unter
 If roDi(i).Pat_ID <> rDi(j).Pat_ID Then GoSub unter
 If roDi(i).DiagDatum <> rDi(j).DiagDatum Then GoSub unter
 If roDi(i).DiagSicherheit <> rDi(j).DiagSicherheit Then GoSub unter
 If roDi(i).DiagText <> rDi(j).DiagText Then GoSub unter
 If roDi(i).DiagSeite <> rDi(j).DiagSeite Then GoSub unter
 If roDi(i).DiagAttr <> rDi(j).DiagAttr Then GoSub unter
 If roDi(i).ICD <> rDi(j).ICD Then GoSub unter
 If roDi(i).obDauer <> rDi(j).obDauer Then GoSub unter
 If roDi(i).intBemerk <> rDi(j).intBemerk Then GoSub unter
 If roDi(i).absPos <> rDi(j).absPos Then GoSub unter
 If roDi(i).aktZeit <> rDi(j).aktZeit Then GoSub unter
 If roDi(i).StByte <> rDi(j).StByte Then GoSub unter
 If roDi(i).AusnBegr <> rDi(j).AusnBegr Then GoSub unter
 If roDi(i).Dggel <> rDi(j).Dggel Then GoSub unter
 If roDi(i).obKasse <> rDi(j).obKasse Then GoSub unter
 If roDi(i).lKasse <> rDi(j).lKasse Then GoSub unter
 If roDi(i).KFdFA <> rDi(j).KFdFA Then GoSub unter
 Exit Function
unter:
 DiZUnt = DiZUnt + 1
 Return
End Function ' DiZUnt

Public Function diagnosenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(ID1,0) ID1,COALESCE(FID,0) FID,COALESCE(Pat_id,0) Pat_id,COALESCE(DiagDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DiagDatum" & _
",COALESCE(DiagSicherheit,'') DiagSicherheit,COALESCE(DiagText,'') DiagText,COALESCE(DiagSeite,'') DiagSeite,COALESCE(DiagAttr,'') DiagAttr" & _
",COALESCE(ICD,'') ICD,COALESCE(obDauer,0) obDauer,COALESCE(intBemerk,'') intBemerk,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte,COALESCE(AusnBegr,'') AusnBegr,COALESCE(Dggel,0) Dggel" & _
",COALESCE(obKasse,0) obKasse,COALESCE(lKasse - INTERVAL 0 DAY,CONVERT('18991230',DATE)) lKasse,COALESCE(KFdFA,'') KFdFA FROM `diagnosen` WHERE Pat_ID=" & pid & " ORDER BY `DiagDatum`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDi(0)
 Else ' rs.EOF Then
  ReDim roDi(1)
  Do While Not rs.EOF
   akt = UBound(roDi)
   roDi(akt).ID1 = rs!ID1
   roDi(akt).FID = rs!FID
   roDi(akt).Pat_ID = rs!Pat_ID
   roDi(akt).DiagDatum = rs!DiagDatum
   roDi(akt).DiagSicherheit = doUmwfSQL(rs!DiagSicherheit, lies.obMySQL, False)
   roDi(akt).DiagText = doUmwfSQL(rs!DiagText, lies.obMySQL, False)
   roDi(akt).DiagSeite = doUmwfSQL(rs!DiagSeite, lies.obMySQL, False)
   roDi(akt).DiagAttr = doUmwfSQL(rs!DiagAttr, lies.obMySQL, False)
   roDi(akt).ICD = doUmwfSQL(rs!ICD, lies.obMySQL, False)
   roDi(akt).obDauer = rs!obDauer
   roDi(akt).intBemerk = doUmwfSQL(rs!intBemerk, lies.obMySQL, False)
   roDi(akt).absPos = rs!absPos
   roDi(akt).aktZeit = rs!aktZeit
   roDi(akt).StByte = rs!StByte
   roDi(akt).AusnBegr = doUmwfSQL(rs!AusnBegr, lies.obMySQL, False)
   roDi(akt).Dggel = rs!Dggel
   roDi(akt).obKasse = rs!obKasse
   roDi(akt).lKasse = rs!lKasse
   roDi(akt).KFdFA = doUmwfSQL(rs!KFdFA, lies.obMySQL, False)
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDi) > 0 Then
  For ri = 1 To UBound(rDi)
   If rDi(ri).DiagDatum >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roDi)
    If roDi(roendpe).DiagDatum >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roDi(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roDi(roendpe + UBound(rDi) - rbeg)
   For ri = rbeg To UBound(rDi)
    Call roDiZuw(roendpe + ri - rbeg, ri)
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

Public Function rDiDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rDidump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rDi)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rDi(" & i & ").ID1:" & String$(33, "."), 33) & rDi(i).ID1
  Print #200, left$("rDi(" & i & ").FID:" & String$(33, "."), 33) & rDi(i).FID
  Print #200, left$("rDi(" & i & ").Pat_id:" & String$(33, "."), 33) & rDi(i).Pat_ID
  Print #200, left$("rDi(" & i & ").DiagDatum:" & String$(33, "."), 33) & rDi(i).DiagDatum
  Print #200, left$("rDi(" & i & ").DiagSicherheit:" & String$(33, "."), 33) & "'" & rDi(i).DiagSicherheit & "'"
  Print #200, left$("rDi(" & i & ").DiagText:" & String$(33, "."), 33) & "'" & rDi(i).DiagText & "'"
  Print #200, left$("rDi(" & i & ").DiagSeite:" & String$(33, "."), 33) & "'" & rDi(i).DiagSeite & "'"
  Print #200, left$("rDi(" & i & ").DiagAttr:" & String$(33, "."), 33) & "'" & rDi(i).DiagAttr & "'"
  Print #200, left$("rDi(" & i & ").ICD:" & String$(33, "."), 33) & "'" & rDi(i).ICD & "'"
  Print #200, left$("rDi(" & i & ").obDauer:" & String$(33, "."), 33) & rDi(i).obDauer
  Print #200, left$("rDi(" & i & ").intBemerk:" & String$(33, "."), 33) & "'" & rDi(i).intBemerk & "'"
  Print #200, left$("rDi(" & i & ").absPos:" & String$(33, "."), 33) & rDi(i).absPos
  Print #200, left$("rDi(" & i & ").AktZeit:" & String$(33, "."), 33) & rDi(i).aktZeit
  Print #200, left$("rDi(" & i & ").StByte:" & String$(33, "."), 33) & rDi(i).StByte
  Print #200, left$("rDi(" & i & ").AusnBegr:" & String$(33, "."), 33) & "'" & rDi(i).AusnBegr & "'"
  Print #200, left$("rDi(" & i & ").Dggel:" & String$(33, "."), 33) & rDi(i).Dggel
  Print #200, left$("rDi(" & i & ").obKasse:" & String$(33, "."), 33) & rDi(i).obKasse
  Print #200, left$("rDi(" & i & ").lKasse:" & String$(33, "."), 33) & rDi(i).lKasse
  Print #200, left$("rDi(" & i & ").KFdFA:" & String$(33, "."), 33) & "'" & rDi(i).KFdFA & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' diagnosenDump

Public Function diagnosenSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "diagnosen")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDi) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_id,DiagDatum," & _
     "DiagSicherheit,DiagText,DiagSeite,DiagAttr,ICD,obDauer,intBemerk,absPos,AktZeit,StByte," & _
     "AusnBegr,Dggel,obKasse,lKasse,KFdFA)               VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rDi)
  rDi(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rDi(i).FID, ",", rDi(i).Pat_ID, ",", DatFor_k(rDi(i).DiagDatum), ",'", rDi(i).DiagSicherheit, "','", rDi(i).DiagText, "','", rDi(i).DiagSeite, "','", rDi(i).DiagAttr, "','", _
   rDi(i).ICD, "',", rDi(i).obDauer, ",'", rDi(i).intBemerk, "',", rDi(i).absPos, ",", DatFor_k(rDi(i).aktZeit), ",", rDi(i).StByte, ",'", rDi(i).AusnBegr, "',", rDi(i).Dggel, ",", rDi(i).obKasse, ",", DatFor_k( _
   rDi(i).lKasse), ",'", rDi(i).KFdFA, "')")
  If SammelInsert <> 0 And i < UBound(rDi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDi) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rDi(i).ID1 = myEFrag("SELECT MAX(ID1)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rDi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rDi(" & i & "/" & UBound(rDi) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""diagnosenSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
  If Len(rDi(k).KFdFA) > maxi(7) Then maxi(7) = Len(rDi(k).KFdFA)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSicherheit: '" & rDi(k).DiagSicherheit & "' -> '" & left$(rDi(k).DiagSicherheit, maxL) & "'", True: rDi(k).DiagSicherheit = left$(rDi(k).DiagSicherheit, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagText: '" & rDi(k).DiagText & "' -> '" & left$(rDi(k).DiagText, maxL) & "'", True: rDi(k).DiagText = left$(rDi(k).DiagText, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagSeite: '" & rDi(k).DiagSeite & "' -> '" & left$(rDi(k).DiagSeite, maxL) & "'", True: rDi(k).DiagSeite = left$(rDi(k).DiagSeite, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDi.DiagAttr: '" & rDi(k).DiagAttr & "' -> '" & left$(rDi(k).DiagAttr, maxL) & "'", True: rDi(k).DiagAttr = left$(rDi(k).DiagAttr, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDi.ICD: '" & rDi(k).ICD & "' -> '" & left$(rDi(k).ICD, maxL) & "'", True: rDi(k).ICD = left$(rDi(k).ICD, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDi.intBemerk: '" & rDi(k).intBemerk & "' -> '" & left$(rDi(k).intBemerk, maxL) & "'", True: rDi(k).intBemerk = left$(rDi(k).intBemerk, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rDi.AusnBegr: '" & rDi(k).AusnBegr & "' -> '" & left$(rDi(k).AusnBegr, maxL) & "'", True: rDi(k).AusnBegr = left$(rDi(k).AusnBegr, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rDi.KFdFA: '" & rDi(k).KFdFA & "' -> '" & left$(rDi(k).KFdFA, maxL) & "'", True: rDi(k).KFdFA = left$(rDi(k).KFdFA, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in diagnosenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' diagnosenSpeichern

Public Function roDoZuw(i&, j&)
 roDo(i).FID = rDo(j).FID
 roDo(i).Pat_ID = rDo(j).Pat_ID
 roDo(i).Zeitpunkt = rDo(j).Zeitpunkt
 roDo(i).DokPfad = rDo(j).DokPfad
 roDo(i).DokArt = rDo(j).DokArt
 roDo(i).DokName = rDo(j).DokName
 roDo(i).Quelldatum = rDo(j).Quelldatum
 roDo(i).absPos = rDo(j).absPos
 roDo(i).aktZeit = rDo(j).aktZeit
 roDo(i).DokGroe = rDo(j).DokGroe
 roDo(i).DokAenD = rDo(j).DokAenD
 roDo(i).QS = rDo(j).QS
 roDo(i).QT = rDo(j).QT
 roDo(i).StByte = rDo(j).StByte
End Function ' roDoZuw

Public Function DoZUnt%(i&, j&)
 If roDo(i).FID <> rDo(j).FID Then GoSub unter
 If roDo(i).Pat_ID <> rDo(j).Pat_ID Then GoSub unter
 If roDo(i).Zeitpunkt <> rDo(j).Zeitpunkt Then GoSub unter
 If roDo(i).DokPfad <> rDo(j).DokPfad Then GoSub unter
 If roDo(i).DokArt <> rDo(j).DokArt Then GoSub unter
 If roDo(i).DokName <> rDo(j).DokName Then GoSub unter
 If roDo(i).Quelldatum <> rDo(j).Quelldatum Then GoSub unter
 If roDo(i).absPos <> rDo(j).absPos Then GoSub unter
 If roDo(i).aktZeit <> rDo(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
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
   roDo(akt).Pat_ID = rs!Pat_ID
   roDo(akt).Zeitpunkt = rs!Zeitpunkt
   roDo(akt).DokPfad = doUmwfSQL(rs!DokPfad, lies.obMySQL, False)
   roDo(akt).DokArt = doUmwfSQL(rs!DokArt, lies.obMySQL, False)
   roDo(akt).DokName = doUmwfSQL(rs!DokName, lies.obMySQL, False)
   roDo(akt).Quelldatum = rs!Quelldatum
   roDo(akt).absPos = rs!absPos
   roDo(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDo) > 0 Then
  For ri = 1 To UBound(rDo)
   If rDo(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roDo)
    If roDo(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roDo(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roDo(roendpe + UBound(rDo) - rbeg)
   For ri = rbeg To UBound(rDo)
    Call roDoZuw(roendpe + ri - rbeg, ri)
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

Public Function rDoDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rDodump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rDo)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rDo(" & i & ").FID:" & String$(33, "."), 33) & rDo(i).FID
  Print #200, left$("rDo(" & i & ").Pat_ID:" & String$(33, "."), 33) & rDo(i).Pat_ID
  Print #200, left$("rDo(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rDo(i).Zeitpunkt
  Print #200, left$("rDo(" & i & ").DokPfad:" & String$(33, "."), 33) & "'" & rDo(i).DokPfad & "'"
  Print #200, left$("rDo(" & i & ").DokArt:" & String$(33, "."), 33) & "'" & rDo(i).DokArt & "'"
  Print #200, left$("rDo(" & i & ").DokName:" & String$(33, "."), 33) & "'" & rDo(i).DokName & "'"
  Print #200, left$("rDo(" & i & ").Quelldatum:" & String$(33, "."), 33) & rDo(i).Quelldatum
  Print #200, left$("rDo(" & i & ").absPos:" & String$(33, "."), 33) & rDo(i).absPos
  Print #200, left$("rDo(" & i & ").AktZeit:" & String$(33, "."), 33) & rDo(i).aktZeit
  Print #200, left$("rDo(" & i & ").DokGroe:" & String$(33, "."), 33) & rDo(i).DokGroe
  Print #200, left$("rDo(" & i & ").DokAenD:" & String$(33, "."), 33) & rDo(i).DokAenD
  Print #200, left$("rDo(" & i & ").QS:" & String$(33, "."), 33) & "'" & rDo(i).QS & "'"
  Print #200, left$("rDo(" & i & ").QT:" & String$(33, "."), 33) & "'" & rDo(i).QT & "'"
  Print #200, left$("rDo(" & i & ").StByte:" & String$(33, "."), 33) & rDo(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' dokumenteDump

Public Function dokumenteSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "dokumente")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDo) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "DokPfad,DokArt,DokName,Quelldatum,absPos,AktZeit,DokGroe,DokAenD,QS,QT," & _
     "StByte)  VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rDo)
  rDo(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rDo(i).FID, ",", rDo(i).Pat_ID, ",", DatFor_k(rDo(i).Zeitpunkt), ",'", rDo(i).DokPfad, "','", rDo(i).DokArt, "','", rDo(i).DokName, "',", DatFor_k(rDo(i).Quelldatum), ",", _
   rDo(i).absPos, ",", DatFor_k(rDo(i).aktZeit), ",", rDo(i).DokGroe, ",", DatFor_k(rDo(i).DokAenD), ",'", rDo(i).QS, "','", rDo(i).QT, "',", rDo(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rDo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDo) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rDo)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rDo(" & i & "/" & UBound(rDo) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""dokumenteSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokPfad: '" & rDo(k).DokPfad & "' -> '" & left$(rDo(k).DokPfad, maxL) & "'", True: rDo(k).DokPfad = left$(rDo(k).DokPfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokArt: '" & rDo(k).DokArt & "' -> '" & left$(rDo(k).DokArt, maxL) & "'", True: rDo(k).DokArt = left$(rDo(k).DokArt, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDo.DokName: '" & rDo(k).DokName & "' -> '" & left$(rDo(k).DokName, maxL) & "'", True: rDo(k).DokName = left$(rDo(k).DokName, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDo.QS: '" & rDo(k).QS & "' -> '" & left$(rDo(k).QS, maxL) & "'", True: rDo(k).QS = left$(rDo(k).QS, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDo.QT: '" & rDo(k).QT & "' -> '" & left$(rDo(k).QT, maxL) & "'", True: rDo(k).QT = left$(rDo(k).QT, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dokumenteSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' dokumenteSpeichern

Public Function roEiZuw(i&, j&)
 roEi(i).FID = rEi(j).FID
 roEi(i).Pat_ID = rEi(j).Pat_ID
 roEi(i).Zeitpunkt = rEi(j).Zeitpunkt
 roEi(i).art = rEi(j).art
 roEi(i).Ersteller = rEi(j).Ersteller
 roEi(i).Änderer = rEi(j).Änderer
 roEi(i).Inhalt = rEi(j).Inhalt
 roEi(i).absPos = rEi(j).absPos
 roEi(i).aktZeit = rEi(j).aktZeit
 roEi(i).QS = rEi(j).QS
 roEi(i).QT = rEi(j).QT
 roEi(i).StByte = rEi(j).StByte
 roEi(i).id = rEi(j).id
 roEi(i).inhNum = rEi(j).inhNum
End Function ' roEiZuw

Public Function EiZUnt%(i&, j&)
 If roEi(i).FID <> rEi(j).FID Then GoSub unter
 If roEi(i).Pat_ID <> rEi(j).Pat_ID Then GoSub unter
 If roEi(i).Zeitpunkt <> rEi(j).Zeitpunkt Then GoSub unter
 If roEi(i).art <> rEi(j).art Then GoSub unter
 If roEi(i).Ersteller <> rEi(j).Ersteller Then GoSub unter
 If roEi(i).Änderer <> rEi(j).Änderer Then GoSub unter
 If roEi(i).Inhalt <> rEi(j).Inhalt Then GoSub unter
 If roEi(i).absPos <> rEi(j).absPos Then GoSub unter
 If roEi(i).aktZeit <> rEi(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Art,'') Art" & _
",COALESCE(Ersteller,'') Ersteller,COALESCE(Änderer,'') Änderer,COALESCE(Inhalt,'') Inhalt,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte" & _
",COALESCE(id,0) id,COALESCE(inhNum,0) inhNum FROM `eintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roEi(0)
 Else ' rs.EOF Then
  ReDim roEi(1)
  Do While Not rs.EOF
   akt = UBound(roEi)
   roEi(akt).FID = rs!FID
   roEi(akt).Pat_ID = rs!Pat_ID
   roEi(akt).Zeitpunkt = rs!Zeitpunkt
   roEi(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roEi(akt).Ersteller = doUmwfSQL(rs!Ersteller, lies.obMySQL, False)
   roEi(akt).Änderer = doUmwfSQL(rs!Änderer, lies.obMySQL, False)
   roEi(akt).Inhalt = doUmwfSQL(rs!Inhalt, lies.obMySQL, False)
   roEi(akt).absPos = rs!absPos
   roEi(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rEi) > 0 Then
  For ri = 1 To UBound(rEi)
   If rEi(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roEi)
    If roEi(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roEi(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roEi(roendpe + UBound(rEi) - rbeg)
   For ri = rbeg To UBound(rEi)
    Call roEiZuw(roendpe + ri - rbeg, ri)
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

Public Function rEiDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rEidump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rEi)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rEi(" & i & ").FID:" & String$(33, "."), 33) & rEi(i).FID
  Print #200, left$("rEi(" & i & ").Pat_ID:" & String$(33, "."), 33) & rEi(i).Pat_ID
  Print #200, left$("rEi(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rEi(i).Zeitpunkt
  Print #200, left$("rEi(" & i & ").Art:" & String$(33, "."), 33) & "'" & rEi(i).art & "'"
  Print #200, left$("rEi(" & i & ").Ersteller:" & String$(33, "."), 33) & "'" & rEi(i).Ersteller & "'"
  Print #200, left$("rEi(" & i & ").Änderer:" & String$(33, "."), 33) & "'" & rEi(i).Änderer & "'"
  Print #200, left$("rEi(" & i & ").Inhalt:" & String$(33, "."), 33) & "'" & rEi(i).Inhalt & "'"
  Print #200, left$("rEi(" & i & ").absPos:" & String$(33, "."), 33) & rEi(i).absPos
  Print #200, left$("rEi(" & i & ").AktZeit:" & String$(33, "."), 33) & rEi(i).aktZeit
  Print #200, left$("rEi(" & i & ").QS:" & String$(33, "."), 33) & "'" & rEi(i).QS & "'"
  Print #200, left$("rEi(" & i & ").QT:" & String$(33, "."), 33) & "'" & rEi(i).QT & "'"
  Print #200, left$("rEi(" & i & ").StByte:" & String$(33, "."), 33) & rEi(i).StByte
  Print #200, left$("rEi(" & i & ").id:" & String$(33, "."), 33) & rEi(i).id
  Print #200, left$("rEi(" & i & ").inhNum:" & String$(33, "."), 33) & rEi(i).inhNum
 Next i
 Close #200
 zeigan ffadat
End Function ' eintraegeDump

Public Function eintraegeSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "eintraege")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rEi) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Ersteller,Änderer,Inhalt,absPos,AktZeit,QS,QT,StByte,inhNum)               VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rEi)
  rEi(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rEi(i).FID, ",", rEi(i).Pat_ID, ",", DatFor_k(rEi(i).Zeitpunkt), ",'", rEi(i).art, "','", rEi(i).Ersteller, "','", rEi(i).Änderer, "','", rEi(i).Inhalt, "',", rEi(i).absPos, ",", DatFor_k( _
   rEi(i).aktZeit), ",'", rEi(i).QS, "','", rEi(i).QT, "',", rEi(i).StByte, ",", REPLACE$(rEi(i).inhNum, ",", "."), ")")
  If SammelInsert <> 0 And i < UBound(rEi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rEi) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rEi(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rEi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rEi(" & i & "/" & UBound(rEi) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""eintraegeSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rEi), i)
  If Len(rEi(k).art) > maxi(0) Then maxi(0) = Len(rEi(k).art)
  If Len(rEi(k).Ersteller) > maxi(1) Then maxi(1) = Len(rEi(k).Ersteller)
  If Len(rEi(k).Änderer) > maxi(2) Then maxi(2) = Len(rEi(k).Änderer)
  If Len(rEi(k).Inhalt) > maxi(3) Then maxi(3) = Len(rEi(k).Inhalt)
  If Len(rEi(k).QS) > maxi(4) Then maxi(4) = Len(rEi(k).QS)
  If Len(rEi(k).QT) > maxi(5) Then maxi(5) = Len(rEi(k).QT)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rEi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rEi.Art: '" & rEi(k).art & "' -> '" & left$(rEi(k).art, maxL) & "'", True: rEi(k).art = left$(rEi(k).art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rEi.Ersteller: '" & rEi(k).Ersteller & "' -> '" & left$(rEi(k).Ersteller, maxL) & "'", True: rEi(k).Ersteller = left$(rEi(k).Ersteller, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rEi.Änderer: '" & rEi(k).Änderer & "' -> '" & left$(rEi(k).Änderer, maxL) & "'", True: rEi(k).Änderer = left$(rEi(k).Änderer, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rEi.Inhalt: '" & rEi(k).Inhalt & "' -> '" & left$(rEi(k).Inhalt, maxL) & "'", True: rEi(k).Inhalt = left$(rEi(k).Inhalt, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rEi.QS: '" & rEi(k).QS & "' -> '" & left$(rEi(k).QS, maxL) & "'", True: rEi(k).QS = left$(rEi(k).QS, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rEi.QT: '" & rEi(k).QT & "' -> '" & left$(rEi(k).QT, maxL) & "'", True: rEi(k).QT = left$(rEi(k).QT, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in eintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' eintraegeSpeichern

Public Function rFoDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rFodump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rFo)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rFo(" & i & ").FormID:" & String$(33, "."), 33) & rFo(i).FormID
  Print #200, left$("rFo(" & i & ").Form_Abk:" & String$(33, "."), 33) & "'" & rFo(i).Form_Abk & "'"
  Print #200, left$("rFo(" & i & ").FormBez:" & String$(33, "."), 33) & "'" & rFo(i).FormBez & "'"
  Print #200, left$("rFo(" & i & ").FormVorl:" & String$(33, "."), 33) & "'" & rFo(i).FormVorl & "'"
  Print #200, left$("rFo(" & i & ").AktZeit:" & String$(33, "."), 33) & rFo(i).aktZeit
  Print #200, left$("rFo(" & i & ").absPos:" & String$(33, "."), 33) & rFo(i).absPos
  Print #200, left$("rFo(" & i & ").StByte:" & String$(33, "."), 33) & rFo(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' formulareDump

Public Function formulareSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "formulare")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 Dim neuFormID&, j&
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFo) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Form_Abk,FormBez,FormVorl," & _
     "AktZeit,absPos,StByte)               VALUES"))
 If Not AllePat Then
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = rFo1 + 1 To UBound(rFo)
 If rFo(i).FormID < 0 Then
  rFo(i).StByte = CStr(AktByte)
setz:
   csql.Append csql0
  csql.AppVar Array("('", rFo(i).Form_Abk, "','", rFo(i).FormBez, "','", rFo(i).FormVorl, "',", DatFor_k(rFo(i).aktZeit), ",", rFo(i).absPos, ",", rFo(i).StByte, ")")
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    neuFormID = myEFrag("SELECT formid FROM formulare WHERE Form_Abk='" & rFo(i).Form_Abk & "' AND FormBez='" & rFo(i).FormBez & "' AND LCASE(FormVorl)='" & LCase$(rFo(i).FormVorl) & "'").Fields(0)
    For j = 1 To UBound(rFr)
     If rFr(j).Form_ID = rFo(i).FormID Then rFr(j).Form_ID = neuFormID
    Next j
    If ErrN Then
     If InStrB(ErrD, "Duplicate") <> 0 Then
      rFo(i).FormID = myEFrag("SELECT MAX(FormID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  DoEvents
 End If ' If rFo(i).FormID < 0 Then
 Next i
 rFo1 = UBound(rFo)
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rFo(" & i & "/" & UBound(rFo) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""formulareSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, rFo1 + 1, i) To IIf(SammelInsert <> 0, UBound(rFo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFo.Form_Abk: '" & rFo(k).Form_Abk & "' -> '" & left$(rFo(k).Form_Abk, maxL) & "'", True: rFo(k).Form_Abk = left$(rFo(k).Form_Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormBez: '" & rFo(k).FormBez & "' -> '" & left$(rFo(k).FormBez, maxL) & "'", True: rFo(k).FormBez = left$(rFo(k).FormBez, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFo.FormVorl: '" & rFo(k).FormVorl & "' -> '" & left$(rFo(k).FormVorl, maxL) & "'", True: rFo(k).FormVorl = left$(rFo(k).FormVorl, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in formulareSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' formulareSpeichern

Public Function roFrZuw(i&, j&)
 roFr(i).Foid = rFr(j).Foid
 roFr(i).FID = rFr(j).FID
 roFr(i).Pat_ID = rFr(j).Pat_ID
 roFr(i).Form_ID = rFr(j).Form_ID
 roFr(i).Zeitpunkt = rFr(j).Zeitpunkt
 roFr(i).absPos = rFr(j).absPos
 roFr(i).aktZeit = rFr(j).aktZeit
 roFr(i).StByte = rFr(j).StByte
 roFr(i).Satzart = rFr(j).Satzart
 roFr(i).Satzlänge = rFr(j).Satzlänge
 roFr(i).lanrid = rFr(j).lanrid
End Function ' roFrZuw

Public Function FrZUnt%(i&, j&)
 If roFr(i).Foid <> rFr(j).Foid Then GoSub unter
 If roFr(i).FID <> rFr(j).FID Then GoSub unter
 If roFr(i).Pat_ID <> rFr(j).Pat_ID Then GoSub unter
 If roFr(i).Form_ID <> rFr(j).Form_ID Then GoSub unter
 If roFr(i).Zeitpunkt <> rFr(j).Zeitpunkt Then GoSub unter
 If roFr(i).absPos <> rFr(j).absPos Then GoSub unter
 If roFr(i).aktZeit <> rFr(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
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
   roFr(akt).Pat_ID = rs!Pat_ID
   roFr(akt).Form_ID = rs!Form_ID
   roFr(akt).Zeitpunkt = rs!Zeitpunkt
   roFr(akt).absPos = rs!absPos
   roFr(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rFr) > 0 Then
  For ri = 1 To UBound(rFr)
   If rFr(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roFr)
    If roFr(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roFr(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roFr(roendpe + UBound(rFr) - rbeg)
   For ri = rbeg To UBound(rFr)
    Call roFrZuw(roendpe + ri - rbeg, ri)
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

Public Function rFrDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rFrdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rFr)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rFr(" & i & ").FoID:" & String$(33, "."), 33) & rFr(i).Foid
  Print #200, left$("rFr(" & i & ").FID:" & String$(33, "."), 33) & rFr(i).FID
  Print #200, left$("rFr(" & i & ").Pat_ID:" & String$(33, "."), 33) & rFr(i).Pat_ID
  Print #200, left$("rFr(" & i & ").Form_ID:" & String$(33, "."), 33) & rFr(i).Form_ID
  Print #200, left$("rFr(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rFr(i).Zeitpunkt
  Print #200, left$("rFr(" & i & ").AbsPos:" & String$(33, "."), 33) & rFr(i).absPos
  Print #200, left$("rFr(" & i & ").AktZeit:" & String$(33, "."), 33) & rFr(i).aktZeit
  Print #200, left$("rFr(" & i & ").StByte:" & String$(33, "."), 33) & rFr(i).StByte
  Print #200, left$("rFr(" & i & ").Satzart:" & String$(33, "."), 33) & "'" & rFr(i).Satzart & "'"
  Print #200, left$("rFr(" & i & ").Satzlänge:" & String$(33, "."), 33) & "'" & rFr(i).Satzlänge & "'"
  Print #200, left$("rFr(" & i & ").LANRid:" & String$(33, "."), 33) & rFr(i).lanrid
 Next i
 Close #200
 zeigan ffadat
End Function ' forminhkopfDump

Public Function forminhkopfSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "forminhkopf")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFr) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FoID,FID,Pat_ID," & _
     "Form_ID,ZeitPunkt,AbsPos,AktZeit,StByte,Satzart,Satzlänge,LANRid)              VALUES"))
 FoIDv = 0
erneut:
 If Not AllePat Then
   myEFrag "DELETE fif FROM forminhfeld fif LEFT JOIN forminhkopf fk USING (foid) WHERE pat_ID=" & CStr(rNa(0).Pat_ID), rAFi
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rFr)
  rFr(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rFr(i).Foid, ",", rFr(i).FID, ",", rFr(i).Pat_ID, ",", rFr(i).Form_ID, ",", DatFor_k(rFr(i).Zeitpunkt), ",", rFr(i).absPos, ",", DatFor_k(rFr(i).aktZeit), ",", rFr(i).StByte, ",'", _
   rFr(i).Satzart, "','", rFr(i).Satzlänge, "',", rFr(i).lanrid, ")")
  If SammelInsert <> 0 And i < UBound(rFr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFr) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rFr(i).Foid = myEFrag("SELECT MAX(FoID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rFr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rFr(" & i & "/" & UBound(rFr) & "):   " & ErrDescr
If ErrNumber = -2147217900 And ErrDescr Like "*Duplicate entry * for key 'PRIMARY'" Then
 Dim schlüssel$, pos&, iiru&, jjru&
 pos = InStr(ErrDescr, "'")
 schlüssel = Mid$(ErrDescr, pos + 1, InStr(pos + 2, ErrDescr, "'") - pos - 1)
' Debug.Print schlüssel
 For iiru = 1 To UBound(rFr)
  If rFr(iiru).Foid = schlüssel Then
   If FoIDv = 0 Then FoIDv = myEFrag("SELECT (MAX(foid)+1) FROM forminhkopf", , DBCn).Fields(0) Else FoIDv = FoIDv + 1
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
 Lese.Ausgeb sfnr & " Fehler in ""forminhkopfSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzart: '" & rFr(k).Satzart & "' -> '" & left$(rFr(k).Satzart, maxL) & "'", True: rFr(k).Satzart = left$(rFr(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFr.Satzlänge: '" & rFr(k).Satzlänge & "' -> '" & left$(rFr(k).Satzlänge, maxL) & "'", True: rFr(k).Satzlänge = left$(rFr(k).Satzlänge, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhkopfSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhkopfSpeichern

Public Function rFmDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rFmdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rFm)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rFm(" & i & ").FoID:" & String$(33, "."), 33) & rFm(i).Foid
  Print #200, left$("rFm(" & i & ").Nr:" & String$(33, "."), 33) & rFm(i).nr
  Print #200, left$("rFm(" & i & ").FeldNr:" & String$(33, "."), 33) & rFm(i).FeldNr
  Print #200, left$("rFm(" & i & ").FeldVW:" & String$(33, "."), 33) & rFm(i).FeldVW
  Print #200, left$("rFm(" & i & ").FeldInhVW:" & String$(33, "."), 33) & rFm(i).FeldInhVW
 Next i
 Close #200
 zeigan ffadat
End Function ' forminhfeldDump

Public Function forminhfeldSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "forminhfeld")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFm) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FoID,Nr,FeldNr," & _
     "FeldVW,FeldInhVW)      VALUES"))
 If Not AllePat Then
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rFm)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rFm(i).Foid, ",", rFm(i).nr, ",", rFm(i).FeldNr, ",", rFm(i).FeldVW, ",", rFm(i).FeldInhVW, ")")
  If SammelInsert <> 0 And i < UBound(rFm) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFm) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rFm)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rFm(" & i & "/" & UBound(rFm) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""forminhfeldSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(-1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFm), i)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in forminhfeldSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' forminhfeldSpeichern

Public Function roKhZuw(i&, j&)
 roKh(i).FID = rKh(j).FID
 roKh(i).Pat_ID = rKh(j).Pat_ID
 roKh(i).Zeitpunkt = rKh(j).Zeitpunkt
 roKh(i).Ziel = rKh(j).Ziel
 roKh(i).obNot = rKh(j).obNot
 roKh(i).obBeleg = rKh(j).obBeleg
 roKh(i).Diagnose = rKh(j).Diagnose
 roKh(i).Befund = rKh(j).Befund
 roKh(i).BisMas = rKh(j).BisMas
 roKh(i).FraStel = rKh(j).FraStel
 roKh(i).MitBef = rKh(j).MitBef
 roKh(i).absPos = rKh(j).absPos
 roKh(i).aktZeit = rKh(j).aktZeit
 roKh(i).StByte = rKh(j).StByte
End Function ' roKhZuw

Public Function KhZUnt%(i&, j&)
 If roKh(i).FID <> rKh(j).FID Then GoSub unter
 If roKh(i).Pat_ID <> rKh(j).Pat_ID Then GoSub unter
 If roKh(i).Zeitpunkt <> rKh(j).Zeitpunkt Then GoSub unter
 If roKh(i).Ziel <> rKh(j).Ziel Then GoSub unter
 If roKh(i).obNot <> rKh(j).obNot Then GoSub unter
 If roKh(i).obBeleg <> rKh(j).obBeleg Then GoSub unter
 If roKh(i).Diagnose <> rKh(j).Diagnose Then GoSub unter
 If roKh(i).Befund <> rKh(j).Befund Then GoSub unter
 If roKh(i).BisMas <> rKh(j).BisMas Then GoSub unter
 If roKh(i).FraStel <> rKh(j).FraStel Then GoSub unter
 If roKh(i).MitBef <> rKh(j).MitBef Then GoSub unter
 If roKh(i).absPos <> rKh(j).absPos Then GoSub unter
 If roKh(i).aktZeit <> rKh(j).aktZeit Then GoSub unter
 If roKh(i).StByte <> rKh(j).StByte Then GoSub unter
 Exit Function
unter:
 KhZUnt = KhZUnt + 1
 Return
End Function ' KhZUnt

Public Function kheinweisLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Ziel,'') Ziel" & _
",COALESCE(obNot,0) obNot,COALESCE(obBeleg,0) obBeleg,COALESCE(Diagnose,'') Diagnose,COALESCE(Befund,'') Befund" & _
",COALESCE(BisMas,'') BisMas,COALESCE(FraStel,'') FraStel,COALESCE(MitBef,'') MitBef,COALESCE(absPos,0) absPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `kheinweis` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roKh(0)
 Else ' rs.EOF Then
  ReDim roKh(1)
  Do While Not rs.EOF
   akt = UBound(roKh)
   roKh(akt).FID = rs!FID
   roKh(akt).Pat_ID = rs!Pat_ID
   roKh(akt).Zeitpunkt = rs!Zeitpunkt
   roKh(akt).Ziel = doUmwfSQL(rs!Ziel, lies.obMySQL, False)
   roKh(akt).obNot = rs!obNot
   roKh(akt).obBeleg = rs!obBeleg
   roKh(akt).Diagnose = doUmwfSQL(rs!Diagnose, lies.obMySQL, False)
   roKh(akt).Befund = doUmwfSQL(rs!Befund, lies.obMySQL, False)
   roKh(akt).BisMas = doUmwfSQL(rs!BisMas, lies.obMySQL, False)
   roKh(akt).FraStel = doUmwfSQL(rs!FraStel, lies.obMySQL, False)
   roKh(akt).MitBef = doUmwfSQL(rs!MitBef, lies.obMySQL, False)
   roKh(akt).absPos = rs!absPos
   roKh(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rKh) > 0 Then
  For ri = 1 To UBound(rKh)
   If rKh(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roKh)
    If roKh(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roKh(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roKh(roendpe + UBound(rKh) - rbeg)
   For ri = rbeg To UBound(rKh)
    Call roKhZuw(roendpe + ri - rbeg, ri)
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

Public Function rKhDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rKhdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rKh)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rKh(" & i & ").FID:" & String$(33, "."), 33) & rKh(i).FID
  Print #200, left$("rKh(" & i & ").Pat_ID:" & String$(33, "."), 33) & rKh(i).Pat_ID
  Print #200, left$("rKh(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rKh(i).Zeitpunkt
  Print #200, left$("rKh(" & i & ").Ziel:" & String$(33, "."), 33) & "'" & rKh(i).Ziel & "'"
  Print #200, left$("rKh(" & i & ").obNot:" & String$(33, "."), 33) & rKh(i).obNot
  Print #200, left$("rKh(" & i & ").obBeleg:" & String$(33, "."), 33) & rKh(i).obBeleg
  Print #200, left$("rKh(" & i & ").Diagnose:" & String$(33, "."), 33) & "'" & rKh(i).Diagnose & "'"
  Print #200, left$("rKh(" & i & ").Befund:" & String$(33, "."), 33) & "'" & rKh(i).Befund & "'"
  Print #200, left$("rKh(" & i & ").BisMas:" & String$(33, "."), 33) & "'" & rKh(i).BisMas & "'"
  Print #200, left$("rKh(" & i & ").FraStel:" & String$(33, "."), 33) & "'" & rKh(i).FraStel & "'"
  Print #200, left$("rKh(" & i & ").MitBef:" & String$(33, "."), 33) & "'" & rKh(i).MitBef & "'"
  Print #200, left$("rKh(" & i & ").absPos:" & String$(33, "."), 33) & rKh(i).absPos
  Print #200, left$("rKh(" & i & ").AktZeit:" & String$(33, "."), 33) & rKh(i).aktZeit
  Print #200, left$("rKh(" & i & ").StByte:" & String$(33, "."), 33) & rKh(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' kheinweisDump

Public Function kheinweisSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "kheinweis")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rKh) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Ziel,obNot,obBeleg,Diagnose,Befund,BisMas,FraStel,MitBef,absPos,AktZeit," & _
     "StByte)  VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rKh)
  rKh(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rKh(i).FID, ",", rKh(i).Pat_ID, ",", DatFor_k(rKh(i).Zeitpunkt), ",'", rKh(i).Ziel, "',", CStr(-(rKh(i).obNot <> 0)), ",", CStr(-(rKh(i).obBeleg <> 0)), ",'", rKh(i).Diagnose, "','", _
   rKh(i).Befund, "','", rKh(i).BisMas, "','", rKh(i).FraStel, "','", rKh(i).MitBef, "',", rKh(i).absPos, ",", DatFor_k(rKh(i).aktZeit), ",", rKh(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rKh) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rKh) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rKh)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rKh(" & i & "/" & UBound(rKh) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""kheinweisSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(5)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKh), i)
  If Len(rKh(k).Ziel) > maxi(0) Then maxi(0) = Len(rKh(k).Ziel)
  If Len(rKh(k).Diagnose) > maxi(1) Then maxi(1) = Len(rKh(k).Diagnose)
  If Len(rKh(k).Befund) > maxi(2) Then maxi(2) = Len(rKh(k).Befund)
  If Len(rKh(k).BisMas) > maxi(3) Then maxi(3) = Len(rKh(k).BisMas)
  If Len(rKh(k).FraStel) > maxi(4) Then maxi(4) = Len(rKh(k).FraStel)
  If Len(rKh(k).MitBef) > maxi(5) Then maxi(5) = Len(rKh(k).MitBef)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKh), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKh.Ziel: '" & rKh(k).Ziel & "' -> '" & left$(rKh(k).Ziel, maxL) & "'", True: rKh(k).Ziel = left$(rKh(k).Ziel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rKh.Diagnose: '" & rKh(k).Diagnose & "' -> '" & left$(rKh(k).Diagnose, maxL) & "'", True: rKh(k).Diagnose = left$(rKh(k).Diagnose, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rKh.Befund: '" & rKh(k).Befund & "' -> '" & left$(rKh(k).Befund, maxL) & "'", True: rKh(k).Befund = left$(rKh(k).Befund, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rKh.BisMas: '" & rKh(k).BisMas & "' -> '" & left$(rKh(k).BisMas, maxL) & "'", True: rKh(k).BisMas = left$(rKh(k).BisMas, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rKh.FraStel: '" & rKh(k).FraStel & "' -> '" & left$(rKh(k).FraStel, maxL) & "'", True: rKh(k).FraStel = left$(rKh(k).FraStel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rKh.MitBef: '" & rKh(k).MitBef & "' -> '" & left$(rKh(k).MitBef, maxL) & "'", True: rKh(k).MitBef = left$(rKh(k).MitBef, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kheinweisSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kheinweisSpeichern

Public Function roLbZuw(i&, j&)
 roLb(i).FID = rLb(j).FID
 roLb(i).Pat_ID = rLb(j).Pat_ID
 roLb(i).Zeitpunkt = rLb(j).Zeitpunkt
 roLb(i).AnfText = rLb(j).AnfText
 roLb(i).absPos = rLb(j).absPos
 roLb(i).aktZeit = rLb(j).aktZeit
 roLb(i).StByte = rLb(j).StByte
End Function ' roLbZuw

Public Function LbZUnt%(i&, j&)
 If roLb(i).FID <> rLb(j).FID Then GoSub unter
 If roLb(i).Pat_ID <> rLb(j).Pat_ID Then GoSub unter
 If roLb(i).Zeitpunkt <> rLb(j).Zeitpunkt Then GoSub unter
 If roLb(i).AnfText <> rLb(j).AnfText Then GoSub unter
 If roLb(i).absPos <> rLb(j).absPos Then GoSub unter
 If roLb(i).aktZeit <> rLb(j).aktZeit Then GoSub unter
 If roLb(i).StByte <> rLb(j).StByte Then GoSub unter
 Exit Function
unter:
 LbZUnt = LbZUnt + 1
 Return
End Function ' LbZUnt

Public Function lbanforderungenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
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
   roLb(akt).Pat_ID = rs!Pat_ID
   roLb(akt).Zeitpunkt = rs!Zeitpunkt
   roLb(akt).AnfText = doUmwfSQL(rs!AnfText, lies.obMySQL, False)
   roLb(akt).absPos = rs!absPos
   roLb(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLb) > 0 Then
  For ri = 1 To UBound(rLb)
   If rLb(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roLb)
    If roLb(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roLb(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roLb(roendpe + UBound(rLb) - rbeg)
   For ri = rbeg To UBound(rLb)
    Call roLbZuw(roendpe + ri - rbeg, ri)
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

Public Function rLbDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLbdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLb)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLb(" & i & ").FID:" & String$(33, "."), 33) & rLb(i).FID
  Print #200, left$("rLb(" & i & ").Pat_ID:" & String$(33, "."), 33) & rLb(i).Pat_ID
  Print #200, left$("rLb(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rLb(i).Zeitpunkt
  Print #200, left$("rLb(" & i & ").AnfText:" & String$(33, "."), 33) & "'" & rLb(i).AnfText & "'"
  Print #200, left$("rLb(" & i & ").absPos:" & String$(33, "."), 33) & rLb(i).absPos
  Print #200, left$("rLb(" & i & ").AktZeit:" & String$(33, "."), 33) & rLb(i).aktZeit
  Print #200, left$("rLb(" & i & ").StByte:" & String$(33, "."), 33) & rLb(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' lbanforderungenDump

Public Function lbanforderungenSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "lbanforderungen")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLb) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "AnfText,absPos,AktZeit,StByte)       VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLb)
  rLb(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLb(i).FID, ",", rLb(i).Pat_ID, ",", DatFor_k(rLb(i).Zeitpunkt), ",'", rLb(i).AnfText, "',", rLb(i).absPos, ",", DatFor_k(rLb(i).aktZeit), ",", rLb(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rLb) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLb) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLb)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLb(" & i & "/" & UBound(rLb) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""lbanforderungenSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLb), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLb.AnfText: '" & rLb(k).AnfText & "' -> '" & left$(rLb(k).AnfText, maxL) & "'", True: rLb(k).AnfText = left$(rLb(k).AnfText, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in lbanforderungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' lbanforderungenSpeichern

Public Function roLaZuw(i&, j&)
 roLa(i).FID = rLa(j).FID
 roLa(i).Pat_ID = rLa(j).Pat_ID
 roLa(i).Zeitpunkt = rLa(j).Zeitpunkt
 roLa(i).FertigStGrad = rLa(j).FertigStGrad
 roLa(i).Abkü = rLa(j).Abkü
 roLa(i).LangtextVW = rLa(j).LangtextVW
 roLa(i).Wert = rLa(j).Wert
 roLa(i).Einheit = rLa(j).Einheit
 roLa(i).obpath = rLa(j).obpath
 roLa(i).AnmkgVW = rLa(j).AnmkgVW
 roLa(i).KommentarVW = rLa(j).KommentarVW
 roLa(i).absPos = rLa(j).absPos
 roLa(i).aktZeit = rLa(j).aktZeit
 roLa(i).RefNr = rLa(j).RefNr
 roLa(i).StByte = rLa(j).StByte
 roLa(i).id = rLa(j).id
 roLa(i).AbschlZlVW = rLa(j).AbschlZlVW
 roLa(i).NormberVW = rLa(j).NormberVW
End Function ' roLaZuw

Public Function LaZUnt%(i&, j&)
 If roLa(i).FID <> rLa(j).FID Then GoSub unter
 If roLa(i).Pat_ID <> rLa(j).Pat_ID Then GoSub unter
 If roLa(i).Zeitpunkt <> rLa(j).Zeitpunkt Then GoSub unter
 If roLa(i).FertigStGrad <> rLa(j).FertigStGrad Then GoSub unter
 If roLa(i).Abkü <> rLa(j).Abkü Then GoSub unter
 If roLa(i).LangtextVW <> rLa(j).LangtextVW Then GoSub unter
 If roLa(i).Wert <> rLa(j).Wert Then GoSub unter
 If roLa(i).Einheit <> rLa(j).Einheit Then GoSub unter
 If roLa(i).obpath <> rLa(j).obpath Then GoSub unter
 If roLa(i).AnmkgVW <> rLa(j).AnmkgVW Then GoSub unter
 If roLa(i).KommentarVW <> rLa(j).KommentarVW Then GoSub unter
 If roLa(i).absPos <> rLa(j).absPos Then GoSub unter
 If roLa(i).aktZeit <> rLa(j).aktZeit Then GoSub unter
 If roLa(i).RefNr <> rLa(j).RefNr Then GoSub unter
 If roLa(i).StByte <> rLa(j).StByte Then GoSub unter
 If roLa(i).id <> rLa(j).id Then GoSub unter
 If roLa(i).AbschlZlVW <> rLa(j).AbschlZlVW Then GoSub unter
 If roLa(i).NormberVW <> rLa(j).NormberVW Then GoSub unter
 Exit Function
unter:
 LaZUnt = LaZUnt + 1
 Return
End Function ' LaZUnt

Public Function laborneuLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(FertigStGrad,'') FertigStGrad" & _
",COALESCE(Abkü,'') Abkü,COALESCE(LangtextVW,0) LangtextVW,COALESCE(Wert,'') Wert,COALESCE(Einheit,'') Einheit" & _
",COALESCE(obpath,'') obpath,COALESCE(AnmkgVW,0) AnmkgVW,COALESCE(KommentarVW,0) KommentarVW,COALESCE(AbsPos,0) AbsPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(Refnr,0) Refnr,COALESCE(StByte,0) StByte,COALESCE(ID,0) ID" & _
",COALESCE(AbschlZlVW,0) AbschlZlVW,COALESCE(NormberVW,0) NormberVW FROM `laborneu` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roLa(0)
 Else ' rs.EOF Then
  ReDim roLa(1)
  Do While Not rs.EOF
   akt = UBound(roLa)
   roLa(akt).FID = rs!FID
   roLa(akt).Pat_ID = rs!Pat_ID
   roLa(akt).Zeitpunkt = rs!Zeitpunkt
   roLa(akt).FertigStGrad = doUmwfSQL(rs!FertigStGrad, lies.obMySQL, False)
   roLa(akt).Abkü = doUmwfSQL(rs!Abkü, lies.obMySQL, False)
   roLa(akt).LangtextVW = rs!LangtextVW
   roLa(akt).Wert = doUmwfSQL(rs!Wert, lies.obMySQL, False)
   roLa(akt).Einheit = doUmwfSQL(rs!Einheit, lies.obMySQL, False)
   roLa(akt).obpath = doUmwfSQL(rs!obpath, lies.obMySQL, False)
   roLa(akt).AnmkgVW = rs!AnmkgVW
   roLa(akt).KommentarVW = rs!KommentarVW
   roLa(akt).absPos = rs!absPos
   roLa(akt).aktZeit = rs!aktZeit
   roLa(akt).RefNr = rs!RefNr
   roLa(akt).StByte = rs!StByte
   roLa(akt).id = rs!id
   roLa(akt).AbschlZlVW = rs!AbschlZlVW
   roLa(akt).NormberVW = rs!NormberVW
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

' in alleSpeichern
Function laborneuEinf()
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLa) > 0 Then
  For ri = 1 To UBound(rLa)
   If rLa(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roLa)
    If roLa(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roLa(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roLa(roendpe + UBound(rLa) - rbeg)
   For ri = rbeg To UBound(rLa)
    Call roLaZuw(roendpe + ri - rbeg, ri)
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

Public Function rLaDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLadump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLa)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLa(" & i & ").FID:" & String$(33, "."), 33) & rLa(i).FID
  Print #200, left$("rLa(" & i & ").Pat_ID:" & String$(33, "."), 33) & rLa(i).Pat_ID
  Print #200, left$("rLa(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rLa(i).Zeitpunkt
  Print #200, left$("rLa(" & i & ").FertigStGrad:" & String$(33, "."), 33) & "'" & rLa(i).FertigStGrad & "'"
  Print #200, left$("rLa(" & i & ").Abkü:" & String$(33, "."), 33) & "'" & rLa(i).Abkü & "'"
  Print #200, left$("rLa(" & i & ").LangtextVW:" & String$(33, "."), 33) & rLa(i).LangtextVW
  Print #200, left$("rLa(" & i & ").Wert:" & String$(33, "."), 33) & "'" & rLa(i).Wert & "'"
  Print #200, left$("rLa(" & i & ").Einheit:" & String$(33, "."), 33) & "'" & rLa(i).Einheit & "'"
  Print #200, left$("rLa(" & i & ").obpath:" & String$(33, "."), 33) & "'" & rLa(i).obpath & "'"
  Print #200, left$("rLa(" & i & ").AnmkgVW:" & String$(33, "."), 33) & rLa(i).AnmkgVW
  Print #200, left$("rLa(" & i & ").KommentarVW:" & String$(33, "."), 33) & rLa(i).KommentarVW
  Print #200, left$("rLa(" & i & ").AbsPos:" & String$(33, "."), 33) & rLa(i).absPos
  Print #200, left$("rLa(" & i & ").AktZeit:" & String$(33, "."), 33) & rLa(i).aktZeit
  Print #200, left$("rLa(" & i & ").Refnr:" & String$(33, "."), 33) & rLa(i).RefNr
  Print #200, left$("rLa(" & i & ").StByte:" & String$(33, "."), 33) & rLa(i).StByte
  Print #200, left$("rLa(" & i & ").ID:" & String$(33, "."), 33) & rLa(i).id
  Print #200, left$("rLa(" & i & ").AbschlZlVW:" & String$(33, "."), 33) & rLa(i).AbschlZlVW
  Print #200, left$("rLa(" & i & ").NormberVW:" & String$(33, "."), 33) & rLa(i).NormberVW
 Next i
 Close #200
 zeigan ffadat
End Function ' laborneuDump

Public Function laborneuSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborneu")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLa) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "FertigStGrad,Abkü,LangtextVW,Wert,Einheit,obpath,AnmkgVW,KommentarVW,AbsPos,AktZeit," & _
     "Refnr,StByte,AbschlZlVW,NormberVW)   VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLa)
  rLa(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLa(i).FID, ",", rLa(i).Pat_ID, ",", DatFor_k(rLa(i).Zeitpunkt), ",'", rLa(i).FertigStGrad, "','", rLa(i).Abkü, "',", rLa(i).LangtextVW, ",'", rLa(i).Wert, "','", rLa(i).Einheit, "','", _
   rLa(i).obpath, "',", rLa(i).AnmkgVW, ",", rLa(i).KommentarVW, ",", rLa(i).absPos, ",", DatFor_k(rLa(i).aktZeit), ",", rLa(i).RefNr, ",", rLa(i).StByte, ",", rLa(i).AbschlZlVW, ",", _
   rLa(i).NormberVW, ")")
  If SammelInsert <> 0 And i < UBound(rLa) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLa) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLa(i).id = myEFrag("SELECT MAX(ID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLa)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLa(" & i & "/" & UBound(rLa) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborneuSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLa), i)
  If Len(rLa(k).FertigStGrad) > maxi(0) Then maxi(0) = Len(rLa(k).FertigStGrad)
  If Len(rLa(k).Abkü) > maxi(1) Then maxi(1) = Len(rLa(k).Abkü)
  If Len(rLa(k).Wert) > maxi(2) Then maxi(2) = Len(rLa(k).Wert)
  If Len(rLa(k).Einheit) > maxi(3) Then maxi(3) = Len(rLa(k).Einheit)
  If Len(rLa(k).obpath) > maxi(4) Then maxi(4) = Len(rLa(k).obpath)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLa), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLa.FertigStGrad: '" & rLa(k).FertigStGrad & "' -> '" & left$(rLa(k).FertigStGrad, maxL) & "'", True: rLa(k).FertigStGrad = left$(rLa(k).FertigStGrad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLa.Abkü: '" & rLa(k).Abkü & "' -> '" & left$(rLa(k).Abkü, maxL) & "'", True: rLa(k).Abkü = left$(rLa(k).Abkü, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLa.Wert: '" & rLa(k).Wert & "' -> '" & left$(rLa(k).Wert, maxL) & "'", True: rLa(k).Wert = left$(rLa(k).Wert, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLa.Einheit: '" & rLa(k).Einheit & "' -> '" & left$(rLa(k).Einheit, maxL) & "'", True: rLa(k).Einheit = left$(rLa(k).Einheit, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLa.obpath: '" & rLa(k).obpath & "' -> '" & left$(rLa(k).obpath, maxL) & "'", True: rLa(k).obpath = left$(rLa(k).obpath, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborneuSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborneuSpeichern

Public Function roLeZuw(i&, j&)
 roLe(i).id = rLe(j).id
 roLe(i).FID = rLe(j).FID
 roLe(i).Pat_ID = rLe(j).Pat_ID
 roLe(i).Zeitpunkt = rLe(j).Zeitpunkt
 roLe(i).Leistung = rLe(j).Leistung
 roLe(i).ArtdUs = rLe(j).ArtdUs
 roLe(i).LAnzl = rLe(j).LAnzl
 roLe(i).LUhrz = rLe(j).LUhrz
 roLe(i).LfBegr = rLe(j).LfBegr
 roLe(i).Med = rLe(j).Med
 roLe(i).LOrgan = rLe(j).LOrgan
 roLe(i).LArztBf = rLe(j).LArztBf
 roLe(i).DtlKbsV = rLe(j).DtlKbsV
 roLe(i).LEntlDt = rLe(j).LEntlDt
 roLe(i).Faktor = rLe(j).Faktor
 roLe(i).LBSNR = rLe(j).LBSNR
 roLe(i).Charge = rLe(j).Charge
 roLe(i).Lanr = rLe(j).Lanr
 roLe(i).letzVorg = rLe(j).letzVorg
 roLe(i).Ausn = rLe(j).Ausn
 roLe(i).beme = rLe(j).beme
 roLe(i).absPos = rLe(j).absPos
 roLe(i).aktZeit = rLe(j).aktZeit
 roLe(i).QS = rLe(j).QS
 roLe(i).QT = rLe(j).QT
 roLe(i).StByte = rLe(j).StByte
 roLe(i).lanrid = rLe(j).lanrid
 roLe(i).Sachkbez = rLe(j).Sachkbez
 roLe(i).Sachkct = rLe(j).Sachkct
 roLe(i).Zone = rLe(j).Zone
 roLe(i).Punkte = rLe(j).Punkte
 roLe(i).Lstgerbnr = rLe(j).Lstgerbnr
 roLe(i).Position = rLe(j).Position
 roLe(i).Eignung = rLe(j).Eignung
 roLe(i).Pruefzeit = rLe(j).Pruefzeit
 roLe(i).Kalkzeit = rLe(j).Kalkzeit
 roLe(i).BSNR = rLe(j).BSNR
 roLe(i).Abrechnungssperre = rLe(j).Abrechnungssperre
 roLe(i).Medikanr = rLe(j).Medikanr
 roLe(i).Laborkosten = rLe(j).Laborkosten
 roLe(i).Sachkostentyp = rLe(j).Sachkostentyp
 roLe(i).Usegebordtext = rLe(j).Usegebordtext
 roLe(i).Arztliste = rLe(j).Arztliste
 roLe(i).Ersteller = rLe(j).Ersteller
 roLe(i).Änderer = rLe(j).Änderer
 roLe(i).Kennzeichen_arzt_technik = rLe(j).Kennzeichen_arzt_technik
 roLe(i).Auftragsschluessel = rLe(j).Auftragsschluessel
 roLe(i).Auftragstext = rLe(j).Auftragstext
End Function ' roLeZuw

Public Function LeZUnt%(i&, j&)
 If roLe(i).id <> rLe(j).id Then GoSub unter
 If roLe(i).FID <> rLe(j).FID Then GoSub unter
 If roLe(i).Pat_ID <> rLe(j).Pat_ID Then GoSub unter
 If roLe(i).Zeitpunkt <> rLe(j).Zeitpunkt Then GoSub unter
 If roLe(i).Leistung <> rLe(j).Leistung Then GoSub unter
 If roLe(i).ArtdUs <> rLe(j).ArtdUs Then GoSub unter
 If roLe(i).LAnzl <> rLe(j).LAnzl Then GoSub unter
 If roLe(i).LUhrz <> rLe(j).LUhrz Then GoSub unter
 If roLe(i).LfBegr <> rLe(j).LfBegr Then GoSub unter
 If roLe(i).Med <> rLe(j).Med Then GoSub unter
 If roLe(i).LOrgan <> rLe(j).LOrgan Then GoSub unter
 If roLe(i).LArztBf <> rLe(j).LArztBf Then GoSub unter
 If roLe(i).DtlKbsV <> rLe(j).DtlKbsV Then GoSub unter
 If roLe(i).LEntlDt <> rLe(j).LEntlDt Then GoSub unter
 If roLe(i).Faktor <> rLe(j).Faktor Then GoSub unter
 If roLe(i).LBSNR <> rLe(j).LBSNR Then GoSub unter
 If roLe(i).Charge <> rLe(j).Charge Then GoSub unter
 If roLe(i).Lanr <> rLe(j).Lanr Then GoSub unter
 If roLe(i).letzVorg <> rLe(j).letzVorg Then GoSub unter
 If roLe(i).Ausn <> rLe(j).Ausn Then GoSub unter
 If roLe(i).beme <> rLe(j).beme Then GoSub unter
 If roLe(i).absPos <> rLe(j).absPos Then GoSub unter
 If roLe(i).aktZeit <> rLe(j).aktZeit Then GoSub unter
 If roLe(i).QS <> rLe(j).QS Then GoSub unter
 If roLe(i).QT <> rLe(j).QT Then GoSub unter
 If roLe(i).StByte <> rLe(j).StByte Then GoSub unter
 If roLe(i).lanrid <> rLe(j).lanrid Then GoSub unter
 If roLe(i).Sachkbez <> rLe(j).Sachkbez Then GoSub unter
 If roLe(i).Sachkct <> rLe(j).Sachkct Then GoSub unter
 If roLe(i).Zone <> rLe(j).Zone Then GoSub unter
 If roLe(i).Punkte <> rLe(j).Punkte Then GoSub unter
 If roLe(i).Lstgerbnr <> rLe(j).Lstgerbnr Then GoSub unter
 If roLe(i).Position <> rLe(j).Position Then GoSub unter
 If roLe(i).Eignung <> rLe(j).Eignung Then GoSub unter
 If roLe(i).Pruefzeit <> rLe(j).Pruefzeit Then GoSub unter
 If roLe(i).Kalkzeit <> rLe(j).Kalkzeit Then GoSub unter
 If roLe(i).BSNR <> rLe(j).BSNR Then GoSub unter
 If roLe(i).Abrechnungssperre <> rLe(j).Abrechnungssperre Then GoSub unter
 If roLe(i).Medikanr <> rLe(j).Medikanr Then GoSub unter
 If roLe(i).Laborkosten <> rLe(j).Laborkosten Then GoSub unter
 If roLe(i).Sachkostentyp <> rLe(j).Sachkostentyp Then GoSub unter
 If roLe(i).Usegebordtext <> rLe(j).Usegebordtext Then GoSub unter
 If roLe(i).Arztliste <> rLe(j).Arztliste Then GoSub unter
 If roLe(i).Ersteller <> rLe(j).Ersteller Then GoSub unter
 If roLe(i).Änderer <> rLe(j).Änderer Then GoSub unter
 If roLe(i).Kennzeichen_arzt_technik <> rLe(j).Kennzeichen_arzt_technik Then GoSub unter
 If roLe(i).Auftragsschluessel <> rLe(j).Auftragsschluessel Then GoSub unter
 If roLe(i).Auftragstext <> rLe(j).Auftragstext Then GoSub unter
 Exit Function
unter:
 LeZUnt = LeZUnt + 1
 Return
End Function ' LeZUnt

Public Function leistungenLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(id,0) id,COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt" & _
",COALESCE(Leistung,'') Leistung,COALESCE(ArtdUs,'') ArtdUs,COALESCE(LAnzl,'') LAnzl,COALESCE(LUhrz,'') LUhrz" & _
",COALESCE(LfBegr,'') LfBegr,COALESCE(Med,'') Med,COALESCE(LOrgan,'') LOrgan,COALESCE(LArztBf,'') LArztBf" & _
",COALESCE(DtlKbsV,'') DtlKbsV,COALESCE(LEntlDt,'') LEntlDt,COALESCE(Faktor,'') Faktor,COALESCE(LBSNR,'') LBSNR" & _
",COALESCE(Charge,'') Charge,COALESCE(LANR,'') LANR,COALESCE(letzVorg - INTERVAL 0 DAY,CONVERT('18991230',DATE)) letzVorg,COALESCE(Ausn,'') Ausn" & _
",COALESCE(Beme,'') Beme,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(QS,'') QS" & _
",COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid,COALESCE(Sachkbez,'') Sachkbez" & _
",COALESCE(Sachkct,0) Sachkct,COALESCE(Zone,'') Zone,COALESCE(Punkte,0) Punkte,COALESCE(Lstgerbnr,0) Lstgerbnr" & _
",COALESCE(Position,0) Position,COALESCE(Eignung,0) Eignung,COALESCE(Pruefzeit,0) Pruefzeit,COALESCE(Kalkzeit,0) Kalkzeit" & _
",COALESCE(Bsnr,0) Bsnr,COALESCE(Abrechnungssperre,0) Abrechnungssperre,COALESCE(Medikanr,0) Medikanr,COALESCE(Laborkosten,0) Laborkosten" & _
",COALESCE(Sachkostentyp,0) Sachkostentyp,COALESCE(Usegebordtext,'') Usegebordtext,COALESCE(Arztliste,'') Arztliste,COALESCE(Ersteller,'') Ersteller" & _
",COALESCE(Änderer,'') Änderer,COALESCE(Kennzeichen_arzt_technik,'') Kennzeichen_arzt_technik,COALESCE(Auftragsschluessel,'') Auftragsschluessel,COALESCE(Auftragstext,'') Auftragstext" & _
" FROM `leistungen` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roLe(0)
 Else ' rs.EOF Then
  ReDim roLe(1)
  Do While Not rs.EOF
   akt = UBound(roLe)
   roLe(akt).id = rs!id
   roLe(akt).FID = rs!FID
   roLe(akt).Pat_ID = rs!Pat_ID
   roLe(akt).Zeitpunkt = rs!Zeitpunkt
   roLe(akt).Leistung = doUmwfSQL(rs!Leistung, lies.obMySQL, False)
   roLe(akt).ArtdUs = doUmwfSQL(rs!ArtdUs, lies.obMySQL, False)
   roLe(akt).LAnzl = doUmwfSQL(rs!LAnzl, lies.obMySQL, False)
   roLe(akt).LUhrz = doUmwfSQL(rs!LUhrz, lies.obMySQL, False)
   roLe(akt).LfBegr = doUmwfSQL(rs!LfBegr, lies.obMySQL, False)
   roLe(akt).Med = doUmwfSQL(rs!Med, lies.obMySQL, False)
   roLe(akt).LOrgan = doUmwfSQL(rs!LOrgan, lies.obMySQL, False)
   roLe(akt).LArztBf = doUmwfSQL(rs!LArztBf, lies.obMySQL, False)
   roLe(akt).DtlKbsV = doUmwfSQL(rs!DtlKbsV, lies.obMySQL, False)
   roLe(akt).LEntlDt = doUmwfSQL(rs!LEntlDt, lies.obMySQL, False)
   roLe(akt).Faktor = doUmwfSQL(rs!Faktor, lies.obMySQL, False)
   roLe(akt).LBSNR = doUmwfSQL(rs!LBSNR, lies.obMySQL, False)
   roLe(akt).Charge = doUmwfSQL(rs!Charge, lies.obMySQL, False)
   roLe(akt).Lanr = doUmwfSQL(rs!Lanr, lies.obMySQL, False)
   roLe(akt).letzVorg = rs!letzVorg
   roLe(akt).Ausn = doUmwfSQL(rs!Ausn, lies.obMySQL, False)
   roLe(akt).beme = doUmwfSQL(rs!beme, lies.obMySQL, False)
   roLe(akt).absPos = rs!absPos
   roLe(akt).aktZeit = rs!aktZeit
   roLe(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roLe(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roLe(akt).StByte = rs!StByte
   roLe(akt).lanrid = rs!lanrid
   roLe(akt).Sachkbez = doUmwfSQL(rs!Sachkbez, lies.obMySQL, False)
   roLe(akt).Sachkct = rs!Sachkct
   roLe(akt).Zone = doUmwfSQL(rs!Zone, lies.obMySQL, False)
   roLe(akt).Punkte = rs!Punkte
   roLe(akt).Lstgerbnr = rs!Lstgerbnr
   roLe(akt).Position = rs!Position
   roLe(akt).Eignung = rs!Eignung
   roLe(akt).Pruefzeit = rs!Pruefzeit
   roLe(akt).Kalkzeit = rs!Kalkzeit
   roLe(akt).BSNR = rs!BSNR
   roLe(akt).Abrechnungssperre = rs!Abrechnungssperre
   roLe(akt).Medikanr = rs!Medikanr
   roLe(akt).Laborkosten = rs!Laborkosten
   roLe(akt).Sachkostentyp = rs!Sachkostentyp
   roLe(akt).Usegebordtext = doUmwfSQL(rs!Usegebordtext, lies.obMySQL, False)
   roLe(akt).Arztliste = doUmwfSQL(rs!Arztliste, lies.obMySQL, False)
   roLe(akt).Ersteller = doUmwfSQL(rs!Ersteller, lies.obMySQL, False)
   roLe(akt).Änderer = doUmwfSQL(rs!Änderer, lies.obMySQL, False)
   roLe(akt).Kennzeichen_arzt_technik = doUmwfSQL(rs!Kennzeichen_arzt_technik, lies.obMySQL, False)
   roLe(akt).Auftragsschluessel = doUmwfSQL(rs!Auftragsschluessel, lies.obMySQL, False)
   roLe(akt).Auftragstext = doUmwfSQL(rs!Auftragstext, lies.obMySQL, False)
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rLe) > 0 Then
  For ri = 1 To UBound(rLe)
   If rLe(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roLe)
    If roLe(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roLe(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roLe(roendpe + UBound(rLe) - rbeg)
   For ri = rbeg To UBound(rLe)
    Call roLeZuw(roendpe + ri - rbeg, ri)
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

Public Function rLeDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLedump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLe)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLe(" & i & ").id:" & String$(33, "."), 33) & rLe(i).id
  Print #200, left$("rLe(" & i & ").FID:" & String$(33, "."), 33) & rLe(i).FID
  Print #200, left$("rLe(" & i & ").Pat_ID:" & String$(33, "."), 33) & rLe(i).Pat_ID
  Print #200, left$("rLe(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rLe(i).Zeitpunkt
  Print #200, left$("rLe(" & i & ").Leistung:" & String$(33, "."), 33) & "'" & rLe(i).Leistung & "'"
  Print #200, left$("rLe(" & i & ").ArtdUs:" & String$(33, "."), 33) & "'" & rLe(i).ArtdUs & "'"
  Print #200, left$("rLe(" & i & ").LAnzl:" & String$(33, "."), 33) & "'" & rLe(i).LAnzl & "'"
  Print #200, left$("rLe(" & i & ").LUhrz:" & String$(33, "."), 33) & "'" & rLe(i).LUhrz & "'"
  Print #200, left$("rLe(" & i & ").LfBegr:" & String$(33, "."), 33) & "'" & rLe(i).LfBegr & "'"
  Print #200, left$("rLe(" & i & ").Med:" & String$(33, "."), 33) & "'" & rLe(i).Med & "'"
  Print #200, left$("rLe(" & i & ").LOrgan:" & String$(33, "."), 33) & "'" & rLe(i).LOrgan & "'"
  Print #200, left$("rLe(" & i & ").LArztBf:" & String$(33, "."), 33) & "'" & rLe(i).LArztBf & "'"
  Print #200, left$("rLe(" & i & ").DtlKbsV:" & String$(33, "."), 33) & "'" & rLe(i).DtlKbsV & "'"
  Print #200, left$("rLe(" & i & ").LEntlDt:" & String$(33, "."), 33) & "'" & rLe(i).LEntlDt & "'"
  Print #200, left$("rLe(" & i & ").Faktor:" & String$(33, "."), 33) & "'" & rLe(i).Faktor & "'"
  Print #200, left$("rLe(" & i & ").LBSNR:" & String$(33, "."), 33) & "'" & rLe(i).LBSNR & "'"
  Print #200, left$("rLe(" & i & ").Charge:" & String$(33, "."), 33) & "'" & rLe(i).Charge & "'"
  Print #200, left$("rLe(" & i & ").LANR:" & String$(33, "."), 33) & "'" & rLe(i).Lanr & "'"
  Print #200, left$("rLe(" & i & ").letzVorg:" & String$(33, "."), 33) & rLe(i).letzVorg
  Print #200, left$("rLe(" & i & ").Ausn:" & String$(33, "."), 33) & "'" & rLe(i).Ausn & "'"
  Print #200, left$("rLe(" & i & ").Beme:" & String$(33, "."), 33) & "'" & rLe(i).beme & "'"
  Print #200, left$("rLe(" & i & ").absPos:" & String$(33, "."), 33) & rLe(i).absPos
  Print #200, left$("rLe(" & i & ").AktZeit:" & String$(33, "."), 33) & rLe(i).aktZeit
  Print #200, left$("rLe(" & i & ").QS:" & String$(33, "."), 33) & "'" & rLe(i).QS & "'"
  Print #200, left$("rLe(" & i & ").QT:" & String$(33, "."), 33) & "'" & rLe(i).QT & "'"
  Print #200, left$("rLe(" & i & ").StByte:" & String$(33, "."), 33) & rLe(i).StByte
  Print #200, left$("rLe(" & i & ").LANRid:" & String$(33, "."), 33) & rLe(i).lanrid
  Print #200, left$("rLe(" & i & ").Sachkbez:" & String$(33, "."), 33) & "'" & rLe(i).Sachkbez & "'"
  Print #200, left$("rLe(" & i & ").Sachkct:" & String$(33, "."), 33) & rLe(i).Sachkct
  Print #200, left$("rLe(" & i & ").Zone:" & String$(33, "."), 33) & "'" & rLe(i).Zone & "'"
  Print #200, left$("rLe(" & i & ").Punkte:" & String$(33, "."), 33) & "'" & rLe(i).Punkte & "'"
  Print #200, left$("rLe(" & i & ").Lstgerbnr:" & String$(33, "."), 33) & rLe(i).Lstgerbnr
  Print #200, left$("rLe(" & i & ").Position:" & String$(33, "."), 33) & rLe(i).Position
  Print #200, left$("rLe(" & i & ").Eignung:" & String$(33, "."), 33) & rLe(i).Eignung
  Print #200, left$("rLe(" & i & ").Pruefzeit:" & String$(33, "."), 33) & rLe(i).Pruefzeit
  Print #200, left$("rLe(" & i & ").Kalkzeit:" & String$(33, "."), 33) & rLe(i).Kalkzeit
  Print #200, left$("rLe(" & i & ").Bsnr:" & String$(33, "."), 33) & rLe(i).BSNR
  Print #200, left$("rLe(" & i & ").Abrechnungssperre:" & String$(33, "."), 33) & rLe(i).Abrechnungssperre
  Print #200, left$("rLe(" & i & ").Medikanr:" & String$(33, "."), 33) & rLe(i).Medikanr
  Print #200, left$("rLe(" & i & ").Laborkosten:" & String$(33, "."), 33) & rLe(i).Laborkosten
  Print #200, left$("rLe(" & i & ").Sachkostentyp:" & String$(33, "."), 33) & rLe(i).Sachkostentyp
  Print #200, left$("rLe(" & i & ").Usegebordtext:" & String$(33, "."), 33) & "'" & rLe(i).Usegebordtext & "'"
  Print #200, left$("rLe(" & i & ").Arztliste:" & String$(33, "."), 33) & "'" & rLe(i).Arztliste & "'"
  Print #200, left$("rLe(" & i & ").Ersteller:" & String$(33, "."), 33) & "'" & rLe(i).Ersteller & "'"
  Print #200, left$("rLe(" & i & ").Änderer:" & String$(33, "."), 33) & "'" & rLe(i).Änderer & "'"
  Print #200, left$("rLe(" & i & ").Kennzeichen_arzt_technik:" & String$(33, "."), 33) & "'" & rLe(i).Kennzeichen_arzt_technik & "'"
  Print #200, left$("rLe(" & i & ").Auftragsschluessel:" & String$(33, "."), 33) & "'" & rLe(i).Auftragsschluessel & "'"
  Print #200, left$("rLe(" & i & ").Auftragstext:" & String$(33, "."), 33) & "'" & rLe(i).Auftragstext & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' leistungenDump

Public Function leistungenSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "leistungen")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLe) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Leistung,ArtdUs,LAnzl,LUhrz,LfBegr,Med,LOrgan,LArztBf,DtlKbsV,LEntlDt," & _
     "Faktor,LBSNR,Charge,LANR,letzVorg,Ausn,Beme,absPos,AktZeit,QS," & _
     "QT,StByte,LANRid,Sachkbez,Sachkct,Zone,Punkte,Lstgerbnr,Position,Eignung," & _
     "Pruefzeit,Kalkzeit,Bsnr,Abrechnungssperre,Medikanr,Laborkosten,Sachkostentyp,Usegebordtext,Arztliste,Ersteller," & _
     "Änderer,Kennzeichen_arzt_technik,Auftragsschluessel,Auftragstext)              VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLe)
  rLe(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLe(i).FID, ",", rLe(i).Pat_ID, ",", DatFor_k(rLe(i).Zeitpunkt), ",'", rLe(i).Leistung, "','", rLe(i).ArtdUs, "','", rLe(i).LAnzl, "','", rLe(i).LUhrz, "','", rLe(i).LfBegr, "','", _
   rLe(i).Med, "','", rLe(i).LOrgan, "','", rLe(i).LArztBf, "','", rLe(i).DtlKbsV, "','", rLe(i).LEntlDt, "','", rLe(i).Faktor, "','", rLe(i).LBSNR, "','", rLe(i).Charge, "','", rLe(i).Lanr, "',", DatFor_k( _
   rLe(i).letzVorg), ",'", rLe(i).Ausn, "','", rLe(i).beme, "',", rLe(i).absPos, ",", DatFor_k(rLe(i).aktZeit), ",'", rLe(i).QS, "','", rLe(i).QT, "',", rLe(i).StByte, ",", _
   rLe(i).lanrid, ",'", rLe(i).Sachkbez, "',", rLe(i).Sachkct, ",'", rLe(i).Zone, "',", rLe(i).Punkte, ",", rLe(i).Lstgerbnr, ",", rLe(i).Position, ",", rLe(i).Eignung, ",", _
   rLe(i).Pruefzeit, ",", rLe(i).Kalkzeit, ",", rLe(i).BSNR, ",", rLe(i).Abrechnungssperre, ",", rLe(i).Medikanr, ",", rLe(i).Laborkosten, ",", rLe(i).Sachkostentyp, ",'", rLe(i).Usegebordtext, "','", _
   rLe(i).Arztliste, "','", rLe(i).Ersteller, "','", rLe(i).Änderer, "','", rLe(i).Kennzeichen_arzt_technik, "','", rLe(i).Auftragsschluessel, "','", rLe(i).Auftragstext, "')")
  If SammelInsert <> 0 And i < UBound(rLe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLe) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLe(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLe(" & i & "/" & UBound(rLe) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""leistungenSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(26)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLe), i)
  If Len(rLe(k).Leistung) > maxi(0) Then maxi(0) = Len(rLe(k).Leistung)
  If Len(rLe(k).ArtdUs) > maxi(1) Then maxi(1) = Len(rLe(k).ArtdUs)
  If Len(rLe(k).LAnzl) > maxi(2) Then maxi(2) = Len(rLe(k).LAnzl)
  If Len(rLe(k).LUhrz) > maxi(3) Then maxi(3) = Len(rLe(k).LUhrz)
  If Len(rLe(k).LfBegr) > maxi(4) Then maxi(4) = Len(rLe(k).LfBegr)
  If Len(rLe(k).Med) > maxi(5) Then maxi(5) = Len(rLe(k).Med)
  If Len(rLe(k).LOrgan) > maxi(6) Then maxi(6) = Len(rLe(k).LOrgan)
  If Len(rLe(k).LArztBf) > maxi(7) Then maxi(7) = Len(rLe(k).LArztBf)
  If Len(rLe(k).DtlKbsV) > maxi(8) Then maxi(8) = Len(rLe(k).DtlKbsV)
  If Len(rLe(k).LEntlDt) > maxi(9) Then maxi(9) = Len(rLe(k).LEntlDt)
  If Len(rLe(k).Faktor) > maxi(10) Then maxi(10) = Len(rLe(k).Faktor)
  If Len(rLe(k).LBSNR) > maxi(11) Then maxi(11) = Len(rLe(k).LBSNR)
  If Len(rLe(k).Charge) > maxi(12) Then maxi(12) = Len(rLe(k).Charge)
  If Len(rLe(k).Lanr) > maxi(13) Then maxi(13) = Len(rLe(k).Lanr)
  If Len(rLe(k).Ausn) > maxi(14) Then maxi(14) = Len(rLe(k).Ausn)
  If Len(rLe(k).beme) > maxi(15) Then maxi(15) = Len(rLe(k).beme)
  If Len(rLe(k).QS) > maxi(16) Then maxi(16) = Len(rLe(k).QS)
  If Len(rLe(k).QT) > maxi(17) Then maxi(17) = Len(rLe(k).QT)
  If Len(rLe(k).Sachkbez) > maxi(18) Then maxi(18) = Len(rLe(k).Sachkbez)
  If Len(rLe(k).Zone) > maxi(19) Then maxi(19) = Len(rLe(k).Zone)
  If Len(rLe(k).Usegebordtext) > maxi(20) Then maxi(20) = Len(rLe(k).Usegebordtext)
  If Len(rLe(k).Arztliste) > maxi(21) Then maxi(21) = Len(rLe(k).Arztliste)
  If Len(rLe(k).Ersteller) > maxi(22) Then maxi(22) = Len(rLe(k).Ersteller)
  If Len(rLe(k).Änderer) > maxi(23) Then maxi(23) = Len(rLe(k).Änderer)
  If Len(rLe(k).Kennzeichen_arzt_technik) > maxi(24) Then maxi(24) = Len(rLe(k).Kennzeichen_arzt_technik)
  If Len(rLe(k).Auftragsschluessel) > maxi(25) Then maxi(25) = Len(rLe(k).Auftragsschluessel)
  If Len(rLe(k).Auftragstext) > maxi(26) Then maxi(26) = Len(rLe(k).Auftragstext)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLe.Leistung: '" & rLe(k).Leistung & "' -> '" & left$(rLe(k).Leistung, maxL) & "'", True: rLe(k).Leistung = left$(rLe(k).Leistung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLe.ArtdUs: '" & rLe(k).ArtdUs & "' -> '" & left$(rLe(k).ArtdUs, maxL) & "'", True: rLe(k).ArtdUs = left$(rLe(k).ArtdUs, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLe.LAnzl: '" & rLe(k).LAnzl & "' -> '" & left$(rLe(k).LAnzl, maxL) & "'", True: rLe(k).LAnzl = left$(rLe(k).LAnzl, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLe.LUhrz: '" & rLe(k).LUhrz & "' -> '" & left$(rLe(k).LUhrz, maxL) & "'", True: rLe(k).LUhrz = left$(rLe(k).LUhrz, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLe.LfBegr: '" & rLe(k).LfBegr & "' -> '" & left$(rLe(k).LfBegr, maxL) & "'", True: rLe(k).LfBegr = left$(rLe(k).LfBegr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLe.Med: '" & rLe(k).Med & "' -> '" & left$(rLe(k).Med, maxL) & "'", True: rLe(k).Med = left$(rLe(k).Med, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLe.LOrgan: '" & rLe(k).LOrgan & "' -> '" & left$(rLe(k).LOrgan, maxL) & "'", True: rLe(k).LOrgan = left$(rLe(k).LOrgan, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLe.LArztBf: '" & rLe(k).LArztBf & "' -> '" & left$(rLe(k).LArztBf, maxL) & "'", True: rLe(k).LArztBf = left$(rLe(k).LArztBf, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLe.DtlKbsV: '" & rLe(k).DtlKbsV & "' -> '" & left$(rLe(k).DtlKbsV, maxL) & "'", True: rLe(k).DtlKbsV = left$(rLe(k).DtlKbsV, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLe.LEntlDt: '" & rLe(k).LEntlDt & "' -> '" & left$(rLe(k).LEntlDt, maxL) & "'", True: rLe(k).LEntlDt = left$(rLe(k).LEntlDt, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLe.Faktor: '" & rLe(k).Faktor & "' -> '" & left$(rLe(k).Faktor, maxL) & "'", True: rLe(k).Faktor = left$(rLe(k).Faktor, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLe.LBSNR: '" & rLe(k).LBSNR & "' -> '" & left$(rLe(k).LBSNR, maxL) & "'", True: rLe(k).LBSNR = left$(rLe(k).LBSNR, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLe.Charge: '" & rLe(k).Charge & "' -> '" & left$(rLe(k).Charge, maxL) & "'", True: rLe(k).Charge = left$(rLe(k).Charge, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLe.LANR: '" & rLe(k).Lanr & "' -> '" & left$(rLe(k).Lanr, maxL) & "'", True: rLe(k).Lanr = left$(rLe(k).Lanr, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ausn: '" & rLe(k).Ausn & "' -> '" & left$(rLe(k).Ausn, maxL) & "'", True: rLe(k).Ausn = left$(rLe(k).Ausn, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLe.Beme: '" & rLe(k).beme & "' -> '" & left$(rLe(k).beme, maxL) & "'", True: rLe(k).beme = left$(rLe(k).beme, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLe.QS: '" & rLe(k).QS & "' -> '" & left$(rLe(k).QS, maxL) & "'", True: rLe(k).QS = left$(rLe(k).QS, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLe.QT: '" & rLe(k).QT & "' -> '" & left$(rLe(k).QT, maxL) & "'", True: rLe(k).QT = left$(rLe(k).QT, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLe.Sachkbez: '" & rLe(k).Sachkbez & "' -> '" & left$(rLe(k).Sachkbez, maxL) & "'", True: rLe(k).Sachkbez = left$(rLe(k).Sachkbez, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLe.Zone: '" & rLe(k).Zone & "' -> '" & left$(rLe(k).Zone, maxL) & "'", True: rLe(k).Zone = left$(rLe(k).Zone, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLe.Usegebordtext: '" & rLe(k).Usegebordtext & "' -> '" & left$(rLe(k).Usegebordtext, maxL) & "'", True: rLe(k).Usegebordtext = left$(rLe(k).Usegebordtext, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLe.Arztliste: '" & rLe(k).Arztliste & "' -> '" & left$(rLe(k).Arztliste, maxL) & "'", True: rLe(k).Arztliste = left$(rLe(k).Arztliste, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLe.Ersteller: '" & rLe(k).Ersteller & "' -> '" & left$(rLe(k).Ersteller, maxL) & "'", True: rLe(k).Ersteller = left$(rLe(k).Ersteller, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLe.Änderer: '" & rLe(k).Änderer & "' -> '" & left$(rLe(k).Änderer, maxL) & "'", True: rLe(k).Änderer = left$(rLe(k).Änderer, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLe.Kennzeichen_arzt_technik: '" & rLe(k).Kennzeichen_arzt_technik & "' -> '" & left$(rLe(k).Kennzeichen_arzt_technik, maxL) & "'", True: rLe(k).Kennzeichen_arzt_technik = left$(rLe(k).Kennzeichen_arzt_technik, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rLe.Auftragsschluessel: '" & rLe(k).Auftragsschluessel & "' -> '" & left$(rLe(k).Auftragsschluessel, maxL) & "'", True: rLe(k).Auftragsschluessel = left$(rLe(k).Auftragsschluessel, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rLe.Auftragstext: '" & rLe(k).Auftragstext & "' -> '" & left$(rLe(k).Auftragstext, maxL) & "'", True: rLe(k).Auftragstext = left$(rLe(k).Auftragstext, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in leistungenSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' leistungenSpeichern

Public Function roMeZuw(i&, j&)
 roMe(i).FID = rMe(j).FID
 roMe(i).Pat_ID = rMe(j).Pat_ID
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
 roMe(i).Menge = rMe(j).Menge
 roMe(i).Nutzer = rMe(j).Nutzer
 roMe(i).absPos = rMe(j).absPos
 roMe(i).aktZeit = rMe(j).aktZeit
 roMe(i).StByte = rMe(j).StByte
 roMe(i).ergaenzt = rMe(j).ergaenzt
 roMe(i).mpart = rMe(j).mpart
End Function ' roMeZuw

Public Function MeZUnt%(i&, j&)
 If roMe(i).FID <> rMe(j).FID Then GoSub unter
 If roMe(i).Pat_ID <> rMe(j).Pat_ID Then GoSub unter
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
 If roMe(i).Menge <> rMe(j).Menge Then GoSub unter
 If roMe(i).Nutzer <> rMe(j).Nutzer Then GoSub unter
 If roMe(i).absPos <> rMe(j).absPos Then GoSub unter
 If roMe(i).aktZeit <> rMe(j).aktZeit Then GoSub unter
 If roMe(i).StByte <> rMe(j).StByte Then GoSub unter
 If roMe(i).ergaenzt <> rMe(j).ergaenzt Then GoSub unter
 If roMe(i).mpart <> rMe(j).mpart Then GoSub unter
 Exit Function
unter:
 MeZUnt = MeZUnt + 1
 Return
End Function ' MeZUnt

Public Function medplanLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(MPNr,0) MPNr,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt" & _
",COALESCE(Datum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Datum,COALESCE(Medikament,'') Medikament,COALESCE(MedAnfang,'') MedAnfang,COALESCE(Wirkstoff,'') Wirkstoff" & _
",COALESCE(PZN,0) PZN,COALESCE(FeldNr,0) FeldNr,COALESCE(mo,'') mo,COALESCE(mi,'') mi" & _
",COALESCE(nm,'') nm,COALESCE(ab,'') ab,COALESCE(zn,'') zn,COALESCE(bBed,0) bBed" & _
",COALESCE(Bemerkung,'') Bemerkung,COALESCE(Grund,'') Grund,COALESCE(Stärke,'') Stärke,COALESCE(Einheit,'') Einheit" & _
",COALESCE(Form,'') Form,COALESCE(Menge,0) Menge,COALESCE(Nutzer,'') Nutzer,COALESCE(AbsPos,0) AbsPos" & _
",COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte,COALESCE(ergaenzt,0) ergaenzt,COALESCE(mpart,0) mpart" & _
" FROM `medplan` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roMe(0)
 Else ' rs.EOF Then
  ReDim roMe(1)
  Do While Not rs.EOF
   akt = UBound(roMe)
   roMe(akt).FID = rs!FID
   roMe(akt).Pat_ID = rs!Pat_ID
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
   roMe(akt).Menge = rs!Menge
   roMe(akt).Nutzer = doUmwfSQL(rs!Nutzer, lies.obMySQL, False)
   roMe(akt).absPos = rs!absPos
   roMe(akt).aktZeit = rs!aktZeit
   roMe(akt).StByte = rs!StByte
   roMe(akt).ergaenzt = rs!ergaenzt
   roMe(akt).mpart = rs!mpart
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rMe) > 0 Then
  For ri = 1 To UBound(rMe)
   If rMe(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roMe)
    If roMe(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roMe(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roMe(roendpe + UBound(rMe) - rbeg)
   For ri = rbeg To UBound(rMe)
    Call roMeZuw(roendpe + ri - rbeg, ri)
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

Public Function rMeDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rMedump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rMe)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rMe(" & i & ").FID:" & String$(33, "."), 33) & rMe(i).FID
  Print #200, left$("rMe(" & i & ").Pat_ID:" & String$(33, "."), 33) & rMe(i).Pat_ID
  Print #200, left$("rMe(" & i & ").MPNr:" & String$(33, "."), 33) & rMe(i).MPNr
  Print #200, left$("rMe(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rMe(i).Zeitpunkt
  Print #200, left$("rMe(" & i & ").Datum:" & String$(33, "."), 33) & rMe(i).Datum
  Print #200, left$("rMe(" & i & ").Medikament:" & String$(33, "."), 33) & "'" & rMe(i).Medikament & "'"
  Print #200, left$("rMe(" & i & ").MedAnfang:" & String$(33, "."), 33) & "'" & rMe(i).MedAnfang & "'"
  Print #200, left$("rMe(" & i & ").Wirkstoff:" & String$(33, "."), 33) & "'" & rMe(i).Wirkstoff & "'"
  Print #200, left$("rMe(" & i & ").PZN:" & String$(33, "."), 33) & rMe(i).PZN
  Print #200, left$("rMe(" & i & ").FeldNr:" & String$(33, "."), 33) & rMe(i).FeldNr
  Print #200, left$("rMe(" & i & ").mo:" & String$(33, "."), 33) & "'" & rMe(i).mo & "'"
  Print #200, left$("rMe(" & i & ").mi:" & String$(33, "."), 33) & "'" & rMe(i).mi & "'"
  Print #200, left$("rMe(" & i & ").nm:" & String$(33, "."), 33) & "'" & rMe(i).nm & "'"
  Print #200, left$("rMe(" & i & ").ab:" & String$(33, "."), 33) & "'" & rMe(i).ab & "'"
  Print #200, left$("rMe(" & i & ").zn:" & String$(33, "."), 33) & "'" & rMe(i).Zn & "'"
  Print #200, left$("rMe(" & i & ").bBed:" & String$(33, "."), 33) & rMe(i).bBed
  Print #200, left$("rMe(" & i & ").Bemerkung:" & String$(33, "."), 33) & "'" & rMe(i).Bemerkung & "'"
  Print #200, left$("rMe(" & i & ").Grund:" & String$(33, "."), 33) & "'" & rMe(i).Grund & "'"
  Print #200, left$("rMe(" & i & ").Stärke:" & String$(33, "."), 33) & "'" & rMe(i).Stärke & "'"
  Print #200, left$("rMe(" & i & ").Einheit:" & String$(33, "."), 33) & "'" & rMe(i).Einheit & "'"
  Print #200, left$("rMe(" & i & ").Form:" & String$(33, "."), 33) & "'" & rMe(i).Form & "'"
  Print #200, left$("rMe(" & i & ").Menge:" & String$(33, "."), 33) & rMe(i).Menge
  Print #200, left$("rMe(" & i & ").Nutzer:" & String$(33, "."), 33) & "'" & rMe(i).Nutzer & "'"
  Print #200, left$("rMe(" & i & ").AbsPos:" & String$(33, "."), 33) & rMe(i).absPos
  Print #200, left$("rMe(" & i & ").AktZeit:" & String$(33, "."), 33) & rMe(i).aktZeit
  Print #200, left$("rMe(" & i & ").StByte:" & String$(33, "."), 33) & rMe(i).StByte
  Print #200, left$("rMe(" & i & ").ergaenzt:" & String$(33, "."), 33) & rMe(i).ergaenzt
  Print #200, left$("rMe(" & i & ").mpart:" & String$(33, "."), 33) & rMe(i).mpart
 Next i
 Close #200
 zeigan ffadat
End Function ' medplanDump

Public Function medplanSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "medplan")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rMe) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,MPNr," & _
     "ZeitPunkt,Datum,Medikament,MedAnfang,Wirkstoff,PZN,FeldNr,mo,mi,nm," & _
     "ab,zn,bBed,Bemerkung,Grund,Stärke,Einheit,Form,Menge,Nutzer," & _
     "AbsPos,AktZeit,StByte,ergaenzt,mpart)              VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rMe)
  rMe(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rMe(i).FID, ",", rMe(i).Pat_ID, ",", rMe(i).MPNr, ",", DatFor_k(rMe(i).Zeitpunkt), ",", DatFor_k(rMe(i).Datum), ",'", rMe(i).Medikament, "','", rMe(i).MedAnfang, "','", rMe(i).Wirkstoff, "',", _
   rMe(i).PZN, ",", rMe(i).FeldNr, ",'", rMe(i).mo, "','", rMe(i).mi, "','", rMe(i).nm, "','", rMe(i).ab, "','", rMe(i).Zn, "',", CStr(-(rMe(i).bBed <> 0)), ",'", rMe(i).Bemerkung, "','", _
   rMe(i).Grund, "','", rMe(i).Stärke, "','", rMe(i).Einheit, "','", rMe(i).Form, "',", rMe(i).Menge, ",'", rMe(i).Nutzer, "',", rMe(i).absPos, ",", DatFor_k(rMe(i).aktZeit), ",", _
   rMe(i).StByte, ",", CStr(-(rMe(i).ergaenzt <> 0)), ",", rMe(i).mpart, ")")
  If SammelInsert <> 0 And i < UBound(rMe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rMe) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rMe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rMe(" & i & "/" & UBound(rMe) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""medplanSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(13)
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
  If Len(rMe(k).Nutzer) > maxi(13) Then maxi(13) = Len(rMe(k).Nutzer)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rMe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rMe.Medikament: '" & rMe(k).Medikament & "' -> '" & left$(rMe(k).Medikament, maxL) & "'", True: rMe(k).Medikament = left$(rMe(k).Medikament, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rMe.MedAnfang: '" & rMe(k).MedAnfang & "' -> '" & left$(rMe(k).MedAnfang, maxL) & "'", True: rMe(k).MedAnfang = left$(rMe(k).MedAnfang, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rMe.Wirkstoff: '" & rMe(k).Wirkstoff & "' -> '" & left$(rMe(k).Wirkstoff, maxL) & "'", True: rMe(k).Wirkstoff = left$(rMe(k).Wirkstoff, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rMe.mo: '" & rMe(k).mo & "' -> '" & left$(rMe(k).mo, maxL) & "'", True: rMe(k).mo = left$(rMe(k).mo, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rMe.mi: '" & rMe(k).mi & "' -> '" & left$(rMe(k).mi, maxL) & "'", True: rMe(k).mi = left$(rMe(k).mi, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rMe.nm: '" & rMe(k).nm & "' -> '" & left$(rMe(k).nm, maxL) & "'", True: rMe(k).nm = left$(rMe(k).nm, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rMe.ab: '" & rMe(k).ab & "' -> '" & left$(rMe(k).ab, maxL) & "'", True: rMe(k).ab = left$(rMe(k).ab, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rMe.zn: '" & rMe(k).Zn & "' -> '" & left$(rMe(k).Zn, maxL) & "'", True: rMe(k).Zn = left$(rMe(k).Zn, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rMe.Bemerkung: '" & rMe(k).Bemerkung & "' -> '" & left$(rMe(k).Bemerkung, maxL) & "'", True: rMe(k).Bemerkung = left$(rMe(k).Bemerkung, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rMe.Grund: '" & rMe(k).Grund & "' -> '" & left$(rMe(k).Grund, maxL) & "'", True: rMe(k).Grund = left$(rMe(k).Grund, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rMe.Stärke: '" & rMe(k).Stärke & "' -> '" & left$(rMe(k).Stärke, maxL) & "'", True: rMe(k).Stärke = left$(rMe(k).Stärke, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rMe.Einheit: '" & rMe(k).Einheit & "' -> '" & left$(rMe(k).Einheit, maxL) & "'", True: rMe(k).Einheit = left$(rMe(k).Einheit, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rMe.Form: '" & rMe(k).Form & "' -> '" & left$(rMe(k).Form, maxL) & "'", True: rMe(k).Form = left$(rMe(k).Form, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rMe.Nutzer: '" & rMe(k).Nutzer & "' -> '" & left$(rMe(k).Nutzer, maxL) & "'", True: rMe(k).Nutzer = left$(rMe(k).Nutzer, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in medplanSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' medplanSpeichern

Public Function roReZuw(i&, j&)
 roRe(i).FID = rRe(j).FID
 roRe(i).Pat_ID = rRe(j).Pat_ID
 roRe(i).Zeitpunkt = rRe(j).Zeitpunkt
 roRe(i).Rezept = rRe(j).Rezept
 roRe(i).RKlnm = rRe(j).RKlnm
 roRe(i).Rezeptklasse = rRe(j).Rezeptklasse
 roRe(i).Rezklkurz = rRe(j).Rezklkurz
 roRe(i).erez = rRe(j).erez
 roRe(i).ered = rRe(j).ered
 roRe(i).Rezkllang = rRe(j).Rezkllang
 roRe(i).rea = rRe(j).rea
 roRe(i).kbez = rRe(j).kbez
 roRe(i).Medikament = rRe(j).Medikament
 roRe(i).auti = rRe(j).auti
 roRe(i).obBTM = rRe(j).obBTM
 roRe(i).noctu = rRe(j).noctu
 roRe(i).ersatzv = rRe(j).ersatzv
 roRe(i).einzdr = rRe(j).einzdr
 roRe(i).freig = rRe(j).freig
 roRe(i).obst = rRe(j).obst
 roRe(i).anzl = rRe(j).anzl
 roRe(i).PZN = rRe(j).PZN
 roRe(i).absPos = rRe(j).absPos
 roRe(i).aktZeit = rRe(j).aktZeit
 roRe(i).QS = rRe(j).QS
 roRe(i).QT = rRe(j).QT
 roRe(i).StByte = rRe(j).StByte
 roRe(i).lanrid = rRe(j).lanrid
 roRe(i).id = rRe(j).id
 roRe(i).FEintragsart = rRe(j).FEintragsart
 roRe(i).Rezeptart = rRe(j).Rezeptart
End Function ' roReZuw

Public Function ReZUnt%(i&, j&)
 If roRe(i).FID <> rRe(j).FID Then GoSub unter
 If roRe(i).Pat_ID <> rRe(j).Pat_ID Then GoSub unter
 If roRe(i).Zeitpunkt <> rRe(j).Zeitpunkt Then GoSub unter
 If roRe(i).Rezept <> rRe(j).Rezept Then GoSub unter
 If roRe(i).RKlnm <> rRe(j).RKlnm Then GoSub unter
 If roRe(i).Rezeptklasse <> rRe(j).Rezeptklasse Then GoSub unter
 If roRe(i).Rezklkurz <> rRe(j).Rezklkurz Then GoSub unter
 If roRe(i).erez <> rRe(j).erez Then GoSub unter
 If roRe(i).ered <> rRe(j).ered Then GoSub unter
 If roRe(i).Rezkllang <> rRe(j).Rezkllang Then GoSub unter
 If roRe(i).rea <> rRe(j).rea Then GoSub unter
 If roRe(i).kbez <> rRe(j).kbez Then GoSub unter
 If roRe(i).Medikament <> rRe(j).Medikament Then GoSub unter
 If roRe(i).auti <> rRe(j).auti Then GoSub unter
 If roRe(i).obBTM <> rRe(j).obBTM Then GoSub unter
 If roRe(i).noctu <> rRe(j).noctu Then GoSub unter
 If roRe(i).ersatzv <> rRe(j).ersatzv Then GoSub unter
 If roRe(i).einzdr <> rRe(j).einzdr Then GoSub unter
 If roRe(i).freig <> rRe(j).freig Then GoSub unter
 If roRe(i).obst <> rRe(j).obst Then GoSub unter
 If roRe(i).anzl <> rRe(j).anzl Then GoSub unter
 If roRe(i).PZN <> rRe(j).PZN Then GoSub unter
 If roRe(i).absPos <> rRe(j).absPos Then GoSub unter
 If roRe(i).aktZeit <> rRe(j).aktZeit Then GoSub unter
 If roRe(i).QS <> rRe(j).QS Then GoSub unter
 If roRe(i).QT <> rRe(j).QT Then GoSub unter
 If roRe(i).StByte <> rRe(j).StByte Then GoSub unter
 If roRe(i).lanrid <> rRe(j).lanrid Then GoSub unter
 If roRe(i).id <> rRe(j).id Then GoSub unter
 If roRe(i).FEintragsart <> rRe(j).FEintragsart Then GoSub unter
 If roRe(i).Rezeptart <> rRe(j).Rezeptart Then GoSub unter
 Exit Function
unter:
 ReZUnt = ReZUnt + 1
 Return
End Function ' ReZUnt

Public Function rezepteintraegeLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(Rezept,'') Rezept" & _
",COALESCE(RKlnm,'') RKlnm,COALESCE(Rezeptklasse,'') Rezeptklasse,COALESCE(Rezklkurz,'') Rezklkurz,COALESCE(erez,0) erez" & _
",COALESCE(ered,0) ered,COALESCE(Rezkllang,'') Rezkllang,COALESCE(rea,'') rea,COALESCE(kbez,'') kbez" & _
",COALESCE(Medikament,'') Medikament,COALESCE(auti,0) auti,COALESCE(obBTM,0) obBTM,COALESCE(noctu,0) noctu" & _
",COALESCE(ersatzv,0) ersatzv,COALESCE(einzdr,0) einzdr,COALESCE(freig,0) freig,COALESCE(obst,0) obst" & _
",COALESCE(anzl,0) anzl,COALESCE(PZN,'') PZN,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit" & _
",COALESCE(QS,'') QS,COALESCE(QT,'') QT,COALESCE(StByte,0) StByte,COALESCE(LANRid,0) LANRid" & _
",COALESCE(id,0) id,COALESCE(FEintragsart,'') FEintragsart,COALESCE(Rezeptart,0) Rezeptart FROM `rezepteintraege` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roRe(0)
 Else ' rs.EOF Then
  ReDim roRe(1)
  Do While Not rs.EOF
   akt = UBound(roRe)
   roRe(akt).FID = rs!FID
   roRe(akt).Pat_ID = rs!Pat_ID
   roRe(akt).Zeitpunkt = rs!Zeitpunkt
   roRe(akt).Rezept = doUmwfSQL(rs!Rezept, lies.obMySQL, False)
   roRe(akt).RKlnm = doUmwfSQL(rs!RKlnm, lies.obMySQL, False)
   roRe(akt).Rezeptklasse = doUmwfSQL(rs!Rezeptklasse, lies.obMySQL, False)
   roRe(akt).Rezklkurz = doUmwfSQL(rs!Rezklkurz, lies.obMySQL, False)
   roRe(akt).erez = rs!erez
   roRe(akt).ered = rs!ered
   roRe(akt).Rezkllang = doUmwfSQL(rs!Rezkllang, lies.obMySQL, False)
   roRe(akt).rea = doUmwfSQL(rs!rea, lies.obMySQL, False)
   roRe(akt).kbez = doUmwfSQL(rs!kbez, lies.obMySQL, False)
   roRe(akt).Medikament = doUmwfSQL(rs!Medikament, lies.obMySQL, False)
   roRe(akt).auti = rs!auti
   roRe(akt).obBTM = rs!obBTM
   roRe(akt).noctu = rs!noctu
   roRe(akt).ersatzv = rs!ersatzv
   roRe(akt).einzdr = rs!einzdr
   roRe(akt).freig = rs!freig
   roRe(akt).obst = rs!obst
   roRe(akt).anzl = rs!anzl
   roRe(akt).PZN = doUmwfSQL(rs!PZN, lies.obMySQL, False)
   roRe(akt).absPos = rs!absPos
   roRe(akt).aktZeit = rs!aktZeit
   roRe(akt).QS = doUmwfSQL(rs!QS, lies.obMySQL, False)
   roRe(akt).QT = doUmwfSQL(rs!QT, lies.obMySQL, False)
   roRe(akt).StByte = rs!StByte
   roRe(akt).lanrid = rs!lanrid
   roRe(akt).id = rs!id
   roRe(akt).FEintragsart = doUmwfSQL(rs!FEintragsart, lies.obMySQL, False)
   roRe(akt).Rezeptart = rs!Rezeptart
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rRe) > 0 Then
  For ri = 1 To UBound(rRe)
   If rRe(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roRe)
    If roRe(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roRe(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roRe(roendpe + UBound(rRe) - rbeg)
   For ri = rbeg To UBound(rRe)
    Call roReZuw(roendpe + ri - rbeg, ri)
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

Public Function rReDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rRedump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rRe)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rRe(" & i & ").FID:" & String$(33, "."), 33) & rRe(i).FID
  Print #200, left$("rRe(" & i & ").Pat_ID:" & String$(33, "."), 33) & rRe(i).Pat_ID
  Print #200, left$("rRe(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rRe(i).Zeitpunkt
  Print #200, left$("rRe(" & i & ").Rezept:" & String$(33, "."), 33) & "'" & rRe(i).Rezept & "'"
  Print #200, left$("rRe(" & i & ").RKlnm:" & String$(33, "."), 33) & "'" & rRe(i).RKlnm & "'"
  Print #200, left$("rRe(" & i & ").Rezeptklasse:" & String$(33, "."), 33) & "'" & rRe(i).Rezeptklasse & "'"
  Print #200, left$("rRe(" & i & ").Rezklkurz:" & String$(33, "."), 33) & "'" & rRe(i).Rezklkurz & "'"
  Print #200, left$("rRe(" & i & ").erez:" & String$(33, "."), 33) & rRe(i).erez
  Print #200, left$("rRe(" & i & ").ered:" & String$(33, "."), 33) & rRe(i).ered
  Print #200, left$("rRe(" & i & ").Rezkllang:" & String$(33, "."), 33) & "'" & rRe(i).Rezkllang & "'"
  Print #200, left$("rRe(" & i & ").rea:" & String$(33, "."), 33) & "'" & rRe(i).rea & "'"
  Print #200, left$("rRe(" & i & ").kbez:" & String$(33, "."), 33) & "'" & rRe(i).kbez & "'"
  Print #200, left$("rRe(" & i & ").Medikament:" & String$(33, "."), 33) & "'" & rRe(i).Medikament & "'"
  Print #200, left$("rRe(" & i & ").auti:" & String$(33, "."), 33) & rRe(i).auti
  Print #200, left$("rRe(" & i & ").obBTM:" & String$(33, "."), 33) & rRe(i).obBTM
  Print #200, left$("rRe(" & i & ").noctu:" & String$(33, "."), 33) & rRe(i).noctu
  Print #200, left$("rRe(" & i & ").ersatzv:" & String$(33, "."), 33) & rRe(i).ersatzv
  Print #200, left$("rRe(" & i & ").einzdr:" & String$(33, "."), 33) & rRe(i).einzdr
  Print #200, left$("rRe(" & i & ").freig:" & String$(33, "."), 33) & rRe(i).freig
  Print #200, left$("rRe(" & i & ").obst:" & String$(33, "."), 33) & rRe(i).obst
  Print #200, left$("rRe(" & i & ").anzl:" & String$(33, "."), 33) & rRe(i).anzl
  Print #200, left$("rRe(" & i & ").PZN:" & String$(33, "."), 33) & "'" & rRe(i).PZN & "'"
  Print #200, left$("rRe(" & i & ").absPos:" & String$(33, "."), 33) & rRe(i).absPos
  Print #200, left$("rRe(" & i & ").AktZeit:" & String$(33, "."), 33) & rRe(i).aktZeit
  Print #200, left$("rRe(" & i & ").QS:" & String$(33, "."), 33) & "'" & rRe(i).QS & "'"
  Print #200, left$("rRe(" & i & ").QT:" & String$(33, "."), 33) & "'" & rRe(i).QT & "'"
  Print #200, left$("rRe(" & i & ").StByte:" & String$(33, "."), 33) & rRe(i).StByte
  Print #200, left$("rRe(" & i & ").LANRid:" & String$(33, "."), 33) & rRe(i).lanrid
  Print #200, left$("rRe(" & i & ").id:" & String$(33, "."), 33) & rRe(i).id
  Print #200, left$("rRe(" & i & ").FEintragsart:" & String$(33, "."), 33) & "'" & rRe(i).FEintragsart & "'"
  Print #200, left$("rRe(" & i & ").Rezeptart:" & String$(33, "."), 33) & rRe(i).Rezeptart
 Next i
 Close #200
 zeigan ffadat
End Function ' rezepteintraegeDump

Public Function rezepteintraegeSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "rezepteintraege")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rRe) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Rezept,RKlnm,Rezeptklasse,Rezklkurz,erez,ered,Rezkllang,rea,kbez,Medikament," & _
     "auti,obBTM,noctu,ersatzv,einzdr,freig,obst,anzl,PZN,absPos," & _
     "AktZeit,QS,QT,StByte,LANRid,FEintragsart,Rezeptart)              VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rRe)
  rRe(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rRe(i).FID, ",", rRe(i).Pat_ID, ",", DatFor_k(rRe(i).Zeitpunkt), ",'", rRe(i).Rezept, "','", rRe(i).RKlnm, "','", rRe(i).Rezeptklasse, "','", rRe(i).Rezklkurz, "',", rRe(i).erez, ",", _
   rRe(i).ered, ",'", rRe(i).Rezkllang, "','", rRe(i).rea, "','", rRe(i).kbez, "','", rRe(i).Medikament, "',", rRe(i).auti, ",", rRe(i).obBTM, ",", rRe(i).noctu, ",", rRe(i).ersatzv, ",", _
   rRe(i).einzdr, ",", rRe(i).freig, ",", rRe(i).obst, ",", rRe(i).anzl, ",'", rRe(i).PZN, "',", rRe(i).absPos, ",", DatFor_k(rRe(i).aktZeit), ",'", rRe(i).QS, "','", rRe(i).QT, "',", _
   rRe(i).StByte, ",", rRe(i).lanrid, ",'", rRe(i).FEintragsart, "',", rRe(i).Rezeptart, ")")
  If SammelInsert <> 0 And i < UBound(rRe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rRe) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rRe(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rRe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rRe(" & i & "/" & UBound(rRe) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""rezepteintraegeSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(11)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRe), i)
  If Len(rRe(k).Rezept) > maxi(0) Then maxi(0) = Len(rRe(k).Rezept)
  If Len(rRe(k).RKlnm) > maxi(1) Then maxi(1) = Len(rRe(k).RKlnm)
  If Len(rRe(k).Rezeptklasse) > maxi(2) Then maxi(2) = Len(rRe(k).Rezeptklasse)
  If Len(rRe(k).Rezklkurz) > maxi(3) Then maxi(3) = Len(rRe(k).Rezklkurz)
  If Len(rRe(k).Rezkllang) > maxi(4) Then maxi(4) = Len(rRe(k).Rezkllang)
  If Len(rRe(k).rea) > maxi(5) Then maxi(5) = Len(rRe(k).rea)
  If Len(rRe(k).kbez) > maxi(6) Then maxi(6) = Len(rRe(k).kbez)
  If Len(rRe(k).Medikament) > maxi(7) Then maxi(7) = Len(rRe(k).Medikament)
  If Len(rRe(k).PZN) > maxi(8) Then maxi(8) = Len(rRe(k).PZN)
  If Len(rRe(k).QS) > maxi(9) Then maxi(9) = Len(rRe(k).QS)
  If Len(rRe(k).QT) > maxi(10) Then maxi(10) = Len(rRe(k).QT)
  If Len(rRe(k).FEintragsart) > maxi(11) Then maxi(11) = Len(rRe(k).FEintragsart)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezept: '" & rRe(k).Rezept & "' -> '" & left$(rRe(k).Rezept, maxL) & "'", True: rRe(k).Rezept = left$(rRe(k).Rezept, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRe.RKlnm: '" & rRe(k).RKlnm & "' -> '" & left$(rRe(k).RKlnm, maxL) & "'", True: rRe(k).RKlnm = left$(rRe(k).RKlnm, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezeptklasse: '" & rRe(k).Rezeptklasse & "' -> '" & left$(rRe(k).Rezeptklasse, maxL) & "'", True: rRe(k).Rezeptklasse = left$(rRe(k).Rezeptklasse, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezklkurz: '" & rRe(k).Rezklkurz & "' -> '" & left$(rRe(k).Rezklkurz, maxL) & "'", True: rRe(k).Rezklkurz = left$(rRe(k).Rezklkurz, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRe.Rezkllang: '" & rRe(k).Rezkllang & "' -> '" & left$(rRe(k).Rezkllang, maxL) & "'", True: rRe(k).Rezkllang = left$(rRe(k).Rezkllang, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rRe.rea: '" & rRe(k).rea & "' -> '" & left$(rRe(k).rea, maxL) & "'", True: rRe(k).rea = left$(rRe(k).rea, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rRe.kbez: '" & rRe(k).kbez & "' -> '" & left$(rRe(k).kbez, maxL) & "'", True: rRe(k).kbez = left$(rRe(k).kbez, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rRe.Medikament: '" & rRe(k).Medikament & "' -> '" & left$(rRe(k).Medikament, maxL) & "'", True: rRe(k).Medikament = left$(rRe(k).Medikament, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rRe.PZN: '" & rRe(k).PZN & "' -> '" & left$(rRe(k).PZN, maxL) & "'", True: rRe(k).PZN = left$(rRe(k).PZN, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rRe.QS: '" & rRe(k).QS & "' -> '" & left$(rRe(k).QS, maxL) & "'", True: rRe(k).QS = left$(rRe(k).QS, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rRe.QT: '" & rRe(k).QT & "' -> '" & left$(rRe(k).QT, maxL) & "'", True: rRe(k).QT = left$(rRe(k).QT, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rRe.FEintragsart: '" & rRe(k).FEintragsart & "' -> '" & left$(rRe(k).FEintragsart, maxL) & "'", True: rRe(k).FEintragsart = left$(rRe(k).FEintragsart, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rezepteintraegeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rezepteintraegeSpeichern

Public Function roRrZuw(i&, j&)
 roRr(i).FID = rRr(j).FID
 roRr(i).Pat_ID = rRr(j).Pat_ID
 roRr(i).Zeitpunkt = rRr(j).Zeitpunkt
 roRr(i).FormTitel = rRr(j).FormTitel
 roRr(i).art = rRr(j).art
 roRr(i).RR = rRr(j).RR
 roRr(i).Puls = rRr(j).Puls
 roRr(i).RRsyst = rRr(j).RRsyst
 roRr(i).RRdiast = rRr(j).RRdiast
 roRr(i).RRzahl = rRr(j).RRzahl
 roRr(i).Quelle = rRr(j).Quelle
 roRr(i).Bemerkung = rRr(j).Bemerkung
 roRr(i).absPos = rRr(j).absPos
 roRr(i).aktZeit = rRr(j).aktZeit
 roRr(i).StByte = rRr(j).StByte
End Function ' roRrZuw

Public Function RrZUnt%(i&, j&)
 If roRr(i).FID <> rRr(j).FID Then GoSub unter
 If roRr(i).Pat_ID <> rRr(j).Pat_ID Then GoSub unter
 If roRr(i).Zeitpunkt <> rRr(j).Zeitpunkt Then GoSub unter
 If roRr(i).FormTitel <> rRr(j).FormTitel Then GoSub unter
 If roRr(i).art <> rRr(j).art Then GoSub unter
 If roRr(i).RR <> rRr(j).RR Then GoSub unter
 If roRr(i).Puls <> rRr(j).Puls Then GoSub unter
 If roRr(i).RRsyst <> rRr(j).RRsyst Then GoSub unter
 If roRr(i).RRdiast <> rRr(j).RRdiast Then GoSub unter
 If roRr(i).RRzahl <> rRr(j).RRzahl Then GoSub unter
 If roRr(i).Quelle <> rRr(j).Quelle Then GoSub unter
 If roRr(i).Bemerkung <> rRr(j).Bemerkung Then GoSub unter
 If roRr(i).absPos <> rRr(j).absPos Then GoSub unter
 If roRr(i).aktZeit <> rRr(j).aktZeit Then GoSub unter
 If roRr(i).StByte <> rRr(j).StByte Then GoSub unter
 Exit Function
unter:
 RrZUnt = RrZUnt + 1
 Return
End Function ' RrZUnt

Public Function rrLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(FormTitel,'') FormTitel" & _
",COALESCE(art,'') art,COALESCE(RR,'') RR,COALESCE(Puls,0) Puls,COALESCE(RRsyst,0) RRsyst" & _
",COALESCE(RRdiast,0) RRdiast,COALESCE(RRzahl,0) RRzahl,COALESCE(Quelle,'') Quelle,COALESCE(Bemerkung,'') Bemerkung" & _
",COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte FROM `rr` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roRr(0)
 Else ' rs.EOF Then
  ReDim roRr(1)
  Do While Not rs.EOF
   akt = UBound(roRr)
   roRr(akt).FID = rs!FID
   roRr(akt).Pat_ID = rs!Pat_ID
   roRr(akt).Zeitpunkt = rs!Zeitpunkt
   roRr(akt).FormTitel = doUmwfSQL(rs!FormTitel, lies.obMySQL, False)
   roRr(akt).art = doUmwfSQL(rs!art, lies.obMySQL, False)
   roRr(akt).RR = doUmwfSQL(rs!RR, lies.obMySQL, False)
   roRr(akt).Puls = rs!Puls
   roRr(akt).RRsyst = rs!RRsyst
   roRr(akt).RRdiast = rs!RRdiast
   roRr(akt).RRzahl = rs!RRzahl
   roRr(akt).Quelle = doUmwfSQL(rs!Quelle, lies.obMySQL, False)
   roRr(akt).Bemerkung = doUmwfSQL(rs!Bemerkung, lies.obMySQL, False)
   roRr(akt).absPos = rs!absPos
   roRr(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rRr) > 0 Then
  For ri = 1 To UBound(rRr)
   If rRr(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roRr)
    If roRr(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roRr(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roRr(roendpe + UBound(rRr) - rbeg)
   For ri = rbeg To UBound(rRr)
    Call roRrZuw(roendpe + ri - rbeg, ri)
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

Public Function rRrDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rRrdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rRr)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rRr(" & i & ").FID:" & String$(33, "."), 33) & rRr(i).FID
  Print #200, left$("rRr(" & i & ").Pat_ID:" & String$(33, "."), 33) & rRr(i).Pat_ID
  Print #200, left$("rRr(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rRr(i).Zeitpunkt
  Print #200, left$("rRr(" & i & ").FormTitel:" & String$(33, "."), 33) & "'" & rRr(i).FormTitel & "'"
  Print #200, left$("rRr(" & i & ").art:" & String$(33, "."), 33) & "'" & rRr(i).art & "'"
  Print #200, left$("rRr(" & i & ").RR:" & String$(33, "."), 33) & "'" & rRr(i).RR & "'"
  Print #200, left$("rRr(" & i & ").Puls:" & String$(33, "."), 33) & rRr(i).Puls
  Print #200, left$("rRr(" & i & ").RRsyst:" & String$(33, "."), 33) & rRr(i).RRsyst
  Print #200, left$("rRr(" & i & ").RRdiast:" & String$(33, "."), 33) & rRr(i).RRdiast
  Print #200, left$("rRr(" & i & ").RRzahl:" & String$(33, "."), 33) & rRr(i).RRzahl
  Print #200, left$("rRr(" & i & ").Quelle:" & String$(33, "."), 33) & "'" & rRr(i).Quelle & "'"
  Print #200, left$("rRr(" & i & ").Bemerkung:" & String$(33, "."), 33) & "'" & rRr(i).Bemerkung & "'"
  Print #200, left$("rRr(" & i & ").absPos:" & String$(33, "."), 33) & rRr(i).absPos
  Print #200, left$("rRr(" & i & ").AktZeit:" & String$(33, "."), 33) & rRr(i).aktZeit
  Print #200, left$("rRr(" & i & ").StByte:" & String$(33, "."), 33) & rRr(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' rrDump

Public Function rrSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "rr")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rRr) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "FormTitel,art,RR,Puls,RRsyst,RRdiast,RRzahl,Quelle,Bemerkung,absPos," & _
     "AktZeit,StByte)        VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rRr)
  rRr(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rRr(i).FID, ",", rRr(i).Pat_ID, ",", DatFor_k(rRr(i).Zeitpunkt), ",'", rRr(i).FormTitel, "','", rRr(i).art, "','", rRr(i).RR, "',", rRr(i).Puls, ",", rRr(i).RRsyst, ",", rRr(i).RRdiast, ",", _
   rRr(i).RRzahl, ",'", rRr(i).Quelle, "','", rRr(i).Bemerkung, "',", rRr(i).absPos, ",", DatFor_k(rRr(i).aktZeit), ",", rRr(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rRr) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rRr) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rRr)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rRr(" & i & "/" & UBound(rRr) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""rrSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(4)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRr), i)
  If Len(rRr(k).FormTitel) > maxi(0) Then maxi(0) = Len(rRr(k).FormTitel)
  If Len(rRr(k).art) > maxi(1) Then maxi(1) = Len(rRr(k).art)
  If Len(rRr(k).RR) > maxi(2) Then maxi(2) = Len(rRr(k).RR)
  If Len(rRr(k).Quelle) > maxi(3) Then maxi(3) = Len(rRr(k).Quelle)
  If Len(rRr(k).Bemerkung) > maxi(4) Then maxi(4) = Len(rRr(k).Bemerkung)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rRr), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rRr.FormTitel: '" & rRr(k).FormTitel & "' -> '" & left$(rRr(k).FormTitel, maxL) & "'", True: rRr(k).FormTitel = left$(rRr(k).FormTitel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rRr.art: '" & rRr(k).art & "' -> '" & left$(rRr(k).art, maxL) & "'", True: rRr(k).art = left$(rRr(k).art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rRr.RR: '" & rRr(k).RR & "' -> '" & left$(rRr(k).RR, maxL) & "'", True: rRr(k).RR = left$(rRr(k).RR, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rRr.Quelle: '" & rRr(k).Quelle & "' -> '" & left$(rRr(k).Quelle, maxL) & "'", True: rRr(k).Quelle = left$(rRr(k).Quelle, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rRr.Bemerkung: '" & rRr(k).Bemerkung & "' -> '" & left$(rRr(k).Bemerkung, maxL) & "'", True: rRr(k).Bemerkung = left$(rRr(k).Bemerkung, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in rrSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' rrSpeichern

Public Function rKvDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rKvdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rKv)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rKv(" & i & ").lfdnr:" & String$(33, "."), 33) & rKv(i).lfdnr
  Print #200, left$("rKv(" & i & ").Pat_ID:" & String$(33, "."), 33) & rKv(i).Pat_ID
  Print #200, left$("rKv(" & i & ").KVNr:" & String$(33, "."), 33) & "'" & rKv(i).KVNr & "'"
  Print #200, left$("rKv(" & i & ").absPos:" & String$(33, "."), 33) & rKv(i).absPos
  Print #200, left$("rKv(" & i & ").AktZeit:" & String$(33, "."), 33) & rKv(i).aktZeit
  Print #200, left$("rKv(" & i & ").StByte:" & String$(33, "."), 33) & rKv(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' kvnrueDump

Public Function kvnrueSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "kvnrue")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rKv) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Pat_ID,KVNr,absPos," & _
     "AktZeit,StByte)        VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rKv)
  rKv(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rKv(i).Pat_ID, ",'", rKv(i).KVNr, "',", rKv(i).absPos, ",", DatFor_k(rKv(i).aktZeit), ",", rKv(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rKv) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rKv) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rKv(i).lfdnr = myEFrag("SELECT MAX(lfdnr)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rKv)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rKv(" & i & "/" & UBound(rKv) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""kvnrueSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rKv), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rKv.KVNr: '" & rKv(k).KVNr & "' -> '" & left$(rKv(k).KVNr, maxL) & "'", True: rKv(k).KVNr = left$(rKv(k).KVNr, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in kvnrueSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' kvnrueSpeichern

Public Function rUnDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rUndump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rUn)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rUn(" & i & ").Kennung:" & String$(33, "."), 33) & "'" & rUn(i).Kennung & "'"
  Print #200, left$("rUn(" & i & ").absPos:" & String$(33, "."), 33) & rUn(i).absPos
  Print #200, left$("rUn(" & i & ").StByte:" & String$(33, "."), 33) & rUn(i).StByte
  Print #200, left$("rUn(" & i & ").Pat_id:" & String$(33, "."), 33) & rUn(i).Pat_ID
  Print #200, left$("rUn(" & i & ").Inhalt:" & String$(33, "."), 33) & "'" & rUn(i).Inhalt & "'"
  Print #200, left$("rUn(" & i & ").Zeitpunkt:" & String$(33, "."), 33) & rUn(i).Zeitpunkt
 Next i
 Close #200
 zeigan ffadat
End Function ' unbek_kennDump

Public Function unbek_kennSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "unbek_kenn")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 syscmd 4, pid & ": Speichere " & UBound(rUn) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Kennung,absPos,StByte," & _
     "Pat_id,Inhalt,Zeitpunkt)             VALUES"))
 If Not AllePat Then
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = rUn1 + 1 To UBound(rUn)
  rUn(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = rUn1 + 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = rUn1 + 1 Then
  csql.AppVar Array("('", rUn(i).Kennung, "',", rUn(i).absPos, ",", rUn(i).StByte, ",", rUn(i).Pat_ID, ",'", rUn(i).Inhalt, "',", DatFor_k(rUn(i).Zeitpunkt), ")")
  If SammelInsert <> 0 And i < UBound(rUn) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUn) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rUn)
  DoEvents
 Next i
 rUn1 = UBound(rUn)
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rUn(" & i & "/" & UBound(rUn) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""unbek_kennSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, rUn1 + 1, i) To IIf(SammelInsert <> 0, UBound(rUn), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUn.Kennung: '" & rUn(k).Kennung & "' -> '" & left$(rUn(k).Kennung, maxL) & "'", True: rUn(k).Kennung = left$(rUn(k).Kennung, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUn.Inhalt: '" & rUn(k).Inhalt & "' -> '" & left$(rUn(k).Inhalt, maxL) & "'", True: rUn(k).Inhalt = left$(rUn(k).Inhalt, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in unbek_kennSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' unbek_kennSpeichern

Public Function roDmZuw(i&, j&)
 roDm(i).Abk = rDm(j).Abk
 roDm(i).DMPArt = rDm(j).DMPArt
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
 roDm(i).Pat_ID = rDm(j).Pat_ID
 roDm(i).StByte = rDm(j).StByte
 roDm(i).aktZeit = rDm(j).aktZeit
 roDm(i).lanrid = rDm(j).lanrid
 roDm(i).Zusatzdaten = rDm(j).Zusatzdaten
 roDm(i).Druckdatum = rDm(j).Druckdatum
 roDm(i).uDat = rDm(j).uDat
 roDm(i).eGFR = rDm(j).eGFR
End Function ' roDmZuw

Public Function DmZUnt%(i&, j&)
 If roDm(i).Abk <> rDm(j).Abk Then GoSub unter
 If roDm(i).DMPArt <> rDm(j).DMPArt Then GoSub unter
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
 If roDm(i).Pat_ID <> rDm(j).Pat_ID Then GoSub unter
 If roDm(i).StByte <> rDm(j).StByte Then GoSub unter
 If roDm(i).aktZeit <> rDm(j).aktZeit Then GoSub unter
 If roDm(i).lanrid <> rDm(j).lanrid Then GoSub unter
 If roDm(i).Zusatzdaten <> rDm(j).Zusatzdaten Then GoSub unter
 If roDm(i).Druckdatum <> rDm(j).Druckdatum Then GoSub unter
 If roDm(i).uDat <> rDm(j).uDat Then GoSub unter
 If roDm(i).eGFR <> rDm(j).eGFR Then GoSub unter
 Exit Function
unter:
 DmZUnt = DmZUnt + 1
 Return
End Function ' DmZUnt

Public Function dmpreiheLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(Abk,'') Abk,COALESCE(DMPArt,0) DMPArt,COALESCE(Art,'') Art,COALESCE(KarteiDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) KarteiDatum" & _
",COALESCE(exportiert - INTERVAL 0 DAY,CONVERT('18991230',DATE)) exportiert,COALESCE(DokuDatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) DokuDatum,COALESCE(obvoll,0) obvoll,COALESCE(ok,0) ok" & _
",COALESCE(ausgedruckt,0) ausgedruckt,COALESCE(NachName,'') NachName,COALESCE(VorName,'') VorName,COALESCE(GebDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) GebDat" & _
",COALESCE(Pat_id,0) Pat_id,COALESCE(StByte,0) StByte,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(lanrid,0) lanrid" & _
",COALESCE(Zusatzdaten,'') Zusatzdaten,COALESCE(Druckdatum - INTERVAL 0 DAY,CONVERT('18991230',DATE)) Druckdatum,COALESCE(uDat - INTERVAL 0 DAY,CONVERT('18991230',DATE)) uDat,COALESCE(eGFR,0) eGFR" & _
" FROM `dmpreihe` WHERE Pat_ID=" & pid & " ORDER BY `Dokudatum`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roDm(0)
 Else ' rs.EOF Then
  ReDim roDm(1)
  Do While Not rs.EOF
   akt = UBound(roDm)
   roDm(akt).Abk = doUmwfSQL(rs!Abk, lies.obMySQL, False)
   roDm(akt).DMPArt = rs!DMPArt
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
   roDm(akt).Pat_ID = rs!Pat_ID
   roDm(akt).StByte = rs!StByte
   roDm(akt).aktZeit = rs!aktZeit
   roDm(akt).lanrid = rs!lanrid
   roDm(akt).Zusatzdaten = doUmwfSQL(rs!Zusatzdaten, lies.obMySQL, False)
   roDm(akt).Druckdatum = rs!Druckdatum
   roDm(akt).uDat = rs!uDat
   roDm(akt).eGFR = rs!eGFR
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDm) > 0 Then
  For ri = 1 To UBound(rDm)
   If rDm(ri).DokuDatum >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roDm)
    If roDm(roendpe).DokuDatum >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roDm(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roDm(roendpe + UBound(rDm) - rbeg)
   For ri = rbeg To UBound(rDm)
    Call roDmZuw(roendpe + ri - rbeg, ri)
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

Public Function rDmDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rDmdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rDm)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rDm(" & i & ").Abk:" & String$(33, "."), 33) & "'" & rDm(i).Abk & "'"
  Print #200, left$("rDm(" & i & ").DMPArt:" & String$(33, "."), 33) & rDm(i).DMPArt
  Print #200, left$("rDm(" & i & ").Art:" & String$(33, "."), 33) & "'" & rDm(i).art & "'"
  Print #200, left$("rDm(" & i & ").KarteiDatum:" & String$(33, "."), 33) & rDm(i).KarteiDatum
  Print #200, left$("rDm(" & i & ").exportiert:" & String$(33, "."), 33) & rDm(i).exportiert
  Print #200, left$("rDm(" & i & ").DokuDatum:" & String$(33, "."), 33) & rDm(i).DokuDatum
  Print #200, left$("rDm(" & i & ").obvoll:" & String$(33, "."), 33) & rDm(i).obvoll
  Print #200, left$("rDm(" & i & ").ok:" & String$(33, "."), 33) & rDm(i).Ok
  Print #200, left$("rDm(" & i & ").ausgedruckt:" & String$(33, "."), 33) & rDm(i).ausgedruckt
  Print #200, left$("rDm(" & i & ").NachName:" & String$(33, "."), 33) & "'" & rDm(i).Nachname & "'"
  Print #200, left$("rDm(" & i & ").VorName:" & String$(33, "."), 33) & "'" & rDm(i).Vorname & "'"
  Print #200, left$("rDm(" & i & ").GebDat:" & String$(33, "."), 33) & rDm(i).GebDat
  Print #200, left$("rDm(" & i & ").Pat_id:" & String$(33, "."), 33) & rDm(i).Pat_ID
  Print #200, left$("rDm(" & i & ").StByte:" & String$(33, "."), 33) & rDm(i).StByte
  Print #200, left$("rDm(" & i & ").AktZeit:" & String$(33, "."), 33) & rDm(i).aktZeit
  Print #200, left$("rDm(" & i & ").lanrid:" & String$(33, "."), 33) & rDm(i).lanrid
  Print #200, left$("rDm(" & i & ").Zusatzdaten:" & String$(33, "."), 33) & "'" & rDm(i).Zusatzdaten & "'"
  Print #200, left$("rDm(" & i & ").Druckdatum:" & String$(33, "."), 33) & rDm(i).Druckdatum
  Print #200, left$("rDm(" & i & ").uDat:" & String$(33, "."), 33) & rDm(i).uDat
  Print #200, left$("rDm(" & i & ").eGFR:" & String$(33, "."), 33) & rDm(i).eGFR
 Next i
 Close #200
 zeigan ffadat
End Function ' dmpreiheDump

Public Function dmpreiheSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "dmpreihe")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDm) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Abk,DMPArt,Art," & _
     "KarteiDatum,exportiert,DokuDatum,obvoll,ok,ausgedruckt,NachName,VorName,GebDat,Pat_id," & _
     "StByte,AktZeit,lanrid,Zusatzdaten,Druckdatum,uDat,eGFR)          VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rDm)
  rDm(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("('", rDm(i).Abk, "',", rDm(i).DMPArt, ",'", rDm(i).art, "',", DatFor_k(rDm(i).KarteiDatum), ",", DatFor_k(rDm(i).exportiert), ",", DatFor_k(rDm(i).DokuDatum), ",", CStr(-(rDm(i).obvoll <> 0)), ",", CStr(-( _
   rDm(i).Ok <> 0)), ",", CStr(-(rDm(i).ausgedruckt <> 0)), ",'", rDm(i).Nachname, "','", rDm(i).Vorname, "',", DatFor_k(rDm(i).GebDat), ",", rDm(i).Pat_ID, ",", rDm(i).StByte, ",", DatFor_k( _
   rDm(i).aktZeit), ",", rDm(i).lanrid, ",'", rDm(i).Zusatzdaten, "',", DatFor_k(rDm(i).Druckdatum), ",", DatFor_k(rDm(i).uDat), ",", rDm(i).eGFR, ")")
  If SammelInsert <> 0 And i < UBound(rDm) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDm) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rDm(i).Pat_ID = myEFrag("SELECT MAX(Pat_id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rDm)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rDm(" & i & "/" & UBound(rDm) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""dmpreiheSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDm), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDm.Abk: '" & rDm(k).Abk & "' -> '" & left$(rDm(k).Abk, maxL) & "'", True: rDm(k).Abk = left$(rDm(k).Abk, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDm.Art: '" & rDm(k).art & "' -> '" & left$(rDm(k).art, maxL) & "'", True: rDm(k).art = left$(rDm(k).art, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDm.NachName: '" & rDm(k).Nachname & "' -> '" & left$(rDm(k).Nachname, maxL) & "'", True: rDm(k).Nachname = left$(rDm(k).Nachname, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDm.VorName: '" & rDm(k).Vorname & "' -> '" & left$(rDm(k).Vorname, maxL) & "'", True: rDm(k).Vorname = left$(rDm(k).Vorname, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDm.Zusatzdaten: '" & rDm(k).Zusatzdaten & "' -> '" & left$(rDm(k).Zusatzdaten, maxL) & "'", True: rDm(k).Zusatzdaten = left$(rDm(k).Zusatzdaten, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in dmpreiheSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
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
 roDe(i).Titel = rDe(j).Titel
 roDe(i).toolTipText = rDe(j).toolTipText
 roDe(i).verankert = rDe(j).verankert
 roDe(i).absPos = rDe(j).absPos
 roDe(i).aktZeit = rDe(j).aktZeit
 roDe(i).StByte = rDe(j).StByte
End Function ' roDeZuw

Public Function DeZUnt%(i&, j&)
 If roDe(i).id <> rDe(j).id Then GoSub unter
 If roDe(i).IDS <> rDe(j).IDS Then GoSub unter
 If roDe(i).Pat_ID <> rDe(j).Pat_ID Then GoSub unter
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
 If roDe(i).aktZeit <> rDe(j).aktZeit Then GoSub unter
 If roDe(i).StByte <> rDe(j).StByte Then GoSub unter
 Exit Function
unter:
 DeZUnt = DeZUnt + 1
 Return
End Function ' DeZUnt

Public Function desktopLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
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
   roDe(akt).Titel = doUmwfSQL(rs!Titel, lies.obMySQL, False)
   roDe(akt).toolTipText = doUmwfSQL(rs!toolTipText, lies.obMySQL, False)
   roDe(akt).verankert = rs!verankert
   roDe(akt).absPos = rs!absPos
   roDe(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rDe) > 0 Then
  For ri = 1 To UBound(rDe)
   If rDe(ri).erstZP >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roDe)
    If roDe(roendpe).erstZP >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roDe(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roDe(roendpe + UBound(rDe) - rbeg)
   For ri = rbeg To UBound(rDe)
    Call roDeZuw(roendpe + ri - rbeg, ri)
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

Public Function rDeDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rDedump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rDe)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rDe(" & i & ").id:" & String$(33, "."), 33) & rDe(i).id
  Print #200, left$("rDe(" & i & ").IDS:" & String$(33, "."), 33) & "'" & rDe(i).IDS & "'"
  Print #200, left$("rDe(" & i & ").Pat_ID:" & String$(33, "."), 33) & rDe(i).Pat_ID
  Print #200, left$("rDe(" & i & ").erstZP:" & String$(33, "."), 33) & rDe(i).erstZP
  Print #200, left$("rDe(" & i & ").exoL:" & String$(33, "."), 33) & "'" & rDe(i).exoL & "'"
  Print #200, left$("rDe(" & i & ").hideT:" & String$(33, "."), 33) & rDe(i).hideT
  Print #200, left$("rDe(" & i & ").iconPath:" & String$(33, "."), 33) & "'" & rDe(i).iconPath & "'"
  Print #200, left$("rDe(" & i & ").noteBkColor:" & String$(33, "."), 33) & rDe(i).noteBkColor
  Print #200, left$("rDe(" & i & ").noteFgColor:" & String$(33, "."), 33) & rDe(i).noteFgColor
  Print #200, left$("rDe(" & i & ").positionBottom:" & String$(33, "."), 33) & rDe(i).positionBottom
  Print #200, left$("rDe(" & i & ").positionLeft:" & String$(33, "."), 33) & rDe(i).positionLeft
  Print #200, left$("rDe(" & i & ").positionRight:" & String$(33, "."), 33) & rDe(i).positionRight
  Print #200, left$("rDe(" & i & ").positionTop:" & String$(33, "."), 33) & rDe(i).positionTop
  Print #200, left$("rDe(" & i & ").showAsNote:" & String$(33, "."), 33) & rDe(i).showAsNote
  Print #200, left$("rDe(" & i & ").syncInfoList:" & String$(33, "."), 33) & "'" & rDe(i).syncInfoList & "'"
  Print #200, left$("rDe(" & i & ").titel:" & String$(33, "."), 33) & "'" & rDe(i).Titel & "'"
  Print #200, left$("rDe(" & i & ").toolTipText:" & String$(33, "."), 33) & "'" & rDe(i).toolTipText & "'"
  Print #200, left$("rDe(" & i & ").verankert:" & String$(33, "."), 33) & rDe(i).verankert
  Print #200, left$("rDe(" & i & ").absPos:" & String$(33, "."), 33) & rDe(i).absPos
  Print #200, left$("rDe(" & i & ").AktZeit:" & String$(33, "."), 33) & rDe(i).aktZeit
  Print #200, left$("rDe(" & i & ").StByte:" & String$(33, "."), 33) & rDe(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' desktopDump

Public Function desktopSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "desktop")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rDe) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (IDS,Pat_ID,erstZP," & _
     "exoL,hideT,iconPath,noteBkColor,noteFgColor,positionBottom,positionLeft,positionRight,positionTop,showAsNote," & _
     "syncInfoList,titel,toolTipText,verankert,absPos,AktZeit,StByte)  VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rDe)
  rDe(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("('", rDe(i).IDS, "',", rDe(i).Pat_ID, ",", DatFor_k(rDe(i).erstZP), ",'", rDe(i).exoL, "',", rDe(i).hideT, ",'", rDe(i).iconPath, "',", rDe(i).noteBkColor, ",", rDe(i).noteFgColor, ",", _
   rDe(i).positionBottom, ",", rDe(i).positionLeft, ",", rDe(i).positionRight, ",", rDe(i).positionTop, ",", rDe(i).showAsNote, ",'", rDe(i).syncInfoList, "','", rDe(i).Titel, "','", _
   rDe(i).toolTipText, "',", rDe(i).verankert, ",", rDe(i).absPos, ",", DatFor_k(rDe(i).aktZeit), ",", rDe(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rDe) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rDe) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rDe(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rDe)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rDe(" & i & "/" & UBound(rDe) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""desktopSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rDe), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rDe.IDS: '" & rDe(k).IDS & "' -> '" & left$(rDe(k).IDS, maxL) & "'", True: rDe(k).IDS = left$(rDe(k).IDS, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rDe.exoL: '" & rDe(k).exoL & "' -> '" & left$(rDe(k).exoL, maxL) & "'", True: rDe(k).exoL = left$(rDe(k).exoL, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rDe.iconPath: '" & rDe(k).iconPath & "' -> '" & left$(rDe(k).iconPath, maxL) & "'", True: rDe(k).iconPath = left$(rDe(k).iconPath, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rDe.syncInfoList: '" & rDe(k).syncInfoList & "' -> '" & left$(rDe(k).syncInfoList, maxL) & "'", True: rDe(k).syncInfoList = left$(rDe(k).syncInfoList, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rDe.titel: '" & rDe(k).Titel & "' -> '" & left$(rDe(k).Titel, maxL) & "'", True: rDe(k).Titel = left$(rDe(k).Titel, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rDe.toolTipText: '" & rDe(k).toolTipText & "' -> '" & left$(rDe(k).toolTipText, maxL) & "'", True: rDe(k).toolTipText = left$(rDe(k).toolTipText, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in desktopSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' desktopSpeichern

Public Function roUsZuw(i&, j&)
 roUs(i).FID = rUs(j).FID
 roUs(i).Pat_ID = rUs(j).Pat_ID
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
 roUs(i).aktZeit = rUs(j).aktZeit
 roUs(i).QS = rUs(j).QS
 roUs(i).QT = rUs(j).QT
 roUs(i).StByte = rUs(j).StByte
 roUs(i).id = rUs(j).id
End Function ' roUsZuw

Public Function UsZUnt%(i&, j&)
 If roUs(i).FID <> rUs(j).FID Then GoSub unter
 If roUs(i).Pat_ID <> rUs(j).Pat_ID Then GoSub unter
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
 If roUs(i).aktZeit <> rUs(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
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
   roUs(akt).Pat_ID = rs!Pat_ID
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
   roUs(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rUs) > 0 Then
  For ri = 1 To UBound(rUs)
   If rUs(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roUs)
    If roUs(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roUs(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roUs(roendpe + UBound(rUs) - rbeg)
   For ri = rbeg To UBound(rUs)
    Call roUsZuw(roendpe + ri - rbeg, ri)
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

Public Function rUsDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rUsdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rUs)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rUs(" & i & ").FID:" & String$(33, "."), 33) & rUs(i).FID
  Print #200, left$("rUs(" & i & ").Pat_ID:" & String$(33, "."), 33) & rUs(i).Pat_ID
  Print #200, left$("rUs(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rUs(i).Zeitpunkt
  Print #200, left$("rUs(" & i & ").Art:" & String$(33, "."), 33) & "'" & rUs(i).art & "'"
  Print #200, left$("rUs(" & i & ").Spritzst:" & String$(33, "."), 33) & "'" & rUs(i).Spritzst & "'"
  Print #200, left$("rUs(" & i & ").Fußbef_re:" & String$(33, "."), 33) & "'" & rUs(i).Fußbef_re & "'"
  Print #200, left$("rUs(" & i & ").Fußbef_li:" & String$(33, "."), 33) & "'" & rUs(i).Fußbef_li & "'"
  Print #200, left$("rUs(" & i & ").Hyperk_re:" & String$(33, "."), 33) & "'" & rUs(i).Hyperk_re & "'"
  Print #200, left$("rUs(" & i & ").Hyperk_li:" & String$(33, "."), 33) & "'" & rUs(i).Hyperk_li & "'"
  Print #200, left$("rUs(" & i & ").Ulcera_re:" & String$(33, "."), 33) & "'" & rUs(i).Ulcera_re & "'"
  Print #200, left$("rUs(" & i & ").Ulcera_li:" & String$(33, "."), 33) & "'" & rUs(i).Ulcera_li & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Zh_re:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Zh_re & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Zh_li:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Zh_li & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Zb_re:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Zb_re & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Zb_li:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Zb_li & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Knie_re:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Knie_re & "'"
  Print #200, left$("rUs(" & i & ").Kraft_Knie_li:" & String$(33, "."), 33) & "'" & rUs(i).Kraft_Knie_li & "'"
  Print #200, left$("rUs(" & i & ").ASR_re:" & String$(33, "."), 33) & "'" & rUs(i).ASR_re & "'"
  Print #200, left$("rUs(" & i & ").ASR_li:" & String$(33, "."), 33) & "'" & rUs(i).ASR_li & "'"
  Print #200, left$("rUs(" & i & ").PSR_re:" & String$(33, "."), 33) & "'" & rUs(i).PSR_re & "'"
  Print #200, left$("rUs(" & i & ").PSR_li:" & String$(33, "."), 33) & "'" & rUs(i).PSR_li & "'"
  Print #200, left$("rUs(" & i & ").Oberfl_re:" & String$(33, "."), 33) & "'" & rUs(i).Oberfl_re & "'"
  Print #200, left$("rUs(" & i & ").Oberfl_li:" & String$(33, "."), 33) & "'" & rUs(i).Oberfl_li & "'"
  Print #200, left$("rUs(" & i & ").MF_re:" & String$(33, "."), 33) & "'" & rUs(i).MF_re & "'"
  Print #200, left$("rUs(" & i & ").MF_li:" & String$(33, "."), 33) & "'" & rUs(i).MF_li & "'"
  Print #200, left$("rUs(" & i & ").KW_re:" & String$(33, "."), 33) & "'" & rUs(i).KW_re & "'"
  Print #200, left$("rUs(" & i & ").KW_li:" & String$(33, "."), 33) & "'" & rUs(i).KW_li & "'"
  Print #200, left$("rUs(" & i & ").Vibr_IK_re:" & String$(33, "."), 33) & "'" & rUs(i).Vibr_IK_re & "'"
  Print #200, left$("rUs(" & i & ").Vibr_IK_li:" & String$(33, "."), 33) & "'" & rUs(i).Vibr_IK_li & "'"
  Print #200, left$("rUs(" & i & ").Vibr_GZ_re:" & String$(33, "."), 33) & "'" & rUs(i).Vibr_GZ_re & "'"
  Print #200, left$("rUs(" & i & ").Vibr_GZ_li:" & String$(33, "."), 33) & "'" & rUs(i).Vibr_GZ_li & "'"
  Print #200, left$("rUs(" & i & ").PulsL_re:" & String$(33, "."), 33) & "'" & rUs(i).PulsL_re & "'"
  Print #200, left$("rUs(" & i & ").PulsL_li:" & String$(33, "."), 33) & "'" & rUs(i).PulsL_li & "'"
  Print #200, left$("rUs(" & i & ").PulsKK_re:" & String$(33, "."), 33) & "'" & rUs(i).PulsKK_re & "'"
  Print #200, left$("rUs(" & i & ").PulsKK_li:" & String$(33, "."), 33) & "'" & rUs(i).PulsKK_li & "'"
  Print #200, left$("rUs(" & i & ").PulsAtp_re:" & String$(33, "."), 33) & "'" & rUs(i).PulsAtp_re & "'"
  Print #200, left$("rUs(" & i & ").PulsAtp_li:" & String$(33, "."), 33) & "'" & rUs(i).PulsAtp_li & "'"
  Print #200, left$("rUs(" & i & ").PulsAdp_re:" & String$(33, "."), 33) & "'" & rUs(i).PulsAdp_re & "'"
  Print #200, left$("rUs(" & i & ").PulsAdp_li:" & String$(33, "."), 33) & "'" & rUs(i).PulsAdp_li & "'"
  Print #200, left$("rUs(" & i & ").Mitarbeiter:" & String$(33, "."), 33) & "'" & rUs(i).Mitarbeiter & "'"
  Print #200, left$("rUs(" & i & ").absPos:" & String$(33, "."), 33) & rUs(i).absPos
  Print #200, left$("rUs(" & i & ").AktZeit:" & String$(33, "."), 33) & rUs(i).aktZeit
  Print #200, left$("rUs(" & i & ").QS:" & String$(33, "."), 33) & "'" & rUs(i).QS & "'"
  Print #200, left$("rUs(" & i & ").QT:" & String$(33, "."), 33) & "'" & rUs(i).QT & "'"
  Print #200, left$("rUs(" & i & ").StByte:" & String$(33, "."), 33) & rUs(i).StByte
  Print #200, left$("rUs(" & i & ").id:" & String$(33, "."), 33) & rUs(i).id
 Next i
 Close #200
 zeigan ffadat
End Function ' usdmDump

Public Function usdmSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "usdm")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rUs) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Spritzst,Fußbef_re,Fußbef_li,Hyperk_re,Hyperk_li,Ulcera_re,Ulcera_li,Kraft_Zh_re,Kraft_Zh_li," & _
     "Kraft_Zb_re,Kraft_Zb_li,Kraft_Knie_re,Kraft_Knie_li,ASR_re,ASR_li,PSR_re,PSR_li,Oberfl_re,Oberfl_li," & _
     "MF_re,MF_li,KW_re,KW_li,Vibr_IK_re,Vibr_IK_li,Vibr_GZ_re,Vibr_GZ_li,PulsL_re,PulsL_li," & _
     "PulsKK_re,PulsKK_li,PulsAtp_re,PulsAtp_li,PulsAdp_re,PulsAdp_li,Mitarbeiter,absPos,AktZeit,QS," & _
     "QT,StByte)             VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rUs)
  rUs(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rUs(i).FID, ",", rUs(i).Pat_ID, ",", DatFor_k(rUs(i).Zeitpunkt), ",'", rUs(i).art, "','", rUs(i).Spritzst, "','", rUs(i).Fußbef_re, "','", rUs(i).Fußbef_li, "','", rUs(i).Hyperk_re, "','", _
   rUs(i).Hyperk_li, "','", rUs(i).Ulcera_re, "','", rUs(i).Ulcera_li, "','", rUs(i).Kraft_Zh_re, "','", rUs(i).Kraft_Zh_li, "','", rUs(i).Kraft_Zb_re, "','", rUs(i).Kraft_Zb_li, "','", _
   rUs(i).Kraft_Knie_re, "','", rUs(i).Kraft_Knie_li, "','", rUs(i).ASR_re, "','", rUs(i).ASR_li, "','", rUs(i).PSR_re, "','", rUs(i).PSR_li, "','", rUs(i).Oberfl_re, "','", rUs(i).Oberfl_li, "','", _
   rUs(i).MF_re, "','", rUs(i).MF_li, "','", rUs(i).KW_re, "','", rUs(i).KW_li, "','", rUs(i).Vibr_IK_re, "','", rUs(i).Vibr_IK_li, "','", rUs(i).Vibr_GZ_re, "','", _
   rUs(i).Vibr_GZ_li, "','", rUs(i).PulsL_re, "','", rUs(i).PulsL_li, "','", rUs(i).PulsKK_re, "','", rUs(i).PulsKK_li, "','", rUs(i).PulsAtp_re, "','", rUs(i).PulsAtp_li, "','", rUs(i).PulsAdp_re, "','", _
   rUs(i).PulsAdp_li, "','", rUs(i).Mitarbeiter, "',", rUs(i).absPos, ",", DatFor_k(rUs(i).aktZeit), ",'", rUs(i).QS, "','", rUs(i).QT, "',", rUs(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rUs) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUs) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rUs(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rUs)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rUs(" & i & "/" & UBound(rUs) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""usdmSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUs.Art: '" & rUs(k).art & "' -> '" & left$(rUs(k).art, maxL) & "'", True: rUs(k).art = left$(rUs(k).art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUs.Spritzst: '" & rUs(k).Spritzst & "' -> '" & left$(rUs(k).Spritzst, maxL) & "'", True: rUs(k).Spritzst = left$(rUs(k).Spritzst, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_re: '" & rUs(k).Fußbef_re & "' -> '" & left$(rUs(k).Fußbef_re, maxL) & "'", True: rUs(k).Fußbef_re = left$(rUs(k).Fußbef_re, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUs.Fußbef_li: '" & rUs(k).Fußbef_li & "' -> '" & left$(rUs(k).Fußbef_li, maxL) & "'", True: rUs(k).Fußbef_li = left$(rUs(k).Fußbef_li, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_re: '" & rUs(k).Hyperk_re & "' -> '" & left$(rUs(k).Hyperk_re, maxL) & "'", True: rUs(k).Hyperk_re = left$(rUs(k).Hyperk_re, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUs.Hyperk_li: '" & rUs(k).Hyperk_li & "' -> '" & left$(rUs(k).Hyperk_li, maxL) & "'", True: rUs(k).Hyperk_li = left$(rUs(k).Hyperk_li, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_re: '" & rUs(k).Ulcera_re & "' -> '" & left$(rUs(k).Ulcera_re, maxL) & "'", True: rUs(k).Ulcera_re = left$(rUs(k).Ulcera_re, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUs.Ulcera_li: '" & rUs(k).Ulcera_li & "' -> '" & left$(rUs(k).Ulcera_li, maxL) & "'", True: rUs(k).Ulcera_li = left$(rUs(k).Ulcera_li, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_re: '" & rUs(k).Kraft_Zh_re & "' -> '" & left$(rUs(k).Kraft_Zh_re, maxL) & "'", True: rUs(k).Kraft_Zh_re = left$(rUs(k).Kraft_Zh_re, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zh_li: '" & rUs(k).Kraft_Zh_li & "' -> '" & left$(rUs(k).Kraft_Zh_li, maxL) & "'", True: rUs(k).Kraft_Zh_li = left$(rUs(k).Kraft_Zh_li, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_re: '" & rUs(k).Kraft_Zb_re & "' -> '" & left$(rUs(k).Kraft_Zb_re, maxL) & "'", True: rUs(k).Kraft_Zb_re = left$(rUs(k).Kraft_Zb_re, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Zb_li: '" & rUs(k).Kraft_Zb_li & "' -> '" & left$(rUs(k).Kraft_Zb_li, maxL) & "'", True: rUs(k).Kraft_Zb_li = left$(rUs(k).Kraft_Zb_li, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_re: '" & rUs(k).Kraft_Knie_re & "' -> '" & left$(rUs(k).Kraft_Knie_re, maxL) & "'", True: rUs(k).Kraft_Knie_re = left$(rUs(k).Kraft_Knie_re, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rUs.Kraft_Knie_li: '" & rUs(k).Kraft_Knie_li & "' -> '" & left$(rUs(k).Kraft_Knie_li, maxL) & "'", True: rUs(k).Kraft_Knie_li = left$(rUs(k).Kraft_Knie_li, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_re: '" & rUs(k).ASR_re & "' -> '" & left$(rUs(k).ASR_re, maxL) & "'", True: rUs(k).ASR_re = left$(rUs(k).ASR_re, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rUs.ASR_li: '" & rUs(k).ASR_li & "' -> '" & left$(rUs(k).ASR_li, maxL) & "'", True: rUs(k).ASR_li = left$(rUs(k).ASR_li, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_re: '" & rUs(k).PSR_re & "' -> '" & left$(rUs(k).PSR_re, maxL) & "'", True: rUs(k).PSR_re = left$(rUs(k).PSR_re, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rUs.PSR_li: '" & rUs(k).PSR_li & "' -> '" & left$(rUs(k).PSR_li, maxL) & "'", True: rUs(k).PSR_li = left$(rUs(k).PSR_li, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_re: '" & rUs(k).Oberfl_re & "' -> '" & left$(rUs(k).Oberfl_re, maxL) & "'", True: rUs(k).Oberfl_re = left$(rUs(k).Oberfl_re, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rUs.Oberfl_li: '" & rUs(k).Oberfl_li & "' -> '" & left$(rUs(k).Oberfl_li, maxL) & "'", True: rUs(k).Oberfl_li = left$(rUs(k).Oberfl_li, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_re: '" & rUs(k).MF_re & "' -> '" & left$(rUs(k).MF_re, maxL) & "'", True: rUs(k).MF_re = left$(rUs(k).MF_re, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rUs.MF_li: '" & rUs(k).MF_li & "' -> '" & left$(rUs(k).MF_li, maxL) & "'", True: rUs(k).MF_li = left$(rUs(k).MF_li, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_re: '" & rUs(k).KW_re & "' -> '" & left$(rUs(k).KW_re, maxL) & "'", True: rUs(k).KW_re = left$(rUs(k).KW_re, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rUs.KW_li: '" & rUs(k).KW_li & "' -> '" & left$(rUs(k).KW_li, maxL) & "'", True: rUs(k).KW_li = left$(rUs(k).KW_li, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_re: '" & rUs(k).Vibr_IK_re & "' -> '" & left$(rUs(k).Vibr_IK_re, maxL) & "'", True: rUs(k).Vibr_IK_re = left$(rUs(k).Vibr_IK_re, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_IK_li: '" & rUs(k).Vibr_IK_li & "' -> '" & left$(rUs(k).Vibr_IK_li, maxL) & "'", True: rUs(k).Vibr_IK_li = left$(rUs(k).Vibr_IK_li, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_re: '" & rUs(k).Vibr_GZ_re & "' -> '" & left$(rUs(k).Vibr_GZ_re, maxL) & "'", True: rUs(k).Vibr_GZ_re = left$(rUs(k).Vibr_GZ_re, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rUs.Vibr_GZ_li: '" & rUs(k).Vibr_GZ_li & "' -> '" & left$(rUs(k).Vibr_GZ_li, maxL) & "'", True: rUs(k).Vibr_GZ_li = left$(rUs(k).Vibr_GZ_li, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_re: '" & rUs(k).PulsL_re & "' -> '" & left$(rUs(k).PulsL_re, maxL) & "'", True: rUs(k).PulsL_re = left$(rUs(k).PulsL_re, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsL_li: '" & rUs(k).PulsL_li & "' -> '" & left$(rUs(k).PulsL_li, maxL) & "'", True: rUs(k).PulsL_li = left$(rUs(k).PulsL_li, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_re: '" & rUs(k).PulsKK_re & "' -> '" & left$(rUs(k).PulsKK_re, maxL) & "'", True: rUs(k).PulsKK_re = left$(rUs(k).PulsKK_re, maxL)
       Case 31: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsKK_li: '" & rUs(k).PulsKK_li & "' -> '" & left$(rUs(k).PulsKK_li, maxL) & "'", True: rUs(k).PulsKK_li = left$(rUs(k).PulsKK_li, maxL)
       Case 32: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_re: '" & rUs(k).PulsAtp_re & "' -> '" & left$(rUs(k).PulsAtp_re, maxL) & "'", True: rUs(k).PulsAtp_re = left$(rUs(k).PulsAtp_re, maxL)
       Case 33: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAtp_li: '" & rUs(k).PulsAtp_li & "' -> '" & left$(rUs(k).PulsAtp_li, maxL) & "'", True: rUs(k).PulsAtp_li = left$(rUs(k).PulsAtp_li, maxL)
       Case 34: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_re: '" & rUs(k).PulsAdp_re & "' -> '" & left$(rUs(k).PulsAdp_re, maxL) & "'", True: rUs(k).PulsAdp_re = left$(rUs(k).PulsAdp_re, maxL)
       Case 35: Lese.Ausgeb "   Verkürze Inhalt von rUs.PulsAdp_li: '" & rUs(k).PulsAdp_li & "' -> '" & left$(rUs(k).PulsAdp_li, maxL) & "'", True: rUs(k).PulsAdp_li = left$(rUs(k).PulsAdp_li, maxL)
       Case 36: Lese.Ausgeb "   Verkürze Inhalt von rUs.Mitarbeiter: '" & rUs(k).Mitarbeiter & "' -> '" & left$(rUs(k).Mitarbeiter, maxL) & "'", True: rUs(k).Mitarbeiter = left$(rUs(k).Mitarbeiter, maxL)
       Case 37: Lese.Ausgeb "   Verkürze Inhalt von rUs.QS: '" & rUs(k).QS & "' -> '" & left$(rUs(k).QS, maxL) & "'", True: rUs(k).QS = left$(rUs(k).QS, maxL)
       Case 38: Lese.Ausgeb "   Verkürze Inhalt von rUs.QT: '" & rUs(k).QT & "' -> '" & left$(rUs(k).QT, maxL) & "'", True: rUs(k).QT = left$(rUs(k).QT, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in usdmSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' usdmSpeichern

Public Function roFuZuw(i&, j&)
 roFu(i).FID = rFu(j).FID
 roFu(i).Pat_ID = rFu(j).Pat_ID
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
 roFu(i).aktZeit = rFu(j).aktZeit
 roFu(i).QS = rFu(j).QS
 roFu(i).QT = rFu(j).QT
 roFu(i).StByte = rFu(j).StByte
 roFu(i).id = rFu(j).id
End Function ' roFuZuw

Public Function FuZUnt%(i&, j&)
 If roFu(i).FID <> rFu(j).FID Then GoSub unter
 If roFu(i).Pat_ID <> rFu(j).Pat_ID Then GoSub unter
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
 If roFu(i).aktZeit <> rFu(j).aktZeit Then GoSub unter
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
 pid = rNa(0).Pat_ID
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
   roFu(akt).Pat_ID = rs!Pat_ID
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
   roFu(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rFu) > 0 Then
  For ri = 1 To UBound(rFu)
   If rFu(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roFu)
    If roFu(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roFu(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roFu(roendpe + UBound(rFu) - rbeg)
   For ri = rbeg To UBound(rFu)
    Call roFuZuw(roendpe + ri - rbeg, ri)
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

Public Function rFuDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rFudump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rFu)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rFu(" & i & ").FID:" & String$(33, "."), 33) & rFu(i).FID
  Print #200, left$("rFu(" & i & ").Pat_ID:" & String$(33, "."), 33) & rFu(i).Pat_ID
  Print #200, left$("rFu(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rFu(i).Zeitpunkt
  Print #200, left$("rFu(" & i & ").Art:" & String$(33, "."), 33) & "'" & rFu(i).art & "'"
  Print #200, left$("rFu(" & i & ").Fußdeform:" & String$(33, "."), 33) & "'" & rFu(i).Fußdeform & "'"
  Print #200, left$("rFu(" & i & ").Hyper_mEin:" & String$(33, "."), 33) & "'" & rFu(i).Hyper_mEin & "'"
  Print #200, left$("rFu(" & i & ").Weiteres:" & String$(33, "."), 33) & "'" & rFu(i).Weiteres & "'"
  Print #200, left$("rFu(" & i & ").Zn_Ulcus:" & String$(33, "."), 33) & "'" & rFu(i).Zn_Ulcus & "'"
  Print #200, left$("rFu(" & i & ").Zn_Amput:" & String$(33, "."), 33) & "'" & rFu(i).Zn_Amput & "'"
  Print #200, left$("rFu(" & i & ").Fuß_ang:" & String$(33, "."), 33) & "'" & rFu(i).Fuß_ang & "'"
  Print #200, left$("rFu(" & i & ").Ulcera:" & String$(33, "."), 33) & "'" & rFu(i).Ulcera & "'"
  Print #200, left$("rFu(" & i & ").Wundinfektion:" & String$(33, "."), 33) & "'" & rFu(i).Wundinfektion & "'"
  Print #200, left$("rFu(" & i & ").nae_US:" & String$(33, "."), 33) & "'" & rFu(i).nae_US & "'"
  Print #200, left$("rFu(" & i & ").Mitarbeiter:" & String$(33, "."), 33) & "'" & rFu(i).Mitarbeiter & "'"
  Print #200, left$("rFu(" & i & ").absPos:" & String$(33, "."), 33) & rFu(i).absPos
  Print #200, left$("rFu(" & i & ").AktZeit:" & String$(33, "."), 33) & rFu(i).aktZeit
  Print #200, left$("rFu(" & i & ").QS:" & String$(33, "."), 33) & "'" & rFu(i).QS & "'"
  Print #200, left$("rFu(" & i & ").QT:" & String$(33, "."), 33) & "'" & rFu(i).QT & "'"
  Print #200, left$("rFu(" & i & ").StByte:" & String$(33, "."), 33) & rFu(i).StByte
  Print #200, left$("rFu(" & i & ").id:" & String$(33, "."), 33) & rFu(i).id
 Next i
 Close #200
 zeigan ffadat
End Function ' fussDump

Public Function fussSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "fuss")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rFu) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Art,Fußdeform,Hyper_mEin,Weiteres,Zn_Ulcus,Zn_Amput,Fuß_ang,Ulcera,Wundinfektion,nae_US," & _
     "Mitarbeiter,absPos,AktZeit,QS,QT,StByte)           VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rFu)
  rFu(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rFu(i).FID, ",", rFu(i).Pat_ID, ",", DatFor_k(rFu(i).Zeitpunkt), ",'", rFu(i).art, "','", rFu(i).Fußdeform, "','", rFu(i).Hyper_mEin, "','", rFu(i).Weiteres, "','", rFu(i).Zn_Ulcus, "','", _
   rFu(i).Zn_Amput, "','", rFu(i).Fuß_ang, "','", rFu(i).Ulcera, "','", rFu(i).Wundinfektion, "','", rFu(i).nae_US, "','", rFu(i).Mitarbeiter, "',", rFu(i).absPos, ",", DatFor_k(rFu(i).aktZeit), ",'", _
   rFu(i).QS, "','", rFu(i).QT, "',", rFu(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rFu) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rFu) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rFu(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rFu)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rFu(" & i & "/" & UBound(rFu) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""fussSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rFu), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rFu.Art: '" & rFu(k).art & "' -> '" & left$(rFu(k).art, maxL) & "'", True: rFu(k).art = left$(rFu(k).art, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fußdeform: '" & rFu(k).Fußdeform & "' -> '" & left$(rFu(k).Fußdeform, maxL) & "'", True: rFu(k).Fußdeform = left$(rFu(k).Fußdeform, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rFu.Hyper_mEin: '" & rFu(k).Hyper_mEin & "' -> '" & left$(rFu(k).Hyper_mEin, maxL) & "'", True: rFu(k).Hyper_mEin = left$(rFu(k).Hyper_mEin, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rFu.Weiteres: '" & rFu(k).Weiteres & "' -> '" & left$(rFu(k).Weiteres, maxL) & "'", True: rFu(k).Weiteres = left$(rFu(k).Weiteres, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Ulcus: '" & rFu(k).Zn_Ulcus & "' -> '" & left$(rFu(k).Zn_Ulcus, maxL) & "'", True: rFu(k).Zn_Ulcus = left$(rFu(k).Zn_Ulcus, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rFu.Zn_Amput: '" & rFu(k).Zn_Amput & "' -> '" & left$(rFu(k).Zn_Amput, maxL) & "'", True: rFu(k).Zn_Amput = left$(rFu(k).Zn_Amput, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rFu.Fuß_ang: '" & rFu(k).Fuß_ang & "' -> '" & left$(rFu(k).Fuß_ang, maxL) & "'", True: rFu(k).Fuß_ang = left$(rFu(k).Fuß_ang, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rFu.Ulcera: '" & rFu(k).Ulcera & "' -> '" & left$(rFu(k).Ulcera, maxL) & "'", True: rFu(k).Ulcera = left$(rFu(k).Ulcera, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rFu.Wundinfektion: '" & rFu(k).Wundinfektion & "' -> '" & left$(rFu(k).Wundinfektion, maxL) & "'", True: rFu(k).Wundinfektion = left$(rFu(k).Wundinfektion, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rFu.nae_US: '" & rFu(k).nae_US & "' -> '" & left$(rFu(k).nae_US, maxL) & "'", True: rFu(k).nae_US = left$(rFu(k).nae_US, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rFu.Mitarbeiter: '" & rFu(k).Mitarbeiter & "' -> '" & left$(rFu(k).Mitarbeiter, maxL) & "'", True: rFu(k).Mitarbeiter = left$(rFu(k).Mitarbeiter, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rFu.QS: '" & rFu(k).QS & "' -> '" & left$(rFu(k).QS, maxL) & "'", True: rFu(k).QS = left$(rFu(k).QS, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rFu.QT: '" & rFu(k).QT & "' -> '" & left$(rFu(k).QT, maxL) & "'", True: rFu(k).QT = left$(rFu(k).QT, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in fussSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' fussSpeichern

Public Function roUlZuw(i&, j&)
 roUl(i).FID = rUl(j).FID
 roUl(i).Pat_ID = rUl(j).Pat_ID
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
 roUl(i).aktZeit = rUl(j).aktZeit
 roUl(i).StByte = rUl(j).StByte
End Function ' roUlZuw

Public Function UlZUnt%(i&, j&)
 If roUl(i).FID <> rUl(j).FID Then GoSub unter
 If roUl(i).Pat_ID <> rUl(j).Pat_ID Then GoSub unter
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
 If roUl(i).aktZeit <> rUl(j).aktZeit Then GoSub unter
 If roUl(i).StByte <> rUl(j).StByte Then GoSub unter
 Exit Function
unter:
 UlZUnt = UlZUnt + 1
 Return
End Function ' UlZUnt

Public Function ulcusLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
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
   roUl(akt).Pat_ID = rs!Pat_ID
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
   roUl(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rUl) > 0 Then
  For ri = 1 To UBound(rUl)
   If rUl(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roUl)
    If roUl(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roUl(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roUl(roendpe + UBound(rUl) - rbeg)
   For ri = rbeg To UBound(rUl)
    Call roUlZuw(roendpe + ri - rbeg, ri)
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

Public Function rUlDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rUldump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rUl)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rUl(" & i & ").FID:" & String$(33, "."), 33) & rUl(i).FID
  Print #200, left$("rUl(" & i & ").Pat_ID:" & String$(33, "."), 33) & rUl(i).Pat_ID
  Print #200, left$("rUl(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rUl(i).Zeitpunkt
  Print #200, left$("rUl(" & i & ").Lokalisation:" & String$(33, "."), 33) & "'" & rUl(i).Lokalisation & "'"
  Print #200, left$("rUl(" & i & ").Seite:" & String$(33, "."), 33) & "'" & rUl(i).Seite & "'"
  Print #200, left$("rUl(" & i & ").Größe:" & String$(33, "."), 33) & "'" & rUl(i).Größe & "'"
  Print #200, left$("rUl(" & i & ").Beläge:" & String$(33, "."), 33) & "'" & rUl(i).Beläge & "'"
  Print #200, left$("rUl(" & i & ").Exsudat:" & String$(33, "."), 33) & "'" & rUl(i).Exsudat & "'"
  Print #200, left$("rUl(" & i & ").Geruch:" & String$(33, "."), 33) & "'" & rUl(i).Geruch & "'"
  Print #200, left$("rUl(" & i & ").Wundrand:" & String$(33, "."), 33) & "'" & rUl(i).Wundrand & "'"
  Print #200, left$("rUl(" & i & ").Wundumgebung:" & String$(33, "."), 33) & "'" & rUl(i).Wundumgebung & "'"
  Print #200, left$("rUl(" & i & ").Temperatur:" & String$(33, "."), 33) & "'" & rUl(i).Temperatur & "'"
  Print #200, left$("rUl(" & i & ").Fotodoku:" & String$(33, "."), 33) & "'" & rUl(i).Fotodoku & "'"
  Print #200, left$("rUl(" & i & ").Wundversorgung:" & String$(33, "."), 33) & "'" & rUl(i).Wundversorgung & "'"
  Print #200, left$("rUl(" & i & ").Mitarbeiter:" & String$(33, "."), 33) & "'" & rUl(i).Mitarbeiter & "'"
  Print #200, left$("rUl(" & i & ").absPos:" & String$(33, "."), 33) & rUl(i).absPos
  Print #200, left$("rUl(" & i & ").AktZeit:" & String$(33, "."), 33) & rUl(i).aktZeit
  Print #200, left$("rUl(" & i & ").StByte:" & String$(33, "."), 33) & rUl(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' ulcusDump

Public Function ulcusSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "ulcus")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rUl) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Lokalisation,Seite,Größe,Beläge,Exsudat,Geruch,Wundrand,Wundumgebung,Temperatur,Fotodoku," & _
     "Wundversorgung,Mitarbeiter,absPos,AktZeit,StByte)  VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rUl)
  rUl(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rUl(i).FID, ",", rUl(i).Pat_ID, ",", DatFor_k(rUl(i).Zeitpunkt), ",'", rUl(i).Lokalisation, "','", rUl(i).Seite, "','", rUl(i).Größe, "','", rUl(i).Beläge, "','", rUl(i).Exsudat, "','", _
   rUl(i).Geruch, "','", rUl(i).Wundrand, "','", rUl(i).Wundumgebung, "','", rUl(i).Temperatur, "','", rUl(i).Fotodoku, "','", rUl(i).Wundversorgung, "','", rUl(i).Mitarbeiter, "',", _
   rUl(i).absPos, ",", DatFor_k(rUl(i).aktZeit), ",", rUl(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rUl) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rUl) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rUl)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rUl(" & i & "/" & UBound(rUl) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""ulcusSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rUl), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rUl.Lokalisation: '" & rUl(k).Lokalisation & "' -> '" & left$(rUl(k).Lokalisation, maxL) & "'", True: rUl(k).Lokalisation = left$(rUl(k).Lokalisation, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rUl.Seite: '" & rUl(k).Seite & "' -> '" & left$(rUl(k).Seite, maxL) & "'", True: rUl(k).Seite = left$(rUl(k).Seite, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rUl.Größe: '" & rUl(k).Größe & "' -> '" & left$(rUl(k).Größe, maxL) & "'", True: rUl(k).Größe = left$(rUl(k).Größe, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rUl.Beläge: '" & rUl(k).Beläge & "' -> '" & left$(rUl(k).Beläge, maxL) & "'", True: rUl(k).Beläge = left$(rUl(k).Beläge, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rUl.Exsudat: '" & rUl(k).Exsudat & "' -> '" & left$(rUl(k).Exsudat, maxL) & "'", True: rUl(k).Exsudat = left$(rUl(k).Exsudat, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rUl.Geruch: '" & rUl(k).Geruch & "' -> '" & left$(rUl(k).Geruch, maxL) & "'", True: rUl(k).Geruch = left$(rUl(k).Geruch, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundrand: '" & rUl(k).Wundrand & "' -> '" & left$(rUl(k).Wundrand, maxL) & "'", True: rUl(k).Wundrand = left$(rUl(k).Wundrand, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundumgebung: '" & rUl(k).Wundumgebung & "' -> '" & left$(rUl(k).Wundumgebung, maxL) & "'", True: rUl(k).Wundumgebung = left$(rUl(k).Wundumgebung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rUl.Temperatur: '" & rUl(k).Temperatur & "' -> '" & left$(rUl(k).Temperatur, maxL) & "'", True: rUl(k).Temperatur = left$(rUl(k).Temperatur, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rUl.Fotodoku: '" & rUl(k).Fotodoku & "' -> '" & left$(rUl(k).Fotodoku, maxL) & "'", True: rUl(k).Fotodoku = left$(rUl(k).Fotodoku, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rUl.Wundversorgung: '" & rUl(k).Wundversorgung & "' -> '" & left$(rUl(k).Wundversorgung, maxL) & "'", True: rUl(k).Wundversorgung = left$(rUl(k).Wundversorgung, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rUl.Mitarbeiter: '" & rUl(k).Mitarbeiter & "' -> '" & left$(rUl(k).Mitarbeiter, maxL) & "'", True: rUl(k).Mitarbeiter = left$(rUl(k).Mitarbeiter, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in ulcusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' ulcusSpeichern

Public Function roVkZuw(i&, j&)
 roVk(i).FID = rVk(j).FID
 roVk(i).Pat_ID = rVk(j).Pat_ID
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
 roVk(i).aktZeit = rVk(j).aktZeit
 roVk(i).StByte = rVk(j).StByte
End Function ' roVkZuw

Public Function VkZUnt%(i&, j&)
 If roVk(i).FID <> rVk(j).FID Then GoSub unter
 If roVk(i).Pat_ID <> rVk(j).Pat_ID Then GoSub unter
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
 If roVk(i).aktZeit <> rVk(j).aktZeit Then GoSub unter
 If roVk(i).StByte <> rVk(j).StByte Then GoSub unter
 Exit Function
unter:
 VkZUnt = VkZUnt + 1
 Return
End Function ' VkZUnt

Public Function vkgdLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
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
   roVk(akt).Pat_ID = rs!Pat_ID
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
   roVk(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rVk) > 0 Then
  For ri = 1 To UBound(rVk)
   If rVk(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roVk)
    If roVk(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roVk(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roVk(roendpe + UBound(rVk) - rbeg)
   For ri = rbeg To UBound(rVk)
    Call roVkZuw(roendpe + ri - rbeg, ri)
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

Public Function rVkDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rVkdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rVk)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rVk(" & i & ").FID:" & String$(33, "."), 33) & rVk(i).FID
  Print #200, left$("rVk(" & i & ").Pat_ID:" & String$(33, "."), 33) & rVk(i).Pat_ID
  Print #200, left$("rVk(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rVk(i).Zeitpunkt
  Print #200, left$("rVk(" & i & ").Wohlempfinden:" & String$(33, "."), 33) & "'" & rVk(i).Wohlempfinden & "'"
  Print #200, left$("rVk(" & i & ").Saettigung:" & String$(33, "."), 33) & "'" & rVk(i).Saettigung & "'"
  Print #200, left$("rVk(" & i & ").Zielwerterreichung:" & String$(33, "."), 33) & "'" & rVk(i).Zielwerterreichung & "'"
  Print #200, left$("rVk(" & i & ").Ketonkörper:" & String$(33, "."), 33) & "'" & rVk(i).Ketonkörper & "'"
  Print #200, left$("rVk(" & i & ").Gynaekologenbefund:" & String$(33, "."), 33) & "'" & rVk(i).Gynaekologenbefund & "'"
  Print #200, left$("rVk(" & i & ").Gewichtsentwicklung:" & String$(33, "."), 33) & "'" & rVk(i).Gewichtsentwicklung & "'"
  Print #200, left$("rVk(" & i & ").HbA1c:" & String$(33, "."), 33) & "'" & rVk(i).HbA1c & "'"
  Print #200, left$("rVk(" & i & ").Bewegung:" & String$(33, "."), 33) & "'" & rVk(i).Bewegung & "'"
  Print #200, left$("rVk(" & i & ").Minuten:" & String$(33, "."), 33) & "'" & rVk(i).Minuten & "'"
  Print #200, left$("rVk(" & i & ").Blutdruck:" & String$(33, "."), 33) & "'" & rVk(i).Blutdruck & "'"
  Print #200, left$("rVk(" & i & ").Puls:" & String$(33, "."), 33) & "'" & rVk(i).Puls & "'"
  Print #200, left$("rVk(" & i & ").Mitarbeiter:" & String$(33, "."), 33) & "'" & rVk(i).Mitarbeiter & "'"
  Print #200, left$("rVk(" & i & ").absPos:" & String$(33, "."), 33) & rVk(i).absPos
  Print #200, left$("rVk(" & i & ").AktZeit:" & String$(33, "."), 33) & rVk(i).aktZeit
  Print #200, left$("rVk(" & i & ").StByte:" & String$(33, "."), 33) & rVk(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' vkgdDump

Public Function vkgdSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "vkgd")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rVk) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "Wohlempfinden,Saettigung,Zielwerterreichung,Ketonkörper,Gynaekologenbefund,Gewichtsentwicklung,HbA1c,Bewegung,Minuten,Blutdruck," & _
     "Puls,Mitarbeiter,absPos,AktZeit,StByte)            VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rVk)
  rVk(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rVk(i).FID, ",", rVk(i).Pat_ID, ",", DatFor_k(rVk(i).Zeitpunkt), ",'", rVk(i).Wohlempfinden, "','", rVk(i).Saettigung, "','", rVk(i).Zielwerterreichung, "','", rVk(i).Ketonkörper, "','", _
   rVk(i).Gynaekologenbefund, "','", rVk(i).Gewichtsentwicklung, "','", rVk(i).HbA1c, "','", rVk(i).Bewegung, "','", rVk(i).Minuten, "','", rVk(i).Blutdruck, "','", rVk(i).Puls, "','", _
   rVk(i).Mitarbeiter, "',", rVk(i).absPos, ",", DatFor_k(rVk(i).aktZeit), ",", rVk(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rVk) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rVk) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rVk)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rVk(" & i & "/" & UBound(rVk) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""vkgdSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rVk), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rVk.Wohlempfinden: '" & rVk(k).Wohlempfinden & "' -> '" & left$(rVk(k).Wohlempfinden, maxL) & "'", True: rVk(k).Wohlempfinden = left$(rVk(k).Wohlempfinden, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rVk.Saettigung: '" & rVk(k).Saettigung & "' -> '" & left$(rVk(k).Saettigung, maxL) & "'", True: rVk(k).Saettigung = left$(rVk(k).Saettigung, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rVk.Zielwerterreichung: '" & rVk(k).Zielwerterreichung & "' -> '" & left$(rVk(k).Zielwerterreichung, maxL) & "'", True: rVk(k).Zielwerterreichung = left$(rVk(k).Zielwerterreichung, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rVk.Ketonkörper: '" & rVk(k).Ketonkörper & "' -> '" & left$(rVk(k).Ketonkörper, maxL) & "'", True: rVk(k).Ketonkörper = left$(rVk(k).Ketonkörper, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gynaekologenbefund: '" & rVk(k).Gynaekologenbefund & "' -> '" & left$(rVk(k).Gynaekologenbefund, maxL) & "'", True: rVk(k).Gynaekologenbefund = left$(rVk(k).Gynaekologenbefund, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rVk.Gewichtsentwicklung: '" & rVk(k).Gewichtsentwicklung & "' -> '" & left$(rVk(k).Gewichtsentwicklung, maxL) & "'", True: rVk(k).Gewichtsentwicklung = left$(rVk(k).Gewichtsentwicklung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rVk.HbA1c: '" & rVk(k).HbA1c & "' -> '" & left$(rVk(k).HbA1c, maxL) & "'", True: rVk(k).HbA1c = left$(rVk(k).HbA1c, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rVk.Bewegung: '" & rVk(k).Bewegung & "' -> '" & left$(rVk(k).Bewegung, maxL) & "'", True: rVk(k).Bewegung = left$(rVk(k).Bewegung, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rVk.Minuten: '" & rVk(k).Minuten & "' -> '" & left$(rVk(k).Minuten, maxL) & "'", True: rVk(k).Minuten = left$(rVk(k).Minuten, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rVk.Blutdruck: '" & rVk(k).Blutdruck & "' -> '" & left$(rVk(k).Blutdruck, maxL) & "'", True: rVk(k).Blutdruck = left$(rVk(k).Blutdruck, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rVk.Puls: '" & rVk(k).Puls & "' -> '" & left$(rVk(k).Puls, maxL) & "'", True: rVk(k).Puls = left$(rVk(k).Puls, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rVk.Mitarbeiter: '" & rVk(k).Mitarbeiter & "' -> '" & left$(rVk(k).Mitarbeiter, maxL) & "'", True: rVk(k).Mitarbeiter = left$(rVk(k).Mitarbeiter, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in vkgdSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' vkgdSpeichern

Public Function roSwZuw(i&, j&)
 roSw(i).FID = rSw(j).FID
 roSw(i).Pat_ID = rSw(j).Pat_ID
 roSw(i).Zeitpunkt = rSw(j).Zeitpunkt
 roSw(i).FormTitel = rSw(j).FormTitel
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
 roSw(i).aktZeit = rSw(j).aktZeit
 roSw(i).StByte = rSw(j).StByte
End Function ' roSwZuw

Public Function SwZUnt%(i&, j&)
 If roSw(i).FID <> rSw(j).FID Then GoSub unter
 If roSw(i).Pat_ID <> rSw(j).Pat_ID Then GoSub unter
 If roSw(i).Zeitpunkt <> rSw(j).Zeitpunkt Then GoSub unter
 If roSw(i).FormTitel <> rSw(j).FormTitel Then GoSub unter
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
 If roSw(i).aktZeit <> rSw(j).aktZeit Then GoSub unter
 If roSw(i).StByte <> rSw(j).StByte Then GoSub unter
 Exit Function
unter:
 SwZUnt = SwZUnt + 1
 Return
End Function ' SwZUnt

Public Function swsLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(FormTitel,'') FormTitel" & _
",COALESCE(LR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) LR,COALESCE(vorET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) vorET,COALESCE(ET - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ET,COALESCE(efLR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) efLR" & _
",COALESCE(erLR - INTERVAL 0 DAY,CONVERT('18991230',DATE)) erLR,COALESCE(kGT - INTERVAL 0 DAY,CONVERT('18991230',DATE)) kGT,COALESCE(MB - INTERVAL 0 DAY,CONVERT('18991230',DATE)) MB,COALESCE(EndeArt,'') EndeArt" & _
",COALESCE(ED - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ED,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
" FROM `sws` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roSw(0)
 Else ' rs.EOF Then
  ReDim roSw(1)
  Do While Not rs.EOF
   akt = UBound(roSw)
   roSw(akt).FID = rs!FID
   roSw(akt).Pat_ID = rs!Pat_ID
   roSw(akt).Zeitpunkt = rs!Zeitpunkt
   roSw(akt).FormTitel = doUmwfSQL(rs!FormTitel, lies.obMySQL, False)
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
   roSw(akt).aktZeit = rs!aktZeit
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
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rSw) > 0 Then
  For ri = 1 To UBound(rSw)
   If rSw(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roSw)
    If roSw(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roSw(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roSw(roendpe + UBound(rSw) - rbeg)
   For ri = rbeg To UBound(rSw)
    Call roSwZuw(roendpe + ri - rbeg, ri)
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

Public Function rSwDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rSwdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rSw)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rSw(" & i & ").FID:" & String$(33, "."), 33) & rSw(i).FID
  Print #200, left$("rSw(" & i & ").Pat_ID:" & String$(33, "."), 33) & rSw(i).Pat_ID
  Print #200, left$("rSw(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rSw(i).Zeitpunkt
  Print #200, left$("rSw(" & i & ").FormTitel:" & String$(33, "."), 33) & "'" & rSw(i).FormTitel & "'"
  Print #200, left$("rSw(" & i & ").LR:" & String$(33, "."), 33) & rSw(i).lR
  Print #200, left$("rSw(" & i & ").vorET:" & String$(33, "."), 33) & rSw(i).vorET
  Print #200, left$("rSw(" & i & ").ET:" & String$(33, "."), 33) & rSw(i).ET
  Print #200, left$("rSw(" & i & ").efLR:" & String$(33, "."), 33) & rSw(i).efLR
  Print #200, left$("rSw(" & i & ").erLR:" & String$(33, "."), 33) & rSw(i).erLR
  Print #200, left$("rSw(" & i & ").kGT:" & String$(33, "."), 33) & rSw(i).kGT
  Print #200, left$("rSw(" & i & ").MB:" & String$(33, "."), 33) & rSw(i).MB
  Print #200, left$("rSw(" & i & ").EndeArt:" & String$(33, "."), 33) & "'" & rSw(i).EndeArt & "'"
  Print #200, left$("rSw(" & i & ").ED:" & String$(33, "."), 33) & rSw(i).ED
  Print #200, left$("rSw(" & i & ").absPos:" & String$(33, "."), 33) & rSw(i).absPos
  Print #200, left$("rSw(" & i & ").AktZeit:" & String$(33, "."), 33) & rSw(i).aktZeit
  Print #200, left$("rSw(" & i & ").StByte:" & String$(33, "."), 33) & rSw(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' swsDump

Public Function swsSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "sws")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rSw) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "FormTitel,LR,vorET,ET,efLR,erLR,kGT,MB,EndeArt,ED," & _
     "absPos,AktZeit,StByte)               VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rSw)
  rSw(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rSw(i).FID, ",", rSw(i).Pat_ID, ",", DatFor_k(rSw(i).Zeitpunkt), ",'", rSw(i).FormTitel, "',", DatFor_k(rSw(i).lR), ",", DatFor_k(rSw(i).vorET), ",", DatFor_k(rSw(i).ET), ",", DatFor_k( _
   rSw(i).efLR), ",", DatFor_k(rSw(i).erLR), ",", DatFor_k(rSw(i).kGT), ",", DatFor_k(rSw(i).MB), ",'", rSw(i).EndeArt, "',", DatFor_k(rSw(i).ED), ",", rSw(i).absPos, ",", DatFor_k(rSw(i).aktZeit), ",", _
   rSw(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rSw) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rSw) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rSw(i).Pat_ID = myEFrag("SELECT MAX(Pat_ID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rSw)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rSw(" & i & "/" & UBound(rSw) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""swsSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rSw), i)
  If Len(rSw(k).FormTitel) > maxi(0) Then maxi(0) = Len(rSw(k).FormTitel)
  If Len(rSw(k).EndeArt) > maxi(1) Then maxi(1) = Len(rSw(k).EndeArt)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rSw), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rSw.FormTitel: '" & rSw(k).FormTitel & "' -> '" & left$(rSw(k).FormTitel, maxL) & "'", True: rSw(k).FormTitel = left$(rSw(k).FormTitel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rSw.EndeArt: '" & rSw(k).EndeArt & "' -> '" & left$(rSw(k).EndeArt, maxL) & "'", True: rSw(k).EndeArt = left$(rSw(k).EndeArt, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in swsSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' swsSpeichern

Public Function roVoZuw(i&, j&)
 roVo(i).FID = rVo(j).FID
 roVo(i).Pat_ID = rVo(j).Pat_ID
 roVo(i).Zeitpunkt = rVo(j).Zeitpunkt
 roVo(i).FormTitel = rVo(j).FormTitel
 roVo(i).Inhalt = rVo(j).Inhalt
 roVo(i).absPos = rVo(j).absPos
 roVo(i).aktZeit = rVo(j).aktZeit
 roVo(i).StByte = rVo(j).StByte
End Function ' roVoZuw

Public Function VoZUnt%(i&, j&)
 If roVo(i).FID <> rVo(j).FID Then GoSub unter
 If roVo(i).Pat_ID <> rVo(j).Pat_ID Then GoSub unter
 If roVo(i).Zeitpunkt <> rVo(j).Zeitpunkt Then GoSub unter
 If roVo(i).FormTitel <> rVo(j).FormTitel Then GoSub unter
 If roVo(i).Inhalt <> rVo(j).Inhalt Then GoSub unter
 If roVo(i).absPos <> rVo(j).absPos Then GoSub unter
 If roVo(i).aktZeit <> rVo(j).aktZeit Then GoSub unter
 If roVo(i).StByte <> rVo(j).StByte Then GoSub unter
 Exit Function
unter:
 VoZUnt = VoZUnt + 1
 Return
End Function ' VoZUnt

Public Function voplLaden()
 Dim pid$, rs As New Recordset, akt&
 On Error GoTo fehler
 pid = rNa(0).Pat_ID
 sql = "SELECT COALESCE(FID,0) FID,COALESCE(Pat_ID,0) Pat_ID,COALESCE(ZeitPunkt - INTERVAL 0 DAY,CONVERT('18991230',DATE)) ZeitPunkt,COALESCE(FormTitel,'') FormTitel" & _
",COALESCE(Inhalt,'') Inhalt,COALESCE(absPos,0) absPos,COALESCE(AktZeit - INTERVAL 0 DAY,CONVERT('18991230',DATE)) AktZeit,COALESCE(StByte,0) StByte" & _
" FROM `vopl` WHERE Pat_ID=" & pid & " ORDER BY `ZeitPunkt`"
 myFrag rs, sql
 If rs.EOF Then
  ReDim roVo(0)
 Else ' rs.EOF Then
  ReDim roVo(1)
  Do While Not rs.EOF
   akt = UBound(roVo)
   roVo(akt).FID = rs!FID
   roVo(akt).Pat_ID = rs!Pat_ID
   roVo(akt).Zeitpunkt = rs!Zeitpunkt
   roVo(akt).FormTitel = doUmwfSQL(rs!FormTitel, lies.obMySQL, False)
   roVo(akt).Inhalt = doUmwfSQL(rs!Inhalt, lies.obMySQL, False)
   roVo(akt).absPos = rs!absPos
   roVo(akt).aktZeit = rs!aktZeit
   roVo(akt).StByte = rs!StByte
   rs.MoveNext
   If Not rs.EOF Then ReDim Preserve roVo(UBound(roVo) + 1)
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
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in voplLaden/" + AnwPfad)
  Case vbAbort: Call MsgBox(" 2Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' voplLaden

Function voplEinf()
 Dim rbeg&, roendpe&, ri&, roi&
 On Error GoTo fehler
 If UBound(rVo) > 0 Then
  For ri = 1 To UBound(rVo)
   If rVo(ri).Zeitpunkt >= qbeg Then ' aktqanf()
    rbeg = ri
    Exit For
   End If
  Next ri
  If rbeg <> 0 Then
   For roendpe = 0 To UBound(roVo)
    If roVo(roendpe).Zeitpunkt >= qbeg Then
     Exit For
    End If
   Next roendpe
   If roendpe <= UBound(roFa) Then
    ReDim Preserve roVo(roendpe - 1)
   End If ' IF roendpe <= UBound(roFa) THEN
   ReDim Preserve roVo(roendpe + UBound(rVo) - rbeg)
   For ri = rbeg To UBound(rVo)
    Call roVoZuw(roendpe + ri - rbeg, ri)
   Next ri
  End If ' IF rbeg <> 0 THEN
 End If ' IF UBound(rFa) > 0 THEN
 rVo = roVo
 Exit Function
fehler:
 Dim AnwPfad$
 #If VBA6 Then
 AnwPfad = CurrentDb.name
 #Else
 AnwPfad = App.path
 #End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in voplEinf/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox(" Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' voplEinf

Public Function rVoDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rVodump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rVo)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rVo(" & i & ").FID:" & String$(33, "."), 33) & rVo(i).FID
  Print #200, left$("rVo(" & i & ").Pat_ID:" & String$(33, "."), 33) & rVo(i).Pat_ID
  Print #200, left$("rVo(" & i & ").ZeitPunkt:" & String$(33, "."), 33) & rVo(i).Zeitpunkt
  Print #200, left$("rVo(" & i & ").FormTitel:" & String$(33, "."), 33) & "'" & rVo(i).FormTitel & "'"
  Print #200, left$("rVo(" & i & ").Inhalt:" & String$(33, "."), 33) & "'" & rVo(i).Inhalt & "'"
  Print #200, left$("rVo(" & i & ").absPos:" & String$(33, "."), 33) & rVo(i).absPos
  Print #200, left$("rVo(" & i & ").AktZeit:" & String$(33, "."), 33) & rVo(i).aktZeit
  Print #200, left$("rVo(" & i & ").StByte:" & String$(33, "."), 33) & rVo(i).StByte
 Next i
 Close #200
 zeigan ffadat
End Function ' voplDump

Public Function voplSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "vopl")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rVo) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (FID,Pat_ID,ZeitPunkt," & _
     "FormTitel,Inhalt,absPos,AktZeit,StByte)            VALUES"))
 If Not AllePat Then
   sql = "DELETE FROM `" & LCase$(Tbnm) & "` WHERE Pat_ID = " & CStr(rNa(0).Pat_ID)
   Call myEFrag(sql)
 End If ' not AllePat
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rVo)
  rVo(i).StByte = CStr(AktByte)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rVo(i).FID, ",", rVo(i).Pat_ID, ",", DatFor_k(rVo(i).Zeitpunkt), ",'", rVo(i).FormTitel, "','", rVo(i).Inhalt, "',", rVo(i).absPos, ",", DatFor_k(rVo(i).aktZeit), ",", rVo(i).StByte, ")")
  If SammelInsert <> 0 And i < UBound(rVo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rVo) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rVo)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rVo(" & i & "/" & UBound(rVo) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""voplSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(1)
 For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rVo), i)
  If Len(rVo(k).FormTitel) > maxi(0) Then maxi(0) = Len(rVo(k).FormTitel)
  If Len(rVo(k).Inhalt) > maxi(1) Then maxi(1) = Len(rVo(k).Inhalt)
 Next k
 If obTrans <> 0 Then If myEFrag("SELECT COUNT(1) FROM information_schema.innodb_trx WHERE trx_mysql_thread_id = CONNECTION_ID()", , DBCn).Fields(0) <> 0 Then ComTrans ' DBCn.CommitTrans: obtrans = 0
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rVo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rVo.FormTitel: '" & rVo(k).FormTitel & "' -> '" & left$(rVo(k).FormTitel, maxL) & "'", True: rVo(k).FormTitel = left$(rVo(k).FormTitel, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rVo.Inhalt: '" & rVo(k).Inhalt & "' -> '" & left$(rVo(k).Inhalt, maxL) & "'", True: rVo(k).Inhalt = left$(rVo(k).Inhalt, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in voplSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' voplSpeichern

Public Function rLsDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLsdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLs)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLs(" & i & ").SatzID:" & String$(33, "."), 33) & rLs(i).SatzID
  Print #200, left$("rLs(" & i & ").DatID:" & String$(33, "."), 33) & rLs(i).DatID
  Print #200, left$("rLs(" & i & ").Satzart:" & String$(33, "."), 33) & "'" & rLs(i).Satzart & "'"
  Print #200, left$("rLs(" & i & ").Satzlänge:" & String$(33, "."), 33) & "'" & rLs(i).Satzlänge & "'"
  Print #200, left$("rLs(" & i & ").SatzlängeSchluss:" & String$(33, "."), 33) & "'" & rLs(i).SatzlängeSchluss & "'"
  Print #200, left$("rLs(" & i & ").VersionSatzb:" & String$(33, "."), 33) & "'" & rLs(i).VersionSatzb & "'"
  Print #200, left$("rLs(" & i & ").Arztnr:" & String$(33, "."), 33) & "'" & rLs(i).Arztnr & "'"
  Print #200, left$("rLs(" & i & ").Arztname:" & String$(33, "."), 33) & "'" & rLs(i).Arztname & "'"
  Print #200, left$("rLs(" & i & ").StraßePraxis:" & String$(33, "."), 33) & "'" & rLs(i).StraßePraxis & "'"
  Print #200, left$("rLs(" & i & ").Arzt:" & String$(33, "."), 33) & "'" & rLs(i).Arzt & "'"
  Print #200, left$("rLs(" & i & ").LANR:" & String$(33, "."), 33) & "'" & rLs(i).Lanr & "'"
  Print #200, left$("rLs(" & i & ").PLZPraxis:" & String$(33, "."), 33) & "'" & rLs(i).PLZPraxis & "'"
  Print #200, left$("rLs(" & i & ").OrtPraxis:" & String$(33, "."), 33) & "'" & rLs(i).OrtPraxis & "'"
  Print #200, left$("rLs(" & i & ").Labor:" & String$(33, "."), 33) & "'" & rLs(i).Labor & "'"
  Print #200, left$("rLs(" & i & ").StraßeLabor:" & String$(33, "."), 33) & "'" & rLs(i).StraßeLabor & "'"
  Print #200, left$("rLs(" & i & ").PLZLabor:" & String$(33, "."), 33) & "'" & rLs(i).PLZLabor & "'"
  Print #200, left$("rLs(" & i & ").OrtLabor:" & String$(33, "."), 33) & "'" & rLs(i).OrtLabor & "'"
  Print #200, left$("rLs(" & i & ").KBVPrüfnr:" & String$(33, "."), 33) & "'" & rLs(i).KBVPrüfnr & "'"
  Print #200, left$("rLs(" & i & ").Zeichensatz:" & String$(33, "."), 33) & "'" & rLs(i).Zeichensatz & "'"
  Print #200, left$("rLs(" & i & ").Kundenarztnr:" & String$(33, "."), 33) & "'" & rLs(i).Kundenarztnr & "'"
  Print #200, left$("rLs(" & i & ").Erstellungsdatum:" & String$(33, "."), 33) & "'" & rLs(i).Erstellungsdatum & "'"
  Print #200, left$("rLs(" & i & ").Gesamtlänge:" & String$(33, "."), 33) & "'" & rLs(i).Gesamtlänge & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxsaetzeDump

Public Function laborxsaetzeSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxsaetze")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLs) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (DatID,Satzart,Satzlänge," & _
     "SatzlängeSchluss,VersionSatzb,Arztnr,Arztname,StraßePraxis,Arzt,LANR,PLZPraxis,OrtPraxis,Labor," & _
     "StraßeLabor,PLZLabor,OrtLabor,KBVPrüfnr,Zeichensatz,Kundenarztnr,Erstellungsdatum,Gesamtlänge)             VALUES"))
sql:
 csql.m_Len = 0
 For i = 0 To UBound(rLs)
setz:
  If SammelInsert = 0 Or i = 0 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 0 Then
  csql.AppVar Array("(", rLs(i).DatID, ",'", rLs(i).Satzart, "','", rLs(i).Satzlänge, "','", rLs(i).SatzlängeSchluss, "','", rLs(i).VersionSatzb, "','", rLs(i).Arztnr, "','", rLs(i).Arztname, "','", _
   rLs(i).StraßePraxis, "','", rLs(i).Arzt, "','", rLs(i).Lanr, "','", rLs(i).PLZPraxis, "','", rLs(i).OrtPraxis, "','", rLs(i).Labor, "','", rLs(i).StraßeLabor, "','", rLs(i).PLZLabor, "','", _
   rLs(i).OrtLabor, "','", rLs(i).KBVPrüfnr, "','", rLs(i).Zeichensatz, "','", rLs(i).Kundenarztnr, "','", rLs(i).Erstellungsdatum, "','", rLs(i).Gesamtlänge, "')")
  If SammelInsert <> 0 And i < UBound(rLs) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLs) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLs(i).SatzID = myEFrag("SELECT MAX(SatzID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLs)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLs(" & i & "/" & UBound(rLs) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxsaetzeSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 0, i) To IIf(SammelInsert <> 0, UBound(rLs), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzart: '" & rLs(k).Satzart & "' -> '" & left$(rLs(k).Satzart, maxL) & "'", True: rLs(k).Satzart = left$(rLs(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLs.Satzlänge: '" & rLs(k).Satzlänge & "' -> '" & left$(rLs(k).Satzlänge, maxL) & "'", True: rLs(k).Satzlänge = left$(rLs(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLs.SatzlängeSchluss: '" & rLs(k).SatzlängeSchluss & "' -> '" & left$(rLs(k).SatzlängeSchluss, maxL) & "'", True: rLs(k).SatzlängeSchluss = left$(rLs(k).SatzlängeSchluss, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLs.VersionSatzb: '" & rLs(k).VersionSatzb & "' -> '" & left$(rLs(k).VersionSatzb, maxL) & "'", True: rLs(k).VersionSatzb = left$(rLs(k).VersionSatzb, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztnr: '" & rLs(k).Arztnr & "' -> '" & left$(rLs(k).Arztnr, maxL) & "'", True: rLs(k).Arztnr = left$(rLs(k).Arztnr, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arztname: '" & rLs(k).Arztname & "' -> '" & left$(rLs(k).Arztname, maxL) & "'", True: rLs(k).Arztname = left$(rLs(k).Arztname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßePraxis: '" & rLs(k).StraßePraxis & "' -> '" & left$(rLs(k).StraßePraxis, maxL) & "'", True: rLs(k).StraßePraxis = left$(rLs(k).StraßePraxis, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLs.Arzt: '" & rLs(k).Arzt & "' -> '" & left$(rLs(k).Arzt, maxL) & "'", True: rLs(k).Arzt = left$(rLs(k).Arzt, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLs.LANR: '" & rLs(k).Lanr & "' -> '" & left$(rLs(k).Lanr, maxL) & "'", True: rLs(k).Lanr = left$(rLs(k).Lanr, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZPraxis: '" & rLs(k).PLZPraxis & "' -> '" & left$(rLs(k).PLZPraxis, maxL) & "'", True: rLs(k).PLZPraxis = left$(rLs(k).PLZPraxis, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtPraxis: '" & rLs(k).OrtPraxis & "' -> '" & left$(rLs(k).OrtPraxis, maxL) & "'", True: rLs(k).OrtPraxis = left$(rLs(k).OrtPraxis, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLs.Labor: '" & rLs(k).Labor & "' -> '" & left$(rLs(k).Labor, maxL) & "'", True: rLs(k).Labor = left$(rLs(k).Labor, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLs.StraßeLabor: '" & rLs(k).StraßeLabor & "' -> '" & left$(rLs(k).StraßeLabor, maxL) & "'", True: rLs(k).StraßeLabor = left$(rLs(k).StraßeLabor, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLs.PLZLabor: '" & rLs(k).PLZLabor & "' -> '" & left$(rLs(k).PLZLabor, maxL) & "'", True: rLs(k).PLZLabor = left$(rLs(k).PLZLabor, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLs.OrtLabor: '" & rLs(k).OrtLabor & "' -> '" & left$(rLs(k).OrtLabor, maxL) & "'", True: rLs(k).OrtLabor = left$(rLs(k).OrtLabor, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLs.KBVPrüfnr: '" & rLs(k).KBVPrüfnr & "' -> '" & left$(rLs(k).KBVPrüfnr, maxL) & "'", True: rLs(k).KBVPrüfnr = left$(rLs(k).KBVPrüfnr, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLs.Zeichensatz: '" & rLs(k).Zeichensatz & "' -> '" & left$(rLs(k).Zeichensatz, maxL) & "'", True: rLs(k).Zeichensatz = left$(rLs(k).Zeichensatz, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLs.Kundenarztnr: '" & rLs(k).Kundenarztnr & "' -> '" & left$(rLs(k).Kundenarztnr, maxL) & "'", True: rLs(k).Kundenarztnr = left$(rLs(k).Kundenarztnr, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLs.Erstellungsdatum: '" & rLs(k).Erstellungsdatum & "' -> '" & left$(rLs(k).Erstellungsdatum, maxL) & "'", True: rLs(k).Erstellungsdatum = left$(rLs(k).Erstellungsdatum, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLs.Gesamtlänge: '" & rLs(k).Gesamtlänge & "' -> '" & left$(rLs(k).Gesamtlänge, maxL) & "'", True: rLs(k).Gesamtlänge = left$(rLs(k).Gesamtlänge, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxsaetzeSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxsaetzeSpeichern

Public Function rLgDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLgdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLg)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLg(" & i & ").DatID:" & String$(33, "."), 33) & rLg(i).DatID
  Print #200, left$("rLg(" & i & ").Pfad:" & String$(33, "."), 33) & "'" & rLg(i).Pfad & "'"
  Print #200, left$("rLg(" & i & ").Name:" & String$(33, "."), 33) & "'" & rLg(i).name & "'"
  Print #200, left$("rLg(" & i & ").Zp:" & String$(33, "."), 33) & rLg(i).Zp
  Print #200, left$("rLg(" & i & ").fertig:" & String$(33, "."), 33) & rLg(i).Fertig
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxeingelDump

Public Function laborxeingelSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxeingel")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLg) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (Pfad,Name,Zp," & _
     "fertig)  VALUES"))
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLg)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("('", rLg(i).Pfad, "','", rLg(i).name, "',", DatFor_k(rLg(i).Zp), ",", CStr(-(rLg(i).Fertig <> 0)), ")")
  If SammelInsert <> 0 And i < UBound(rLg) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLg) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLg(i).DatID = myEFrag("SELECT MAX(DatID)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLg)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLg(" & i & "/" & UBound(rLg) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxeingelSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLg), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLg.Pfad: '" & rLg(k).Pfad & "' -> '" & left$(rLg(k).Pfad, maxL) & "'", True: rLg(k).Pfad = left$(rLg(k).Pfad, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLg.Name: '" & rLg(k).name & "' -> '" & left$(rLg(k).name, maxL) & "'", True: rLg(k).name = left$(rLg(k).name, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxeingelSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxeingelSpeichern

Public Function rLuDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLudump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLu)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLu(" & i & ").RefNr:" & String$(33, "."), 33) & rLu(i).RefNr
  Print #200, left$("rLu(" & i & ").DatID:" & String$(33, "."), 33) & rLu(i).DatID
  Print #200, left$("rLu(" & i & ").SatzID:" & String$(33, "."), 33) & rLu(i).SatzID
  Print #200, left$("rLu(" & i & ").Satzart:" & String$(33, "."), 33) & "'" & rLu(i).Satzart & "'"
  Print #200, left$("rLu(" & i & ").Satzlänge:" & String$(33, "."), 33) & "'" & rLu(i).Satzlänge & "'"
  Print #200, left$("rLu(" & i & ").Auftragsnummer:" & String$(33, "."), 33) & "'" & rLu(i).Auftragsnummer & "'"
  Print #200, left$("rLu(" & i & ").Auftragsschlüssel:" & String$(33, "."), 33) & "'" & rLu(i).Auftragsschlüssel & "'"
  Print #200, left$("rLu(" & i & ").Eingang:" & String$(33, "."), 33) & rLu(i).Eingang
  Print #200, left$("rLu(" & i & ").Berichtsdatum:" & String$(33, "."), 33) & "'" & rLu(i).Berichtsdatum & "'"
  Print #200, left$("rLu(" & i & ").Pat_ID:" & String$(33, "."), 33) & rLu(i).Pat_ID
  Print #200, left$("rLu(" & i & ").TM_Pat_id:" & String$(33, "."), 33) & rLu(i).TM_Pat_ID
  Print #200, left$("rLu(" & i & ").Nachname:" & String$(33, "."), 33) & "'" & rLu(i).Nachname & "'"
  Print #200, left$("rLu(" & i & ").Vorname:" & String$(33, "."), 33) & "'" & rLu(i).Vorname & "'"
  Print #200, left$("rLu(" & i & ").GebDat:" & String$(33, "."), 33) & "'" & rLu(i).GebDat & "'"
  Print #200, left$("rLu(" & i & ").Titel:" & String$(33, "."), 33) & "'" & rLu(i).Titel & "'"
  Print #200, left$("rLu(" & i & ").NVorsatz:" & String$(33, "."), 33) & "'" & rLu(i).NVorsatz & "'"
  Print #200, left$("rLu(" & i & ").BefArt:" & String$(33, "."), 33) & "'" & rLu(i).BefArt & "'"
  Print #200, left$("rLu(" & i & ").Abrechnungstyp:" & String$(33, "."), 33) & "'" & rLu(i).Abrechnungstyp & "'"
  Print #200, left$("rLu(" & i & ").GebüOrd:" & String$(33, "."), 33) & "'" & rLu(i).GebüOrd & "'"
  Print #200, left$("rLu(" & i & ").Auftraggeber:" & String$(33, "."), 33) & "'" & rLu(i).Auftraggeber & "'"
  Print #200, left$("rLu(" & i & ").Patienteninformation:" & String$(33, "."), 33) & "'" & rLu(i).Patienteninformation & "'"
  Print #200, left$("rLu(" & i & ").Geschlecht:" & String$(33, "."), 33) & "'" & rLu(i).geschlecht & "'"
  Print #200, left$("rLu(" & i & ").AuftrHinw:" & String$(33, "."), 33) & "'" & rLu(i).AuftrHinw & "'"
  Print #200, left$("rLu(" & i & ").Pat_idUrsp:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idUrsp & "'"
  Print #200, left$("rLu(" & i & ").Pat_idErwVNG:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idErwVNG & "'"
  Print #200, left$("rLu(" & i & ").Pat_idErwVN:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idErwVN & "'"
  Print #200, left$("rLu(" & i & ").Pat_idErwG:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idErwG & "'"
  Print #200, left$("rLu(" & i & ").Pat_idErwGB:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idErwGB & "'"
  Print #200, left$("rLu(" & i & ").Pat_idErwGL:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idErwGL & "'"
  Print #200, left$("rLu(" & i & ").Pat_idLaborNeu:" & String$(33, "."), 33) & "'" & rLu(i).Pat_idLaborNeu & "'"
  Print #200, left$("rLu(" & i & ").ZeitpunktLaborneu:" & String$(33, "."), 33) & rLu(i).ZeitpunktLaborneu
  Print #200, left$("rLu(" & i & ").ZdüP:" & String$(33, "."), 33) & rLu(i).ZdüP
  Print #200, left$("rLu(" & i & ").ZdiP:" & String$(33, "."), 33) & rLu(i).ZdiP
  Print #200, left$("rLu(" & i & ").LWerte:" & String$(33, "."), 33) & "'" & rLu(i).LWerte & "'"
  Print #200, left$("rLu(" & i & ").verglichen:" & String$(33, "."), 33) & rLu(i).verglichen
  Print #200, left$("rLu(" & i & ").AfN:" & String$(33, "."), 33) & rLu(i).AfN
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxusDump

Public Function laborxusSpeichern(SammelInsert%, BezfSp%, j&, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxus")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLu) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (DatID,SatzID,Satzart," & _
     "Satzlänge,Auftragsnummer,Auftragsschlüssel,Eingang,Berichtsdatum,Pat_ID,TM_Pat_id,Nachname,Vorname,GebDat," & _
     "Titel,NVorsatz,BefArt,Abrechnungstyp,GebüOrd,Auftraggeber,Patienteninformation,Geschlecht,AuftrHinw,Pat_idUrsp," & _
     "Pat_idErwVNG,Pat_idErwVN,Pat_idErwG,Pat_idErwGB,Pat_idErwGL,Pat_idLaborNeu,ZeitpunktLaborneu,ZdüP,ZdiP,LWerte," & _
     "verglichen,AfN)        VALUES"))
sql:
 csql.m_Len = 0
 For i = j To j
setz:
  If SammelInsert = 0 Or i = j Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = j Then
  csql.AppVar Array("(", rLu(i).DatID, ",", rLu(i).SatzID, ",'", rLu(i).Satzart, "','", rLu(i).Satzlänge, "','", rLu(i).Auftragsnummer, "','", rLu(i).Auftragsschlüssel, "',", DatFor_k(rLu(i).Eingang), ",'", _
   rLu(i).Berichtsdatum, "',", rLu(i).Pat_ID, ",", rLu(i).TM_Pat_ID, ",'", rLu(i).Nachname, "','", rLu(i).Vorname, "','", rLu(i).GebDat, "','", rLu(i).Titel, "','", rLu(i).NVorsatz, "','", _
   rLu(i).BefArt, "','", rLu(i).Abrechnungstyp, "','", rLu(i).GebüOrd, "','", rLu(i).Auftraggeber, "','", rLu(i).Patienteninformation, "','", rLu(i).geschlecht, "','", rLu(i).AuftrHinw, "','", _
   rLu(i).Pat_idUrsp, "','", rLu(i).Pat_idErwVNG, "','", rLu(i).Pat_idErwVN, "','", rLu(i).Pat_idErwG, "','", rLu(i).Pat_idErwGB, "','", rLu(i).Pat_idErwGL, "','", rLu(i).Pat_idLaborNeu, "',", DatFor_k( _
   rLu(i).ZeitpunktLaborneu), ",", rLu(i).ZdüP, ",", rLu(i).ZdiP, ",'", rLu(i).LWerte, "',", DatFor_k(rLu(i).verglichen), ",", rLu(i).AfN, ")")
  If SammelInsert <> 0 And i < j Then csql.Append ","
  If SammelInsert = 0 Or i = j Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLu(i).RefNr = myEFrag("SELECT MAX(RefNr)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = j
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLu(" & i & "/" & UBound(rLu) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxusSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, j, i) To IIf(SammelInsert <> 0, j, i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzart: '" & rLu(k).Satzart & "' -> '" & left$(rLu(k).Satzart, maxL) & "'", True: rLu(k).Satzart = left$(rLu(k).Satzart, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLu.Satzlänge: '" & rLu(k).Satzlänge & "' -> '" & left$(rLu(k).Satzlänge, maxL) & "'", True: rLu(k).Satzlänge = left$(rLu(k).Satzlänge, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsnummer: '" & rLu(k).Auftragsnummer & "' -> '" & left$(rLu(k).Auftragsnummer, maxL) & "'", True: rLu(k).Auftragsnummer = left$(rLu(k).Auftragsnummer, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftragsschlüssel: '" & rLu(k).Auftragsschlüssel & "' -> '" & left$(rLu(k).Auftragsschlüssel, maxL) & "'", True: rLu(k).Auftragsschlüssel = left$(rLu(k).Auftragsschlüssel, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLu.Berichtsdatum: '" & rLu(k).Berichtsdatum & "' -> '" & left$(rLu(k).Berichtsdatum, maxL) & "'", True: rLu(k).Berichtsdatum = left$(rLu(k).Berichtsdatum, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLu.Nachname: '" & rLu(k).Nachname & "' -> '" & left$(rLu(k).Nachname, maxL) & "'", True: rLu(k).Nachname = left$(rLu(k).Nachname, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLu.Vorname: '" & rLu(k).Vorname & "' -> '" & left$(rLu(k).Vorname, maxL) & "'", True: rLu(k).Vorname = left$(rLu(k).Vorname, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebDat: '" & rLu(k).GebDat & "' -> '" & left$(rLu(k).GebDat, maxL) & "'", True: rLu(k).GebDat = left$(rLu(k).GebDat, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLu.Titel: '" & rLu(k).Titel & "' -> '" & left$(rLu(k).Titel, maxL) & "'", True: rLu(k).Titel = left$(rLu(k).Titel, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLu.NVorsatz: '" & rLu(k).NVorsatz & "' -> '" & left$(rLu(k).NVorsatz, maxL) & "'", True: rLu(k).NVorsatz = left$(rLu(k).NVorsatz, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLu.BefArt: '" & rLu(k).BefArt & "' -> '" & left$(rLu(k).BefArt, maxL) & "'", True: rLu(k).BefArt = left$(rLu(k).BefArt, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLu.Abrechnungstyp: '" & rLu(k).Abrechnungstyp & "' -> '" & left$(rLu(k).Abrechnungstyp, maxL) & "'", True: rLu(k).Abrechnungstyp = left$(rLu(k).Abrechnungstyp, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLu.GebüOrd: '" & rLu(k).GebüOrd & "' -> '" & left$(rLu(k).GebüOrd, maxL) & "'", True: rLu(k).GebüOrd = left$(rLu(k).GebüOrd, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLu.Auftraggeber: '" & rLu(k).Auftraggeber & "' -> '" & left$(rLu(k).Auftraggeber, maxL) & "'", True: rLu(k).Auftraggeber = left$(rLu(k).Auftraggeber, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLu.Patienteninformation: '" & rLu(k).Patienteninformation & "' -> '" & left$(rLu(k).Patienteninformation, maxL) & "'", True: rLu(k).Patienteninformation = left$(rLu(k).Patienteninformation, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLu.Geschlecht: '" & rLu(k).geschlecht & "' -> '" & left$(rLu(k).geschlecht, maxL) & "'", True: rLu(k).geschlecht = left$(rLu(k).geschlecht, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLu.AuftrHinw: '" & rLu(k).AuftrHinw & "' -> '" & left$(rLu(k).AuftrHinw, maxL) & "'", True: rLu(k).AuftrHinw = left$(rLu(k).AuftrHinw, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idUrsp: '" & rLu(k).Pat_idUrsp & "' -> '" & left$(rLu(k).Pat_idUrsp, maxL) & "'", True: rLu(k).Pat_idUrsp = left$(rLu(k).Pat_idUrsp, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVNG: '" & rLu(k).Pat_idErwVNG & "' -> '" & left$(rLu(k).Pat_idErwVNG, maxL) & "'", True: rLu(k).Pat_idErwVNG = left$(rLu(k).Pat_idErwVNG, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwVN: '" & rLu(k).Pat_idErwVN & "' -> '" & left$(rLu(k).Pat_idErwVN, maxL) & "'", True: rLu(k).Pat_idErwVN = left$(rLu(k).Pat_idErwVN, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwG: '" & rLu(k).Pat_idErwG & "' -> '" & left$(rLu(k).Pat_idErwG, maxL) & "'", True: rLu(k).Pat_idErwG = left$(rLu(k).Pat_idErwG, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGB: '" & rLu(k).Pat_idErwGB & "' -> '" & left$(rLu(k).Pat_idErwGB, maxL) & "'", True: rLu(k).Pat_idErwGB = left$(rLu(k).Pat_idErwGB, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idErwGL: '" & rLu(k).Pat_idErwGL & "' -> '" & left$(rLu(k).Pat_idErwGL, maxL) & "'", True: rLu(k).Pat_idErwGL = left$(rLu(k).Pat_idErwGL, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLu.Pat_idLaborNeu: '" & rLu(k).Pat_idLaborNeu & "' -> '" & left$(rLu(k).Pat_idLaborNeu, maxL) & "'", True: rLu(k).Pat_idLaborNeu = left$(rLu(k).Pat_idLaborNeu, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLu.LWerte: '" & rLu(k).LWerte & "' -> '" & left$(rLu(k).LWerte, maxL) & "'", True: rLu(k).LWerte = left$(rLu(k).LWerte, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxusSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxusSpeichern

Public Function rLoDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLodump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLo)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLo(" & i & ").RefNr:" & String$(33, "."), 33) & rLo(i).RefNr
  Print #200, left$("rLo(" & i & ").Verf:" & String$(33, "."), 33) & "'" & rLo(i).Verf & "'"
  Print #200, left$("rLo(" & i & ").KuQu:" & String$(33, "."), 33) & "'" & rLo(i).KuQu & "'"
  Print #200, left$("rLo(" & i & ").Quelle:" & String$(33, "."), 33) & "'" & rLo(i).Quelle & "'"
  Print #200, left$("rLo(" & i & ").QSpez:" & String$(33, "."), 33) & "'" & rLo(i).QSpez & "'"
  Print #200, left$("rLo(" & i & ").AbnDat:" & String$(33, "."), 33) & rLo(i).AbnDat
  Print #200, left$("rLo(" & i & ").Kommentar:" & String$(33, "."), 33) & "'" & rLo(i).Kommentar & "'"
  Print #200, left$("rLo(" & i & ").Erklärung:" & String$(33, "."), 33) & "'" & rLo(i).Erklärung & "'"
  Print #200, left$("rLo(" & i & ").Keimzahl:" & String$(33, "."), 33) & "'" & rLo(i).Keimzahl & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxbaktDump

Public Function laborxbaktSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxbakt")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLo) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (RefNr,Verf,KuQu," & _
     "Quelle,QSpez,AbnDat,Kommentar,Erklärung,Keimzahl)  VALUES"))
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLo)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLo(i).RefNr, ",'", rLo(i).Verf, "','", rLo(i).KuQu, "','", rLo(i).Quelle, "','", rLo(i).QSpez, "',", DatFor_k(rLo(i).AbnDat), ",'", rLo(i).Kommentar, "','", rLo(i).Erklärung, "','", _
   rLo(i).Keimzahl, "')")
  If SammelInsert <> 0 And i < UBound(rLo) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLo) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLo)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLo(" & i & "/" & UBound(rLo) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxbaktSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLo), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLo.Verf: '" & rLo(k).Verf & "' -> '" & left$(rLo(k).Verf, maxL) & "'", True: rLo(k).Verf = left$(rLo(k).Verf, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLo.KuQu: '" & rLo(k).KuQu & "' -> '" & left$(rLo(k).KuQu, maxL) & "'", True: rLo(k).KuQu = left$(rLo(k).KuQu, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLo.Quelle: '" & rLo(k).Quelle & "' -> '" & left$(rLo(k).Quelle, maxL) & "'", True: rLo(k).Quelle = left$(rLo(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLo.QSpez: '" & rLo(k).QSpez & "' -> '" & left$(rLo(k).QSpez, maxL) & "'", True: rLo(k).QSpez = left$(rLo(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLo.Kommentar: '" & rLo(k).Kommentar & "' -> '" & left$(rLo(k).Kommentar, maxL) & "'", True: rLo(k).Kommentar = left$(rLo(k).Kommentar, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLo.Erklärung: '" & rLo(k).Erklärung & "' -> '" & left$(rLo(k).Erklärung, maxL) & "'", True: rLo(k).Erklärung = left$(rLo(k).Erklärung, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLo.Keimzahl: '" & rLo(k).Keimzahl & "' -> '" & left$(rLo(k).Keimzahl, maxL) & "'", True: rLo(k).Keimzahl = left$(rLo(k).Keimzahl, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxbaktSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxbaktSpeichern

Public Function rLwDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLwdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLw)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLw(" & i & ").RefNr:" & String$(33, "."), 33) & rLw(i).RefNr
  Print #200, left$("rLw(" & i & ").Abkü:" & String$(33, "."), 33) & "'" & rLw(i).Abkü & "'"
  Print #200, left$("rLw(" & i & ").Langname:" & String$(33, "."), 33) & "'" & rLw(i).Langname & "'"
  Print #200, left$("rLw(" & i & ").Quelle:" & String$(33, "."), 33) & "'" & rLw(i).Quelle & "'"
  Print #200, left$("rLw(" & i & ").QSpez:" & String$(33, "."), 33) & "'" & rLw(i).QSpez & "'"
  Print #200, left$("rLw(" & i & ").AbnDat:" & String$(33, "."), 33) & rLw(i).AbnDat
  Print #200, left$("rLw(" & i & ").Wert:" & String$(33, "."), 33) & "'" & rLw(i).Wert & "'"
  Print #200, left$("rLw(" & i & ").Einheit:" & String$(33, "."), 33) & "'" & rLw(i).Einheit & "'"
  Print #200, left$("rLw(" & i & ").Grenzwerti:" & String$(33, "."), 33) & "'" & rLw(i).Grenzwerti & "'"
  Print #200, left$("rLw(" & i & ").Kommentar:" & String$(33, "."), 33) & "'" & rLw(i).Kommentar & "'"
  Print #200, left$("rLw(" & i & ").Teststatus:" & String$(33, "."), 33) & "'" & rLw(i).Teststatus & "'"
  Print #200, left$("rLw(" & i & ").Erklärung:" & String$(33, "."), 33) & "'" & rLw(i).Erklärung & "'"
  Print #200, left$("rLw(" & i & ").AuftrHinw:" & String$(33, "."), 33) & "'" & rLw(i).AuftrHinw & "'"
  Print #200, left$("rLw(" & i & ").nbid:" & String$(33, "."), 33) & rLw(i).nbid
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxwertDump

Public Function laborxwertSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxwert")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLw) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (RefNr,Abkü,Langname," & _
     "Quelle,QSpez,AbnDat,Wert,Einheit,Grenzwerti,Kommentar,Teststatus,Erklärung,AuftrHinw," & _
     "nbid)    VALUES"))
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLw)
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
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLw(i).RefNr, ",'", rLw(i).Abkü, "','", rLw(i).Langname, "','", rLw(i).Quelle, "','", rLw(i).QSpez, "',", DatFor_k(rLw(i).AbnDat), ",'", rLw(i).Wert, "','", rLw(i).Einheit, "','", _
   rLw(i).Grenzwerti, "','", rLw(i).Kommentar, "','", rLw(i).Teststatus, "','", rLw(i).Erklärung, "','", rLw(i).AuftrHinw, "',", rLw(i).nbid, ")")
  If SammelInsert <> 0 And i < UBound(rLw) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLw) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLw)
nexti:
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLw(" & i & "/" & UBound(rLw) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxwertSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLw), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLw.Abkü: '" & rLw(k).Abkü & "' -> '" & left$(rLw(k).Abkü, maxL) & "'", True: rLw(k).Abkü = left$(rLw(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLw.Langname: '" & rLw(k).Langname & "' -> '" & left$(rLw(k).Langname, maxL) & "'", True: rLw(k).Langname = left$(rLw(k).Langname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLw.Quelle: '" & rLw(k).Quelle & "' -> '" & left$(rLw(k).Quelle, maxL) & "'", True: rLw(k).Quelle = left$(rLw(k).Quelle, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLw.QSpez: '" & rLw(k).QSpez & "' -> '" & left$(rLw(k).QSpez, maxL) & "'", True: rLw(k).QSpez = left$(rLw(k).QSpez, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLw.Wert: '" & rLw(k).Wert & "' -> '" & left$(rLw(k).Wert, maxL) & "'", True: rLw(k).Wert = left$(rLw(k).Wert, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLw.Einheit: '" & rLw(k).Einheit & "' -> '" & left$(rLw(k).Einheit, maxL) & "'", True: rLw(k).Einheit = left$(rLw(k).Einheit, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLw.Grenzwerti: '" & rLw(k).Grenzwerti & "' -> '" & left$(rLw(k).Grenzwerti, maxL) & "'", True: rLw(k).Grenzwerti = left$(rLw(k).Grenzwerti, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLw.Kommentar: '" & rLw(k).Kommentar & "' -> '" & left$(rLw(k).Kommentar, maxL) & "'", True: rLw(k).Kommentar = left$(rLw(k).Kommentar, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLw.Teststatus: '" & rLw(k).Teststatus & "' -> '" & left$(rLw(k).Teststatus, maxL) & "'", True: rLw(k).Teststatus = left$(rLw(k).Teststatus, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLw.Erklärung: '" & rLw(k).Erklärung & "' -> '" & left$(rLw(k).Erklärung, maxL) & "'", True: rLw(k).Erklärung = left$(rLw(k).Erklärung, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLw.AuftrHinw: '" & rLw(k).AuftrHinw & "' -> '" & left$(rLw(k).AuftrHinw, maxL) & "'", True: rLw(k).AuftrHinw = left$(rLw(k).AuftrHinw, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxwertSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxwertSpeichern

Public Function rLLDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLLdump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLL)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLL(" & i & ").RefNr:" & String$(33, "."), 33) & rLL(i).RefNr
  Print #200, left$("rLL(" & i & ").Abkü:" & String$(33, "."), 33) & "'" & rLL(i).Abkü & "'"
  Print #200, left$("rLL(" & i & ").Verf:" & String$(33, "."), 33) & "'" & rLL(i).Verf & "'"
  Print #200, left$("rLL(" & i & ").EBM:" & String$(33, "."), 33) & "'" & rLL(i).EBM & "'"
  Print #200, left$("rLL(" & i & ").goä:" & String$(33, "."), 33) & "'" & rLL(i).goä & "'"
  Print #200, left$("rLL(" & i & ").Anzahl:" & String$(33, "."), 33) & "'" & rLL(i).Anzahl & "'"
  Print #200, left$("rLL(" & i & ").abrd:" & String$(33, "."), 33) & "'" & rLL(i).abrd & "'"
 Next i
 Close #200
 zeigan ffadat
End Function ' laborxleistDump

Public Function laborxleistSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "laborxleist")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLL) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (RefNr,Abkü,Verf," & _
     "EBM,goä,Anzahl,abrd)   VALUES"))
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLL)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("(", rLL(i).RefNr, ",'", rLL(i).Abkü, "','", rLL(i).Verf, "','", rLL(i).EBM, "','", rLL(i).goä, "','", rLL(i).Anzahl, "','", rLL(i).abrd, "')")
  If SammelInsert <> 0 And i < UBound(rLL) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLL) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLL)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLL(" & i & "/" & UBound(rLL) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""laborxleistSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
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
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLL), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLL.Abkü: '" & rLL(k).Abkü & "' -> '" & left$(rLL(k).Abkü, maxL) & "'", True: rLL(k).Abkü = left$(rLL(k).Abkü, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLL.Verf: '" & rLL(k).Verf & "' -> '" & left$(rLL(k).Verf, maxL) & "'", True: rLL(k).Verf = left$(rLL(k).Verf, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLL.EBM: '" & rLL(k).EBM & "' -> '" & left$(rLL(k).EBM, maxL) & "'", True: rLL(k).EBM = left$(rLL(k).EBM, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLL.goä: '" & rLL(k).goä & "' -> '" & left$(rLL(k).goä, maxL) & "'", True: rLL(k).goä = left$(rLL(k).goä, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLL.Anzahl: '" & rLL(k).Anzahl & "' -> '" & left$(rLL(k).Anzahl, maxL) & "'", True: rLL(k).Anzahl = left$(rLL(k).Anzahl, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLL.abrd: '" & rLL(k).abrd & "' -> '" & left$(rLL(k).abrd, maxL) & "'", True: rLL(k).abrd = left$(rLL(k).abrd, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in laborxleistSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
  Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
 End Select
End Function ' laborxleistSpeichern

Public Function rLiDump()
 Dim i&, ffa&
 Const ffadat$ = "\\linux1\daten\down\rLidump.txt"
 Open ffadat For Output As #200
 For i = 1 To UBound(rLi)
  Print #200, vbCrLf & "i: " & i
  Print #200, left$("rLi(" & i & ").name:" & String$(33, "."), 33) & "'" & rLi(i).name & "'"
  Print #200, left$("rLi(" & i & ").vorname:" & String$(33, "."), 33) & "'" & rLi(i).Vorname & "'"
  Print #200, left$("rLi(" & i & ").titelt:" & String$(33, "."), 33) & "'" & rLi(i).titelt & "'"
  Print #200, left$("rLi(" & i & ").fachgruppe:" & String$(33, "."), 33) & "'" & rLi(i).fachgruppe & "'"
  Print #200, left$("rLi(" & i & ").strasse:" & String$(33, "."), 33) & "'" & rLi(i).strasse & "'"
  Print #200, left$("rLi(" & i & ").plz:" & String$(33, "."), 33) & "'" & rLi(i).plz & "'"
  Print #200, left$("rLi(" & i & ").ort:" & String$(33, "."), 33) & "'" & rLi(i).ort & "'"
  Print #200, left$("rLi(" & i & ").telefon:" & String$(33, "."), 33) & "'" & rLi(i).telefon & "'"
  Print #200, left$("rLi(" & i & ").fax:" & String$(33, "."), 33) & "'" & rLi(i).fax & "'"
  Print #200, left$("rLi(" & i & ").kvnr:" & String$(33, "."), 33) & "'" & rLi(i).KVNr & "'"
  Print #200, left$("rLi(" & i & ").LANR:" & String$(33, "."), 33) & "'" & rLi(i).Lanr & "'"
  Print #200, left$("rLi(" & i & ").id:" & String$(33, "."), 33) & rLi(i).id
  Print #200, left$("rLi(" & i & ").überschrift:" & String$(33, "."), 33) & "'" & rLi(i).Überschrift & "'"
  Print #200, left$("rLi(" & i & ").dbnr:" & String$(33, "."), 33) & "'" & rLi(i).DBNr & "'"
  Print #200, left$("rLi(" & i & ").bstelle:" & String$(33, "."), 33) & "'" & rLi(i).BStelle & "'"
  Print #200, left$("rLi(" & i & ").anrede:" & String$(33, "."), 33) & "'" & rLi(i).anrede & "'"
  Print #200, left$("rLi(" & i & ").tel1:" & String$(33, "."), 33) & "'" & rLi(i).tel1 & "'"
  Print #200, left$("rLi(" & i & ").tel2:" & String$(33, "."), 33) & "'" & rLi(i).tel2 & "'"
  Print #200, left$("rLi(" & i & ").tel3:" & String$(33, "."), 33) & "'" & rLi(i).tel3 & "'"
  Print #200, left$("rLi(" & i & ").tel4:" & String$(33, "."), 33) & "'" & rLi(i).tel4 & "'"
  Print #200, left$("rLi(" & i & ").fax1:" & String$(33, "."), 33) & "'" & rLi(i).fax1 & "'"
  Print #200, left$("rLi(" & i & ").fax2:" & String$(33, "."), 33) & "'" & rLi(i).fax2 & "'"
  Print #200, left$("rLi(" & i & ").fax3:" & String$(33, "."), 33) & "'" & rLi(i).fax3 & "'"
  Print #200, left$("rLi(" & i & ").email:" & String$(33, "."), 33) & "'" & rLi(i).email & "'"
  Print #200, left$("rLi(" & i & ").zulg:" & String$(33, "."), 33) & "'" & rLi(i).zulg & "'"
  Print #200, left$("rLi(" & i & ").arzttyp:" & String$(33, "."), 33) & "'" & rLi(i).arzttyp & "'"
  Print #200, left$("rLi(" & i & ").gemmit:" & String$(33, "."), 33) & "'" & rLi(i).gemmit & "'"
  Print #200, left$("rLi(" & i & ").beme:" & String$(33, "."), 33) & "'" & rLi(i).beme & "'"
  Print #200, left$("rLi(" & i & ").dmpt2:" & String$(33, "."), 33) & rLi(i).dmpt2
  Print #200, left$("rLi(" & i & ").dmpt1:" & String$(33, "."), 33) & rLi(i).dmpt1
  Print #200, left$("rLi(" & i & ").geschlecht:" & String$(33, "."), 33) & "'" & rLi(i).geschlecht & "'"
  Print #200, left$("rLi(" & i & ").titel:" & String$(33, "."), 33) & "'" & rLi(i).Titel & "'"
  Print #200, left$("rLi(" & i & ").zusatz:" & String$(33, "."), 33) & "'" & rLi(i).zusatz & "'"
  Print #200, left$("rLi(" & i & ").ursp:" & String$(33, "."), 33) & "'" & rLi(i).ursp & "'"
  Print #200, left$("rLi(" & i & ").aktzeit:" & String$(33, "."), 33) & rLi(i).aktZeit
 Next i
 Close #200
 zeigan ffadat
End Function ' liuezDump

Public Function liuezSpeichern(SammelInsert%, BezfSp%, Optional rAf&, Optional sfkco%, Optional Tbnm = "liuez")
 Dim i&, rAFi&, pid$, m%, sfnr%, altmode$, ErrD$, ErrN& ',sql0$
 Dim csql0 As New CString, csql As New CString
 Dim rs As New ADODB.Recordset
 t1 = Timer
 On Error Resume Next
 pid = rNa(0).Pat_ID
 On Error GoTo fehler
 syscmd 4, pid & ": Speichere " & UBound(rLi) + 0 & " Sätze in `" & Tbnm & "`"
 Call csql0.AppVar(Array(" INSERT ", sqlIGNORE, "INTO `" & LCase$(Tbnm) & "` (name,vorname,titelt," & _
     "fachgruppe,strasse,plz,ort,telefon,fax,kvnr,LANR,überschrift,dbnr," & _
     "bstelle,anrede,tel1,tel2,tel3,tel4,fax1,fax2,fax3,email," & _
     "zulg,arzttyp,gemmit,beme,dmpt2,dmpt1,geschlecht,titel,zusatz,ursp," & _
     "aktzeit)               VALUES"))
sql:
 csql.m_Len = 0
 For i = 1 To UBound(rLi)
setz:
  If SammelInsert = 0 Or i = 1 Then
   csql.Append csql0
  End If ' SammelInsert = 0 Or i = 1 Then
  csql.AppVar Array("('", rLi(i).name, "','", rLi(i).Vorname, "','", rLi(i).titelt, "','", rLi(i).fachgruppe, "','", rLi(i).strasse, "','", rLi(i).plz, "','", rLi(i).ort, "','", rLi(i).telefon, "','", _
   rLi(i).fax, "','", rLi(i).KVNr, "','", rLi(i).Lanr, "','", rLi(i).Überschrift, "','", rLi(i).DBNr, "','", rLi(i).BStelle, "','", rLi(i).anrede, "','", rLi(i).tel1, "','", rLi(i).tel2, "','", _
   rLi(i).tel3, "','", rLi(i).tel4, "','", rLi(i).fax1, "','", rLi(i).fax2, "','", rLi(i).fax3, "','", rLi(i).email, "','", rLi(i).zulg, "','", rLi(i).arzttyp, "','", rLi(i).gemmit, "','", _
   rLi(i).beme, "',", rLi(i).dmpt2, ",", rLi(i).dmpt1, ",'", rLi(i).geschlecht, "','", rLi(i).Titel, "','", rLi(i).zusatz, "','", rLi(i).ursp, "',", DatFor_k(rLi(i).aktZeit), ")")
  If SammelInsert <> 0 And i < UBound(rLi) Then csql.Append ","
  If SammelInsert = 0 Or i = UBound(rLi) Then
    altmode = myEFrag("SELECT @@global.sql_mode", , DBCn).Fields(0)
    myEFrag "SET GLOBAL sql_mode='STRICT_TRANS_TABLES'", , DBCn ' NO_ENGINE_SUBSTITUTION
    InsKorr DBCn, csql.Value, rAf, ErrD, , ErrN, sfkco
    If ErrN Then
     If SammelInsert = 1 And ErrN = -2147217900 Then
      SammelInsert = 0
      GoTo sql
     ElseIf InStrB(ErrD, "Duplicate") <> 0 Then
      rLi(i).id = myEFrag("SELECT MAX(id)+1 FROM `" & LCase$(Tbnm) & "`", , DBCn).Fields(0)
      csql.m_Len = 0
      GoTo setz
     Else ' SammelInsert = 1 And ErrN = -2147217900 Then elseif
      Error ErrN
     End If ' SammelInsert = 1 And ErrN = -2147217900 Then else
    End If ' ErrN
   csql.Clear
   If obForK Then
    Call ForeignYes0
    Call ForeignYes1
   End If ' obforK THEN
  End If ' SammelInsert = 0 OR i = ubound(rLi)
  DoEvents
 Next i
 syscmd 5
 Exit Function
fehler:
If ErrN = 0 Then
 ErrDescr = Err.Description
 ErrNumber = Err.Number
Else ' ErrN = 0 Then
 ErrDescr = ErrD
 ErrNumber = ErrN
End If ' ErrN = 0 Then else
syscmd 4, "rLi(" & i & "/" & UBound(rLi) & "):   " & ErrDescr
sfnr = sfnr + 1
If sfnr > 10 Then
 Lese.Ausgeb sfnr & " Fehler in ""liuezSpeichern()"" bei Pat. " & rNa(0).Pat_ID & ", gebe auf (ErrDes: " & ErrDescr & ")", True
 sfnr = 0
 Resume Next
End If ' sfnr > 10
If ErrNumber = -2147217900 And (InStrB(ErrDescr, "Doppelter Eintrag") <> 0 Or InStrB(ErrDescr, "Duplicate") <> 0) Then
 Call Shell(App.path + "\..\nachricht\nachricht.exe " & App.EXEName & " Doppelter Eintrag bei: " & vbCrLf & csql.Value)
 Resume Next
ElseIf ErrNumber = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-5.1.32-log]Cannot add OR update a child row: a FOREIGN KEY constraint fails
 If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAFi
 Else
  Call doBezFeh(csql.Value, BezfSp, ErrDescr)
 End If
 Resume
ElseIf ErrNumber = -2147217833 Or InStrB(ErrDescr, "Daten zu lang") <> 0 Or InStrB(ErrDescr, "Data too long") <> 0 Then
 Dim rsc As ADODB.Recordset, maxi%(), k%
 ReDim maxi(30)
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
  If Len(rLi(k).Lanr) > maxi(10) Then maxi(10) = Len(rLi(k).Lanr)
  If Len(rLi(k).Überschrift) > maxi(11) Then maxi(11) = Len(rLi(k).Überschrift)
  If Len(rLi(k).DBNr) > maxi(12) Then maxi(12) = Len(rLi(k).DBNr)
  If Len(rLi(k).BStelle) > maxi(13) Then maxi(13) = Len(rLi(k).BStelle)
  If Len(rLi(k).anrede) > maxi(14) Then maxi(14) = Len(rLi(k).anrede)
  If Len(rLi(k).tel1) > maxi(15) Then maxi(15) = Len(rLi(k).tel1)
  If Len(rLi(k).tel2) > maxi(16) Then maxi(16) = Len(rLi(k).tel2)
  If Len(rLi(k).tel3) > maxi(17) Then maxi(17) = Len(rLi(k).tel3)
  If Len(rLi(k).tel4) > maxi(18) Then maxi(18) = Len(rLi(k).tel4)
  If Len(rLi(k).fax1) > maxi(19) Then maxi(19) = Len(rLi(k).fax1)
  If Len(rLi(k).fax2) > maxi(20) Then maxi(20) = Len(rLi(k).fax2)
  If Len(rLi(k).fax3) > maxi(21) Then maxi(21) = Len(rLi(k).fax3)
  If Len(rLi(k).email) > maxi(22) Then maxi(22) = Len(rLi(k).email)
  If Len(rLi(k).zulg) > maxi(23) Then maxi(23) = Len(rLi(k).zulg)
  If Len(rLi(k).arzttyp) > maxi(24) Then maxi(24) = Len(rLi(k).arzttyp)
  If Len(rLi(k).gemmit) > maxi(25) Then maxi(25) = Len(rLi(k).gemmit)
  If Len(rLi(k).beme) > maxi(26) Then maxi(26) = Len(rLi(k).beme)
  If Len(rLi(k).geschlecht) > maxi(27) Then maxi(27) = Len(rLi(k).geschlecht)
  If Len(rLi(k).Titel) > maxi(28) Then maxi(28) = Len(rLi(k).Titel)
  If Len(rLi(k).zusatz) > maxi(29) Then maxi(29) = Len(rLi(k).zusatz)
  If Len(rLi(k).ursp) > maxi(30) Then maxi(30) = Len(rLi(k).ursp)
 Next k
nochmal:
 Set rsc = New ADODB.Recordset
 Set rsc = DBCnOSchema(adSchemaColumns, Array(Empty, Empty, LCase$(Tbnm), Empty))
 m = 0
 Do While Not rsc.EOF
  Select Case rsc!data_type
   Case "varchar", "longtext", "mediumtext", "char", "text", "varbinary", 8, 129, 130, 200, 201, 202, 203, "set", "enum", "blob", "longblob", 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
    maxL = SpMod(maxi(m), LCase$(Tbnm), rsc)
    If maxL > 0 Then
     For k = IIf(SammelInsert <> 0, 1, i) To IIf(SammelInsert <> 0, UBound(rLi), i)
      Select Case m
       Case 0: Lese.Ausgeb "   Verkürze Inhalt von rLi.name: '" & rLi(k).name & "' -> '" & left$(rLi(k).name, maxL) & "'", True: rLi(k).name = left$(rLi(k).name, maxL)
       Case 1: Lese.Ausgeb "   Verkürze Inhalt von rLi.vorname: '" & rLi(k).Vorname & "' -> '" & left$(rLi(k).Vorname, maxL) & "'", True: rLi(k).Vorname = left$(rLi(k).Vorname, maxL)
       Case 2: Lese.Ausgeb "   Verkürze Inhalt von rLi.titelt: '" & rLi(k).titelt & "' -> '" & left$(rLi(k).titelt, maxL) & "'", True: rLi(k).titelt = left$(rLi(k).titelt, maxL)
       Case 3: Lese.Ausgeb "   Verkürze Inhalt von rLi.fachgruppe: '" & rLi(k).fachgruppe & "' -> '" & left$(rLi(k).fachgruppe, maxL) & "'", True: rLi(k).fachgruppe = left$(rLi(k).fachgruppe, maxL)
       Case 4: Lese.Ausgeb "   Verkürze Inhalt von rLi.strasse: '" & rLi(k).strasse & "' -> '" & left$(rLi(k).strasse, maxL) & "'", True: rLi(k).strasse = left$(rLi(k).strasse, maxL)
       Case 5: Lese.Ausgeb "   Verkürze Inhalt von rLi.plz: '" & rLi(k).plz & "' -> '" & left$(rLi(k).plz, maxL) & "'", True: rLi(k).plz = left$(rLi(k).plz, maxL)
       Case 6: Lese.Ausgeb "   Verkürze Inhalt von rLi.ort: '" & rLi(k).ort & "' -> '" & left$(rLi(k).ort, maxL) & "'", True: rLi(k).ort = left$(rLi(k).ort, maxL)
       Case 7: Lese.Ausgeb "   Verkürze Inhalt von rLi.telefon: '" & rLi(k).telefon & "' -> '" & left$(rLi(k).telefon, maxL) & "'", True: rLi(k).telefon = left$(rLi(k).telefon, maxL)
       Case 8: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax: '" & rLi(k).fax & "' -> '" & left$(rLi(k).fax, maxL) & "'", True: rLi(k).fax = left$(rLi(k).fax, maxL)
       Case 9: Lese.Ausgeb "   Verkürze Inhalt von rLi.kvnr: '" & rLi(k).KVNr & "' -> '" & left$(rLi(k).KVNr, maxL) & "'", True: rLi(k).KVNr = left$(rLi(k).KVNr, maxL)
       Case 10: Lese.Ausgeb "   Verkürze Inhalt von rLi.LANR: '" & rLi(k).Lanr & "' -> '" & left$(rLi(k).Lanr, maxL) & "'", True: rLi(k).Lanr = left$(rLi(k).Lanr, maxL)
       Case 11: Lese.Ausgeb "   Verkürze Inhalt von rLi.überschrift: '" & rLi(k).Überschrift & "' -> '" & left$(rLi(k).Überschrift, maxL) & "'", True: rLi(k).Überschrift = left$(rLi(k).Überschrift, maxL)
       Case 12: Lese.Ausgeb "   Verkürze Inhalt von rLi.dbnr: '" & rLi(k).DBNr & "' -> '" & left$(rLi(k).DBNr, maxL) & "'", True: rLi(k).DBNr = left$(rLi(k).DBNr, maxL)
       Case 13: Lese.Ausgeb "   Verkürze Inhalt von rLi.bstelle: '" & rLi(k).BStelle & "' -> '" & left$(rLi(k).BStelle, maxL) & "'", True: rLi(k).BStelle = left$(rLi(k).BStelle, maxL)
       Case 14: Lese.Ausgeb "   Verkürze Inhalt von rLi.anrede: '" & rLi(k).anrede & "' -> '" & left$(rLi(k).anrede, maxL) & "'", True: rLi(k).anrede = left$(rLi(k).anrede, maxL)
       Case 15: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel1: '" & rLi(k).tel1 & "' -> '" & left$(rLi(k).tel1, maxL) & "'", True: rLi(k).tel1 = left$(rLi(k).tel1, maxL)
       Case 16: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel2: '" & rLi(k).tel2 & "' -> '" & left$(rLi(k).tel2, maxL) & "'", True: rLi(k).tel2 = left$(rLi(k).tel2, maxL)
       Case 17: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel3: '" & rLi(k).tel3 & "' -> '" & left$(rLi(k).tel3, maxL) & "'", True: rLi(k).tel3 = left$(rLi(k).tel3, maxL)
       Case 18: Lese.Ausgeb "   Verkürze Inhalt von rLi.tel4: '" & rLi(k).tel4 & "' -> '" & left$(rLi(k).tel4, maxL) & "'", True: rLi(k).tel4 = left$(rLi(k).tel4, maxL)
       Case 19: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax1: '" & rLi(k).fax1 & "' -> '" & left$(rLi(k).fax1, maxL) & "'", True: rLi(k).fax1 = left$(rLi(k).fax1, maxL)
       Case 20: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax2: '" & rLi(k).fax2 & "' -> '" & left$(rLi(k).fax2, maxL) & "'", True: rLi(k).fax2 = left$(rLi(k).fax2, maxL)
       Case 21: Lese.Ausgeb "   Verkürze Inhalt von rLi.fax3: '" & rLi(k).fax3 & "' -> '" & left$(rLi(k).fax3, maxL) & "'", True: rLi(k).fax3 = left$(rLi(k).fax3, maxL)
       Case 22: Lese.Ausgeb "   Verkürze Inhalt von rLi.email: '" & rLi(k).email & "' -> '" & left$(rLi(k).email, maxL) & "'", True: rLi(k).email = left$(rLi(k).email, maxL)
       Case 23: Lese.Ausgeb "   Verkürze Inhalt von rLi.zulg: '" & rLi(k).zulg & "' -> '" & left$(rLi(k).zulg, maxL) & "'", True: rLi(k).zulg = left$(rLi(k).zulg, maxL)
       Case 24: Lese.Ausgeb "   Verkürze Inhalt von rLi.arzttyp: '" & rLi(k).arzttyp & "' -> '" & left$(rLi(k).arzttyp, maxL) & "'", True: rLi(k).arzttyp = left$(rLi(k).arzttyp, maxL)
       Case 25: Lese.Ausgeb "   Verkürze Inhalt von rLi.gemmit: '" & rLi(k).gemmit & "' -> '" & left$(rLi(k).gemmit, maxL) & "'", True: rLi(k).gemmit = left$(rLi(k).gemmit, maxL)
       Case 26: Lese.Ausgeb "   Verkürze Inhalt von rLi.beme: '" & rLi(k).beme & "' -> '" & left$(rLi(k).beme, maxL) & "'", True: rLi(k).beme = left$(rLi(k).beme, maxL)
       Case 27: Lese.Ausgeb "   Verkürze Inhalt von rLi.geschlecht: '" & rLi(k).geschlecht & "' -> '" & left$(rLi(k).geschlecht, maxL) & "'", True: rLi(k).geschlecht = left$(rLi(k).geschlecht, maxL)
       Case 28: Lese.Ausgeb "   Verkürze Inhalt von rLi.titel: '" & rLi(k).Titel & "' -> '" & left$(rLi(k).Titel, maxL) & "'", True: rLi(k).Titel = left$(rLi(k).Titel, maxL)
       Case 29: Lese.Ausgeb "   Verkürze Inhalt von rLi.zusatz: '" & rLi(k).zusatz & "' -> '" & left$(rLi(k).zusatz, maxL) & "'", True: rLi(k).zusatz = left$(rLi(k).zusatz, maxL)
       Case 30: Lese.Ausgeb "   Verkürze Inhalt von rLi.ursp: '" & rLi(k).ursp & "' -> '" & left$(rLi(k).ursp, maxL) & "'", True: rLi(k).ursp = left$(rLi(k).ursp, maxL)
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
 Resume sql
ElseIf InStrB(1, ErrDescr, "gone away", vbTextCompare) <> 0 Or InStrB(ErrDescr, "ost connection") <> 0 Then
 DBCnOpen
 Resume
End If ' ErrNumber =
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in liuezSpeichern/" + AnwPfad)
  Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume sql: If ErrNumber = 998 Then Resume sql Else Resume
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
 Call voplLaden
End Function ' tuLaden

' in alleSpeichern
Public Function tuSpeichern(frm As Lese, SI%, BfS%, Optional ohneAktDat%, Optional ohneLabor%) ' frm.dlg.SammelInsert, frm.dlg.BeziehungsfehlerSpeichern
 Dim rAf&, altsi$, altsam%
 altsi = sqlIGNORE
 sqlIGNORE = ""
 altsam = SI
 sqlIGNORE = ""
 On Error GoTo fehler
 Call namenSpeichern(SI, BfS)
 SI = 0
 ComTrans
 Call faelleSpeichern(SI, BfS)
 SI = altsam
 wechsTrans
 Call auSpeichern(SI, BfS)
 Call briefeSpeichern(SI, BfS)
 ComTrans
 Call diagnosenSpeichern(SI, BfS)
 Call dokumenteSpeichern(SI, BfS)
 Call eintraegeSpeichern(SI, BfS)
 sqlIGNORE = " IGNORE "
 Call ForeignNo0 ' 3.2.25
 Call ForeignNo1 ' 3.2.25
 Call formulareSpeichern(SI, BfS)
 Call ForeignYes0 ' 3.2.25
 Call ForeignYes1 ' 3.2.25
 sqlIGNORE = ""
 Call forminhkopfSpeichern(SI, BfS)
 Call forminhfeldSpeichern(SI, BfS)
 Call kheinweisSpeichern(SI, BfS)
 Call lbanforderungenSpeichern(SI, BfS)
 If Not ohneLabor Then
   tg2 = GetTickCount
   tg0 = tg0 + tg2 - tg1
   tg1 = tg2
   Call laborneuSpeichern(SI, BfS)
   tg2 = GetTickCount
 End If ' Not ohneLabor
 ComTrans
 Call leistungenSpeichern(SI, BfS)
 ComTrans
 Call medplanSpeichern(SI, BfS)
 Call rezepteintraegeSpeichern(SI, BfS)
 Call rrSpeichern(SI, BfS)
 Call kvnrueSpeichern(SI, BfS)
 Call unbek_kennSpeichern(SI, BfS)
 Call dmpreiheSpeichern(SI, BfS)
 Call desktopSpeichern(SI, BfS)
 Call usdmSpeichern(SI, BfS)
 Call fussSpeichern(SI, BfS)
 Call ulcusSpeichern(SI, BfS)
 Call vkgdSpeichern(SI, BfS)
 Call swsSpeichern(SI, BfS)
 Call voplSpeichern(SI, BfS)
 If Not ohneAktDat Then
  Call myEFrag("UPDATE `namen` SET aktZeit = " & DatFor_k(rNa(0).aktZeit) & " WHERE pat_id = " & rNa(0).Pat_ID, rAf)
  If rAf <> 1 Then
   frm.Ausgeb "Fehler bei der Setzung des Aktualisierungsdatum bei " & rNa(0).Pat_ID & " " & rNa(0).Nachname & " " & rNa(0).Vorname, True
  End If ' rAf <> 0
 End If ' not ohneAktDat
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

Public Function tuStutzen()
 Dim ErrNr&, ErrD$, rAf&
 Call Lese.ProgStart
 Call myEFrag("DELETE FROM `namen` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `faelle` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `au` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `briefe` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `diagnosen` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `dokumente` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `eintraege` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `forminhkopf` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `kheinweis` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `lbanforderungen` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `laborneu` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `leistungen` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `medplan` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `rezepteintraege` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `rr` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `kvnrue` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `unbek_kenn` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `dmpreihe` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `desktop` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `usdm` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `fuss` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `ulcus` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `vkgd` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `sws` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `vopl` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `laborxus` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
 Call myEFrag("DELETE FROM `anamnesebogen` WHERE pat_id>100000", rAf, DBCn, , ErrNr, ErrD)
End Function ' tuStutzen(frm AS Lese, SI%, BfS%)
